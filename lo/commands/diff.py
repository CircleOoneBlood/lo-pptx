"""
lo.commands.diff — PPTX change detector via snapshot comparison.

Compares current PPTX against a baseline snapshot using structured shape data,
not git diff. Reports text, position, size, font, and fill changes per shape.

Usage:
    python3 -m lo diff
    python3 -m lo diff --current my.pptx
    python3 -m lo snapshot
    python3 -m lo snapshot --name my.pptx
"""

from __future__ import annotations

import os
import re
import shutil
import zipfile
from dataclasses import dataclass
from xml.etree import ElementTree as ET

import click

DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

PML = 'http://schemas.openxmlformats.org/presentationml/2006/main'
DML = 'http://schemas.openxmlformats.org/drawingml/2006/main'


# ── Data structures ───────────────────────────────────────────────────────────

@dataclass(frozen=True)
class ShapeData:
    """Immutable shape property snapshot."""
    text: str = ""
    x: int | None = None
    y: int | None = None
    cx: int | None = None
    cy: int | None = None
    font_color: str | None = None
    font_size: int | None = None   # hundredths of a point (7200 = 72pt)
    bold: bool = False
    italic: bool = False
    fill_color: str | None = None


# ── Shape map extraction ──────────────────────────────────────────────────────

def _build_shape_map(pptx_path: str) -> dict[int, dict[str, ShapeData]]:
    """
    Extract full shape data from a PPTX file.
    Returns {slide_num: {shape_name: ShapeData}}.
    """
    result: dict[int, dict[str, ShapeData]] = {}

    def dtag(local: str) -> str:
        return f'{{{DML}}}{local}'

    def ptag(local: str) -> str:
        return f'{{{PML}}}{local}'

    def _font_color(fill_el) -> str | None:
        if fill_el is None:
            return None
        srgb = fill_el.find(f'.//{dtag("srgbClr")}')
        if srgb is not None:
            val = srgb.get('val', '')
            return f'#{val.upper()}' if val else None
        return None

    def _font_size(rpr_el) -> int | None:
        if rpr_el is None:
            return None
        sz = rpr_el.get('sz')
        return int(sz) if sz else None

    def _fill_color(sp_el) -> str | None:
        spPr = sp_el.find(f'./{ptag("spPr")}')
        if spPr is None:
            return None
        solidFill = spPr.find(f'.//{dtag("solidFill")}')
        return _font_color(solidFill)

    with zipfile.ZipFile(pptx_path, 'r') as z:
        slide_files = sorted(
            [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)],
            key=lambda n: int(re.search(r'\d+', os.path.basename(n)).group())
        )
        for slide_path in slide_files:
            slide_num = int(re.search(r'\d+', os.path.basename(slide_path)).group())
            result[slide_num] = {}

            raw = z.read(slide_path)
            try:
                root = ET.fromstring(raw)
            except ET.ParseError:
                continue

            def _extract_xfrm(el):
                """Extract position/size from any element containing <a:xfrm>."""
                xfrm = el.find(f'.//{dtag("xfrm")}')
                x = y = cx = cy = None
                if xfrm is not None:
                    off = xfrm.find(dtag('off'))
                    ext = xfrm.find(dtag('ext'))
                    if off is not None:
                        x = int(off.get('x', 0)) // 9525
                        y = int(off.get('y', 0)) // 9525
                    if ext is not None:
                        cx = int(ext.get('cx', 0)) // 9525
                        cy = int(ext.get('cy', 0)) // 9525
                return x, y, cx, cy

            # ── Standard shapes (<p:sp>) ──
            for sp_el in root.iter(ptag('sp')):
                cNvPr = sp_el.find(f'.//{ptag("cNvPr")}')
                shape_id = cNvPr.get('id', '') if cNvPr is not None else ''
                shape_name = cNvPr.get('name', '') if cNvPr is not None else ''
                # Use id-based placeholder so anonymous shapes can be tracked
                if not shape_name:
                    shape_name = f"__unnamed_{shape_id}__"

                # Text: concatenate all <a:t>
                texts = [t.text or '' for t in sp_el.iter(dtag('t'))]

                # Position/size from <a:xfrm>
                x, y, cx, cy = _extract_xfrm(sp_el)

                # Font properties from <a:rPr>
                font_color: str | None = None
                font_size: int | None = None
                bold = italic = False

                for rpr_el in sp_el.iter(dtag('rPr')):
                    if rpr_el.get('b') == '1':
                        bold = True
                    if rpr_el.get('i') == '1':
                        italic = True
                    sz = _font_size(rpr_el)
                    if sz is not None and font_size is None:
                        font_size = sz
                    # Color: try direct child solidFill first
                    for child in rpr_el:
                        if child.tag == dtag('solidFill'):
                            c = _font_color(child)
                            if c and font_color is None:
                                font_color = c

                fill_color = _fill_color(sp_el)

                result[slide_num][shape_name] = ShapeData(
                    text=''.join(texts),
                    x=x, y=y, cx=cx, cy=cy,
                    font_color=font_color,
                    font_size=font_size,
                    bold=bold,
                    italic=italic,
                    fill_color=fill_color,
                )

            # ── Non-standard elements: ink, pictures, connectors, etc. ──
            MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

            # mc:AlternateContent (ink/drawing objects)
            for ac_el in root.iter(f'{{{MC}}}AlternateContent'):
                # Find cNvPr by searching all descendants
                cNvPr = None
                for el in ac_el.iter():
                    if el.tag.endswith('}cNvPr'):
                        cNvPr = el
                        break
                shape_id = cNvPr.get('id', '') if cNvPr is not None else ''
                shape_name = cNvPr.get('name', '') if cNvPr is not None else ''
                if not shape_name:
                    shape_name = f"__ink_{shape_id}__"

                x, y, cx, cy = _extract_xfrm(ac_el)
                # Skip LibreOffice ink artifacts — they are noise added on every save
                if shape_name.startswith("__ink_"):
                    continue
                result[slide_num][shape_name] = ShapeData(
                    text=f'[{shape_name}]',
                    x=x, y=y, cx=cx, cy=cy,
                )
            for pic_el in root.iter(ptag('pic')):
                cNvPr = pic_el.find(f'.//{ptag("cNvPr")}')
                shape_id = cNvPr.get('id', '') if cNvPr is not None else ''
                shape_name = cNvPr.get('name', '') if cNvPr is not None else ''
                if not shape_name:
                    shape_name = f"__pic_{shape_id}__"

                x, y, cx, cy = _extract_xfrm(pic_el)
                result[slide_num][shape_name] = ShapeData(
                    text=f'[{shape_name}]',
                    x=x, y=y, cx=cx, cy=cy,
                )

            # <p:cxnSp> (connector lines)
            for cxn_el in root.iter(ptag('cxnSp')):
                cNvPr = cxn_el.find(f'.//{ptag("cNvPr")}')
                shape_id = cNvPr.get('id', '') if cNvPr is not None else ''
                shape_name = cNvPr.get('name', '') if cNvPr is not None else ''
                if not shape_name:
                    shape_name = f"__connector_{shape_id}__"

                x, y, cx, cy = _extract_xfrm(cxn_el)
                result[slide_num][shape_name] = ShapeData(
                    text=f'[{shape_name}]',
                    x=x, y=y, cx=cx, cy=cy,
                )

    return result


# ── Comment extraction ───────────────────────────────────────────────────────

@dataclass(frozen=True)
class CommentData:
    """A single comment on a slide."""
    comment_id: str
    author: str
    text: str
    shape_id: str = ""      # target shape id, if available


def _build_comment_map(pptx_path: str) -> dict[int, list[CommentData]]:
    """
    Extract comments from PPTX.
    Returns {slide_num: [CommentData, ...]}.

    Modern PowerPoint uses ppt/comments/modernComment_*.xml with sldMk[@sldId]
    to link comments to slides. sldId maps to slide order via presentation.xml.
    """
    result: dict[int, list[CommentData]] = {}

    with zipfile.ZipFile(pptx_path, 'r') as z:
        # Build sldId → slide_num mapping from presentation.xml
        sldid_to_num: dict[str, int] = {}
        if 'ppt/presentation.xml' in z.namelist():
            pres_root = ET.fromstring(z.read('ppt/presentation.xml'))
            for i, sld_el in enumerate(pres_root.iter(f'{{{PML}}}sldId'), start=1):
                sid = sld_el.get('id', '')
                if sid:
                    sldid_to_num[sid] = i

        # Build author map
        authors: dict[str, str] = {}
        P188 = 'http://schemas.microsoft.com/office/powerpoint/2018/8/main'
        for name in z.namelist():
            if 'authors' in name.lower():
                try:
                    aroot = ET.fromstring(z.read(name))
                    for a_el in aroot.iter(f'{{{P188}}}author'):
                        aid = a_el.get('id', '')
                        aname = a_el.get('name', '')
                        if aid:
                            authors[aid] = aname
                except ET.ParseError:
                    pass

        # Parse comment files
        PC = 'http://schemas.microsoft.com/office/powerpoint/2013/main/command'
        for name in sorted(z.namelist()):
            if not re.match(r'ppt/comments/.*\.xml$', name):
                continue
            try:
                croot = ET.fromstring(z.read(name))
            except ET.ParseError:
                continue

            for cm in croot.iter(f'{{{P188}}}cm'):
                cid = cm.get('id', '')
                author_id = cm.get('authorId', '')
                author_name = authors.get(author_id, author_id)

                # Extract text
                texts = [t.text or '' for t in cm.iter(f'{{{DML}}}t')]
                text = ''.join(texts)

                # Find target slide via sldMk
                slide_num = 1  # fallback
                sld_mk = cm.find(f'.//{{{PC}}}sldMk')
                if sld_mk is not None:
                    sid = sld_mk.get('sldId', '')
                    slide_num = sldid_to_num.get(sid, 1)

                # Find target shape
                AC = 'http://schemas.microsoft.com/office/drawing/2013/main/command'
                sp_mk = cm.find(f'.//{{{AC}}}spMk')
                shape_id = sp_mk.get('id', '') if sp_mk is not None else ''

                result.setdefault(slide_num, []).append(CommentData(
                    comment_id=cid,
                    author=author_name,
                    text=text,
                    shape_id=shape_id,
                ))

    return result


def _compare_comments(
    baseline_comments: dict[int, list[CommentData]],
    current_comments: dict[int, list[CommentData]],
) -> dict[int, list[Change]]:
    """Compare comments between baseline and current."""
    changes: dict[int, list[Change]] = {}

    all_slides = set(baseline_comments.keys()) | set(current_comments.keys())
    for sn in sorted(all_slides):
        old_set = {(c.comment_id, c.text) for c in baseline_comments.get(sn, [])}
        new_set = {(c.comment_id, c.text) for c in current_comments.get(sn, [])}

        old_texts = {c.text for c in baseline_comments.get(sn, [])}
        new_texts = {c.text for c in current_comments.get(sn, [])}

        for c in current_comments.get(sn, []):
            if (c.comment_id, c.text) not in old_set:
                label = f"💬 {c.author}" if c.author else "💬 注释"
                changes.setdefault(sn, []).append(
                    Change(label, "comment_added", None, c.text))

        for c in baseline_comments.get(sn, []):
            if (c.comment_id, c.text) not in new_set and c.text not in new_texts:
                label = f"💬 {c.author}" if c.author else "💬 注释"
                changes.setdefault(sn, []).append(
                    Change(label, "comment_deleted", c.text, None))

    return changes


# ── Shape comparison ─────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Change:
    shape: str
    prop: str
    old_val: str | bool | int | None
    new_val: str | bool | int | None

    def format_old(self) -> str:
        return _fmt_val(self.old_val)

    def format_new(self) -> str:
        return _fmt_val(self.new_val)


def _fmt_val(v) -> str:
    if v is None:
        return "unset"
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return str(v)
    return str(v)[:40]


_CHANGE_TOLERANCE_PX = 5


def _compare_shapes(
    baseline_map: dict[int, dict[str, ShapeData]],
    current_map: dict[int, dict[str, ShapeData]],
) -> tuple[dict[int, list[Change]], dict[int, set[str]]]:
    """
    Compare baseline vs current shape maps.
    Returns (changes_by_slide, shapes_needing_image).
    changes_by_slide: {slide_num: [Change(...), ...]}
    shapes_needing_image: {slide_num: {shape_name, ...}}  — shapes with layout/color changes
    """
    changes: dict[int, list[Change]] = {}
    needs_image: dict[int, set[str]] = {}

    all_slides = set(baseline_map.keys()) | set(current_map.keys())

    for sn in sorted(all_slides):
        baseline_shapes = baseline_map.get(sn, {})
        current_shapes = current_map.get(sn, {})
        all_names = set(baseline_shapes.keys()) | set(current_shapes.keys())

        for name in sorted(all_names):
            o = baseline_shapes.get(name)
            n = current_shapes.get(name)

            if o is None:
                # New shape added — include text content for anonymous shapes
                text = n.text if n else ''
                changes.setdefault(sn, []).append(Change(name, "shape_added", None, text[:40] if text else '(空)'))
                continue

            if n is None:
                # Shape deleted
                text = o.text if o else ''
                changes.setdefault(sn, []).append(Change(name, "shape_deleted", text[:40] if text else '(空)', None))
                continue

            if o == n:
                continue

            shape_changes: list[Change] = []
            layout_changed = False

            # Text
            if o.text != n.text:
                shape_changes.append(Change(name, "text", o.text, n.text))

            # Font size
            if o.font_size != n.font_size:
                old_sz = f"{o.font_size // 100}pt" if o.font_size else "unset"
                new_sz = f"{n.font_size // 100}pt" if n.font_size else "unset"
                shape_changes.append(Change(name, "font_size", old_sz, new_sz))
                layout_changed = True

            # Font color
            if o.font_color != n.font_color:
                shape_changes.append(Change(name, "font_color", o.font_color, n.font_color))
                layout_changed = True

            # Bold
            if o.bold != n.bold:
                shape_changes.append(Change(name, "bold", o.bold, n.bold))
                layout_changed = True

            # Italic
            if o.italic != n.italic:
                shape_changes.append(Change(name, "italic", o.italic, n.italic))
                layout_changed = True

            # Position (with ±5px tolerance)
            pos_changed = False
            for attr in ('x', 'y'):
                ov = getattr(o, attr)
                nv = getattr(n, attr)
                if (ov is None) != (nv is None):
                    pos_changed = True
                elif ov is not None and nv is not None and abs(ov - nv) > _CHANGE_TOLERANCE_PX:
                    pos_changed = True
            if pos_changed:
                shape_changes.append(Change(name, "position",
                                           f"({o.x},{o.y})" if o.x is not None else "unset",
                                           f"({n.x},{n.y})" if n.x is not None else "unset"))
                layout_changed = True

            # Size (with ±5px tolerance)
            size_changed = False
            for attr in ('cx', 'cy'):
                ov = getattr(o, attr)
                nv = getattr(n, attr)
                if (ov is None) != (nv is None):
                    size_changed = True
                elif ov is not None and nv is not None and abs(ov - nv) > _CHANGE_TOLERANCE_PX:
                    size_changed = True
            if size_changed:
                shape_changes.append(Change(name, "size",
                                            f"{o.cx}×{o.cy}" if o.cx is not None else "unset",
                                            f"{n.cx}×{n.cy}" if n.cx is not None else "unset"))
                layout_changed = True

            # Fill color
            if o.fill_color != n.fill_color:
                shape_changes.append(Change(name, "fill_color", o.fill_color, n.fill_color))
                layout_changed = True

            if shape_changes:
                changes.setdefault(sn, []).extend(shape_changes)
            if layout_changed:
                needs_image.setdefault(sn, set()).add(name)

    return changes, needs_image


# ── Formatting ────────────────────────────────────────────────────────────────

_PROP_LABELS: dict[str, str] = {
    "text": "文本",
    "font_size": "字号",
    "font_color": "字体颜色",
    "bold": "粗体",
    "italic": "斜体",
    "position": "位置",
    "size": "尺寸",
    "fill_color": "填充色",
    "shape_added": "新增形状",
    "shape_deleted": "删除形状",
    "comment_added": "新增注释",
    "comment_deleted": "删除注释",
}


def _format_changes(
    changes: dict[int, list[Change]],
    needs_image: dict[int, set[str]],
) -> str:
    if not changes:
        return "(no changes detected)"

    total = sum(len(v) for v in changes.values())
    lines = [f"## PPTX Changes ({total} change(s), {len(changes)} slide(s))\n"]

    for sn in sorted(changes.keys()):
        chs = changes[sn]
        if not chs:
            continue
        lines.append(f"### Slide {sn}")

        # Group by shape
        by_shape: dict[str, list[Change]] = {}
        for ch in chs:
            by_shape.setdefault(ch.shape, []).append(ch)

        img_shapes = needs_image.get(sn, set())

        for name, shape_changes in sorted(by_shape.items()):
            lines.append(f"**`{name}`**")
            for ch in shape_changes:
                label = _PROP_LABELS.get(ch.prop, ch.prop)
                old_s = ch.format_old()
                new_s = ch.format_new()
                if ch.prop in ("text", "font_size", "font_color", "bold", "italic", "position", "size", "fill_color"):
                    lines.append(f"  - {label}: `{old_s}` → `{new_s}`")
                elif ch.prop in ("shape_added", "shape_deleted"):
                    lines.append(f"  - {label}: `{name}` (文本: {ch.new_val if ch.prop == 'shape_added' else ch.old_val})")
                elif ch.prop == "comment_added":
                    lines.append(f"  - {label}: 「{ch.new_val}」")
                elif ch.prop == "comment_deleted":
                    lines.append(f"  - {label}: 「{ch.old_val}」")

            if name in img_shapes:
                lines.append(f"  ⚠️ 位置/尺寸/填充变化，**需重新生成图片**")

        lines.append("")

    return "\n".join(lines)


# ── Config helpers ────────────────────────────────────────────────────────────

def _resolve_paths(
    current_path: str | None,
    baseline_path: str | None,
) -> tuple[str, str]:
    """Resolve current and baseline PPTX paths from workflow config or explicit args."""
    if current_path is None:
        import json
        wf = os.path.join(DIR, "pptx-workflow.json")
        if os.path.exists(wf):
            with open(wf) as f:
                d = json.load(f)
            current_path = os.path.join(DIR, d.get("pptx", ""))

    if not current_path or not os.path.exists(current_path):
        raise click.ClickException(f"✗ PPTX not found: {current_path}")

    # Baseline: if not provided, look for {name}.baseline.pptx next to current
    if baseline_path is None:
        base = os.path.splitext(current_path)[0]
        candidate = f"{base}.baseline.pptx"
        if os.path.exists(candidate):
            baseline_path = candidate
        else:
            raise click.ClickException(
                f"✗ No baseline snapshot found.\n"
                f"  Run: python3 -m lo snapshot\n"
                f"  Or:  python3 -m lo diff --baseline /path/to/baseline.pptx"
            )

    return current_path, baseline_path


# ── CLI commands ─────────────────────────────────────────────────────────────

@click.command(name="diff")
@click.option("--current", "current_path", default=None, help="Current PPTX file")
@click.option("--baseline", "baseline_path", default=None, help="Baseline PPTX snapshot")
@click.option("--no-snapshot", "no_snapshot", is_flag=True, default=False,
              help="Skip auto-updating baseline after diff")
def diff(current_path: str | None, baseline_path: str | None, no_snapshot: bool):
    """
    Compare current PPTX against baseline snapshot, report all changes.

    After reporting, automatically updates the baseline to the current state
    (so the next diff shows only subsequent changes).

    Use --no-snapshot to skip the baseline update.
    """
    current_path, baseline_path = _resolve_paths(current_path, baseline_path)

    click.echo(f"Baseline:  {baseline_path}")
    click.echo(f"Current:   {current_path}\n")

    baseline_map = _build_shape_map(baseline_path)
    current_map = _build_shape_map(current_path)

    changes, needs_image = _compare_shapes(baseline_map, current_map)

    # Comments
    baseline_comments = _build_comment_map(baseline_path)
    current_comments = _build_comment_map(current_path)
    comment_changes = _compare_comments(baseline_comments, current_comments)
    for sn, chs in comment_changes.items():
        changes.setdefault(sn, []).extend(chs)

    summary = _format_changes(changes, needs_image)
    click.echo(summary)

    if needs_image:
        click.echo("\n⚠️  **需要运行 `python3 -m lo export png` 生成新图片**")

    # Auto-update baseline to current state
    if not no_snapshot:
        new_baseline = baseline_path
        shutil.copy2(current_path, new_baseline)
        click.echo(f"\n✓ Baseline auto-updated → {new_baseline}")


@click.command(name="snapshot")
@click.option("--name", "pptx_name", default=None, help="PPTX file to snapshot (default: from workflow)")
def snapshot(pptx_name: str | None):
    """
    Save current PPTX as baseline snapshot.

    The baseline is saved as {name}.baseline.pptx next to the original file.
    After confirming changes are correct, commit it with git.
    """
    if pptx_name is None:
        import json
        wf = os.path.join(DIR, "pptx-workflow.json")
        if os.path.exists(wf):
            with open(wf) as f:
                d = json.load(f)
            pptx_name = os.path.join(DIR, d.get("pptx", ""))

    if not pptx_name or not os.path.exists(pptx_name):
        raise click.ClickException(f"✗ PPTX not found: {pptx_name}")

    baseline_path = os.path.splitext(pptx_name)[0] + ".baseline.pptx"
    shutil.copy2(pptx_name, baseline_path)

    click.echo(f"✓ Snapshot saved: {baseline_path}")
