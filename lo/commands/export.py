"""
lo.commands.export — Export slides as PNG or PDF using LibreOffice headless.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile

import click

from ..core.config import load_config


def _export_pdf(pptx_path: str, out_path: str) -> None:
    """Export PPTX as PDF using LibreOffice headless."""
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")

    with tempfile.TemporaryDirectory() as tmp:
        soffice_cmd = [
            "soffice",
            f"-env:UserInstallation=file://{tmp}",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmp,
            os.path.abspath(pptx_path),
        ]
        result = subprocess.run(soffice_cmd, capture_output=True, timeout=60)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice convert failed: {result.stderr.decode()}")

        pdf_name = os.path.basename(pptx_path).replace(".pptx", ".pdf")
        src = os.path.join(tmp, pdf_name)
        if not os.path.exists(src):
            raise RuntimeError("PDF was not created")

        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
        shutil.copy(src, out_path)
        click.echo(f"  ✓ Exported PDF: {out_path}")


def _export_png(
    pptx_path: str,
    slides_dir: str,
    slide_num: int | None,
    out_path: str,
) -> None:
    """Export slide(s) as PNG using LibreOffice headless + pymupdf."""
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")

    with tempfile.TemporaryDirectory() as tmp:
        soffice_cmd = [
            "soffice",
            f"-env:UserInstallation=file://{tmp}",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", tmp,
            os.path.abspath(pptx_path),
        ]
        result = subprocess.run(soffice_cmd, capture_output=True, timeout=60)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice convert failed: {result.stderr.decode()}")

        pdf_name = os.path.basename(pptx_path).replace(".pptx", ".pdf")
        pdf_path = os.path.join(tmp, pdf_name)
        if not os.path.exists(pdf_path):
            raise RuntimeError("PDF was not created")

        os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

        # Try pymupdf first (high quality, vector-accurate)
        try:
            import fitz  # noqa: F401
        except ImportError:
            fitz = None

        if fitz:
            doc = __import__("fitz").open(pdf_path)
            if slide_num is not None and 0 < slide_num <= len(doc):
                page = doc[slide_num - 1]
                mat = __import__("fitz").Matrix(2, 2)
                pix = page.get_pixmap(matrix=mat)
                pix.save(out_path)
            else:
                # Export all slides
                for i, page in enumerate(doc):
                    mat = __import__("fitz").Matrix(2, 2)
                    pix = page.get_pixmap(matrix=mat)
                    out = out_path.replace(".png", f"_{i+1}.png")
                    pix.save(out)
                    click.echo(f"  ✓ {out}")
            doc.close()
        else:
            # Fallback: soffice direct PNG convert
            soffice_cmd2 = [
                "soffice",
                f"-env:UserInstallation=file://{tmp}",
                "--headless",
                "--convert-to", "png",
                "--outdir", tmp,
                pdf_path,
            ]
            result2 = subprocess.run(soffice_cmd2, capture_output=True, timeout=60)
            if result2.returncode != 0:
                raise RuntimeError(f"LibreOffice PNG convert failed: {result2.stderr.decode()}")

            png_files = sorted(f for f in os.listdir(tmp) if f.endswith(".png"))
            if not png_files:
                raise RuntimeError("No PNG files generated")

            if slide_num is not None and 0 < slide_num <= len(png_files):
                shutil.copy(os.path.join(tmp, png_files[slide_num - 1]), out_path)
            else:
                for i, pf in enumerate(png_files):
                    out = out_path.replace(".png", f"_{i+1}.png")
                    shutil.copy(os.path.join(tmp, pf), out)
                    click.echo(f"  ✓ {out}")

        click.echo(f"  ✓ Exported: {out_path}")


@click.group(name="export")
def export_group():
    """Export PPTX slides as image/PDF."""
    pass


@export_group.command(name="png")
@click.option("--slide", "slide_num", type=int, default=None, help="Slide number (1-based). Omit for all slides.")
@click.option("--out", "out_path", required=True, help="Output path. Use placeholder {n} for multiple files.")
@click.option("--file", "pptx_path", default=None, help="PPTX file path")
def export_png(slide_num: int | None, out_path: str, pptx_path: str | None):
    """Export slide(s) as PNG."""
    cfg = load_config(pptx_path)
    try:
        _export_png(cfg.pptx_path, cfg.slides_dir, slide_num, out_path)
        click.echo(f"✓ PNG exported")
    except Exception as e:
        click.echo(f"✗ Export failed: {e}", err=True)
        raise SystemExit(1)


@export_group.command(name="pdf")
@click.option("--out", "out_path", required=True, help="Output PDF path.")
@click.option("--file", "pptx_path", default=None, help="PPTX file path")
def export_pdf(out_path: str, pptx_path: str | None):
    """Export PPTX as PDF."""
    cfg = load_config(pptx_path)
    try:
        _export_pdf(cfg.pptx_path, out_path)
        click.echo(f"✓ PDF exported")
    except Exception as e:
        click.echo(f"✗ Export failed: {e}", err=True)
        raise SystemExit(1)
