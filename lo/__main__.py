"""
lo — CLI tool for Agent to surgically edit PPTX shapes by name.

Usage:
    python -m lo shape set-text --name cover_title_main --text "Hello"
    python -m lo shape set-fill --name cover_accent_bar --color "#FF3B5C"
    python -m lo shape move --name cover_title_main --x 60 --y 200
    python -m lo shape resize --name cover_title_main --w 960 --h 120
    python -m lo export png --slide 1 --out slides/slide1.png
    python -m lo diff       # Show PPTX changes vs baseline snapshot
    python -m lo reload     # Hot-reload in LibreOffice
"""

from __future__ import annotations

import sys

from .commands.shape import shape_group
from .commands.reload import reload as reload_cmd
from .commands.export import export_group


def main():
    import argparse
    parser = argparse.ArgumentParser(prog="lo", description="PPTX surgical editor for agents")
    parser.add_argument("--version", action="store_true", help="Show version")
    sub = parser.add_subparsers(dest="command")

    # Register shape subcommands
    shape_sub = sub.add_parser("shape", help="Shape operations")
    shape_sub.add_argument("op", choices=["set-text", "set-fill", "set-text-color", "set-font-size", "move", "resize", "get-info"])
    shape_sub.add_argument("--name")
    shape_sub.add_argument("--text", dest="text")
    shape_sub.add_argument("--color", dest="color")
    shape_sub.add_argument("--font-size", type=int, dest="font_size")
    shape_sub.add_argument("--x", type=int)
    shape_sub.add_argument("--y", type=int)
    shape_sub.add_argument("--w", type=int)
    shape_sub.add_argument("--h", type=int)
    shape_sub.add_argument("--file", dest="pptx_path")

    # Register reload
    reload_sub = sub.add_parser("reload", help="Hot-reload in LibreOffice")
    reload_sub.add_argument("--open", action="store_true", dest="do_open")
    reload_sub.add_argument("--file", dest="pptx_path")

    # Register export
    export_sub = sub.add_parser("export", help="Export slides")
    export_sub.add_argument("fmt", choices=["png", "pdf"])
    export_sub.add_argument("--slide", type=int, dest="slide_num")
    export_sub.add_argument("--out", required=True, dest="out_path")
    export_sub.add_argument("--file", dest="pptx_path")

    # Register diff
    diff_sub = sub.add_parser("diff", help="Show PPTX changes vs baseline snapshot")
    diff_sub.add_argument("--current", dest="current_path")
    diff_sub.add_argument("--baseline", dest="baseline_path")
    diff_sub.add_argument("--no-snapshot", dest="no_snapshot", action="store_true")

    # Register snapshot
    snap_sub = sub.add_parser("snapshot", help="Save current PPTX as baseline snapshot")
    snap_sub.add_argument("--name", dest="pptx_name")

    # Register init
    init_sub = sub.add_parser("init", help="Generate a new PPTX template")
    init_sub.add_argument("-o", "--output", dest="output_path", default=None)

    args = parser.parse_args(sys.argv[1:])

    if args.version:
        print("lo v0.1")
        return

    if args.command == "shape":
        _run_shape(args)
    elif args.command == "reload":
        _run_reload(args)
    elif args.command == "export":
        _run_export(args)
    elif args.command == "diff":
        _run_diff(args)
    elif args.command == "snapshot":
        _run_snapshot(args)
    elif args.command == "init":
        _run_init(args)
    else:
        parser.print_help()


def _run_shape(args):
    from pptx import Presentation
    from .core.config import load_config
    from .core.pptx_ops import apply_operation

    cfg = load_config(args.pptx_path)
    prs = Presentation(cfg.pptx_path)

    name = args.name
    op = args.op
    value = None

    if op == "set-text":
        value = args.text
    elif op == "set-fill":
        value = args.color
    elif op == "set-text-color":
        value = args.color
    elif op == "set-font-size":
        value = args.font_size
    elif op == "move":
        value = (args.x, args.y)
    elif op == "resize":
        value = (args.w, args.h)

    if op == "get-info":
        from .core.shape_finder import find_shape_by_name
        from .core.pptx_ops import get_shape_info
        shape = find_shape_by_name(prs, name)
        if shape is None:
            print(f"✗ Shape not found: {name!r}", file=sys.stderr)
            sys.exit(1)
        info = get_shape_info(shape)
        for k, v in info.items():
            print(f"  {k}: {v}")
        return

    try:
        apply_operation(prs, cfg.pptx_path, name, op, value)
        print(f"✓ {name}: {op} → {value}")
    except KeyError as e:
        print(f"✗ {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"✗ {e}", file=sys.stderr)
        sys.exit(1)


def _run_reload(args):
    from .commands.reload import reload as do_reload
    from .core.config import load_config
    cfg = load_config(args.pptx_path)
    cmd_args = []
    if args.do_open:
        cmd_args.append("--open")
    if cfg.pptx_path:
        cmd_args.extend(["--file", cfg.pptx_path])
    ctx = do_reload.make_context("reload", cmd_args)
    do_reload.invoke(ctx)


def _run_export(args):
    from .commands.export import _export_png, _export_pdf
    from .core.config import load_config
    cfg = load_config(args.pptx_path)
    try:
        if args.fmt == "pdf":
            _export_pdf(cfg.pptx_path, args.out_path)
            print(f"✓ PDF: {args.out_path}")
        else:
            _export_png(cfg.pptx_path, cfg.slides_dir, args.slide_num, args.out_path)
            print(f"✓ PNG: {args.out_path}")
    except Exception as e:
        print(f"✗ {e}", file=sys.stderr)
        sys.exit(1)


def _run_diff(args):
    from .commands.diff import diff as do_diff
    cmd_args = []
    if args.current_path:
        cmd_args.extend(["--current", args.current_path])
    if args.baseline_path:
        cmd_args.extend(["--baseline", args.baseline_path])
    if args.no_snapshot:
        cmd_args.append("--no-snapshot")
    ctx = do_diff.make_context("diff", cmd_args)
    do_diff.invoke(ctx)


def _run_snapshot(args):
    from .commands.diff import snapshot as do_snapshot
    cmd_args = []
    if args.pptx_name:
        cmd_args.extend(["--name", args.pptx_name])
    ctx = do_snapshot.make_context("snapshot", cmd_args)
    do_snapshot.invoke(ctx)


def _run_init(args):
    from .commands.init import init as do_init
    cmd_args = []
    if args.output_path is not None:
        cmd_args.extend(["-o", args.output_path])
    ctx = do_init.make_context("init", cmd_args)
    do_init.invoke(ctx)


if __name__ == "__main__":
    main()
