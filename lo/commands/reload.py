"""
lo.commands.reload — Hot-reload PPTX in running LibreOffice instance.
"""

from __future__ import annotations

import os
import subprocess
import sys
import time

import click

DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
HOST = "localhost"
PORT = 2002


def launch_libreoffice(pptx_path: str) -> None:
    """Launch LibreOffice Impress with UNO socket enabled."""
    cmd = [
        "soffice",
        f"--accept=socket,host={HOST},port={PORT};urp;StarOffice.ServiceManager",
        "--impress",
        os.path.abspath(pptx_path),
    ]
    click.echo(f"  Launching: {' '.join(cmd)}")
    subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    click.echo(f"  Waiting for LibreOffice to start...")
    time.sleep(4)


def reload_via_uno(pptx_path: str) -> bool:
    """Reload file in running LibreOffice via UNO. Returns True on success."""
    try:
        import uno
        from com.sun.star.beans import PropertyValue
    except ImportError:
        click.echo("  python-uno not installed; skipping hot-reload.")
        return False

    try:
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext
        )
        ctx = resolver.resolve(
            f"uno:socket,host={HOST},port={PORT};urp;StarOffice.ComponentContext"
        )
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        abs_path = os.path.abspath(pptx_path)
        url = uno.systemPathToFileUrl(abs_path)

        # Close existing document if open
        for comp in desktop.Components:
            try:
                if hasattr(comp, "getURL") and comp.getURL() == url:
                    comp.close(True)
                    break
            except Exception:
                pass

        # Reopen
        props = []
        pv = PropertyValue()
        pv.Name = "MacroExecutionMode"
        pv.Value = 4
        props.append(pv)
        desktop.loadComponentFromURL(url, "_blank", 0, tuple(props))
        click.echo(f"  ✓ Reloaded: {os.path.basename(pptx_path)}")
        return True

    except Exception as e:
        click.echo(f"  UNO connection failed ({e}); skipping hot-reload.")
        click.echo(f"  To enable, start LibreOffice with:")
        click.echo(f'    soffice --accept="socket,host={HOST},port={PORT};urp;StarOffice.ServiceManager" --impress {pptx_path}')
        return False


@click.command(name="reload")
@click.option("--open", "do_open", is_flag=True, help="Launch LibreOffice if not running")
@click.option("--file", "pptx_path", default=None, help="PPTX file path")
def reload(do_open: bool, pptx_path: str | None):
    """Hot-reload PPTX in running LibreOffice."""
    if pptx_path is None:
        from ..core.config import load_config
        cfg = load_config()
        pptx_path = cfg.pptx_path

    if do_open:
        launch_libreoffice(pptx_path)
    else:
        ok = reload_via_uno(pptx_path)
        if not ok:
            click.echo("  Hint: use --open to launch LibreOffice first.")
