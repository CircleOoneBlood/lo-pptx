"""
lo.core.config — PPTX path resolution.

Resolution order:
  1. CLI --file argument
  2. pptx-workflow.json "pptx" field
  3. glob *.pptx excluding *baseline* (first match)
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass


DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKFLOW_JSON = os.path.join(DIR, "..", "pptx-workflow.json")


@dataclass(frozen=True)
class Config:
    pptx_path: str
    slides_dir: str

    def resolve(self) -> tuple[str, str]:
        """Return resolved absolute paths."""
        pptx = os.path.abspath(self.pptx_path)
        slides = os.path.abspath(self.slides_dir)
        return pptx, slides


def load_config(cli_path: str | None = None) -> Config:
    if cli_path:
        pptx_path = cli_path
        slides_dir = os.path.join(os.path.dirname(pptx_path), "slides")
        return Config(pptx_path=pptx_path, slides_dir=slides_dir)

    # Try pptx-workflow.json
    if os.path.exists(WORKFLOW_JSON):
        with open(WORKFLOW_JSON) as f:
            d = json.load(f)
        pptx_path = os.path.join(DIR, "..", d.get("pptx", ""))
        slides_dir = os.path.join(DIR, "..", d.get("slides_dir", "slides"))
        return Config(pptx_path=pptx_path, slides_dir=slides_dir)

    # Fallback: glob for *.pptx excluding baseline
    parent = os.path.dirname(DIR)
    for f in os.listdir(parent):
        if f.endswith(".pptx") and "baseline" not in f:
            pptx_path = os.path.join(parent, f)
            slides_dir = os.path.join(parent, "slides")
            return Config(pptx_path=pptx_path, slides_dir=slides_dir)

    raise FileNotFoundError(
        "No PPTX file found. Create pptx-workflow.json or pass --file."
    )
