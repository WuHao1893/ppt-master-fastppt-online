#!/usr/bin/env python3
"""Render an editable PPTX with the installed Microsoft PowerPoint COM server.

This worker intentionally fails when PowerPoint/pywin32 is unavailable. The API
then records an SVG fallback instead of fabricating an authoritative render.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Render PPTX slides using PowerPoint COM.")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output-dir", required=True)
    args = parser.parse_args(argv)
    input_path = Path(args.input).resolve()
    output_dir = Path(args.output_dir).resolve()
    if not input_path.is_file():
        print(f"PPTX input does not exist: {input_path}", file=sys.stderr)
        return 2
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        print("PowerPoint COM renderer requires pywin32 on a Windows PowerPoint worker.", file=sys.stderr)
        return 2
    application = None
    presentation = None
    try:
        application = win32com.client.DispatchEx("PowerPoint.Application")
        presentation = application.Presentations.Open(str(input_path), WithWindow=False, ReadOnly=True)
        renders = []
        for index in range(1, presentation.Slides.Count + 1):
            output_path = output_dir / f"slide-{index:03d}.png"
            presentation.Slides(index).Export(str(output_path), "PNG", 1920, 1080)
            renders.append({"slide_index": index, "path": str(output_path), "width": 1920, "height": 1080})
        qa = {
            "pptx": str(input_path),
            "renderer": "microsoft-powerpoint-com",
            "status": "passed",
            "slide_count": len(renders),
            "renders": renders,
        }
        qa_path = output_dir / "powerpoint.qa.json"
        qa_path.write_text(json.dumps(qa, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"renderer": qa["renderer"], "qa": str(qa_path), "renders": renders}, ensure_ascii=False))
        return 0
    except Exception as exc:  # COM errors need to be surfaced to the API.
        print(f"PowerPoint COM render failed: {exc}", file=sys.stderr)
        return 1
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
