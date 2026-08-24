#!/usr/bin/env python3
"""Run DOCX extraction, PPTX generation, and QA as one safe command."""

from __future__ import annotations

import argparse
import subprocess
import sys
import tempfile
from pathlib import Path


def run_step(label: str, command: list[str]) -> int:
    print(f"\n[{label}] {' '.join(command)}")
    result = subprocess.run(command, check=False)
    if result.returncode:
        print(f"[ERROR] {label} failed with exit code {result.returncode}", file=sys.stderr)
    return result.returncode


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert a structured DOCX report to PPTX and run QA"
    )
    parser.add_argument("--input", required=True, help="Source DOCX path")
    parser.add_argument("--template", required=True, help="PPTX template path")
    parser.add_argument("--output", required=True, help="Destination PPTX path")
    parser.add_argument("--workdir", default=None, help="Optional empty work directory")
    parser.add_argument("--force", action="store_true", help="Allow replacing the output PPTX")
    parser.add_argument(
        "--render-qa",
        action="store_true",
        help="Render slides during QA; requires soffice and pdftoppm",
    )
    args = parser.parse_args()

    source = Path(args.input).expanduser().resolve()
    template = Path(args.template).expanduser().resolve()
    output = Path(args.output).expanduser().resolve()
    script_dir = Path(__file__).resolve().parent

    if not source.is_file() or source.suffix.lower() != ".docx":
        print(f"[ERROR] DOCX source not found or invalid: {source}", file=sys.stderr)
        return 2
    if not template.is_file() or template.suffix.lower() != ".pptx":
        print(f"[ERROR] PPTX template not found or invalid: {template}", file=sys.stderr)
        return 2
    if output.exists() and not args.force:
        print(f"[ERROR] Output exists; pass --force only if replacement is intended: {output}", file=sys.stderr)
        return 2

    output.parent.mkdir(parents=True, exist_ok=True)
    if args.workdir:
        workdir = Path(args.workdir).expanduser().resolve()
        if workdir.exists() and any(workdir.iterdir()):
            print(f"[ERROR] Work directory must be empty: {workdir}", file=sys.stderr)
            return 2
        workdir.mkdir(parents=True, exist_ok=True)
    else:
        workdir = Path(
            tempfile.mkdtemp(prefix=f"{output.stem}_docx_to_pptx_", dir=output.parent)
        )

    extracted = workdir / "extracted"
    python = sys.executable

    print(f"[INFO] Work directory: {workdir}")
    extract_cmd = [
        python,
        str(script_dir / "extract_docx.py"),
        "--input",
        str(source),
        "--output",
        str(extracted),
    ]
    if run_step("1/3 Extract DOCX", extract_cmd):
        return 1

    content_json = extracted / "content.json"
    if not content_json.is_file() or content_json.stat().st_size <= 2:
        print(f"[ERROR] Extraction produced no usable content: {content_json}", file=sys.stderr)
        return 1

    build_cmd = [
        python,
        str(script_dir / "build_pptx.py"),
        "--template",
        str(template),
        "--content",
        str(extracted),
        "--docx",
        str(source),
        "--output",
        str(output),
    ]
    if run_step("2/3 Build PPTX", build_cmd):
        return 1

    qa_cmd = [python, str(script_dir / "qa_pptx.py"), "--input", str(output)]
    if args.render_qa:
        qa_cmd.extend(["--render", "--render-dir", str(workdir / "qa_render")])
    if run_step("3/3 QA", qa_cmd):
        return 1

    print(f"\n[OK] PPTX created: {output}")
    print(f"[OK] Review artifacts: {extracted}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
