---
name: docx-to-pptx
description: Convert Chinese Word weekly reports, monthly reports, and research reports (.docx) into PowerPoint decks using a user-supplied .pptx template while preserving template styling, chapter structure, body emphasis, charts, and images. Use when Codex is asked to perform docx-to-ppt, Word-to-PPT, research-report-to-PPT, weekly-report-to-PPT, or to map a structured Chinese report into an existing PowerPoint template.
---

# DOCX to PPTX

Convert a structured Chinese report into a presentation by editing the supplied PPTX template at the OOXML level. Preserve template styling and run QA after generation.

## Required inputs

Obtain these paths before running:

- Source `.docx` report.
- Destination `.pptx` template.
- Output `.pptx` path.

Do not assume the Word document or PPT template is generic. Read [references/docx_structure.md](references/docx_structure.md) when headings or chart mapping may be incompatible. Read [references/pptx_template_strategies.md](references/pptx_template_strategies.md) before using a new template.

## Primary workflow

1. Resolve this skill directory from the location of `SKILL.md`.
2. Confirm the source DOCX and template PPTX exist.
3. Confirm the output file will not overwrite an existing file unless the user explicitly allows replacement.
4. Run the one-command driver:

```powershell
python scripts/run_conversion.py `
  --input "D:\path\report.docx" `
  --template "D:\path\template.pptx" `
  --output "D:\path\report.pptx"
```

5. Review the command exit code and the printed work-directory path.
6. If successful, report the output PPTX and the work directory containing `content.md` and `content.json`.
7. If QA reports a warning, inspect the generated PPTX and consult [references/troubleshooting.md](references/troubleshooting.md). Do not claim success while QA is failing.

Use `--force` only when the user authorized overwriting an existing output file. Use `--render-qa` only when LibreOffice `soffice` and Poppler `pdftoppm` are installed.

## Manual workflow

Use the manual commands when debugging an individual phase:

```powershell
python scripts/extract_docx.py `
  --input "D:\path\report.docx" `
  --output "D:\path\work\extracted"

python scripts/build_pptx.py `
  --template "D:\path\template.pptx" `
  --content "D:\path\work\extracted" `
  --docx "D:\path\report.docx" `
  --output "D:\path\report.pptx"

python scripts/qa_pptx.py `
  --input "D:\path\report.pptx"
```

Treat `scripts/plan_slides.py` and `scripts/convert_media.py` as diagnostic helpers, not required stages of the normal workflow. `build_pptx.py` performs its own pagination and consumes the media paths from `content.json`.

## Dependency policy

Require Python 3.9 or newer with `python-pptx` and `Pillow`. Treat `markitdown` and `cairosvg` as optional enhancements. Check imports before running; if required packages are missing, explain the missing packages and obtain approval before installing them.

## Input invariants

- Require at least one Word paragraph with style ID `a0`.
- Interpret `a0`, `a1`, and `a2` as L1, L2, and L3 headings.
- Recognize chart captions by text beginning with `图表 N`.
- Preserve the report's chart-caption-to-image order; do not merge captions.
- Require the template to contain recognizable cover, section, content, QR-code, and closing slides.
- Require the template shape names documented in [references/pptx_template_strategies.md](references/pptx_template_strategies.md).

If these invariants are not met, stop and explain the incompatibility instead of silently producing a partial deck.

## Output expectations

The driver produces:

- Final `.pptx` at the requested output path.
- A uniquely named work directory beside the output file.
- `extracted/content.json` for structured content.
- `extracted/content.md` for human review.
- `extracted/media/` for extracted images.

The output workflow must pass structural QA. Visual QA remains necessary for unfamiliar templates because text-height estimation is heuristic.
