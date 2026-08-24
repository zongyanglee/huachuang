# Troubleshooting

## Missing required packages

The core workflow requires Python 3.9+, `python-pptx`, and `Pillow`.

```text
ModuleNotFoundError: No module named 'pptx'
ModuleNotFoundError: No module named 'PIL'
```

Check the Python executable used by Codex and install packages only after obtaining approval. `markitdown` improves text QA but the checker falls back to `python-pptx`. `cairosvg` is needed only for optional SVG conversion.

## No usable content extracted

Symptoms:

- The extractor reports `未找到一级标题（a0 样式）`.
- `content.json` is empty.

Cause: the source Word template does not use the expected `a0`, `a1`, and `a2` style IDs. Inspect Word paragraph styles or adapt `heading_style_ids` in `extract_docx.py` for the new template.

## Template role missing

Symptom: `模板缺少必要的幻灯片类型`.

Open the template Selection Pane and compare the slide text and shape names with [pptx_template_strategies.md](pptx_template_strategies.md). The template must provide cover, section, content, QR-code, and closing role slides.

## Blank title or body

The slide role may have been recognized while its fill shape has a different name. Check:

- `内容占位符 9` for the cover title.
- `文本框 15` and `Rectangle 1` for section slides.
- `标题 1` and `矩形 8` for content slides.
- `内容占位符 11` for chart-title styling.

## Chart captions do not match images

Review `extracted/content.md` and `extracted/content.json`. Caption and image blocks are paired in source order. Do not merge captions. If an unrelated item breaks the block, adjust the Word order or the mapping logic before rebuilding.

## Images are missing or overflow

- Confirm the media path in `content.json` exists.
- Confirm Pillow can read the image dimensions.
- Use `convert_media.py` only when SVG or legacy vector formats cannot be handled directly.
- Remember that `convert_media.py` does not automatically rewrite media paths in `content.json`.
- For overflow, reduce body content, split the slide, or adjust the fixed 16:9 layout constants.

## PowerPoint asks to repair the file

Likely causes include malformed OOXML relationships, namespace prefixes, or duplicated picture IDs. Do not bypass `fix_ns()`, notes-slide relationship removal, or the global picture-ID counter in `build_pptx.py`.

## QA fails

Run:

```powershell
python scripts/qa_pptx.py --input "D:\path\output.pptx"
```

Add `--render` only when both `soffice` and `pdftoppm` are available. A render warning is treated as QA failure. Structural QA cannot guarantee visual correctness; inspect rendered slides or open the presentation for unfamiliar templates.
