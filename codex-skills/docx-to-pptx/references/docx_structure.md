# DOCX input contract

Read this reference before converting a document whose Word styles or chart layout are unfamiliar.

## Required structure

The extractor reads the DOCX package directly from `word/document.xml` and `word/_rels/document.xml.rels`. It does not use `python-docx` for paragraph extraction because drawing containers and legacy image formats can be omitted by that high-level API.

Require these paragraph style IDs:

| Style ID | Meaning |
|---|---|
| `a0` | Level-1 chapter heading |
| `a1` | Level-2 heading |
| `a2` | Level-3 heading |

At least one `a0` paragraph is mandatory. The extractor returns a failure when none is found.

## Text classification

Classification uses this priority:

1. Text beginning with `图表` followed by a number becomes `chart_title`.
2. Text beginning with `资料来源：` or `资料来源:` becomes `source`.
3. Text beginning with `注：` or `注:` is omitted.
4. Style IDs `a0`, `a1`, and `a2` become headings.
5. Other text longer than five characters becomes body text.

Run formatting is preserved for bold and italic text where it is represented by Word run properties.

## Images

The extractor copies images referenced by `a:blip r:embed="rIdN"` into `extracted/media/`. PNG, JPEG, SVG, EMF, and WMF files can be extracted. Images inside Word tables and formulas are not explicitly modeled.

Chart captions and images are paired by their sequence in `content.json`. Keep the original caption and image order. A source or unrelated paragraph inserted between a caption block and its image block can prevent automatic pairing; review `content.md` when the counts differ.

## Content filtering

The current workflow:

- Omits everything before the first `a0` heading.
- Omits chapters whose L1 title contains `复盘`.
- Omits `TOC1`, `TOC2`, and `TOC3` paragraphs.
- Retains the risk-warning heading and its first following content paragraph, then omits subsequent trailing material.
- Can move a final paragraph beginning with `下周关注` onto the risk-warning slide.

Always review `extracted/content.md` before trusting a new Word layout.

## Diagnostic command

```powershell
python scripts/extract_docx.py `
  --input "D:\path\report.docx" `
  --output "D:\path\work\extracted"
```

Check heading, body, chart-caption, and image counts printed by the command and compare them with the source report.
