# PPTX template contract

Read this reference before using a new PowerPoint template. The generator edits copied slide XML and depends on recognizable slide roles and shape names.

## Slide roles

The template must contain at least one slide for each role:

| Role | Recognition rule |
|---|---|
| Cover | Contains `周冠南` or `SAC`, or contains shape `文本占位符 6` |
| Section | Slide text contains `SECTION` |
| Content | Contains shapes `矩形 8` and `标题 1` |
| QR code | Slide text contains `欢迎关注` or `公众号` |
| Closing | Slide text contains `免责声明` or `本材料仅供` |

Generation stops when any role is missing.

## Required shape names

The selected role slides should contain:

| Slide | Shape | Purpose |
|---|---|---|
| Cover | `内容占位符 9` | Multiline report title |
| Section | `文本框 15` | `SECTION N` label |
| Section | `Rectangle 1` | Chapter title |
| Content | `标题 1` | Page title |
| Content | `矩形 8` | Body text |
| Content | `内容占位符 11` | Chart-title style to clone |

Missing fill shapes may produce blank content without a classification error. Inspect shape names in PowerPoint's Selection Pane when adapting a template.

## Geometry

The layout code assumes a 16:9 slide measuring approximately 13.333 × 7.5 inches. It uses fixed page margins, a maximum of two charts per content slide, and heuristic body-text height estimation. A different slide size or substantially different typography requires code adjustments and visual QA.

## OOXML behavior

The generator:

1. Extracts the PPTX ZIP package into a unique operating-system temporary directory.
2. Copies an existing slide XML for each required page.
3. Replaces text while preserving run properties.
4. Removes copied template pictures from generated content pages.
5. Inserts report images with unique picture IDs.
6. Removes copied notes-slide relationships.
7. Normalizes OOXML namespace prefixes before packaging.

Do not replace this workflow with `python-pptx add_slide()` or generic text boxes when template fidelity is required.

## Preflight checklist

- Confirm all five slide roles can be recognized.
- Confirm the required shape names exist on the selected role slides.
- Confirm the deck is 16:9.
- Keep a separate source template; write the generated presentation to a different path.
- Run structural QA and visually inspect unfamiliar templates.
