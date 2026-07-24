# Table-to-Shapes: what it supports

This feature turns a selected PowerPoint table into a group of plain shapes (boxes and lines)
that looks the same as the original table, then removes the table. This page lists, in plain
terms, what carries over faithfully, what has a known limitation, and what isn't handled yet.

**Status**
- **Yes** - reproduced faithfully.
- **Partial** - works, with a known limitation noted.
- **No** - not handled yet (planned).

**How it's checked**
- **Automated test** - there is a test in the build that fails if this stops working. Two kinds:
  logic tests (run anywhere) and picture-comparison tests that convert a real table and compare
  the before/after images pixel for pixel.
- **Manual / expected** - the code handles it and it has been checked by eye, but there isn't a
  dedicated automated test asserting it yet.

---

## Text and its formatting

| Capability | Status | How it's checked | Notes |
|---|---|---|---|
| Cell text and wording | Yes | Automated (picture) | |
| Multiple styles within one cell (e.g. one word bold or red) | Yes | Automated (picture) | |
| Line breaks / multiple paragraphs in a cell | Partial | Automated (picture) | Text and breaks preserved; a very tall auto-grown row can occasionally under-size (see Caveats). |
| Bold | Yes | Automated (picture) | |
| Strikethrough | Yes | Automated (picture) | |
| Font name (e.g. Arial vs Consolas) | Yes | Automated (picture) | The font must be installed on the machine. |
| Font colour | Yes | Automated (picture) | |
| Highlight colour (text marker) | Yes | Automated (picture) | |
| Italic, underline, font size | Yes | Manual / expected | Handled in code; not individually asserted by a test. |
| Paragraph alignment, spacing, indent | Yes | Manual / expected | |
| Bullets and numbered lists | Yes | Automated (picture) | Glyph, numbering style/start, size, font and colour reproduced; a custom hanging-indent ruler falls back to PowerPoint's default for the level. |
| Superscript / subscript, small caps, letter spacing, underline colour | No | - | Planned. |
| Hyperlinks | No | - | Link styling may show, but the clickable link is not kept. |

## Colours and fills

| Capability | Status | How it's checked | Notes |
|---|---|---|---|
| Solid cell background colour | Yes | Automated (picture) | |
| Cell with no fill | Yes | Manual / expected | |
| Fill transparency | Yes | Manual / expected | |
| Gradient / picture / pattern fills | Partial | - | Shown as a single solid colour, not the gradient/image. |
| Theme colours | Yes | Manual / expected | Captured as fixed colours; the result will not restyle if the deck's theme changes later. |

## Borders and gridlines

| Capability | Status | How it's checked | Notes |
|---|---|---|---|
| Cell borders: on/off, colour, thickness | Yes | Automated (logic + picture) | |
| Shared line between two cells drawn once | Yes | Automated (logic) | |
| Cells with borders switched off ("no border") | Yes | Automated (logic + picture) | |
| Correct border where a merged cell meets normal cells | Yes | Automated (logic) | |
| Dashed / dotted line styles | Yes | Manual / expected | |
| Diagonal cell borders | No | - | Planned. |

## Merged cells and layout

| Capability | Status | How it's checked | Notes |
|---|---|---|---|
| Horizontally merged cells | Yes | Automated (logic + picture) | |
| Vertically merged cells | Yes | Automated (logic + picture) | |
| Block merges (e.g. 2x2) | Yes | Automated (picture) | |
| Column widths and single-line row heights | Yes | Automated (picture) | |
| Table position on the slide | Yes | Automated (picture) | |
| Very tall (auto-grown, multi-line) rows | Yes | Automated (picture) | Row heights come from the actual laid-out cell rectangles; guarded by a multi-paragraph cell in the test fixture. |
| Rotated tables | No | - | Planned. |
| Pictures / charts / other shapes inside a cell | No | - | Only text, fill and borders are carried over. |

## The conversion itself

| Capability | Status | How it's checked | Notes |
|---|---|---|---|
| Original table removed and replaced by one grouped shape | Yes | Automated (picture) | |
| Clean up if conversion fails partway | Yes | Manual / expected | The slide is left as it was found. |
| Doing nothing (with a prompt) when a non-table is selected | Yes | Manual / expected | The "reject a non-table" guard is covered by an E2E test; the prompt itself is checked by eye. |
| Undo in one step | No | - | Reverting may take more than one Ctrl+Z. |
| Editing the result as a table afterwards | No | - | The result is shapes, not a table; conversion is one-directional. |

---

## Caveats worth knowing

- **One-way conversion.** The output is a group of shapes, not a table; you cannot turn it back
  into a table or edit rows/columns.
- **Fonts and themes are "baked in".** Missing fonts are substituted by PowerPoint; colours are
  fixed at conversion time and won't follow later theme changes.
- **Only fill, text and borders are copied** from each cell - embedded pictures, charts or shapes
  inside a cell are dropped.
- **Autofit "shrink text to fit" is not reproduced** - text keeps its original sizes.
- **Custom bullet indentation** (a hand-adjusted ruler) falls back to PowerPoint's default hanging
  indent for that level; the bullet glyph, numbering, size, font and colour are reproduced.

## How thoroughly it's verified

"Automated (picture)" items are exercised by a test that builds a real table containing that
feature, converts it, and compares the before/after images. "Automated (logic)" items are
covered by fast tests of the border and layout rules that run without PowerPoint. "Manual /
expected" items are implemented and checked by eye but do not yet have a dedicated automated
assertion - the natural next step is to add one per row.

For the engineering-level detail (the exact border-resolution rules and a technical
handled/partial/not-yet matrix) see `FIDELITY_RULES.md`.
