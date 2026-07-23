# Table -> Shapes: Conversion Fidelity Rules

This document specifies how the converter reproduces a PowerPoint table as a group of shapes.
It concentrates on the part that is genuinely subtle - **border resolution** - and then
catalogues, dimension by dimension, which visual properties are reproduced today and which are
not yet.

---

## 1. Borders

### 1.1 How PowerPoint models a border

A table is a grid of cells; every cell has four independently styleable edges
(`lnL/lnR/lnT/lnB` in OOXML, `Cell.Borders[...]` in Interop), each with
`Visible, Weight, Color, DashStyle, Transparency`. A line between two cells is a **single
physical edge** that *both* adjacent cells describe. The two descriptions can disagree - most
often after a merge - and PowerPoint renders **one** resolved line per physical edge.

The correct unit of work is therefore the **physical edge segment**, not the cell. Emitting a
border per cell and reconciling the duplicates afterwards makes the result depend on emission
order; resolving per edge does not.

### 1.2 Algorithm

Reconstruct the grid tracks (`colOffsets`, `rowOffsets`) and precompute an **anchor map**: for
every grid slot `(i,j)`, the top-left slot of the merge it belongs to. Then enumerate every
grid-line segment:

- **Vertical**: for each column boundary `j in [0..cols]` and each row `i in [0..rows-1]`,
  a segment at `x = colOffset[j]`, `y in [rowOffset[i], rowOffset[i+1]]`.
  - negative-side cell = anchor of `(i, j-1)` if `j > 0`, else none (table edge);
  - positive-side cell = anchor of `(i, j)` if `j < cols`, else none.
  - If both sides resolve to the **same anchor**, the boundary is interior to a merge -> skip.
  - negative opinion = negCell.`BorderRight`; positive opinion = posCell.`BorderLeft`.
- **Horizontal**: symmetric, using `BorderTop` / `BorderBottom`.

`Resolve(negOpinion, posOpinion)` returns 0 or 1 rendered segment. Each physical segment is
visited exactly once, so there is no de-duplication and no order dependence.

### 1.3 Rules

An *opinion* is `{ exists, fromMerged, isPositiveSide, border }`.

- **R1 - Interior merge edges produce nothing.** A grid line inside a merged cell is not a border.
- **R2 - Table-boundary edges use the single existing side.** Visible -> draw it; off -> nothing.
- **R3 - Plain beats merged.** If one side is a merged cell and the other is plain, the plain side
  is authoritative; discard the merged opinion (including when the plain opinion is "off"). A
  merged cell reports a stray "automatic" border (usually black, still `Visible`) on its outer
  edges after a merge, while the plain neighbour holds the real value. When the two were instead
  edited by the user, PowerPoint keeps them in sync, so they already agree and picking the plain
  side is harmless.
- **R4 - Visible beats off (same tier).** Among the surviving opinions, if any is visible a line
  is drawn; only if *every* surviving opinion is off is there no line.
- **R5 - Heavier wins (same tier).** Among visible survivors the greatest `Weight` wins; it
  carries its own colour / dash / transparency.
- **R6 - Deterministic tie-break.** Equal weight and tier -> take the negative (left/top) side.
  This removes any dependence on iteration order.

Worked examples (merged cell reports a stray black border; plain neighbour holds the real value):

| Negative opinion | Positive opinion | Result | Rule |
|---|---|---|---|
| plain, off | merged, black | no line | R3 -> plain off |
| plain, white on | merged, black | white | R3 -> plain on |
| merged, black | plain, white on | white | R3 -> plain on |
| merged, black | plain, black on | black | R3 -> plain on |
| plain, black on | none (boundary) | black | R2 |

### 1.4 "Off" vs "on" between two plain cells

PowerPoint keeps the shared border of two adjacent **plain** cells in sync as you edit either
one (observed behaviour: "the latest change wins"), so a stable disagreement between two plain
cells does not occur - both store the last-applied value. The only genuine disagreements come
from merges, which leave a stale value on the merged side. Consequently:

- There is **no suppressor rule** (an early design that let "off" delete a neighbour's line was
  wrong; its failure mode is erasing a real border).
- For any residual plain-vs-plain disagreement, **visible wins** (R4) - never erase a border a
  cell actually defines.
- The substance is carried by **R3 (plain beats merged)**.

### 1.5 Caveats

- **R3 assumes merged-vs-plain edits stay in sync.** If a deck ever presents a merged cell whose
  outer border was intentionally set different from its neighbour *and* PowerPoint did not sync
  the neighbour, R3 would drop it. The narrow fix would be to discard a merged opinion only when
  it *conflicts* with the plain side, rather than unconditionally.
- **Two merged cells meeting on one edge** have no plain side and fall through to R5/R6; both
  could be stray borders. Rare, currently not handled specially.
- **R6's tie-break is for determinism only.** It fires only for two same-tier visible borders of
  equal weight but different style, which does not occur for synced plain cells. The negative
  side is chosen arbitrarily; it is not a claim about PowerPoint's own preference.
- **Coordinates are floats.** Track offsets are accumulated sums and `EdgePlacement` comparisons
  use an epsilon; read-time rounding keeps merge keys and edges consistent.
- **The stray merged border appears on a merged cell's interior-facing edges**, where R3 handles
  it. Outer/table-boundary edges of a merged cell read correctly and are used as-is (R2).

---

## 2. Property coverage

Legend: **[H]** handled - **[P]** partial - **[M]** not yet.

### 2.1 Geometry & merges
- Table origin from the minimum cell edge. **[H]**
- Column widths / row heights reconstructed from actual cell rectangles, falling back to declared
  track sizes if an interior boundary is hidden by a merge. **[P]** - auto-grown (multi-line) rows
  are a known risk: if Interop under-reports a cell's `Height`, a row can collapse to its declared
  minimum. This should be confirmed from a real multi-line table before changing the
  reconstruction.
- Merge detection via a geometry hash of `Shape.Left/Top/Width/Height` (all covered cells report
  identical geometry; `Shape.Id` is `E_NOTIMPL` for table cells). Works but is sensitive to
  rounding. **[P]** - a span-based detection with the geometry hash as fallback would be sturdier.

### 2.2 Fills
- Solid fill: visible, RGB (theme resolved to literal RGB), transparency. **[H]**
- Gradient / pattern / picture / texture fills. **[M]** - approximated as a solid colour.

### 2.3 Borders
- Per physical edge, resolved as in section 1 (visible, weight, colour, dash, transparency). **[H]**
- Cap / compound line styles. **[M]** - low priority.

### 2.4 Text frame / cell
- Margins (L/R/T/B), vertical anchor, word wrap, `AutoSize = none`. **[H]**
- Text rotation / vertical text / right-to-left direction. **[M]**
- Multiple text columns in a cell. **[M]** - rare.

### 2.5 Paragraphs
- Alignment, space before/after/within, indent level. **[H]**
- Bullets / numbering (list format, glyph, start-at, colour, size). **[M]** - visible and common.
- Line-spacing rule (multiple vs exact points): `SpaceWithin` is copied but the rule type is not
  distinguished. **[P]**
- Right-to-left paragraphs. **[M]**

### 2.6 Runs (character)
- Text (written verbatim so line/paragraph breaks survive), font name plus complex-script and
  East-Asian faces, size, bold, italic, underline style, strikethrough, fill colour, highlight.
  **[H]**
- Underline colour (distinct from text colour), subscript / superscript, all-caps / small-caps,
  character spacing / kerning, hyperlinks. **[M]**
- Text effects (glow, shadow, reflection, soft edge). **[M]** - low priority.

### 2.7 Grouping & placement
- All generated shapes are grouped into one (a single-shape result is kept as-is, since grouping
  needs two or more); the group origin is re-asserted to the content's true top-left; the
  original table is deleted only **after** the replacement is built and positioned, so a failure
  never leaves the slide without a table. Fills are painted first, then border lines on top;
  on failure the partial or grouped output is removed. **[H]**

---

## 3. Prioritised gaps

1. **Bullets / numbering** in paragraphs (common, currently not reproduced).
2. **Merge detection** hardening (span-based, with the geometry hash as fallback).
3. **Auto-grown row heights** for multi-line cells (confirm the `Height` behaviour first).
4. **Underline colour, sub/superscript, caps, character spacing.**
5. **Hyperlinks.**
6. **Gradient / pattern / picture fills** (approximate faithfully and flag).
7. **Text rotation / RTL / text effects.**

---

## 4. Validation strategy

- **Pure-logic unit tests.** The border-resolution function and all geometry are testable
  without Office; there is a test per rule (R1-R6) plus combinations of merges, borderless cells
  and stray borders.
- **Golden-image E2E.** A fixture builds a table exercising the fidelity-sensitive features
  together (styled runs, highlight, merges, borderless cells), converts it, and asserts a
  near-zero pixel diff between the original and the converted group. New feature families extend
  this fixture.
- **One place for precedence.** All border precedence lives in `Resolve`, so a question about
  "what should PowerPoint draw here" is answered by a small targeted table rather than scattered
  special cases.
- **Optional diagnostics.** With `TABLETOSHAPES_DIAGNOSTICS` set, each conversion logs the parsed
  cells, runs and resolved edges - useful when reconciling a fidelity difference against what the
  reader actually captured.
