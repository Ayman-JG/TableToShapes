# Table → Shapes: Conversion Fidelity Rules

Design spec for making the add-in generically correct, not case-by-case. Covers the
principled border model, a full rule set per fidelity dimension, a gap analysis of the
current code, and an implementation plan.

---

## 1. Why this document

The conversion has been fixed incident by incident (dropped text runs, highlight, fonts,
strikethrough, a phantom merged-border line, a merged-vs-plain border conflict). The text
and fill fixes are sound and general. The **border logic, however, has become a stack of
special cases** layered on one another:

- borders emitted per merge-anchor, decomposed into per-track segments;
- style sourced from the anchor cell (continuation cells are unreliable);
- a de-dup that keeps the *heavier* line, tie broken by *insertion order*;
- a `FromMerged` tag so a plain border beats a merged one;
- a *suppressor* pass that deletes any edge coincident with an explicit "off" border.

These rules interact through **list-insertion order**, which is fragile, and at least one
of them — the blanket suppressor — is **not how PowerPoint actually behaves** (see §2.4).
They happen to pass the two decks tested because those decks have symmetric border
settings. A third deck will likely break them.

> **Status:** the emit → de-dup → suppress pipeline described below has been **replaced** by
> the single, order-independent **edge-resolution model** in `LayoutEngine`
> (`Calculate` → `ResolveEdges` → `Resolve`). This section is the spec for that model. The
> remaining fidelity gaps in text and geometry (§3–§4) are still open.

---

## 2. Borders — the core redesign

### 2.1 How PowerPoint models a border

A table is a grid of cells; every cell has four independently styleable edges
(`lnL/lnR/lnT/lnB` in OOXML, `Cell.Borders[...]` in Interop), each with
`Visible, Weight, Color, DashStyle, Transparency`. A line between two cells is a **single
physical edge** that *both* adjacent cells describe. The two descriptions can disagree
(e.g. after a merge, or after "No Border" on one side); PowerPoint renders **one** resolved
line per physical edge.

So the correct unit of work is the **physical edge segment**, not the cell. Everything the
current code does with de-dup and suppression is an attempt to reconstruct this after the
fact from per-cell emissions — which is exactly what causes the order dependence.

### 2.2 Physical-edge algorithm

Reconstruct the grid tracks (already done: `colOffsets`, `rowOffsets`) and precompute an
**anchor map**: for every grid slot `(i,j)`, the top-left slot of the merge it belongs to.

Enumerate every candidate segment on the grid lines:

- **Vertical**: for each column boundary `j ∈ [0..cols]` and each row `i ∈ [0..rows-1]`,
  segment at `x = colOffset[j]`, `y ∈ [rowOffset[i], rowOffset[i+1]]`.
  - negative-side cell = anchor of `(i, j-1)` if `j > 0`, else none (table edge);
  - positive-side cell = anchor of `(i, j)` if `j < cols`, else none.
  - If both sides resolve to the **same anchor**, the boundary is *interior to a merge* → skip.
  - negative opinion = negCell.`BorderRight`; positive opinion = posCell.`BorderLeft`.
- **Horizontal**: symmetric, using `BorderTop` / `BorderBottom`.

Then `resolve(negOpinion, posOpinion)` → 0 or 1 rendered segment. No de-dup, no suppressors,
no insertion-order effects: each physical segment is visited exactly once.

### 2.3 The rule set

An *opinion* is `{ exists, fromMerged, border }` where `border` carries
`Visible/Weight/Color/Dash/Transparency`.

- **R1 — Interior merge edges produce nothing.** A grid line inside a merged cell is not a
  border.
- **R2 — Table-boundary edges use the single existing side.** Visible → draw it; off → nothing.
- **R3 — Plain beats merged.** If one side is a merged cell and the other is plain, the plain
  side is authoritative (merged cells emit stray/automatic borders on their outer edges).
  Discard the merged opinion entirely — including when the plain opinion is "off".
- **R4 — Visible beats off (same tier).** Among the surviving opinions, if any is a visible
  border, a line is drawn; only if *all* surviving opinions are "off" is there no line.
- **R5 — Heavier wins (same tier).** Among visible survivors, the greatest `Weight` wins;
  it carries its own colour/dash/transparency.
- **R6 — Deterministic tie-break.** Equal weight and tier → take the **negative side**
  (left/top owner). This removes the current dependence on list order.

This reproduces both tested decks:

| Case (test) | neg opinion | pos opinion | Result | Rule |
|---|---|---|---|---|
| Merged left vs borderless cell (T1 row 2) | off (plain) | black (merged) | no line | R3 → plain off |
| Merged left vs bordered cell (T1 row 1) | white on (plain) | black (merged) | white | R3 → plain on |
| Merged bottom vs "No border cell" top (T2) | black (merged) | white on (plain) | white | R3 → plain on |
| Merged bottom vs "Different font" top (T2) | black (merged) | black on (plain) | black | R3 → plain on |
| Full-border cell edges (T2) | black on | (boundary/plain) | black | R2/R5 |

Note both decks are now satisfied **without** the suppressor rule — R3 alone handles them.

### 2.4 Plain "off" vs plain "on" — RESOLVED by testing

Tested directly: cell A borderless, adjacent cell B fully bordered. Observed behaviour in
PowerPoint is **"the latest change wins"** — the shared line reflects whichever cell was edited
last. The practical implication is that PowerPoint keeps the two plain cells' shared-edge borders
**in sync** on edit, so a stable *disagreement* between two plain cells does not occur; both cells
store the last-applied value. The disagreements we actually hit came only from **merges**, which
leave a stale value on the merged side.

Decisions (implemented):
- **Deleted the suppressor rule.** It assumed "off wins" and its failure mode is deleting a real
  border — the opposite of what we want.
- For any residual plain-vs-plain disagreement (shouldn't happen, but be safe), **visible wins**
  (R4). Never erase a border a cell actually defines.
- The real work is done by **R3 (plain beats merged)**, which both tested decks confirm.

A later test refined the picture further: setting a **merged** cell's shared border also behaves
"latest change wins", i.e. PowerPoint syncs it onto the plain neighbour too. So when a merged
border was *explicitly edited* the two sides agree, and R3 (which returns the plain side) yields
that agreed value — harmless. R3 only changes anything for **un-edited merge artifacts**, which is
exactly its purpose.

### 2.5 Caveats (know these before extending)

- **R3 assumes merged-vs-plain edits stay in sync.** If a deck ever surfaces a merged cell whose
  outer border was *intentionally* set different from its neighbour **and** PowerPoint did not sync
  the neighbour, R3 would wrongly drop it. The fix, if it ever happens: discard a merged opinion
  only when it *conflicts* with the plain side, not unconditionally. Confirm any such case from a
  diagnostics dump (`merged R:… , neighbour L:…`) before changing the rule.
- **Two merged cells meeting on one edge** have no plain side, so they fall through to R5/R6.
  Both could be artifacts; rare, currently unhandled specially.
- **R6's tie-break is for determinism only.** It fires only for two same-tier visible borders of
  equal weight but different style — which does not occur for synced plain cells. We take the
  negative (left/top) side; this is *not* a claim about what PowerPoint would render.
- **Coordinates are floats.** Track offsets are accumulated sums; `EdgePlacement` comparisons use
  an epsilon. Read-time rounding (in `TableReader`) keeps merge keys and edges consistent.
- **The merge artifact appears on the merged cell's *interior-facing* edges** (adjacent to other
  cells), where R3 handles it. Its outer/table-boundary edges read correctly, so R2 uses them
  as-is. If a boundary artifact ever shows up, it would need separate handling.

---

## 3. Full rule set by dimension

Legend: **[H]** handled today · **[P]** partial · **[M]** missing.

### 3.1 Geometry & merges
- Table origin from the min cell edge. **[H]**
- Column widths / row heights reconstructed from actual cell rectangles, falling back to declared
  track sizes only if an interior boundary is hidden. **[P]** — auto-grown (multi-line) rows are
  a known risk: if PowerPoint under-reports a cell's `Height`, the row can collapse to its
  declared minimum. Needs confirming from a real multi-line diagnostics dump before changing
  (an earlier position-based attempt was reverted because it didn't fix the observed issue, which
  turned out to be the text writer — see §3.6).
- Merge detection: currently a **geometry-hash of `Shape.Left/Top/Width/Height`**. Works but is
  fragile (rounding; assumes every covered cell reports identical geometry). **[P]** — prefer a
  span-based detection, and keep geometry hashing only as a fallback.
- Rule: round all coordinates to a fixed precision **once** at read time and use those
  everywhere (offsets, edges, merge keys) so comparisons are consistent.

### 3.2 Fills
- Solid fill: visible, RGB (theme resolved to literal RGB), transparency. **[H]**
- Gradient / pattern / picture / texture fills. **[M]** — fall back to a solid approximation
  and flag; full fidelity needs replicating the `FillFormat`.
- Rule: a cell with no fill must produce a no-fill rectangle (already handled), not white.

### 3.3 Borders
- See §2. Replace pipeline with edge-resolution. **[P → H]**
- Also carry `DashStyle` and `Transparency` through resolution (already read; make sure the
  resolved winner keeps them). **[H]**
- Cap/compound line styles. **[M]** — low priority.

### 3.4 Text frame / cell
- Margins (L/R/T/B), vertical anchor, word wrap, `AutoSize = none`. **[H]**
- Text rotation / vertical text / RTL direction. **[M]**
- Multiple text columns in a cell. **[M]** — rare.

### 3.5 Paragraphs
- Alignment, space before/after/within, indent level. **[H]**
- Bullets / numbering (list format, glyph, start-at, colour, size). **[M]** — visible and common.
- Line-spacing rule (multiple vs exact points). **[P]** — `SpaceWithin` is copied but the rule
  type is not distinguished.
- Right-to-left paragraph. **[M]**

### 3.6 Runs (character)
- Text, font name + complex-script + far-east names, size, bold, italic, underline style,
  strike, fill colour, highlight. **[H]**
- Underline colour (separate from text colour). **[M]**
- Subscript / superscript (baseline offset). **[M]**
- All-caps / small-caps. **[M]**
- Character spacing / kerning. **[M]**
- Language id (affects spell-check only; skip). **[skip]**
- Hyperlinks. **[M]** — preserve the link, not just the styling.
- Text effects: glow, shadow, reflection, soft edge. **[M]** — low priority.
- Rule: write run text **verbatim** (concatenate `run.Text` as read) then apply run formatting
  over exact character ranges. A paragraph's final run already carries its break character, so the
  concatenation reproduces the original and the offsets stay aligned. Do **not** synthesise extra
  `\r` separators between paragraphs — doing so double-inserts breaks (displacing later lines) and
  shifts every subsequent offset (dropping runs). This was the multi-line regression seen in the
  3rd test deck.

### 3.7 Grouping & placement
- All generated shapes grouped into one; original table deleted; group origin re-asserted to
  the table's original Left/Top. **[H]**
- Paint order: fills first, then border lines on top. **[H]** — text lives on the fill rects.
- Rule: on any write failure, delete partial output by name prefix so the slide is untouched
  (already done). **[H]**

---

## 4. Gap summary (priority order)

1. ~~Border edge-resolution refactor (§2)~~ — **DONE** (`LayoutEngine`).
2. ~~Settle the plain off-vs-on rule (§2.4)~~ — **DONE** (tested: latest-wins/sync → visible wins,
   suppressor deleted).
3. **Bullets / numbering** in paragraphs. **High** (common, currently silently dropped).
4. **Merge detection hardening** (§3.1). **Medium.**
5. **Underline colour, sub/superscript, caps, char spacing.** **Medium.**
6. **Hyperlinks.** **Medium.**
7. **Gradient/pattern/picture fills.** **Medium** (approximate + flag).
8. **Text rotation / RTL / text effects.** **Low.**

---

## 5. Implementation plan

1. ~~Refactor `LayoutEngine` borders to the physical-edge model~~ — **DONE**: `Calculate` →
   `ResolveEdges` → `Resolve`; anchor map built up front; `EdgePlacement` kept, order-dependent
   de-dup / suppressor removed.
2. ~~Rework the border unit tests~~ — **DONE**: `LayoutEngineBorderTests` has one test per rule
   R1–R6 plus the two real-deck regressions; the old "suppressor" test became the R4
   "visible wins" test.
3. **Extend the model/reader/writer** for the High/Medium character + paragraph gaps
   (bullets, underline colour, sub/superscript, caps, spacing, hyperlinks) — each behind a
   round-trip unit test and a line in the diagnostics dump. *(next)*
4. **Harden merge detection** with span-based logic + geometry fallback.
5. **Fills**: detect non-solid fills, approximate, and record in diagnostics.

---

## 6. Validation strategy

- **Diagnostics-first**: the per-cell / per-run / per-edge dump is the source of truth. Every
  new rule is validated against a real dump before it ships (this is how the border bugs were
  actually pinned down).
- **Golden-image E2E**: one fixture per feature family (styled runs, merges, borderless cells,
  border conflicts, bullets, …), each asserting a near-zero pixel diff. Grow the fixture as
  gaps close.
- **Pure-logic unit tests**: the edge-resolution function is fully testable without Office —
  cover every rule and every tier combination.
- **One rule, one place**: all border precedence lives in `Resolve`; all "what does PowerPoint
  do here" questions get answered by a targeted 2-cell deck, never by assumption.
