# TableToShapes — PowerPoint Table ? Shape Group Converter

Converts a selected PowerPoint table into a pixel-identical group of shapes
(cell rectangles + border lines), deletes the original table, and replaces it
with a single grouped shape.

## Solution layout

| Project | Purpose |
|---|---|
| `TableToShapes.Core` | **Pure, COM-free logic**: table snapshot model (`TableModel`), geometry (`LayoutEngine`), merged-cell span detection, border edge de-duplication. Fully unit-testable without Office. |
| `TableToShapes.Interop` | Thin Interop layer: `TableReader` (table ? model), `ShapeWriter` (layout ? shapes), `TableConverter` (orchestration: read ? layout ? write ? delete ? group). |
| `TableToShapes.AddIn` | VSTO add-in shell: `Ribbon.xml` + `Ribbon.cs` add a "Convert Table to Shapes" button to the Home tab. Requires the VSTO project system in Visual Studio (see below). |
| `TableToShapes.Tests.Unit` | NUnit + FluentAssertions. Covers geometry, merges, edge de-dup, conflict resolution. Runs anywhere. |
| `TableToShapes.Tests.E2E` | Drives a **real PowerPoint** instance: builds styled/merged tables, converts them, and asserts renders are pixel-identical via `slide.Export` + `ImageComparer`. Tagged `[Category("E2E")]`; requires PowerPoint installed. |

## How fidelity is achieved

- **One rectangle per merge-anchor cell** at exact cumulative row/column offsets;
  merged cells detected via shared cell `Shape.Id` and rendered as a single spanning rectangle.
- **Borders drawn as separate line shapes** on top of fills (matching table paint order).
  Adjacent cells both report shared edges — these are de-duplicated with an
  epsilon-tolerant, endpoint-order-independent comparison; heavier weight wins on conflict.
- **Text copied run-by-run** (font, size, bold/italic, underline, colour) and
  paragraph-by-paragraph (alignment, spacing, indent), plus text-frame margins,
  vertical anchor and word wrap. `AutoSize` is forced off so rectangles never resize.
- **Theme colours are resolved to literal RGB** via `ForeColor.RGB` — required for
  pixel fidelity (trade-off: the group won't re-theme).
- After `Group()`, `Left`/`Top` are re-asserted because grouping can nudge coordinates.

## Building & running

```bash
dotnet build            # Core, Interop, both test projects (no Office needed)
dotnet test TableToShapes.Tests.Unit
dotnet test TableToShapes.Tests.E2E     # requires PowerPoint installed, UI session
```

The add-in project (`TableToShapes.AddIn`) is a VSTO customisation:
1. In Visual Studio (with the *Office/SharePoint development* workload), create a
   "PowerPoint VSTO Add-in" project named `TableToShapes.AddIn`.
2. Add the existing `Ribbon.cs` / `Ribbon.xml` (as Ribbon XML) and reference
   `TableToShapes.Core` and `TableToShapes.Interop`.
3. F5 launches PowerPoint with the add-in registered; select a table and click
   **Convert Table to Shapes** on the Home tab.

## AI workflow used (challenge requirement)

1. **Spec extraction** — prompted the model to enumerate every visual property of a
   PowerPoint table that affects rendering (fills, per-edge borders, merges, margins,
   autofit, mixed runs, theme colours). Turned the answer into a fidelity checklist.
2. **Architecture first** — asked for an architecture that isolates fidelity logic
   from COM so it's unit-testable; adopted the model/layout/interop split.
3. **Generate ? review ? own** — each class was AI-generated, then manually reviewed;
   notable manual corrections: `TextRange2.Runs` is an indexer not a method, and
   `TextFrame2` on shapes is the PowerPoint-namespaced type, not `Office.TextFrame2`.
4. **Test-driven fidelity loop** — unit tests generated from the checklist
   (e.g. "shared edges render once", "heavier border wins"); E2E image-diff tests
   close the loop on true pixel fidelity.
5. **Everything AI-produced was verified by a green build and passing test run.**

## Known limitations / next steps

- Gradient/picture cell fills copied as solid `ForeColor` only (extend `FillModel`).
- Diagonal cell borders (`ppBorderDiagonalDown/Up`) not yet rendered.
- Rotated tables: rotation should be captured and applied to the final group.
- E2E tests need a UI-capable agent; excluded from headless CI via `--filter TestCategory!=E2E`.
