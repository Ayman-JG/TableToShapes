# TableToShapes — PowerPoint Table ? Shape Group Converter

Converts a selected PowerPoint table into a pixel-identical group of shapes
(cell rectangles + border lines), deletes the original table, and replaces it
with a single grouped shape.

## Solution layout

| Project | Purpose |
|---|---|
| `TableToShapes.Core` | **Pure, COM-free logic**: table snapshot model (`TableModel`), geometry (`LayoutEngine`), merged-cell span detection, border edge de-duplication. Fully unit-testable without Office. |
| `TableToShapes.Interop` | Thin Interop layer: `TableReader` (table ? model), `ShapeWriter` (layout ? shapes), `TableConverter` (orchestration: read ? layout ? write ? delete ? group). |
| `TableToShapes.AddIn` | COM add-in (`Connect.cs` implements `IDTExtensibility2` + `IRibbonExtensibility`): adds a "Convert Table" button to the Home tab. Registered per-user via `install-addin.ps1` — no VSTO runtime or special VS workload required. |
| `TableToShapes.Tests.Unit` | NUnit + FluentAssertions. Covers geometry, merges, edge de-dup, conflict resolution. Runs anywhere. |
| `TableToShapes.Tests.E2E` | Drives a **real PowerPoint** instance: builds styled/merged tables, converts them, and asserts renders are pixel-identical via `slide.Export` + `ImageComparer`. Tagged `[Category("E2E")]`; requires PowerPoint installed. |

## How fidelity is achieved

- **One rectangle per merge-anchor cell** at exact cumulative row/column offsets;
  merged cells share one underlying shape, so they are detected by identical cell-shape
  geometry (`Left/Top/Width/Height`; `Shape.Id` is `E_NOTIMPL` for table cells) and
  rendered as a single spanning rectangle.
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

### Installing the add-in

```powershell
dotnet build
cd TableToShapes.AddIn
.\install-addin.ps1     # RegAsm + HKCU PowerPoint\Addins entry (current user, no admin)
```

Restart PowerPoint, select a table, and click **Convert Table** in the
**Table to Shapes** group on the Home tab. Remove with `.\uninstall-addin.ps1`.
Note: 32-bit Office needs the 32-bit RegAsm — change `Framework64` to `Framework`
in the scripts.

## AI workflow used (challenge requirement)

1. **Spec extraction** — prompted the model to enumerate every visual property of a
   PowerPoint table that affects rendering (fills, per-edge borders, merges, margins,
   autofit, mixed runs, theme colours). Turned the answer into a fidelity checklist.
2. **Architecture first** — asked for an architecture that isolates fidelity logic
   from COM so it's unit-testable; adopted the model/layout/interop split.
3. **Generate ? review ? own** — each class was AI-generated, then manually reviewed;
   notable manual corrections: `TextRange2.Runs` is an indexer not a method,
   `TextFrame2` on shapes is the PowerPoint-namespaced type, not `Office.TextFrame2`,
   and `Shape.Id` throws `E_NOTIMPL` for table-cell shapes (found via E2E runs against
   real PowerPoint; fixed with geometry-based merge detection).
4. **Test-driven fidelity loop** — unit tests generated from the checklist
   (e.g. "shared edges render once", "heavier border wins"); E2E image-diff tests
   close the loop on true pixel fidelity.
5. **Everything AI-produced was verified by a green build and passing test run.**

## Known limitations / next steps

- Gradient/picture cell fills copied as solid `ForeColor` only (extend `FillModel`).
- Diagonal cell borders (`ppBorderDiagonalDown/Up`) not yet rendered.
- Rotated tables: rotation should be captured and applied to the final group.
- E2E tests need a UI-capable agent; excluded from headless CI via `--filter TestCategory!=E2E`.
