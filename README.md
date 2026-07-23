# TableToShapes - PowerPoint Table to Shape-Group Converter

A PowerPoint add-in that converts a selected table into a visually identical group
of shapes (one rectangle per cell plus border lines), removes the original table, and
replaces it with a single grouped shape.

For a plain-language summary of what is and isn't supported (suitable for non-engineers),
see [`docs/CAPABILITIES.md`](docs/CAPABILITIES.md). For the engineering detail of the border
rules, see [`docs/FIDELITY_RULES.md`](docs/FIDELITY_RULES.md).

## Scope and status

This was built as an **MVP within an ~8-hour timebox**. The aim was a faithful, well-tested
conversion of the table features that come up most often, with the architecture and the known
gaps clearly documented rather than hidden. It handles the common cases end to end; some edge
cases and fine-tuning (listed in [`docs/CAPABILITIES.md`](docs/CAPABILITIES.md) and *Known
limitations* below) would need further work to fully meet production and customer expectations.

## Solution layout

| Project | Purpose |
|---|---|
| `TableToShapes.Core` | **Pure, COM-free logic**: the table snapshot model (`TableModel`), geometry and border resolution (`LayoutEngine`), merged-cell span detection, and the resolved `CellPlacement` / `EdgePlacement` outputs. Fully unit-testable without Office. |
| `TableToShapes.Interop` | Thin Interop layer: `TableReader` (live table -> model), `ShapeWriter` (layout -> shapes), `TableConverter` (orchestration: read -> layout -> write -> delete -> group), and an opt-in `ConversionDiagnostics` log. |
| `TableToShapes.AddIn` | COM add-in (`Connect.cs` implements `IDTExtensibility2` + `IRibbonExtensibility`): adds a "Convert Table" button to the Home tab. Registered per-user via `install-addin.ps1` - no VSTO runtime or special Visual Studio workload required. |
| `TableToShapes.Tests.Unit` | NUnit + FluentAssertions. Covers geometry, merges, and the border-resolution rules. Runs anywhere (no Office). |
| `TableToShapes.Tests.E2E` | Drives a **real PowerPoint** instance: builds styled/merged tables, converts them, and asserts the renders are pixel-identical via `slide.Export` + `ImageComparer`. Tagged `[Category("E2E")]`; requires PowerPoint installed. |

## How fidelity is achieved

- **One rectangle per merge-anchor cell** at exact cumulative row/column offsets. Merged
  cells share one underlying shape, so they are detected by identical cell-shape geometry
  (`Left/Top/Width/Height`; `Shape.Id` throws `E_NOTIMPL` for table cells) and rendered as a
  single spanning rectangle.
- **Borders are resolved per physical grid-line edge, not per cell.** A line between two cells
  is one edge that both cells describe and can disagree about (typically after a merge). The
  engine walks every grid-line segment once, gathers the two adjacent cells' border settings,
  and resolves them with a fixed precedence (rules R1-R6; see `docs/FIDELITY_RULES.md`). The
  resolved segments are drawn as separate line shapes on top of the fills, matching PowerPoint's
  paint order.
- **Text is copied run-by-run** - font name (Latin, complex-script and East-Asian faces), size,
  bold, italic, underline, strikethrough, colour and highlight - and paragraph-by-paragraph
  (alignment, spacing, indent), plus text-frame margins, vertical anchor and word wrap. Run text
  is written back verbatim so line and paragraph breaks are preserved, and `AutoSize` is forced
  off so the rectangles never resize.
- **Theme colours are resolved to literal RGB** via `ForeColor.RGB` - required for pixel fidelity
  (trade-off: the group will not re-theme).
- After `Group()`, `Left` / `Top` are re-asserted because grouping can nudge coordinates. If
  writing fails partway, the partial shapes are removed by name prefix so the slide is left as it
  was found.

## Behaviour and caveats

- **The result is a shape group, not a table.** After conversion the content is a group of
  rectangles and lines; it can no longer be edited as a table (add/remove rows, etc.). The
  operation is one-directional.
- **Undo.** The conversion performs several COM operations (create shapes, delete the table,
  group) and is not wrapped in a single undo unit, so reverting in PowerPoint may take more than
  one Ctrl+Z.
- **Only fills, text and borders are carried over.** Anything else living inside a cell -
  embedded pictures, charts or other shapes - is not reproduced.
- **Theme colours are captured as literal RGB**, so the converted group will not follow later
  theme or colour-scheme changes.
- **Autofit "shrink text on overflow" is not reproduced.** `AutoSize` is forced off so rectangles
  never resize; a cell that had shrunk its text to fit will render at the original run sizes.
- **Fonts must be installed** on the machine running the conversion; a missing font is substituted
  by PowerPoint as normal.
- **Selection.** The first table in the selection is converted; selecting a non-table shows a
  prompt and does nothing.
- **Merge detection** relies on all covered cells reporting identical shape geometry; it is
  sensitive to rounding and could misgroup in unusual layouts (see `docs/FIDELITY_RULES.md`).
- **Auto-grown (multi-line) rows** depend on Interop reporting accurate cell heights; if a row is
  under-reported it can render at its declared minimum. Tracked in `docs/FIDELITY_RULES.md`.

## Building and running

```bash
dotnet build                             # Core, Interop, both test projects (no Office needed)
dotnet test TableToShapes.Tests.Unit     # pure-logic tests, run anywhere
dotnet test TableToShapes.Tests.E2E      # requires PowerPoint installed, in a UI session
```

### Installing the add-in

```powershell
dotnet build
cd TableToShapes.AddIn
.\install-addin.ps1      # writes the COM + PowerPoint add-in registry entries (current user)
```

`install-addin.ps1` registers the class **directly in the per-user COM hive**
(`HKCU\Software\Classes`, loaded via the `mscoree.dll` .NET shim) and adds the
`PowerPoint\Addins` entry - it does **not** use RegAsm, so no administrator rights are needed.
It reads the `bin\Debug\net48` build, so build first (a Release build needs the path in the
script adjusted).

Restart PowerPoint, select a table, and click **Convert Table** in the **Table to Shapes**
group on the Home tab. Remove with `.\uninstall-addin.ps1`.

Bitness note: the script writes to `HKCU\Software\Classes`; it has been used with 64-bit Office.
32-bit Office on 64-bit Windows may resolve the class under `Wow6432Node`, which the script does
not currently write - adjust the registry path if you run 32-bit Office.

### Troubleshooting

Set the `TABLETOSHAPES_DIAGNOSTICS` environment variable before launching PowerPoint to have
each conversion append a human-readable snapshot (parsed cells, borders, runs and resolved
edges) to `%TEMP%\TableToShapes.Diagnostics.log`. It is off by default.

## AI workflow used

1. **Spec extraction** - enumerated every visual property of a PowerPoint table that affects
   rendering (fills, per-edge borders, merges, margins, autofit, mixed runs, highlight,
   strikethrough, theme colours) and turned it into a fidelity checklist.
2. **Architecture first** - isolated all fidelity logic from COM into `Core` so it is
   unit-testable; adopted the model / layout / interop split.
3. **Generate, review, own** - each class was AI-assisted then manually reviewed and corrected
   against the real Office API. Notable corrections found this way: `TextRange2.Runs` is an
   indexer whose `Length` defaults to a single run (so all runs must be enumerated explicitly);
   `Shape.Id` throws `E_NOTIMPL` for table-cell shapes (merges detected by geometry instead); and
   run text already carries its own break characters (so it must be written verbatim, not
   re-joined).
4. **Test-driven fidelity loop** - unit tests cover the geometry and every border-resolution
   rule; E2E image-diff tests close the loop on true pixel fidelity. A rules document
   (`docs/FIDELITY_RULES.md`) records the border model and open gaps.

## Known limitations / next steps

- Gradient / pattern / picture cell fills are approximated as a solid `ForeColor`.
- Paragraph bullets and numbering are not yet reproduced.
- Some character attributes are not yet copied: underline colour, sub/superscript, caps, and
  character spacing; hyperlinks are not preserved.
- Diagonal cell borders (`ppBorderDiagonalDown/Up`) and rotated tables are not handled.
- E2E tests require a UI-capable session with PowerPoint; exclude them from headless CI with
  `--filter TestCategory!=E2E`.

`docs/FIDELITY_RULES.md` has the full property-coverage matrix (handled / partial / not yet) and
the border-resolution caveats.
```
