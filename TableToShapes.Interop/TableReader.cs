using System.Collections.Generic;
using System.Linq;
using TableToShapes.Core.Model;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace TableToShapes.Interop
{
    /// <summary>
    /// Reads a live PowerPoint table into a COM-free <see cref="TableModel"/>.
    /// All fidelity decisions happen downstream on the model.
    /// </summary>
    public sealed class TableReader
    {
        public TableModel Read(PowerPoint.Shape tableShape)
        {
            var table = tableShape.Table;
            int rows = table.Rows.Count;
            int cols = table.Columns.Count;

            var model = new TableModel
            {
                Cells = new CellModel[rows, cols]
            };

            // Cells in the same merge share one underlying shape, so all of them report
            // identical shape geometry. Shape.Id is not implemented for table cell shapes
            // (throws E_NOTIMPL), so we derive a synthetic merge id from the geometry.
            var mergeIds = new Dictionary<string, int>();

            // Row/column declared sizes are only *minimums*: PowerPoint grows rows to fit
            // text, so the rendered table is taller than the sum of Row.Height. We instead
            // reconstruct boundaries from the actual cell rectangles. Collecting BOTH edges
            // of every cell means interior boundaries hidden by a merge on one row still
            // appear from the non-merged cells on that row.
            var xEdges = new SortedSet<float>();
            var yEdges = new SortedSet<float>();

            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    var cell = table.Cell(r, c);
                    var shape = cell.Shape;

                    xEdges.Add(Round(shape.Left));
                    xEdges.Add(Round(shape.Left + shape.Width));
                    yEdges.Add(Round(shape.Top));
                    yEdges.Add(Round(shape.Top + shape.Height));

                    model.Cells[r - 1, c - 1] = ReadCell(cell, shape, mergeIds);
                }
            }

            model.Left = xEdges.Min;
            model.Top = yEdges.Min;
            model.ColumnWidths = EdgesToSizes(xEdges, cols, () => ReadSizes(i => table.Columns[i].Width, cols));
            model.RowHeights = EdgesToSizes(yEdges, rows, () => ReadSizes(i => table.Rows[i].Height, rows));

            return model;
        }

        private static float Round(float value) => (float)System.Math.Round(value, 2);

        // Distinct cell edges form the track boundaries; consecutive differences are the
        // per-track sizes. Falls back to declared sizes if an interior boundary is missing
        // (only when every cell spanning it is merged across it - a rare case).
        private static IReadOnlyList<float> EdgesToSizes(
            SortedSet<float> edges, int expected, System.Func<IReadOnlyList<float>> fallback)
        {
            if (edges.Count != expected + 1) return fallback();

            var ordered = edges.ToList();
            var sizes = new float[expected];
            for (int i = 0; i < expected; i++) sizes[i] = ordered[i + 1] - ordered[i];
            return sizes;
        }

        private static IReadOnlyList<float> ReadSizes(System.Func<int, float> sizeOf, int count)
        {
            var sizes = new float[count];
            for (int i = 1; i <= count; i++) sizes[i - 1] = sizeOf(i);
            return sizes;
        }

        private static CellModel ReadCell(PowerPoint.Cell cell, PowerPoint.Shape shape, Dictionary<string, int> mergeIds)
        {
            string geometryKey = $"{shape.Left:F2}|{shape.Top:F2}|{shape.Width:F2}|{shape.Height:F2}";
            if (!mergeIds.TryGetValue(geometryKey, out int mergeId))
            {
                mergeId = mergeIds.Count + 1;
                mergeIds[geometryKey] = mergeId;
            }

            return new CellModel
            {
                MergeId = mergeId,
                Fill = ReadFill(shape.Fill),
                Text = ReadText(shape.TextFrame2),
                BorderTop = ReadBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderTop]),
                BorderBottom = ReadBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderBottom]),
                BorderLeft = ReadBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderLeft]),
                BorderRight = ReadBorder(cell.Borders[PowerPoint.PpBorderType.ppBorderRight])
            };
        }

        private static FillModel ReadFill(PowerPoint.FillFormat fill)
        {
            bool visible = fill.Visible == Office.MsoTriState.msoTrue;
            return new FillModel
            {
                Visible = visible,
                // ForeColor.RGB resolves theme colours to literal RGB � exactly what we
                // want for pixel fidelity (at the cost of losing theme re-colouring).
                ColorRgb = visible ? fill.ForeColor.RGB : 0,
                Transparency = visible ? fill.Transparency : 0f
            };
        }

        private static BorderModel ReadBorder(PowerPoint.LineFormat border)
        {
            bool visible = border.Visible == Office.MsoTriState.msoTrue;
            return new BorderModel
            {
                Visible = visible,
                Weight = visible ? border.Weight : 0f,
                ColorRgb = visible ? border.ForeColor.RGB : 0,
                DashStyle = visible ? (int)border.DashStyle : 0,
                Transparency = visible ? border.Transparency : 0f
            };
        }

        private static TextModel ReadText(PowerPoint.TextFrame2 frame)
        {
            var text = new TextModel
            {
                HasText = frame.HasText == Office.MsoTriState.msoTrue,
                MarginLeft = frame.MarginLeft,
                MarginRight = frame.MarginRight,
                MarginTop = frame.MarginTop,
                MarginBottom = frame.MarginBottom,
                VerticalAnchor = (int)frame.VerticalAnchor,
                WordWrap = frame.WordWrap == Office.MsoTriState.msoTrue
            };

            if (!text.HasText) return text;

            var paragraphs = new List<ParagraphModel>();
            foreach (Office.TextRange2 para in frame.TextRange.Paragraphs)
            {
                var runs = new List<RunModel>();

                // BUG FIX: the old code iterated `para.Runs[1, -1]`. Per the Office docs,
                // when Start is given but Length is omitted (-1 is the omitted sentinel),
                // Runs returns exactly ONE run - so every run after the first was silently
                // dropped, taking its text and formatting with it. We instead count the runs
                // and read each one individually.
                Office.TextRange2 allRuns = para.Runs[-1, -1]; // both omitted => the whole paragraph's runs
                int runCount = allRuns.Count;
                for (int ri = 1; ri <= runCount; ri++)
                {
                    Office.TextRange2 run = para.Runs[ri, 1];
                    runs.Add(ReadRun(run));
                }

                paragraphs.Add(new ParagraphModel
                {
                    Alignment = (int)para.ParagraphFormat.Alignment,
                    SpaceBefore = para.ParagraphFormat.SpaceBefore,
                    SpaceAfter = para.ParagraphFormat.SpaceAfter,
                    SpaceWithin = para.ParagraphFormat.SpaceWithin,
                    IndentLevel = para.ParagraphFormat.IndentLevel,
                    Runs = runs
                });
            }

            text.Paragraphs = paragraphs;
            return text;
        }

        private static RunModel ReadRun(Office.TextRange2 run)
        {
            var font = run.Font;
            var model = new RunModel
            {
                Text = run.Text,
                FontName = font.Name,
                // Copy the script-specific faces too: PowerPoint renders each character with
                // the face that matches its script, so a run can display a substituted font
                // if only Name is reproduced. Guarded because some builds throw on these.
                FontNameComplexScript = TryGet(() => font.NameComplexScript),
                FontNameFarEast = TryGet(() => font.NameFarEast),
                FontSize = font.Size,
                Bold = font.Bold == Office.MsoTriState.msoTrue,
                Italic = font.Italic == Office.MsoTriState.msoTrue,
                UnderlineStyle = (int)font.UnderlineStyle,
                Strike = TryGetInt(() => (int)font.Strike),
                ColorRgb = font.Fill.ForeColor.RGB
            };

            // Highlight (marker) colour. Unhighlighted runs report msoColorTypeMixed;
            // a real highlight reports a concrete RGB. Reading is wrapped because older
            // Office builds throw E_NOTIMPL on Font2.Highlight.
            try
            {
                var highlight = font.Highlight;
                if (highlight != null && highlight.Type == Office.MsoColorType.msoColorTypeRGB)
                {
                    model.HasHighlight = true;
                    model.HighlightColorRgb = highlight.RGB;
                }
            }
            catch { /* highlight unsupported on this build */ }

            return model;
        }

        private static string TryGet(System.Func<string> get)
        {
            try { return get(); }
            catch { return null; }
        }

        private static int TryGetInt(System.Func<int> get)
        {
            try { return get(); }
            catch { return 0; }
        }
    }
}
