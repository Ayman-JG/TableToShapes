using System;
using System.Globalization;
using System.IO;
using System.Text;
using TableToShapes.Core.Layout;
using TableToShapes.Core.Model;

namespace TableToShapes.Interop
{
    /// <summary>
    /// Writes a human-readable snapshot of what the reader saw and what the layout emitted.
    /// Used to diagnose fidelity issues (e.g. a border drawn where the source shows none)
    /// without a debugger attached to PowerPoint. Best-effort: never throws.
    /// </summary>
    public static class ConversionDiagnostics
    {
        public static string DefaultPath =>
            Path.Combine(Path.GetTempPath(), "TableToShapes.Diagnostics.log");

        public static void Dump(TableModel model, LayoutResult layout, string path = null)
        {
            try
            {
                path = path ?? DefaultPath;
                var sb = new StringBuilder();
                sb.AppendLine("=== TableToShapes diagnostics @ " +
                    DateTime.Now.ToString("O", CultureInfo.InvariantCulture) + " ===");
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    "Grid {0}x{1}  Left={2:F2} Top={3:F2}",
                    model.RowCount, model.ColumnCount, model.Left, model.Top));
                sb.AppendLine("ColWidths=[" + Join(model.ColumnWidths) + "]  RowHeights=[" + Join(model.RowHeights) + "]");
                sb.AppendLine();

                for (int r = 0; r < model.RowCount; r++)
                {
                    for (int c = 0; c < model.ColumnCount; c++)
                    {
                        var cell = model.Cells[r, c];
                        sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                            "cell[{0},{1}] merge={2} fill={3} text=\"{4}\"",
                            r, c, cell.MergeId, Fmt(cell.Fill), FirstLine(cell.Text)));
                        sb.AppendLine("    T:" + Fmt(cell.BorderTop) + "  B:" + Fmt(cell.BorderBottom) +
                                      "  L:" + Fmt(cell.BorderLeft) + "  R:" + Fmt(cell.BorderRight));
                        DumpRuns(sb, cell.Text);
                    }
                }

                sb.AppendLine();
                sb.AppendLine("--- emitted edges (" + layout.Edges.Count + ") ---");
                foreach (var e in layout.Edges)
                {
                    sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                        "  ({0:F2},{1:F2})->({2:F2},{3:F2}) w={4:F2} color=0x{5:X6} dash={6}",
                        e.X1, e.Y1, e.X2, e.Y2, e.Weight, e.ColorRgb & 0xFFFFFF, e.DashStyle));
                }

                File.AppendAllText(path, sb.ToString() + Environment.NewLine);
            }
            catch { /* diagnostics must never break conversion */ }
        }

        private static string Fmt(BorderModel b)
        {
            if (b == null) return "null";
            return b.Visible
                ? string.Format(CultureInfo.InvariantCulture, "ON w={0:F2} 0x{1:X6}", b.Weight, b.ColorRgb & 0xFFFFFF)
                : "off";
        }

        private static string Fmt(FillModel f)
        {
            if (f == null) return "null";
            return f.Visible ? string.Format(CultureInfo.InvariantCulture, "0x{0:X6}", f.ColorRgb & 0xFFFFFF) : "none";
        }

        private static void DumpRuns(StringBuilder sb, TextModel t)
        {
            if (t == null || !t.HasText) return;
            for (int p = 0; p < t.Paragraphs.Count; p++)
            {
                var para = t.Paragraphs[p];
                for (int i = 0; i < para.Runs.Count; i++)
                {
                    var run = para.Runs[i];
                    sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                        "      run[p{0}.{1}] \"{2}\" font=\"{3}\" cs=\"{4}\" fe=\"{5}\" sz={6:F1} " +
                        "B={7} I={8} U={9} strike={10} color=0x{11:X6} highlight={12}",
                        p, i, run.Text, run.FontName, run.FontNameComplexScript, run.FontNameFarEast,
                        run.FontSize, run.Bold ? 1 : 0, run.Italic ? 1 : 0, run.UnderlineStyle,
                        run.Strike, run.ColorRgb & 0xFFFFFF,
                        run.HasHighlight ? ("0x" + (run.HighlightColorRgb & 0xFFFFFF).ToString("X6", CultureInfo.InvariantCulture)) : "off"));
                }
            }
        }

        private static string FirstLine(TextModel t)
        {
            if (t == null || !t.HasText || t.Paragraphs.Count == 0) return "";
            var sb = new StringBuilder();
            foreach (var run in t.Paragraphs[0].Runs) sb.Append(run.Text);
            var s = sb.ToString();
            return s.Length > 30 ? s.Substring(0, 30) : s;
        }

        private static string Join(System.Collections.Generic.IReadOnlyList<float> values)
        {
            if (values == null) return "";
            var sb = new StringBuilder();
            for (int i = 0; i < values.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append(values[i].ToString("F2", CultureInfo.InvariantCulture));
            }
            return sb.ToString();
        }
    }
}
