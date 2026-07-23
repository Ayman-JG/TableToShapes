using System.Collections.Generic;
using System.Text;
using TableToShapes.Core.Layout;
using TableToShapes.Core.Model;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace TableToShapes.Interop
{
    /// <summary>
    /// Materialises a <see cref="LayoutResult"/> as real shapes on a slide.
    /// </summary>
    public sealed class ShapeWriter
    {
        /// <summary>Prefix applied to every shape we create, so partial output can be cleaned up on failure.</summary>
        public const string CreatedShapePrefix = "T2S_";

        /// <summary>Creates all cell rectangles and border lines; returns their shape names.</summary>
        public List<string> Write(PowerPoint.Slide slide, TableModel table, LayoutResult layout)
        {
            var names = new List<string>();

            // Paint order matters: fills first, borders on top (mirrors table rendering).
            foreach (var cell in layout.Cells)
                names.Add(WriteCell(slide, table.Cells[cell.Row, cell.Column], cell).Name);

            foreach (var edge in layout.Edges)
                names.Add(WriteEdge(slide, edge).Name);

            return names;
        }

        private static PowerPoint.Shape WriteCell(PowerPoint.Slide slide, CellModel cell, CellPlacement p)
        {
            var rect = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle, p.Left, p.Top, p.Width, p.Height);
            // Name immediately, before any further property set, so partial-failure cleanup by
            // name prefix can always find this shape.
            rect.Name = CreatedShapePrefix + rect.Name;

            rect.Line.Visible = Office.MsoTriState.msoFalse; // borders are separate shapes

            if (cell.Fill.Visible)
            {
                rect.Fill.Visible = Office.MsoTriState.msoTrue;
                rect.Fill.Solid();
                rect.Fill.ForeColor.RGB = cell.Fill.ColorRgb;
                rect.Fill.Transparency = Clamp01(cell.Fill.Transparency);
            }
            else
            {
                rect.Fill.Visible = Office.MsoTriState.msoFalse;
            }

            WriteText(rect.TextFrame2, cell.Text);
            return rect;
        }

        private static PowerPoint.Shape WriteEdge(PowerPoint.Slide slide, EdgePlacement edge)
        {
            var line = slide.Shapes.AddLine(edge.X1, edge.Y1, edge.X2, edge.Y2);
            line.Name = CreatedShapePrefix + line.Name; // name before other sets (see WriteCell)
            if (edge.Weight > 0f) line.Line.Weight = edge.Weight;
            line.Line.ForeColor.RGB = edge.ColorRgb;
            // Valid dash styles start at 1 (solid); skip 0/negative ("Mixed") sentinels.
            if (edge.DashStyle > 0) line.Line.DashStyle = (Office.MsoLineDashStyle)edge.DashStyle;
            line.Line.Transparency = Clamp01(edge.Transparency);
            return line;
        }

        // PowerPoint requires Transparency in [0, 1]; read-back values can drift slightly.
        private static float Clamp01(float value)
        {
            if (value < 0f) return 0f;
            if (value > 1f) return 1f;
            return value;
        }

        private static void WriteText(PowerPoint.TextFrame2 frame, TextModel text)
        {
            frame.MarginLeft = text.MarginLeft;
            frame.MarginRight = text.MarginRight;
            frame.MarginTop = text.MarginTop;
            frame.MarginBottom = text.MarginBottom;
            // Guard against a "Mixed" sentinel (negative) that would throw when cast back.
            if (text.VerticalAnchor > 0)
                frame.VerticalAnchor = (Office.MsoVerticalAnchor)text.VerticalAnchor;
            frame.WordWrap = text.WordWrap ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            // Never let the rectangle resize itself, or text will shift.
            frame.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;

            if (!text.HasText) return;

            // Insert the full plain text, then re-apply run formatting over exact character
            // ranges. This avoids the clipboard (unreliable in automation).
            //
            // The run text is written back VERBATIM. Each run's text as read already contains any
            // paragraph/line-break characters (a paragraph's final run carries its terminator), so
            // concatenating the runs reproduces the original exactly - and keeps the character
            // offsets below aligned. Do NOT synthesise extra separators between paragraphs: that
            // double-inserts breaks (displacing lines) and shifts every subsequent offset.
            var builder = new StringBuilder();
            foreach (var para in text.Paragraphs)
                foreach (var run in para.Runs)
                    if (run.Text != null) builder.Append(run.Text);

            // Hoist TextRange: it is used for the text, every character range and the paragraph
            // formatting, so fetching it once avoids re-creating the COM object each time.
            var textRange = frame.TextRange;
            textRange.Text = builder.ToString();

            int charIndex = 1;
            int paraIndex = 1;
            foreach (var para in text.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    int length = run.Text?.Length ?? 0;
                    // Characters[index, 0] throws "value out of range"; skip empty runs.
                    if (length == 0) continue;

                    var range = textRange.Characters[charIndex, length];
                    range.Font.Name = run.FontName;
                    if (!string.IsNullOrEmpty(run.FontNameComplexScript))
                        range.Font.NameComplexScript = run.FontNameComplexScript;
                    if (!string.IsNullOrEmpty(run.FontNameFarEast))
                        range.Font.NameFarEast = run.FontNameFarEast;
                    range.Font.Size = run.FontSize;
                    range.Font.Bold = run.Bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    range.Font.Italic = run.Italic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    // 0 = none is valid; a negative "Mixed" sentinel is skipped.
                    if (run.UnderlineStyle >= 0)
                        range.Font.UnderlineStyle = (Office.MsoTextUnderlineType)run.UnderlineStyle;
                    if (run.Strike >= 0)
                        range.Font.Strike = (Office.MsoTextStrike)run.Strike;
                    range.Font.Fill.ForeColor.RGB = run.ColorRgb;
                    // Setting the Highlight colour turns marker highlighting on for the range.
                    if (run.HasHighlight)
                        range.Font.Highlight.RGB = run.HighlightColorRgb;
                    charIndex += length;
                }

                // Guard against paragraph-count drift between model and live text.
                if (paraIndex <= textRange.Paragraphs.Count)
                {
                    var paraRange = textRange.Paragraphs[paraIndex, 1];
                    paraRange.ParagraphFormat.Alignment = (Office.MsoParagraphAlignment)para.Alignment;
                    paraRange.ParagraphFormat.SpaceBefore = para.SpaceBefore;
                    paraRange.ParagraphFormat.SpaceAfter = para.SpaceAfter;
                    paraRange.ParagraphFormat.SpaceWithin = para.SpaceWithin;
                    paraRange.ParagraphFormat.IndentLevel = para.IndentLevel;
                }
                paraIndex++;
            }
        }
    }
}
