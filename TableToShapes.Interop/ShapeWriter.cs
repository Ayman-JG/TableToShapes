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
        /// <summary>Creates all cell rectangles and border lines; returns their shape names.</summary>
        public List<string> Write(PowerPoint.Slide slide, TableModel table, LayoutResult layout)
        {
            var names = new List<string>();

            // Paint order matters: fills first, borders on top (mirrors table rendering).
            foreach (var cell in layout.Cells)
            {
                names.Add(WriteCell(slide, table.Cells[cell.Row, cell.Column], cell).Name);
            }

            foreach (var edge in layout.Edges)
            {
                names.Add(WriteEdge(slide, edge).Name);
            }

            return names;
        }

        private static PowerPoint.Shape WriteCell(PowerPoint.Slide slide, CellModel cell, CellPlacement p)
        {
            var rect = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle, p.Left, p.Top, p.Width, p.Height);

            rect.Line.Visible = Office.MsoTriState.msoFalse; // borders are separate shapes

            if (cell.Fill.Visible)
            {
                rect.Fill.Visible = Office.MsoTriState.msoTrue;
                rect.Fill.Solid();
                rect.Fill.ForeColor.RGB = cell.Fill.ColorRgb;
                rect.Fill.Transparency = cell.Fill.Transparency;
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
            line.Line.Weight = edge.Weight;
            line.Line.ForeColor.RGB = edge.ColorRgb;
            line.Line.DashStyle = (Office.MsoLineDashStyle)edge.DashStyle;
            line.Line.Transparency = edge.Transparency;
            return line;
        }

        private static void WriteText(PowerPoint.TextFrame2 frame, TextModel text)
        {
            frame.MarginLeft = text.MarginLeft;
            frame.MarginRight = text.MarginRight;
            frame.MarginTop = text.MarginTop;
            frame.MarginBottom = text.MarginBottom;
            frame.VerticalAnchor = (Office.MsoVerticalAnchor)text.VerticalAnchor;
            frame.WordWrap = text.WordWrap ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            // Never let the rectangle resize itself, or text will shift.
            frame.AutoSize = Office.MsoAutoSize.msoAutoSizeNone;

            if (!text.HasText) return;

            // Insert the full plain text, then re-apply run formatting over exact character
            // ranges. This avoids the clipboard (unreliable in automation).
            var builder = new StringBuilder();
            foreach (var para in text.Paragraphs)
                foreach (var run in para.Runs)
                    builder.Append(run.Text);

            frame.TextRange.Text = builder.ToString();

            int charIndex = 1;
            int paraIndex = 1;
            foreach (var para in text.Paragraphs)
            {
                foreach (var run in para.Runs)
                {
                    int length = run.Text.Length;
                    var range = frame.TextRange.Characters[charIndex, length];
                    range.Font.Name = run.FontName;
                    range.Font.Size = run.FontSize;
                    range.Font.Bold = run.Bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    range.Font.Italic = run.Italic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    range.Font.UnderlineStyle = (Office.MsoTextUnderlineType)run.UnderlineStyle;
                    range.Font.Fill.ForeColor.RGB = run.ColorRgb;
                    charIndex += length;
                }

                var paraRange = frame.TextRange.Paragraphs[paraIndex, 1];
                paraRange.ParagraphFormat.Alignment = (Office.MsoParagraphAlignment)para.Alignment;
                paraRange.ParagraphFormat.SpaceBefore = para.SpaceBefore;
                paraRange.ParagraphFormat.SpaceAfter = para.SpaceAfter;
                paraRange.ParagraphFormat.SpaceWithin = para.SpaceWithin;
                paraRange.ParagraphFormat.IndentLevel = para.IndentLevel;
                paraIndex++;
            }
        }
    }
}
