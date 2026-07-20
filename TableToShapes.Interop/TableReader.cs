using System.Collections.Generic;
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
                Left = tableShape.Left,
                Top = tableShape.Top,
                RowHeights = ReadSizes(i => table.Rows[i].Height, rows),
                ColumnWidths = ReadSizes(i => table.Columns[i].Width, cols),
                Cells = new CellModel[rows, cols]
            };

            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    model.Cells[r - 1, c - 1] = ReadCell(table.Cell(r, c));
                }
            }

            return model;
        }

        private static IReadOnlyList<float> ReadSizes(System.Func<int, float> sizeOf, int count)
        {
            var sizes = new float[count];
            for (int i = 1; i <= count; i++) sizes[i - 1] = sizeOf(i);
            return sizes;
        }

        private static CellModel ReadCell(PowerPoint.Cell cell)
        {
            var shape = cell.Shape;
            return new CellModel
            {
                // Cells in the same merge share the underlying shape; its Id is our merge key.
                MergeId = shape.Id,
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
                // ForeColor.RGB resolves theme colours to literal RGB — exactly what we
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
                foreach (Office.TextRange2 run in para.Runs[1, -1])
                {
                    runs.Add(new RunModel
                    {
                        Text = run.Text,
                        FontName = run.Font.Name,
                        FontSize = run.Font.Size,
                        Bold = run.Font.Bold == Office.MsoTriState.msoTrue,
                        Italic = run.Font.Italic == Office.MsoTriState.msoTrue,
                        UnderlineStyle = (int)run.Font.UnderlineStyle,
                        ColorRgb = run.Font.Fill.ForeColor.RGB
                    });
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
    }
}
