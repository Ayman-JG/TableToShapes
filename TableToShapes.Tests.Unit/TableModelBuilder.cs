using System.Collections.Generic;
using TableToShapes.Core.Model;

namespace TableToShapes.Tests.Unit
{
    /// <summary>Builds TableModel instances for tests without COM.</summary>
    internal static class TableModelBuilder
    {
        public static TableModel Grid(int rows, int cols, float rowHeight = 20f, float colWidth = 100f,
                                      float left = 0f, float top = 0f)
        {
            var cells = new CellModel[rows, cols];
            int id = 1;
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    cells[r, c] = DefaultCell(id++);

            return new TableModel
            {
                Left = left,
                Top = top,
                RowHeights = Repeat(rowHeight, rows),
                ColumnWidths = Repeat(colWidth, cols),
                Cells = cells
            };
        }

        public static CellModel DefaultCell(int mergeId)
        {
            return new CellModel
            {
                MergeId = mergeId,
                Fill = new FillModel { Visible = true, ColorRgb = 0xFFFFFF },
                Text = new TextModel(),
                BorderTop = VisibleBorder(),
                BorderBottom = VisibleBorder(),
                BorderLeft = VisibleBorder(),
                BorderRight = VisibleBorder()
            };
        }

        public static BorderModel VisibleBorder(float weight = 1f, int color = 0)
            => new BorderModel { Visible = true, Weight = weight, ColorRgb = color };

        private static IReadOnlyList<float> Repeat(float value, int count)
        {
            var list = new float[count];
            for (int i = 0; i < count; i++) list[i] = value;
            return list;
        }
    }
}
