using System.Collections.Generic;
using System.Linq;
using TableToShapes.Core.Model;

namespace TableToShapes.Core.Layout
{
    /// <summary>
    /// Pure geometry: turns a <see cref="TableModel"/> into cell rectangles and a
    /// de-duplicated set of border segments. No Interop dependencies.
    /// </summary>
    public sealed class LayoutEngine
    {
        public LayoutResult Calculate(TableModel table)
        {
            var rowOffsets = Accumulate(table.RowHeights, table.Top);
            var colOffsets = Accumulate(table.ColumnWidths, table.Left);

            var cells = new List<CellPlacement>();
            var edges = new List<EdgePlacement>();

            for (int r = 0; r < table.RowCount; r++)
            {
                for (int c = 0; c < table.ColumnCount; c++)
                {
                    if (!IsMergeAnchor(table, r, c)) continue;

                    int rowSpan = GetRowSpan(table, r, c);
                    int colSpan = GetColumnSpan(table, r, c);

                    var placement = new CellPlacement
                    {
                        Row = r,
                        Column = c,
                        RowSpan = rowSpan,
                        ColumnSpan = colSpan,
                        Left = colOffsets[c],
                        Top = rowOffsets[r],
                        Width = colOffsets[c + colSpan] - colOffsets[c],
                        Height = rowOffsets[r + rowSpan] - rowOffsets[r]
                    };
                    cells.Add(placement);

                    AddEdges(edges, table.Cells[r, c], placement);
                }
            }

            return new LayoutResult(cells, edges);
        }

        private static float[] Accumulate(IReadOnlyList<float> sizes, float origin)
        {
            var offsets = new float[sizes.Count + 1];
            offsets[0] = origin;
            for (int i = 0; i < sizes.Count; i++)
                offsets[i + 1] = offsets[i] + sizes[i];
            return offsets;
        }

        private static bool IsMergeAnchor(TableModel table, int r, int c)
        {
            int id = table.Cells[r, c].MergeId;
            if (r > 0 && table.Cells[r - 1, c].MergeId == id) return false;
            if (c > 0 && table.Cells[r, c - 1].MergeId == id) return false;
            return true;
        }

        private static int GetRowSpan(TableModel table, int r, int c)
        {
            int id = table.Cells[r, c].MergeId, span = 1;
            while (r + span < table.RowCount && table.Cells[r + span, c].MergeId == id) span++;
            return span;
        }

        private static int GetColumnSpan(TableModel table, int r, int c)
        {
            int id = table.Cells[r, c].MergeId, span = 1;
            while (c + span < table.ColumnCount && table.Cells[r, c + span].MergeId == id) span++;
            return span;
        }

        private static void AddEdges(List<EdgePlacement> edges, CellModel cell, CellPlacement p)
        {
            AddEdge(edges, cell.BorderTop, p.Left, p.Top, p.Left + p.Width, p.Top);
            AddEdge(edges, cell.BorderBottom, p.Left, p.Top + p.Height, p.Left + p.Width, p.Top + p.Height);
            AddEdge(edges, cell.BorderLeft, p.Left, p.Top, p.Left, p.Top + p.Height);
            AddEdge(edges, cell.BorderRight, p.Left + p.Width, p.Top, p.Left + p.Width, p.Top + p.Height);
        }

        private static void AddEdge(List<EdgePlacement> edges, BorderModel border,
                                    float x1, float y1, float x2, float y2)
        {
            if (border == null || !border.Visible) return;

            var edge = new EdgePlacement
            {
                X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
                Weight = border.Weight,
                ColorRgb = border.ColorRgb,
                DashStyle = border.DashStyle,
                Transparency = border.Transparency
            };

            // Adjacent cells both report the shared edge; keep only one, preferring
            // the heavier line if styles conflict (matches PowerPoint's paint order).
            var existing = edges.FirstOrDefault(e => e.GeometricallyEquals(edge));
            if (existing == null)
            {
                edges.Add(edge);
            }
            else if (edge.Weight > existing.Weight)
            {
                edges.Remove(existing);
                edges.Add(edge);
            }
        }
    }

    public sealed class LayoutResult
    {
        public LayoutResult(IReadOnlyList<CellPlacement> cells, IReadOnlyList<EdgePlacement> edges)
        {
            Cells = cells;
            Edges = edges;
        }

        public IReadOnlyList<CellPlacement> Cells { get; }
        public IReadOnlyList<EdgePlacement> Edges { get; }
    }
}
