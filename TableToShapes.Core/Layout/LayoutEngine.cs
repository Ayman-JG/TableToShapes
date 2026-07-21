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

                    AddCellEdges(edges, table, placement, rowOffsets, colOffsets);
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

        // A merged cell's border is not a single long line: PowerPoint stores and paints it
        // as one segment per crossed row/column, and each segment can come from a different
        // sub-cell. Emitting the border as one full-span line (the old behaviour) meant a
        // merged cell's border never lined up with the individual borders of the cells on the
        // far side, so both were drawn and the heavier/darker one bled through (the stray
        // vertical line in the test render). Decomposing into per-track unit segments lets the
        // shared-edge de-duplication in AddEdge collapse them correctly.
        private static void AddCellEdges(
            List<EdgePlacement> edges, TableModel table, CellPlacement p,
            float[] rowOffsets, float[] colOffsets)
        {
            int r = p.Row, c = p.Column, rs = p.RowSpan, cs = p.ColumnSpan;
            float top = rowOffsets[r], bottom = rowOffsets[r + rs];
            float left = colOffsets[c], right = colOffsets[c + cs];

            // Top / bottom: one segment per spanned column, sourced from the edge-most sub-cell.
            for (int cc = c; cc < c + cs; cc++)
            {
                AddEdge(edges, table.Cells[r, cc].BorderTop,
                    colOffsets[cc], top, colOffsets[cc + 1], top);
                AddEdge(edges, table.Cells[r + rs - 1, cc].BorderBottom,
                    colOffsets[cc], bottom, colOffsets[cc + 1], bottom);
            }

            // Left / right: one segment per spanned row.
            for (int rr = r; rr < r + rs; rr++)
            {
                AddEdge(edges, table.Cells[rr, c].BorderLeft,
                    left, rowOffsets[rr], left, rowOffsets[rr + 1]);
                AddEdge(edges, table.Cells[rr, c + cs - 1].BorderRight,
                    right, rowOffsets[rr], right, rowOffsets[rr + 1]);
            }
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
