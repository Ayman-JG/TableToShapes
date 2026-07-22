using System.Collections.Generic;
using System.Linq;
using TableToShapes.Core.Model;

namespace TableToShapes.Core.Layout
{
    /// <summary>
    /// Pure geometry: turns a <see cref="TableModel"/> into cell rectangles (merge anchors)
    /// and border line segments. No Interop dependencies, so it is fully unit-testable.
    ///
    /// <para><b>Border model.</b> A grid line between two cells is a single <i>physical edge</i>
    /// that both adjacent cells describe, and the two descriptions can disagree (most often
    /// after a merge). So we do not emit borders per cell and reconcile afterwards - that made
    /// the result depend on emission order. Instead we walk every physical grid-line unit once,
    /// gather the (up to two) adjacent cells' border "opinions", and resolve them to a single
    /// line. All precedence lives in one place, <see cref="Resolve"/>.</para>
    ///
    /// <para><b>Rules</b> (see docs/FIDELITY_RULES.md for the full write-up):</para>
    /// <list type="number">
    ///   <item>R1 - A grid line interior to a merged cell is not a border (skipped).</item>
    ///   <item>R2 - On the table boundary only one cell exists; use its opinion.</item>
    ///   <item>R3 - A plain cell's border beats a merged cell's. Merged cells report a stray
    ///         "automatic" (usually black, still Visible) border on their outer edges after a
    ///         merge; the plain neighbour holds the real value. When the two were instead edited
    ///         by the user they are kept in sync by PowerPoint, so they already agree and picking
    ///         the plain side is harmless.</item>
    ///   <item>R4 - Among same-tier opinions a visible border beats "off"; a line is dropped only
    ///         when every surviving opinion is off.</item>
    ///   <item>R5 - Among same-tier visible opinions the heavier weight wins (it carries its own
    ///         colour / dash / transparency).</item>
    ///   <item>R6 - Ties break to the negative (left/top) side, purely for determinism.</item>
    /// </list>
    ///
    /// <para><b>Caveats.</b>
    /// (a) R3 assumes an explicitly-set merged border is synced onto its neighbour (observed in
    /// PowerPoint); if a future case shows an un-synced, intentionally-different merged border,
    /// R3 would drop it - handle that by only discarding merged opinions that conflict with the
    /// plain side rather than all of them.
    /// (b) Two merged cells meeting on a shared edge (no plain side) fall through to R5/R6; both
    /// could be artifacts, but this is rare.
    /// (c) R6's tie only fires for two same-tier visible borders of equal weight but different
    /// style - which essentially never happens for synced plain cells; it exists so output is
    /// deterministic, not because PowerPoint is known to prefer that side.</para>
    /// </summary>
    public sealed class LayoutEngine
    {
        private const float WeightEpsilon = 0.001f;

        public LayoutResult Calculate(TableModel table)
        {
            var rowOffsets = Accumulate(table.RowHeights, table.Top);
            var colOffsets = Accumulate(table.ColumnWidths, table.Left);

            var cells = BuildPlacements(table, rowOffsets, colOffsets);

            int[,] anchorRow, anchorCol;
            BuildAnchorMap(table, out anchorRow, out anchorCol);

            var edges = ResolveEdges(table, anchorRow, anchorCol, rowOffsets, colOffsets);

            return new LayoutResult(cells, edges);
        }

        // ---- placements (merge anchors only) ----

        private static List<CellPlacement> BuildPlacements(TableModel table, float[] rowOffsets, float[] colOffsets)
        {
            var cells = new List<CellPlacement>();
            for (int r = 0; r < table.RowCount; r++)
            {
                for (int c = 0; c < table.ColumnCount; c++)
                {
                    if (!IsMergeAnchor(table, r, c)) continue;

                    int rowSpan = GetRowSpan(table, r, c);
                    int colSpan = GetColumnSpan(table, r, c);

                    cells.Add(new CellPlacement
                    {
                        Row = r,
                        Column = c,
                        RowSpan = rowSpan,
                        ColumnSpan = colSpan,
                        Left = colOffsets[c],
                        Top = rowOffsets[r],
                        Width = colOffsets[c + colSpan] - colOffsets[c],
                        Height = rowOffsets[r + rowSpan] - rowOffsets[r]
                    });
                }
            }
            return cells;
        }

        // For every grid slot, the (row,col) of the top-left cell of the merge it belongs to.
        private static void BuildAnchorMap(TableModel table, out int[,] anchorRow, out int[,] anchorCol)
        {
            int rows = table.RowCount, cols = table.ColumnCount;
            anchorRow = new int[rows, cols];
            anchorCol = new int[rows, cols];
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    int id = table.Cells[r, c].MergeId;
                    int ar = r; while (ar > 0 && table.Cells[ar - 1, c].MergeId == id) ar--;
                    int ac = c; while (ac > 0 && table.Cells[r, ac - 1].MergeId == id) ac--;
                    anchorRow[r, c] = ar;
                    anchorCol[r, c] = ac;
                }
            }
        }

        // ---- edge resolution ----

        private enum Side { Left, Right, Top, Bottom }

        private struct Opinion
        {
            public int AnchorRow;
            public int AnchorCol;
            public bool FromMerged;
            public bool IsPositiveSide; // owner is the right/bottom cell (used only for tie-breaks)
            public BorderModel Border;
        }

        private List<EdgePlacement> ResolveEdges(
            TableModel table, int[,] anchorRow, int[,] anchorCol, float[] rowOffsets, float[] colOffsets)
        {
            int rows = table.RowCount, cols = table.ColumnCount;
            var edges = new List<EdgePlacement>();

            // Vertical grid lines: boundary column j, over row i.
            for (int j = 0; j <= cols; j++)
            {
                for (int i = 0; i < rows; i++)
                {
                    Opinion? neg = j > 0 ? MakeOpinion(table, anchorRow, anchorCol, i, j - 1, Side.Right) : (Opinion?)null;
                    Opinion? pos = j < cols ? MakeOpinion(table, anchorRow, anchorCol, i, j, Side.Left) : (Opinion?)null;
                    if (IsInteriorToMerge(neg, pos)) continue;

                    var border = Resolve(neg, pos);
                    if (border != null)
                        edges.Add(MakeEdge(border, colOffsets[j], rowOffsets[i], colOffsets[j], rowOffsets[i + 1]));
                }
            }

            // Horizontal grid lines: boundary row i, over column j.
            for (int i = 0; i <= rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    Opinion? neg = i > 0 ? MakeOpinion(table, anchorRow, anchorCol, i - 1, j, Side.Bottom) : (Opinion?)null;
                    Opinion? pos = i < rows ? MakeOpinion(table, anchorRow, anchorCol, i, j, Side.Top) : (Opinion?)null;
                    if (IsInteriorToMerge(neg, pos)) continue;

                    var border = Resolve(neg, pos);
                    if (border != null)
                        edges.Add(MakeEdge(border, colOffsets[j], rowOffsets[i], colOffsets[j + 1], rowOffsets[i]));
                }
            }

            return edges;
        }

        private Opinion MakeOpinion(TableModel table, int[,] anchorRow, int[,] anchorCol, int i, int j, Side side)
        {
            int ar = anchorRow[i, j], ac = anchorCol[i, j];
            var cell = table.Cells[ar, ac];
            bool merged = GetRowSpan(table, ar, ac) > 1 || GetColumnSpan(table, ar, ac) > 1;
            return new Opinion
            {
                AnchorRow = ar,
                AnchorCol = ac,
                FromMerged = merged,
                IsPositiveSide = side == Side.Left || side == Side.Top,
                Border = BorderOf(cell, side)
            };
        }

        // Both sides belong to the same merged cell => the line is interior to that cell.
        private static bool IsInteriorToMerge(Opinion? neg, Opinion? pos)
        {
            return neg.HasValue && pos.HasValue
                && neg.Value.AnchorRow == pos.Value.AnchorRow
                && neg.Value.AnchorCol == pos.Value.AnchorCol;
        }

        /// <summary>Resolves the two adjacent cells' opinions into a single rendered border (or null).</summary>
        private static BorderModel Resolve(Opinion? neg, Opinion? pos)
        {
            var ops = new List<Opinion>(2);
            if (neg.HasValue) ops.Add(neg.Value);
            if (pos.HasValue) ops.Add(pos.Value);
            if (ops.Count == 0) return null;

            // R3 - a plain cell's border is authoritative; merged cells emit stray/automatic
            // borders on their outer edges, so drop them when a plain opinion is present.
            if (ops.Any(o => !o.FromMerged))
                ops = ops.Where(o => !o.FromMerged).ToList();

            // R4 - a visible border beats "off"; only if every surviving opinion is off is there
            // no line.
            var visible = ops.Where(o => o.Border != null && o.Border.Visible).ToList();
            if (visible.Count == 0) return null;

            // R5 - heavier wins; R6 - on a tie prefer the negative (left/top) owner, so the
            // result never depends on iteration order.
            Opinion winner = visible[0];
            for (int k = 1; k < visible.Count; k++)
            {
                var o = visible[k];
                if (o.Border.Weight > winner.Border.Weight + WeightEpsilon)
                    winner = o;
                else if (System.Math.Abs(o.Border.Weight - winner.Border.Weight) <= WeightEpsilon
                         && !o.IsPositiveSide && winner.IsPositiveSide)
                    winner = o;
            }
            return winner.Border;
        }

        private static BorderModel BorderOf(CellModel cell, Side side)
        {
            switch (side)
            {
                case Side.Left: return cell.BorderLeft;
                case Side.Right: return cell.BorderRight;
                case Side.Top: return cell.BorderTop;
                default: return cell.BorderBottom;
            }
        }

        private static EdgePlacement MakeEdge(BorderModel border, float x1, float y1, float x2, float y2)
        {
            return new EdgePlacement
            {
                X1 = x1, Y1 = y1, X2 = x2, Y2 = y2,
                Weight = border.Weight,
                ColorRgb = border.ColorRgb,
                DashStyle = border.DashStyle,
                Transparency = border.Transparency
            };
        }

        // ---- shared helpers ----

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
