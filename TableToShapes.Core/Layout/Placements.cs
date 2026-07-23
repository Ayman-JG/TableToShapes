namespace TableToShapes.Core.Layout
{
    /// <summary>A resolved cell rectangle on the slide (merge anchors only).</summary>
    public sealed class CellPlacement
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int RowSpan { get; set; }
        public int ColumnSpan { get; set; }
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
    }

    /// <summary>
    /// One resolved border line segment to render. Each segment corresponds to a single
    /// physical grid-line unit (one cell edge on one track); the <see cref="LayoutEngine"/>
    /// resolves the two adjacent cells' borders down to this before emitting it, so no two
    /// emitted segments overlap.
    /// </summary>
    public sealed class EdgePlacement
    {
        public const float Epsilon = 0.01f;

        public float X1 { get; set; }
        public float Y1 { get; set; }
        public float X2 { get; set; }
        public float Y2 { get; set; }
        public float Weight { get; set; }
        public int ColorRgb { get; set; }
        public int DashStyle { get; set; }
        public float Transparency { get; set; }

        /// <summary>
        /// Endpoint comparison that ignores direction and tolerates sub-pixel drift. Used by
        /// tests to locate a segment regardless of which end is listed first.
        /// </summary>
        public bool GeometricallyEquals(EdgePlacement other)
        {
            if (other == null) return false;
            return (SamePoint(X1, Y1, other.X1, other.Y1) && SamePoint(X2, Y2, other.X2, other.Y2))
                || (SamePoint(X1, Y1, other.X2, other.Y2) && SamePoint(X2, Y2, other.X1, other.Y1));
        }

        private static bool SamePoint(float ax, float ay, float bx, float by)
            => System.Math.Abs(ax - bx) < Epsilon && System.Math.Abs(ay - by) < Epsilon;
    }
}
