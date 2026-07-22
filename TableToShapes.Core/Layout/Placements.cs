using System;

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
    /// emitted segments overlap. The geometric helpers below remain for tests and any caller
    /// that needs to compare segments.
    /// </summary>
    public sealed class EdgePlacement : IEquatable<EdgePlacement>
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

        public bool GeometricallyEquals(EdgePlacement other)
        {
            if (other == null) return false;
            return (SamePoint(X1, Y1, other.X1, other.Y1) && SamePoint(X2, Y2, other.X2, other.Y2))
                || (SamePoint(X1, Y1, other.X2, other.Y2) && SamePoint(X2, Y2, other.X1, other.Y1));
        }

        public bool Equals(EdgePlacement other)
        {
            return other != null
                && GeometricallyEquals(other)
                && Math.Abs(Weight - other.Weight) < Epsilon
                && ColorRgb == other.ColorRgb
                && DashStyle == other.DashStyle;
        }

        public override bool Equals(object obj) => Equals(obj as EdgePlacement);

        public override int GetHashCode()
        {
            // Order-independent endpoint hash so reversed segments collide.
            int p1 = Quantize(X1) * 397 ^ Quantize(Y1);
            int p2 = Quantize(X2) * 397 ^ Quantize(Y2);
            return (p1 ^ p2) * 31 + ColorRgb;
        }

        private static int Quantize(float v) => (int)Math.Round(v / Epsilon);

        private static bool SamePoint(float ax, float ay, float bx, float by)
            => Math.Abs(ax - bx) < Epsilon && Math.Abs(ay - by) < Epsilon;
    }
}
