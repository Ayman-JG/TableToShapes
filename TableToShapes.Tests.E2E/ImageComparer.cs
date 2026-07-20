using System;
using System.Drawing;

namespace TableToShapes.Tests.E2E
{
    /// <summary>Pixel-level bitmap comparison used to verify visual fidelity.</summary>
    public static class ImageComparer
    {
        /// <summary>Returns the fraction (0..1) of pixels that differ beyond the tolerance.</summary>
        public static double DiffPercentage(Bitmap a, Bitmap b, int channelTolerance = 2)
        {
            if (a.Width != b.Width || a.Height != b.Height)
                throw new ArgumentException("Bitmaps must be the same size to compare.");

            long differing = 0;
            for (int y = 0; y < a.Height; y++)
            {
                for (int x = 0; x < a.Width; x++)
                {
                    var p = a.GetPixel(x, y);
                    var q = b.GetPixel(x, y);
                    if (Math.Abs(p.R - q.R) > channelTolerance ||
                        Math.Abs(p.G - q.G) > channelTolerance ||
                        Math.Abs(p.B - q.B) > channelTolerance)
                    {
                        differing++;
                    }
                }
            }

            return differing / (double)(a.Width * (long)a.Height);
        }
    }
}
