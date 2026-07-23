using System;
using TableToShapes.Core.Layout;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TableToShapes.Interop
{
    /// <summary>
    /// End-to-end conversion: read table -> compute layout -> write shapes -> group ->
    /// delete original. The original table is deleted only after the replacement group has
    /// been built and positioned, so a failure never destroys the table without a replacement.
    /// </summary>
    public sealed class TableConverter
    {
        private readonly TableReader _reader = new TableReader();
        private readonly LayoutEngine _layoutEngine = new LayoutEngine();
        private readonly ShapeWriter _writer = new ShapeWriter();

        /// <exception cref="InvalidOperationException">The shape is not a table.</exception>
        public PowerPoint.Shape Convert(PowerPoint.Shape tableShape)
        {
            if (tableShape.HasTable != Microsoft.Office.Core.MsoTriState.msoTrue)
                throw new InvalidOperationException("Selected shape is not a table.");

            var slide = (PowerPoint.Slide)tableShape.Parent;

            var model = _reader.Read(tableShape);
            var layout = _layoutEngine.Calculate(model);

            // Opt-in troubleshooting output (off unless TABLETOSHAPES_DIAGNOSTICS is set).
            if (ConversionDiagnostics.Enabled)
                ConversionDiagnostics.Dump(model, layout);

            // The generated shapes sit in absolute slide coordinates taken from the actual cell
            // rectangles, so the replacement's origin is the model's own top-left (which can
            // differ from the table shape's bounding box).
            float left = model.Left, top = model.Top;

            PowerPoint.Shape result = null;
            try
            {
                var createdNames = _writer.Write(slide, model, layout);

                // Group the pieces - or, for a degenerate single-shape result (e.g. a 1x1 cell
                // with no borders), keep the single shape, since Group() requires two or more.
                result = createdNames.Count > 1
                    ? slide.Shapes.Range(createdNames.ToArray()).Group()
                    : slide.Shapes[createdNames[0]];
                result.Name = "ConvertedTable";
                // Grouping can nudge coordinates; re-assert the intended origin.
                result.Left = left;
                result.Top = top;
            }
            catch
            {
                // Nothing has been deleted yet, so leave the slide exactly as we found it:
                // remove the partial/grouped output. The original table is still intact.
                if (result != null) result.Delete();               // group (with its children) or single shape
                else RemoveByNamePrefix(slide, ShapeWriter.CreatedShapePrefix);
                throw;
            }

            tableShape.Delete();
            return result;
        }

        private static void RemoveByNamePrefix(PowerPoint.Slide slide, string prefix)
        {
            // Iterate backwards: deleting shifts the 1-based collection indices.
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                var shape = slide.Shapes[i];
                if (shape.Name != null && shape.Name.StartsWith(prefix, StringComparison.Ordinal))
                    shape.Delete();
            }
        }
    }
}
