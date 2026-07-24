using System;
using TableToShapes.Core.Layout;
using TableToShapes.Core.Logging;
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
        private readonly TableReader _reader;
        private readonly LayoutEngine _layoutEngine = new LayoutEngine();
        private readonly ShapeWriter _writer;
        private readonly ILogger _log;

        public TableConverter() : this(NullLogger.Instance) { }

        public TableConverter(ILogger logger)
        {
            _log = logger ?? NullLogger.Instance;
            _reader = new TableReader(_log);
            _writer = new ShapeWriter(_log);
        }

        /// <exception cref="InvalidOperationException">The shape is not a table.</exception>
        public PowerPoint.Shape Convert(PowerPoint.Shape tableShape)
        {
            if (tableShape.HasTable != Microsoft.Office.Core.MsoTriState.msoTrue)
                throw new InvalidOperationException("Selected shape is not a table.");

            var slide = (PowerPoint.Slide)tableShape.Parent;

            var model = _reader.Read(tableShape);
            var layout = _layoutEngine.Calculate(model);

            _log.Info(string.Format("Converting {0}x{1} table ({2} cells, {3} border segments).",
                model.RowCount, model.ColumnCount, layout.Cells.Count, layout.Edges.Count));
            // Only build the (expensive) snapshot string when Debug logging is actually on.
            if (_log.IsEnabled(LogLevel.Debug))
                _log.Debug(ConversionDiagnostics.Describe(model, layout));

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
            catch (Exception ex)
            {
                _log.Error("Conversion failed; rolling back the partial output.", ex);
                // Nothing has been deleted yet, so leave the slide exactly as we found it by
                // removing the partial/grouped output. Cleanup is itself guarded so a failure
                // here cannot mask the original exception (which is rethrown below).
                try
                {
                    if (result != null) result.Delete();           // group (with its children) or single shape
                    else RemoveByNamePrefix(slide, ShapeWriter.CreatedShapePrefix);
                }
                catch (Exception cleanupEx)
                {
                    _log.Warning("Cleanup after a failed conversion did not complete.", cleanupEx);
                }
                throw;
            }

            tableShape.Delete();
            _log.Info("Conversion succeeded.");
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
