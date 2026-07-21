using System;
using System;
using System.Collections.Generic;
using TableToShapes.Core.Layout;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TableToShapes.Interop
{
    /// <summary>
    /// End-to-end conversion: read table ? compute layout ? write shapes ?
    /// delete original ? group replacements.
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

            float left = tableShape.Left, top = tableShape.Top;

            List<string> createdNames;
            try
            {
                createdNames = _writer.Write(slide, model, layout);
            }
            catch
            {
                // Writing failed part-way; remove any orphan shapes so the slide is
                // left exactly as we found it (the original table is still intact).
                RemoveByNamePrefix(slide, ShapeWriter.CreatedShapePrefix);
                throw;
            }

            tableShape.Delete();

            var group = slide.Shapes.Range(createdNames.ToArray()).Group();
            group.Name = "ConvertedTable";
            // Grouping can nudge coordinates; re-assert the original position.
            group.Left = left;
            group.Top = top;
            return group;
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
