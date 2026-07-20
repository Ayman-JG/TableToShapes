using System;
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
            var createdNames = _writer.Write(slide, model, layout);

            float left = tableShape.Left, top = tableShape.Top;
            tableShape.Delete();

            var group = slide.Shapes.Range(createdNames.ToArray()).Group();
            group.Name = "ConvertedTable";
            // Grouping can nudge coordinates; re-assert the original position.
            group.Left = left;
            group.Top = top;
            return group;
        }
    }
}
