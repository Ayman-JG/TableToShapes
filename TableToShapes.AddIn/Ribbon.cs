using System;
using System.Windows.Forms;
using TableToShapes.Interop;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace TableToShapes.AddIn
{
    /// <summary>
    /// Ribbon callback handler. Wire this up from Ribbon.xml (button id="ConvertTableButton").
    /// </summary>
    public partial class Ribbon
    {
        private readonly TableConverter _converter = new TableConverter();

        public void OnConvertTableClicked(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.ActiveWindow.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes ||
                    selection.ShapeRange.Count != 1 ||
                    selection.ShapeRange[1].HasTable != Office.MsoTriState.msoTrue)
                {
                    MessageBox.Show("Please select a single table first.", "Table to Shapes",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var group = _converter.Convert(selection.ShapeRange[1]);
                group.Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Conversion failed: " + ex.Message, "Table to Shapes",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
