using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using TableToShapes.Core.Logging;
using TableToShapes.Interop;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TableToShapes.AddIn
{
    /// <summary>
    /// COM add-in entry point. PowerPoint discovers it through the registry
    /// (see install-addin.ps1), instantiates it via COM, and calls the
    /// IDTExtensibility2 lifecycle methods. GetCustomUI supplies the Ribbon
    /// button; OnConvertClicked runs the conversion on the current selection.
    /// </summary>
    [ComVisible(true)]
    [Guid("6F9B0C64-3C1A-4E6E-9B7D-2D3E8A11F0AB")]
    [ProgId("TableToShapes.AddIn.Connect")]
    // AutoDispatch exposes public methods via IDispatch so the Ribbon can resolve
    // the onAction callback ("OnConvertClicked") by name at runtime.
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private PowerPoint.Application _app;

        // Composition root: build the single logger the pipeline uses. The minimum level comes
        // from the TABLETOSHAPES_LOGLEVEL environment variable (Debug/Info/Warning/Error/None),
        // defaulting to Info. Everything is written to %TEMP%\TableToShapes.log, which
        // self-rotates so it never grows without bound.
        private static readonly ILogger Log = CreateLogger();

        private static ILogger CreateLogger()
        {
            var level = LogLevelParser.Parse(
                Environment.GetEnvironmentVariable("TABLETOSHAPES_LOGLEVEL"), LogLevel.Info);
            var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "TableToShapes.log");
            return new FileLogger(path, level);
        }

        // ---- IDTExtensibility2 ----

        public void OnConnection(object application, ext_ConnectMode connectMode,
                                 object addInInst, ref Array custom)
        {
            try
            {
                Log.Info($"OnConnection: mode={connectMode}");
                _app = (PowerPoint.Application)application;
            }
            catch (Exception ex)
            {
                Log.Error("OnConnection failed.", ex);
                throw;
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ---- IRibbonExtensibility ----

        public string GetCustomUI(string ribbonID)
        {
            Log.Debug($"GetCustomUI: ribbonID={ribbonID}");
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab idMso='TabHome'>
        <group id='TableToShapesGroup' label='Table to Shapes'>
          <button id='ConvertTableButton'
                  label='Convert Table'
                  size='large'
                  imageMso='TableDrawTable'
                  onAction='OnConvertClicked'
                  screentip='Convert Table to Shapes'
                  supertip='Replaces the selected table with a visually identical group of shapes.'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // ---- Ribbon callback (dispatched by name via COM late binding) ----

        public void OnConvertClicked(IRibbonControl control)
        {
            try
            {
                Log.Debug("Convert clicked.");
                var shape = GetSelectedShape();
                if (shape == null || shape.HasTable != Office.MsoTriState.msoTrue)
                {
                    Log.Info("Convert clicked with no table selected.");
                    MessageBox.Show("Please select a table first.", "Table to Shapes",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                new TableConverter(Log).Convert(shape);
            }
            catch (Exception ex)
            {
                Log.Error("Conversion failed.", ex);
                MessageBox.Show("Conversion failed: " + ex.Message, "Table to Shapes",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private PowerPoint.Shape GetSelectedShape()
        {
            var window = _app?.ActiveWindow;
            if (window == null) return null;

            var selection = window.Selection;
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
                return null;

            var range = selection.ShapeRange;
            return range.Count >= 1 ? range[1] : null;
        }
    }
}
