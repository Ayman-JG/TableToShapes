using System;
using System.Drawing;
using System.IO;
using FluentAssertions;
using NUnit.Framework;
using TableToShapes.Interop;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace TableToShapes.Tests.E2E
{
    /// <summary>
    /// Drives a real PowerPoint instance. Requires PowerPoint installed;
    /// run locally or on a UI-capable agent. [Category("E2E")] lets CI filter these out.
    /// </summary>
    [TestFixture, Category("E2E")]
    public class ConversionE2ETests
    {
        private PowerPoint.Application _app;
        private string _outDir;

        [OneTimeSetUp]
        public void StartPowerPoint()
        {
            _app = new PowerPoint.Application();
            _outDir = Path.Combine(TestContext.CurrentContext.WorkDirectory, "renders");
            Directory.CreateDirectory(_outDir);
        }

        [OneTimeTearDown]
        public void QuitPowerPoint()
        {
            _app?.Quit();
        }

        [Test]
        public void GivenStyledTable_WhenConverted_ThenSlideRendersPixelIdentical()
        {
            var pres = _app.Presentations.Add(Office.MsoTriState.msoFalse);
            try
            {
                var slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                var tableShape = BuildStyledTable(slide);

                string before = Render(slide, "before.png");

                var group = new TableConverter().Convert(tableShape);

                group.Type.Should().Be(Office.MsoShapeType.msoGroup);
                slide.Shapes.Count.Should().Be(1, "the original table must be removed and replaced by one group");

                string after = Render(slide, "after.png");

                using (var b = new Bitmap(before))
                using (var a = new Bitmap(after))
                {
                    ImageComparer.DiffPercentage(b, a)
                        .Should().BeLessThan(0.001, "conversion must be visually identical");
                }
            }
            finally
            {
                pres.Close();
            }
        }

        [Test]
        public void GivenTableWithMergedCells_WhenConverted_ThenSlideRendersPixelIdentical()
        {
            var pres = _app.Presentations.Add(Office.MsoTriState.msoFalse);
            try
            {
                var slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                var tableShape = slide.Shapes.AddTable(3, 3, 60, 60, 480, 180);
                tableShape.Table.Cell(1, 1).Merge(tableShape.Table.Cell(1, 3)); // header row merge
                tableShape.Table.Cell(2, 1).Merge(tableShape.Table.Cell(3, 1)); // vertical merge
                tableShape.Table.Cell(1, 1).Shape.TextFrame2.TextRange.Text = "Merged Header";

                string before = Render(slide, "merged-before.png");

                new TableConverter().Convert(tableShape);

                string after = Render(slide, "merged-after.png");

                using (var b = new Bitmap(before))
                using (var a = new Bitmap(after))
                {
                    ImageComparer.DiffPercentage(b, a).Should().BeLessThan(0.001);
                }
            }
            finally
            {
                pres.Close();
            }
        }

        [Test]
        public void GivenNonTableShape_WhenConverted_ThenThrows()
        {
            var pres = _app.Presentations.Add(Office.MsoTriState.msoFalse);
            try
            {
                var slide = pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                var rect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 10, 10, 100, 100);

                Action act = () => new TableConverter().Convert(rect);

                act.Should().Throw<InvalidOperationException>();
            }
            finally
            {
                pres.Close();
            }
        }

        // Builds a table that exercises the fidelity-sensitive features together, so the
        // pixel-diff guards them all: multiple runs in one cell, a highlighted run, a coloured
        // run, a 2x2 merge, and a cell with its borders switched off.
        private static PowerPoint.Shape BuildStyledTable(PowerPoint.Slide slide)
        {
            const int Red = 0x000000FF;    // RGB(255,0,0) in BGR
            const int Yellow = 0x0000FFFF; // RGB(255,255,0)
            const int Teal = 0x00654321;   // arbitrary header fill
            const int Orange = 0x00227EE6; // RGB(230,126,34)

            var shape = slide.Shapes.AddTable(4, 3, 40, 60, 860, 260);
            var table = shape.Table;

            // Header row - bold, "test" coloured red.
            SetText(table.Cell(1, 1), "he1", bold: true, fill: Teal);
            SetText(table.Cell(1, 2), "he2", bold: true, fill: Teal);
            var testHeader = table.Cell(1, 3);
            testHeader.Shape.Fill.ForeColor.RGB = Teal;
            var testRange = testHeader.Shape.TextFrame2.TextRange;
            testRange.Text = "test";
            testRange.Font.Bold = Office.MsoTriState.msoTrue;
            testRange.Font.Fill.ForeColor.RGB = Red;

            // "Test ME" - the "ME" run is bold and highlighted yellow (two runs in one cell).
            var testMe = table.Cell(2, 1).Shape.TextFrame2.TextRange;
            testMe.Text = "Test ME";
            var meRun = testMe.Characters[6, 2]; // "ME"
            meRun.Font.Bold = Office.MsoTriState.msoTrue;
            meRun.Font.Highlight.RGB = Yellow;

            // Merge columns 2-3 across rows 2-3; "MERGED" run coloured red (two runs in one cell).
            table.Cell(2, 2).Merge(table.Cell(3, 3));
            var merged = table.Cell(2, 2).Shape.TextFrame2.TextRange;
            merged.Text = "This is MERGED";
            merged.Characters[9, 6].Font.Fill.ForeColor.RGB = Red; // "MERGED"

            // "No border cell" over two paragraphs - all four borders turned off. The second
            // paragraph guards multi-line / auto-grown-row fidelity.
            var noBorder = table.Cell(3, 1);
            noBorder.Shape.TextFrame2.TextRange.Text = "No border cell\rsecond line";
            foreach (PowerPoint.PpBorderType side in new[]
            {
                PowerPoint.PpBorderType.ppBorderTop, PowerPoint.PpBorderType.ppBorderBottom,
                PowerPoint.PpBorderType.ppBorderLeft, PowerPoint.PpBorderType.ppBorderRight
            })
                noBorder.Borders[side].Visible = Office.MsoTriState.msoFalse;

            // Bottom row - "ro2a" struck through and in a distinct font; last cell filled orange.
            var ro2a = table.Cell(4, 1).Shape.TextFrame2.TextRange;
            ro2a.Text = "ro2a";
            ro2a.Font.Strike = Office.MsoTextStrike.msoSingleStrike;
            ro2a.Font.Name = "Consolas";
            SetText(table.Cell(4, 2), "ed2");
            var rCell = table.Cell(4, 3);
            rCell.Shape.Fill.ForeColor.RGB = Orange;
            rCell.Shape.TextFrame2.TextRange.Text = "r";

            return shape;
        }

        private static void SetText(PowerPoint.Cell cell, string text, bool bold = false, int? fill = null)
        {
            if (fill.HasValue) cell.Shape.Fill.ForeColor.RGB = fill.Value;
            var range = cell.Shape.TextFrame2.TextRange;
            range.Text = text;
            if (bold) range.Font.Bold = Office.MsoTriState.msoTrue;
        }

        private string Render(PowerPoint.Slide slide, string fileName)
        {
            string path = Path.Combine(_outDir, fileName);
            slide.Export(path, "PNG", 1920, 1080);
            return path;
        }
    }
}
