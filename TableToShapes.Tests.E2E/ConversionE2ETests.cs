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

        private static PowerPoint.Shape BuildStyledTable(PowerPoint.Slide slide)
        {
            var shape = slide.Shapes.AddTable(2, 2, 50, 50, 400, 120);
            var table = shape.Table;

            var header = table.Cell(1, 1).Shape;
            header.TextFrame2.TextRange.Text = "Name";
            header.TextFrame2.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            header.Fill.ForeColor.RGB = 0x00CC6600; // BGR blue-ish
            table.Cell(1, 2).Shape.TextFrame2.TextRange.Text = "Amount";
            table.Cell(2, 1).Shape.TextFrame2.TextRange.Text = "Coffee fund";
            table.Cell(2, 2).Shape.TextFrame2.TextRange.Text = "£42.00";

            return shape;
        }

        private string Render(PowerPoint.Slide slide, string fileName)
        {
            string path = Path.Combine(_outDir, fileName);
            slide.Export(path, "PNG", 1920, 1080);
            return path;
        }
    }
}
