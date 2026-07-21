using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using TableToShapes.Core.Layout;

namespace TableToShapes.Tests.Unit
{
    [TestFixture]
    public class LayoutEngineTests
    {
        private readonly LayoutEngine _engine = new LayoutEngine();

        [Test]
        public void GivenSimpleGrid_WhenCalculating_ThenEveryCellGetsAPlacement()
        {
            var table = TableModelBuilder.Grid(2, 3);

            var layout = _engine.Calculate(table);

            layout.Cells.Should().HaveCount(6);
        }

        [Test]
        public void GivenSimpleGrid_WhenCalculating_ThenPlacementsMatchCumulativeOffsets()
        {
            var table = TableModelBuilder.Grid(2, 2, rowHeight: 30f, colWidth: 120f, left: 50f, top: 10f);

            var layout = _engine.Calculate(table);

            var bottomRight = layout.Cells.Single(c => c.Row == 1 && c.Column == 1);
            bottomRight.Left.Should().Be(170f);   // 50 + 120
            bottomRight.Top.Should().Be(40f);     // 10 + 30
            bottomRight.Width.Should().Be(120f);
            bottomRight.Height.Should().Be(30f);
        }

        [Test]
        public void GivenHorizontallyMergedCells_WhenCalculating_ThenSinglePlacementSpansColumns()
        {
            var table = TableModelBuilder.Grid(1, 3, colWidth: 100f);
            table.Cells[0, 1].MergeId = table.Cells[0, 0].MergeId; // merge first two cells

            var layout = _engine.Calculate(table);

            layout.Cells.Should().HaveCount(2);
            var merged = layout.Cells.Single(c => c.Column == 0);
            merged.ColumnSpan.Should().Be(2);
            merged.Width.Should().Be(200f);
        }

        [Test]
        public void GivenVerticallyMergedCells_WhenCalculating_ThenSinglePlacementSpansRows()
        {
            var table = TableModelBuilder.Grid(3, 1, rowHeight: 20f);
            table.Cells[1, 0].MergeId = table.Cells[0, 0].MergeId;
            table.Cells[2, 0].MergeId = table.Cells[0, 0].MergeId;

            var layout = _engine.Calculate(table);

            layout.Cells.Should().HaveCount(1);
            layout.Cells[0].RowSpan.Should().Be(3);
            layout.Cells[0].Height.Should().Be(60f);
        }

        [Test]
        public void GivenAdjacentCells_WhenCalculating_ThenSharedEdgeIsRenderedOnce()
        {
            var table = TableModelBuilder.Grid(1, 2);

            var layout = _engine.Calculate(table);

            // 1x2 grid: 7 unique segments (6 outer + 1 shared inner), not 8.
            layout.Edges.Should().HaveCount(7);
        }

        [Test]
        public void GivenConflictingSharedEdgeWeights_WhenCalculating_ThenHeavierWins()
        {
            var table = TableModelBuilder.Grid(1, 2);
            table.Cells[0, 0].BorderRight = TableModelBuilder.VisibleBorder(weight: 3f);
            table.Cells[0, 1].BorderLeft = TableModelBuilder.VisibleBorder(weight: 1f);

            var layout = _engine.Calculate(table);

            var shared = layout.Edges.Single(e => e.X1 == 100f && e.X2 == 100f);
            shared.Weight.Should().Be(3f);
        }

        [Test]
        public void GivenInvisibleBorders_WhenCalculating_ThenNoEdgesAreEmitted()
        {
            var table = TableModelBuilder.Grid(1, 1);
            var cell = table.Cells[0, 0];
            cell.BorderTop.Visible = false;
            cell.BorderBottom.Visible = false;
            cell.BorderLeft.Visible = false;
            cell.BorderRight.Visible = false;

            var layout = _engine.Calculate(table);

            layout.Edges.Should().BeEmpty();
        }

        [Test]
        public void GivenHorizontallyMergedCells_WhenCalculating_ThenNoBorderRunsThroughTheMerge()
        {
            var table = TableModelBuilder.Grid(1, 3, colWidth: 100f); // x edges at 0,100,200,300
            table.Cells[0, 1].MergeId = table.Cells[0, 0].MergeId;    // merge first two columns

            var layout = _engine.Calculate(table);

            // The interior boundary at x=100 is now inside the merged cell and must not be drawn.
            layout.Edges.Should().NotContain(e => e.X1 == 100f && e.X2 == 100f,
                "a merged cell has no internal grid line");
            // Merge removes exactly that one interior vertical: 10 edges for a plain 1x3 -> 9.
            layout.Edges.Should().HaveCount(9);
        }

        [Test]
        public void GivenHorizontallyMergedCells_WhenCalculating_ThenOuterBorderStillSpansFullWidth()
        {
            var table = TableModelBuilder.Grid(1, 3, colWidth: 100f);
            table.Cells[0, 1].MergeId = table.Cells[0, 0].MergeId;

            var layout = _engine.Calculate(table);

            // Top border of the merge is emitted as aligned per-column segments (0..100, 100..200).
            layout.Edges.Should().Contain(e => e.Y1 == 0f && e.Y2 == 0f && e.X1 == 0f && e.X2 == 100f);
            layout.Edges.Should().Contain(e => e.Y1 == 0f && e.Y2 == 0f && e.X1 == 100f && e.X2 == 200f);
        }

        [Test]
        public void GivenVerticallyMergedCells_WhenCalculating_ThenNoBorderRunsThroughTheMerge()
        {
            var table = TableModelBuilder.Grid(3, 1, rowHeight: 20f); // y edges at 0,20,40,60
            table.Cells[1, 0].MergeId = table.Cells[0, 0].MergeId;
            table.Cells[2, 0].MergeId = table.Cells[0, 0].MergeId;

            var layout = _engine.Calculate(table);

            layout.Edges.Should().NotContain(e => e.Y1 == 20f && e.Y2 == 20f);
            layout.Edges.Should().NotContain(e => e.Y1 == 40f && e.Y2 == 40f);
            layout.Edges.Should().HaveCount(8); // 1 top + 1 bottom + 3 left + 3 right
        }

        [Test]
        public void GivenMergedCellNextToBorderlessCell_WhenCalculating_ThenSharedEdgeUsesTheBorderedNeighbour()
        {
            // Mirrors the test-pic case: a vertically merged right column beside a top cell
            // that has a border and a bottom "no border" cell.
            var table = TableModelBuilder.Grid(2, 2, rowHeight: 20f, colWidth: 100f);
            table.Cells[0, 1].MergeId = 99;
            table.Cells[1, 1].MergeId = 99;                         // right column merged across both rows
            // Left-bottom cell has no borders at all.
            table.Cells[1, 0].BorderTop = new BorderModel { Visible = false };
            table.Cells[1, 0].BorderBottom = new BorderModel { Visible = false };
            table.Cells[1, 0].BorderLeft = new BorderModel { Visible = false };
            table.Cells[1, 0].BorderRight = new BorderModel { Visible = false };

            var layout = _engine.Calculate(table);

            // The row-0 portion of the x=100 divider comes from the bordered top-left cell,
            // which is added before the merged cell, so the merged cell does not override it.
            var row0Divider = layout.Edges.Single(e => e.X1 == 100f && e.X2 == 100f && e.Y1 == 0f && e.Y2 == 20f);
            row0Divider.Should().NotBeNull();
        }
    }

    [TestFixture]
    public class EdgePlacementTests
    {
        [Test]
        public void GivenReversedEndpoints_WhenComparing_ThenEdgesAreGeometricallyEqual()
        {
            var a = new EdgePlacement { X1 = 0, Y1 = 0, X2 = 100, Y2 = 0 };
            var b = new EdgePlacement { X1 = 100, Y1 = 0, X2 = 0, Y2 = 0 };

            a.GeometricallyEquals(b).Should().BeTrue();
        }

        [Test]
        public void GivenEndpointsWithinEpsilon_WhenComparing_ThenEdgesAreGeometricallyEqual()
        {
            var a = new EdgePlacement { X1 = 0, Y1 = 0, X2 = 100, Y2 = 0 };
            var b = new EdgePlacement { X1 = 0.005f, Y1 = 0, X2 = 100.005f, Y2 = 0 };

            a.GeometricallyEquals(b).Should().BeTrue();
        }

        [Test]
        public void GivenDifferentEndpoints_WhenComparing_ThenEdgesAreNotEqual()
        {
            var a = new EdgePlacement { X1 = 0, Y1 = 0, X2 = 100, Y2 = 0 };
            var b = new EdgePlacement { X1 = 0, Y1 = 20, X2 = 100, Y2 = 20 };

            a.GeometricallyEquals(b).Should().BeFalse();
        }
    }
}
