using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using TableToShapes.Core.Layout;
using TableToShapes.Core.Model;

namespace TableToShapes.Tests.Unit
{
    /// <summary>
    /// Geometry / placement behaviour of <see cref="LayoutEngine"/>.
    /// </summary>
    [TestFixture]
    public class LayoutEnginePlacementTests
    {
        private readonly LayoutEngine _engine = new LayoutEngine();

        [Test]
        public void GivenSimpleGrid_WhenCalculating_ThenEveryCellGetsAPlacement()
        {
            var layout = _engine.Calculate(TableModelBuilder.Grid(2, 3));

            layout.Cells.Should().HaveCount(6);
        }

        [Test]
        public void GivenSimpleGrid_WhenCalculating_ThenPlacementsMatchCumulativeOffsets()
        {
            var table = TableModelBuilder.Grid(2, 2, rowHeight: 30f, colWidth: 120f, left: 50f, top: 10f);

            var layout = _engine.Calculate(table);

            var bottomRight = layout.Cells.Single(c => c.Row == 1 && c.Column == 1);
            bottomRight.Left.Should().Be(170f);   // 50 + 120
            bottomRight.Top.Should().Be(40f);      // 10 + 30
            bottomRight.Width.Should().Be(120f);
            bottomRight.Height.Should().Be(30f);
        }

        [Test]
        public void GivenHorizontallyMergedCells_WhenCalculating_ThenOnePlacementSpansColumns()
        {
            var table = TableModelBuilder.Grid(1, 3, colWidth: 100f);
            table.Cells[0, 1].MergeId = table.Cells[0, 0].MergeId;

            var layout = _engine.Calculate(table);

            layout.Cells.Should().HaveCount(2);
            var merged = layout.Cells.Single(c => c.Column == 0);
            merged.ColumnSpan.Should().Be(2);
            merged.Width.Should().Be(200f);
        }

        [Test]
        public void GivenVerticallyMergedCells_WhenCalculating_ThenOnePlacementSpansRows()
        {
            var table = TableModelBuilder.Grid(3, 1, rowHeight: 20f);
            table.Cells[1, 0].MergeId = table.Cells[0, 0].MergeId;
            table.Cells[2, 0].MergeId = table.Cells[0, 0].MergeId;

            var layout = _engine.Calculate(table);

            layout.Cells.Should().HaveCount(1);
            layout.Cells[0].RowSpan.Should().Be(3);
            layout.Cells[0].Height.Should().Be(60f);
        }
    }

    /// <summary>
    /// Border resolution. One test (or small group) per rule R1-R6 from
    /// <see cref="LayoutEngine"/> / docs/FIDELITY_RULES.md, plus a couple of end-to-end
    /// combinations of merges, artifacts and borderless cells.
    /// </summary>
    [TestFixture]
    public class LayoutEngineBorderTests
    {
        private readonly LayoutEngine _engine = new LayoutEngine();

        private const int Black = 0x000000;
        private const int White = 0xFFFFFF;
        private const int Red = 0xFF0000;
        private const int Blue = 0x0000FF;

        // ---- baseline ----

        [Test]
        public void GivenAdjacentCells_WhenCalculating_ThenSharedEdgeIsRenderedOnce()
        {
            // 1x2 grid: 6 outer segments + 1 shared interior = 7 (not 8).
            var layout = _engine.Calculate(TableModelBuilder.Grid(1, 2));

            layout.Edges.Should().HaveCount(7);
        }

        // ---- R1: a grid line interior to a merged cell is not a border ----

        [Test]
        public void R1_GivenHorizontalMerge_WhenCalculating_ThenNoLineRunsThroughTheMerge()
        {
            var table = TableModelBuilder.Grid(1, 3, colWidth: 100f); // x edges 0,100,200,300
            table.Cells[0, 1].MergeId = table.Cells[0, 0].MergeId;    // merge columns 0-1

            var layout = _engine.Calculate(table);

            layout.Edges.Should().NotContain(e => e.X1 == 100f && e.X2 == 100f,
                "x=100 is now interior to the merged cell");
            layout.Edges.Should().HaveCount(9); // plain 1x3 has 10; the one interior vertical is gone
        }

        [Test]
        public void R1_GivenVerticalMerge_WhenCalculating_ThenNoLineRunsThroughTheMerge()
        {
            var table = TableModelBuilder.Grid(3, 1, rowHeight: 20f); // y edges 0,20,40,60
            table.Cells[1, 0].MergeId = table.Cells[0, 0].MergeId;
            table.Cells[2, 0].MergeId = table.Cells[0, 0].MergeId;

            var layout = _engine.Calculate(table);

            layout.Edges.Should().NotContain(e => e.Y1 == 20f && e.Y2 == 20f);
            layout.Edges.Should().NotContain(e => e.Y1 == 40f && e.Y2 == 40f);
            layout.Edges.Should().HaveCount(8); // 1 top + 1 bottom + 3 left + 3 right
        }

        // ---- R2: table-boundary edge uses the single existing side ----

        [Test]
        public void R2_GivenSingleCell_WhenCalculating_ThenAllFourBoundaryEdgesAreDrawn()
        {
            var layout = _engine.Calculate(TableModelBuilder.Grid(1, 1));

            layout.Edges.Should().HaveCount(4);
        }

        [Test]
        public void R2_GivenMergedCell_WhenCalculating_ThenItsOuterBorderSpansEveryTrack()
        {
            var table = TableModelBuilder.Grid(1, 3, colWidth: 100f);
            table.Cells[0, 1].MergeId = table.Cells[0, 0].MergeId;

            var layout = _engine.Calculate(table);

            // Top border of the merge is emitted as aligned per-column segments (0..100, 100..200).
            layout.Edges.Should().Contain(e => e.Y1 == 0f && e.Y2 == 0f && e.X1 == 0f && e.X2 == 100f);
            layout.Edges.Should().Contain(e => e.Y1 == 0f && e.Y2 == 0f && e.X1 == 100f && e.X2 == 200f);
        }

        // ---- R3: a plain cell's border beats a merged cell's ----

        [Test]
        public void R3_GivenMergedArtifactBorder_WhenPlainNeighbourDisagrees_ThenPlainWins()
        {
            // Merged cell (rows 0-1) reports a black bottom border - the merge artifact - while
            // the plain cell below reports white. PowerPoint renders the plain cell's border.
            var table = TableModelBuilder.Grid(3, 1, rowHeight: 20f); // y edges 0,20,40,60
            table.Cells[1, 0].MergeId = table.Cells[0, 0].MergeId;    // merge rows 0-1
            table.Cells[0, 0].BorderBottom = TableModelBuilder.VisibleBorder(color: Black);
            table.Cells[2, 0].BorderTop = TableModelBuilder.VisibleBorder(color: White);

            var layout = _engine.Calculate(table);

            var shared = layout.Edges.Single(e => e.Y1 == 40f && e.Y2 == 40f);
            shared.ColorRgb.Should().Be(White);
        }

        [Test]
        public void R3_GivenMergedArtifactBorder_WhenPlainNeighbourIsOff_ThenNoLine()
        {
            // A vertically merged right column reports a stray black left border, beside a fully
            // borderless bottom-left cell. The plain "off" wins, so no divider is drawn.
            var table = TableModelBuilder.Grid(2, 2, rowHeight: 20f, colWidth: 100f);
            MergeRightColumn(table);
            table.Cells[0, 1].BorderLeft = TableModelBuilder.VisibleBorder(color: Black); // artifact
            table.Cells[1, 1].BorderLeft = TableModelBuilder.VisibleBorder(color: Black);
            table.Cells[0, 0].BorderRight = TableModelBuilder.VisibleBorder(color: White);
            SetBorderless(table.Cells[1, 0]); // bottom-left cell: no borders

            var layout = _engine.Calculate(table);

            // Upper row: plain white wins over the merged black; lower row: plain "off" wins -> gone.
            layout.Edges.Should().Contain(e => Vertical(e, 100f, 0f, 20f) && e.ColorRgb == White);
            layout.Edges.Should().NotContain(e => Vertical(e, 100f, 20f, 40f));
            // The stray black merged border never appears on this divider.
            layout.Edges.Should().NotContain(e => e.X1 == 100f && e.X2 == 100f && e.ColorRgb == Black);
        }

        // ---- R4: a visible border beats "off" (same tier) ----

        [Test]
        public void R4_GivenBorderlessCell_WhenPlainNeighbourHasBorder_ThenBorderIsKept()
        {
            // Two plain cells: left borderless, right keeps its (default) left border. A visible
            // border is never erased by a borderless neighbour - PowerPoint keeps the shared edge
            // in sync, so "visible wins" is the safe, non-destructive resolution.
            var table = TableModelBuilder.Grid(1, 2, colWidth: 100f);
            SetBorderless(table.Cells[0, 0]);

            var layout = _engine.Calculate(table);

            layout.Edges.Should().Contain(e => e.X1 == 100f && e.X2 == 100f);
        }

        [Test]
        public void R4_GivenBothSidesOff_WhenCalculating_ThenNoLine()
        {
            var table = TableModelBuilder.Grid(1, 2, colWidth: 100f);
            table.Cells[0, 0].BorderRight = new BorderModel { Visible = false };
            table.Cells[0, 1].BorderLeft = new BorderModel { Visible = false };

            var layout = _engine.Calculate(table);

            layout.Edges.Should().NotContain(e => e.X1 == 100f && e.X2 == 100f);
        }

        // ---- R5: heavier weight wins (same tier) ----

        [Test]
        public void R5_GivenConflictingWeights_WhenCalculating_ThenHeavierWins()
        {
            var table = TableModelBuilder.Grid(1, 2);
            table.Cells[0, 0].BorderRight = TableModelBuilder.VisibleBorder(weight: 3f);
            table.Cells[0, 1].BorderLeft = TableModelBuilder.VisibleBorder(weight: 1f);

            var layout = _engine.Calculate(table);

            layout.Edges.Single(e => e.X1 == 100f && e.X2 == 100f).Weight.Should().Be(3f);
        }

        // ---- R6: deterministic tie-break (negative/left-top side) ----

        [Test]
        public void R6_GivenEqualWeightDifferentColour_WhenCalculating_ThenNegativeSideWinsDeterministically()
        {
            // This tie basically never occurs for real (synced) plain cells; the rule exists only
            // so the output is deterministic. The negative (left) side owns the tie.
            var table = TableModelBuilder.Grid(1, 2, colWidth: 100f);
            table.Cells[0, 0].BorderRight = TableModelBuilder.VisibleBorder(color: Red);  // negative side
            table.Cells[0, 1].BorderLeft = TableModelBuilder.VisibleBorder(color: Blue);  // positive side

            var layout = _engine.Calculate(table);

            layout.Edges.Single(e => e.X1 == 100f && e.X2 == 100f).ColorRgb.Should().Be(Red);
        }

        // ---- combination: merged bottom border beside mixed plain top borders ----

        [Test]
        public void GivenMergedBottomBesideMixedPlainTops_WhenCalculating_ThenEachSegmentResolvesIndependently()
        {
            // Top row is a 2-column merged cell with a black (artifact) bottom border. Below it,
            // one plain cell has a white top ("no border cell") and the other a black top
            // ("full border" cell). Each column's shared segment resolves independently.
            var table = TableModelBuilder.Grid(2, 2, rowHeight: 20f, colWidth: 100f);
            table.Cells[0, 1].MergeId = table.Cells[0, 0].MergeId; // merge the top row
            table.Cells[0, 0].BorderBottom = TableModelBuilder.VisibleBorder(color: Black);
            table.Cells[0, 1].BorderBottom = TableModelBuilder.VisibleBorder(color: Black);
            table.Cells[1, 0].BorderTop = TableModelBuilder.VisibleBorder(color: White);
            table.Cells[1, 1].BorderTop = TableModelBuilder.VisibleBorder(color: Black);

            var layout = _engine.Calculate(table);

            layout.Edges.Single(e => Horizontal(e, 20f, 0f, 100f)).ColorRgb.Should().Be(White);
            layout.Edges.Single(e => Horizontal(e, 20f, 100f, 200f)).ColorRgb.Should().Be(Black);
        }

        // ---- helpers ----

        private static void MergeRightColumn(TableModel table)
        {
            table.Cells[0, 1].MergeId = 99;
            table.Cells[1, 1].MergeId = 99;
        }

        private static void SetBorderless(CellModel cell)
        {
            cell.BorderTop = new BorderModel { Visible = false };
            cell.BorderBottom = new BorderModel { Visible = false };
            cell.BorderLeft = new BorderModel { Visible = false };
            cell.BorderRight = new BorderModel { Visible = false };
        }

        private static bool Vertical(EdgePlacement e, float x, float y1, float y2)
            => e.X1 == x && e.X2 == x && e.Y1 == y1 && e.Y2 == y2;

        private static bool Horizontal(EdgePlacement e, float y, float x1, float x2)
            => e.Y1 == y && e.Y2 == y && e.X1 == x1 && e.X2 == x2;
    }

    /// <summary>Geometry helpers on <see cref="EdgePlacement"/>.</summary>
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
