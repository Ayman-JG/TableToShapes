using System;
using System.Collections.Generic;

namespace TableToShapes.Core.Model
{
    /// <summary>
    /// COM-free snapshot of a PowerPoint table. All fidelity logic operates on this
    /// model so it can be unit tested without Office running.
    /// </summary>
    public sealed class TableModel
    {
        public float Left { get; set; }
        public float Top { get; set; }
        public IReadOnlyList<float> RowHeights { get; set; }
        public IReadOnlyList<float> ColumnWidths { get; set; }

        /// <summary>Indexed [row, column], 0-based. Continuation cells of a merge share the anchor's MergeId.</summary>
        public CellModel[,] Cells { get; set; }

        public int RowCount => Cells.GetLength(0);
        public int ColumnCount => Cells.GetLength(1);
    }

    public sealed class CellModel
    {
        /// <summary>Cells belonging to the same merge share this id.</summary>
        public int MergeId { get; set; }

        public FillModel Fill { get; set; }
        public TextModel Text { get; set; }

        public BorderModel BorderTop { get; set; }
        public BorderModel BorderBottom { get; set; }
        public BorderModel BorderLeft { get; set; }
        public BorderModel BorderRight { get; set; }
    }

    public sealed class FillModel
    {
        public bool Visible { get; set; }
        public int ColorRgb { get; set; }
        public float Transparency { get; set; }
    }

    public sealed class BorderModel
    {
        public bool Visible { get; set; }
        public float Weight { get; set; }
        public int ColorRgb { get; set; }
        public int DashStyle { get; set; }
        public float Transparency { get; set; }
    }

    public sealed class TextModel
    {
        public bool HasText { get; set; }
        public float MarginLeft { get; set; }
        public float MarginRight { get; set; }
        public float MarginTop { get; set; }
        public float MarginBottom { get; set; }
        public int VerticalAnchor { get; set; }
        public bool WordWrap { get; set; }
        public IReadOnlyList<ParagraphModel> Paragraphs { get; set; } = Array.Empty<ParagraphModel>();
    }

    public sealed class ParagraphModel
    {
        public int Alignment { get; set; }
        public float SpaceBefore { get; set; }
        public float SpaceAfter { get; set; }
        public float SpaceWithin { get; set; }
        public int IndentLevel { get; set; }
        public IReadOnlyList<RunModel> Runs { get; set; } = Array.Empty<RunModel>();
    }

    public sealed class RunModel
    {
        public string Text { get; set; }
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public int UnderlineStyle { get; set; }
        public int ColorRgb { get; set; }
    }
}
