// ShapeData.cs
using System.Collections.Generic;

namespace PowerPointAddIn2
{
    public class ShapeData
    {
        // Existing properties
        public string ShapeType { get; set; }
        public int Id { get; set; }
        public string Name { get; set; }
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }

        // Margin Properties
        public double? MarginLeft { get; set; }
        public double? MarginRight { get; set; }
        public double? MarginTop { get; set; }
        public double? MarginBottom { get; set; }

        // Rotation
        public double Rotation { get; set; }

        // Fill Properties
        public string FillColor { get; set; }
        public string FillType { get; set; }

        // Line Properties
        public string LineColor { get; set; }
        public double? LineWidth { get; set; }
        public string LineDashStyle { get; set; }
        public string LineStyle { get; set; }
        public double? LineTransparency { get; set; }
        public string LineBeginArrowheadStyle { get; set; }
        public string LineEndArrowheadStyle { get; set; }

        // Font Properties
        public string FontName { get; set; }
        public double FontSize { get; set; }
        public string FontColor { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }

        // Text Alignment Properties
        public string allignment { get; set; } // Left, Center, Right, Justify
        public string VerticalAlignment { get; set; }

        // Content
        public List<ContentItem> Content { get; set; }

        // Table Properties
        public TableData Table { get; set; } // Null if not a table
    }

    public class ContentItem
    {
        public string Text { get; set; }
        public string Bullet { get; set; }
    }

    public class TableData
    {
        public int Rows { get; set; }
        public int Columns { get; set; }
        public List<TableRow> TableRows { get; set; }
    }

    public class TableRow
    {
        public List<TableCell> Cells { get; set; }
    }

    public class TableCell
    {
        public string Text { get; set; }
    }

    public class UpdatedTemplate
    {
        public List<ShapeData> Shapes { get; set; }
    }

    public class SlideData
    {
        public string Filename { get; set; }
        public double SimilarityScore { get; set; }
        public UpdatedTemplate updated_template { get; set; }
    }


    public class slideMaster
    {
       
        public string FillColor { get; set; }
        public string FontColor { get; set; }
    }
        
}

