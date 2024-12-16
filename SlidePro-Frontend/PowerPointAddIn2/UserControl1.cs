using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using Newtonsoft.Json;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading.Tasks;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Drawing.Drawing2D;
using System.Net;
using static System.Net.WebRequestMethods;
using System.Threading;
using System.Reflection;


namespace PowerPointAddIn2
{
    public partial class UserControl1 : UserControl
    {
        private List<SlideData> slides;
        private int limit = 0;
        string json_api;
        int MAX_COUNT = 10;
        string base_dir_path = "C:\\Users\\user\\PPT_Output";
        List<string> tempImagePaths = new List<string>();
        bool firstPreview = false;

        public UserControl1()
        {
            InitializeComponent();
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {
        }

        private void buttonLoadJSON_Click(object sender, EventArgs e,int index)
        {
            try
            {
                // Get the active presentation
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Slide activeSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                if (presentation == null)
                {
                    MessageBox.Show("No active presentation found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                // Add a new slide right after the active slide
                int slideIndex = activeSlide.SlideIndex;
                PowerPoint.Slide newSlide = presentation.Slides.Add(slideIndex + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

                // Extract the list of shapes from the dictionary
                List<ShapeData> shapesData = this.slides[index].updated_template.Shapes;

                // Iterate through each shapeData and create shapes accordingly
                foreach (ShapeData shapeData in shapesData)
                {
                    // Create shape based on shapeData.ShapeType
                    PowerPoint.Shape shape = null;

                    // Determine the MsoAutoShapeType or MsoShapeType based on shapeData.ShapeType
                    shape = CreateShapeFromType(newSlide, shapeData);

                    if (shape != null)
                    {
                        // Set properties
                        shape.Left = (float)shapeData.Left;
                        shape.Top = (float)shapeData.Top;
                        shape.Width = (float)shapeData.Width;
                        shape.Height = (float)shapeData.Height;
                        shape.Rotation = (float)shapeData.Rotation;
                        shape.Name = shapeData.Name;

                        // Margins
                        try
                        {
                            if (shapeData.MarginLeft.HasValue)
                            {
                                shape.TextFrame.MarginLeft = (float)shapeData.MarginLeft.Value;
                            }
                            if (shapeData.MarginRight.HasValue)
                            {
                                shape.TextFrame.MarginRight = (float)shapeData.MarginRight.Value;
                            }
                            if (shapeData.MarginTop.HasValue)
                            {
                                shape.TextFrame.MarginTop = (float)shapeData.MarginTop.Value;
                            }
                            if (shapeData.MarginBottom.HasValue)
                            {
                                shape.TextFrame.MarginBottom = (float)shapeData.MarginBottom.Value;
                            }
                        }
                        catch (Exception ex)
                        {
                            // Ignore margin setting errors
                        }

                        // Fill properties
                        try
                        {
                            if (shapeData.FillType != "none")
                            {
                                shape.Fill.Visible = Office.MsoTriState.msoTrue;
                                shape.Fill.ForeColor.RGB = ColorTranslator.FromHtml(shapeData.FillColor).ToArgb();
                                shape.Fill.Solid();
                            }
                            else
                            {
                                shape.Fill.Visible = Office.MsoTriState.msoFalse;
                            }
                        }
                        catch (Exception ex)
                        {
                            // Handle fill errors
                        }

                        // Line properties
                        try
                        {
                            if (shapeData.LineColor != "none")
                            {
                                shape.Line.Visible = Office.MsoTriState.msoTrue;
                                shape.Line.ForeColor.RGB = ColorTranslator.FromHtml(shapeData.LineColor).ToArgb();
                                shape.Line.Weight = (float)shapeData.LineWidth;
                                // Set line dash style
                                shape.Line.DashStyle = GetLineDashStyleFromString(shapeData.LineDashStyle);
                                // Set line style
                                shape.Line.Style = GetLineStyleFromString(shapeData.LineStyle);
                                shape.Line.Transparency = (float)shapeData.LineTransparency;
                                shape.Line.BeginArrowheadStyle = GetArrowheadStyleFromString(shapeData.LineBeginArrowheadStyle);
                                shape.Line.EndArrowheadStyle = GetArrowheadStyleFromString(shapeData.LineEndArrowheadStyle);
                            }
                            else
                            {
                                shape.Line.Visible = Office.MsoTriState.msoFalse;
                            }
                        }
                        catch (Exception ex)
                        {
                            // Handle line errors
                        }

                        // Font properties and content
                        if (shapeData.Content != null && shapeData.Content.Count > 0)
                        {
                            if (shape.TextFrame != null)
                            {
                                PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                                textRange.Text = String.Empty; // Reset text

                                foreach (ContentItem contentItem in shapeData.Content)
                                {
                                    if (!string.IsNullOrEmpty(contentItem.Text))
                                    {
                                        PowerPoint.TextRange newText = textRange.InsertAfter(contentItem.Text);
                                        // Set font properties for this range
                                        SetFontProperties(newText.Font, shapeData);
                                    }
                                    else if (!string.IsNullOrEmpty(contentItem.Bullet))
                                    {
                                        PowerPoint.TextRange newText = textRange.InsertAfter("\r" + contentItem.Bullet + "\r");
                                        newText.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
                                        // Set font properties for this range
                                        SetFontProperties(newText.Font, shapeData);
                                    }
                                }

                                // Set alignment
                                textRange.ParagraphFormat.Alignment = GetParagraphAlignment(shapeData.allignment);
                                shape.TextFrame2.VerticalAnchor = GetVerticalAlignment(shapeData.VerticalAlignment);
                            }
                        }

                        // Table properties
                        if (shapeData.Table != null)
                        {
                            // Since we have already created the table shape, we can populate the cells
                            if (shape.HasTable == Office.MsoTriState.msoTrue)
                            {
                                PowerPoint.Table table = shape.Table;
                                for (int i = 1; i <= table.Rows.Count; i++)
                                {
                                    for (int j = 1; j <= table.Columns.Count; j++)
                                    {
                                        PowerPoint.Cell cell = table.Cell(i, j);
                                        if (shapeData.Table.TableRows.Count >= i && shapeData.Table.TableRows[i - 1].Cells.Count >= j)
                                        {
                                            TableCell tableCellData = shapeData.Table.TableRows[i - 1].Cells[j - 1];
                                            cell.Shape.TextFrame.TextRange.Text = tableCellData.Text;
                                            // Set more properties if needed
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // Optionally, select the new slide
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(newSlide.SlideIndex);
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Helper method to create shape from shape type
        private PowerPoint.Shape CreateShapeFromType(PowerPoint.Slide slide, ShapeData shapeData)
        {
            PowerPoint.Shape shape = null;

            // First, try to map to MsoAutoShapeType
            Office.MsoAutoShapeType autoShapeType = GetAutoShapeTypeFromString(shapeData.ShapeType);



            if (autoShapeType != Office.MsoAutoShapeType.msoShapeMixed)
            {
                shape = slide.Shapes.AddShape(autoShapeType, (float)shapeData.Left, (float)shapeData.Top, (float)shapeData.Width, (float)shapeData.Height);
            }
            else
            {



                // Try other types
                if (shapeData.ShapeType == "TextBox")
                {
                    shape = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        (float)shapeData.Left,
                        (float)shapeData.Top,
                        (float)shapeData.Width,
                        (float)shapeData.Height);
                    // After creating the shape
                    SetFontProperties(shape.TextFrame.TextRange.Font, shapeData);
                    shape.TextFrame.TextRange.ParagraphFormat.Alignment = GetParagraphAlignment(shapeData.allignment);

                }
                else if (shapeData.ShapeType == "Placeholder")
                {
                    shape = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        (float)shapeData.Left,
                        (float)shapeData.Top,
                        (float)shapeData.Width,
                        (float)shapeData.Height);
                }
                else if (shapeData.ShapeType == "Line")
                {
                    shape = slide.Shapes.AddLine((float)shapeData.Left, (float)shapeData.Top, (float)(shapeData.Left + shapeData.Width), (float)(shapeData.Top + shapeData.Height));
                }
                else if (shapeData.ShapeType == "Table" && shapeData.Table != null)
                {
                    int rows = shapeData.Table.Rows;
                    int columns = shapeData.Table.Columns;
                    shape = slide.Shapes.AddTable(rows, columns, (float)shapeData.Left, (float)shapeData.Top, (float)shapeData.Width, (float)shapeData.Height);
                }
                else if (shapeData.ShapeType == "Picture")
                {

                    shape = slide.Shapes.AddPicture(base_dir_path + "\\icons\\icon.png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, (float)shapeData.Left, (float)shapeData.Top, (float)shapeData.Width, (float)shapeData.Height);
                }
                else if (shapeData.ShapeType == "Picture_freeform")
                {
                    string shape_name = shapeData.Name;
                    shape = slide.Shapes.AddPicture(base_dir_path + "\\icons\\" + shape_name, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, (float)shapeData.Left, (float)shapeData.Top, (float)shapeData.Width, (float)shapeData.Height);
                }


                // If shape type is Unkown and name contains Graphic, load an icon
                else if (shapeData.ShapeType == "Unknown" && shapeData.Name.Contains("Graphic"))
                {
                    // Load an icon
                    shape = slide.Shapes.AddPicture(base_dir_path +"\\icons\\icon.png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, (float)shapeData.Left, (float)shapeData.Top, (float)shapeData.Width, (float)shapeData.Height);
                }

                else
                {
                    // Default to rectangle
                    //shape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, (float)shapeData.Left, (float)shapeData.Top, (float)shapeData.Width, (float)shapeData.Height);
                }
            }

            return shape;
        }

        // Helper method to set font properties
        private void SetFontProperties(PowerPoint.Font font, ShapeData shapeData)
        {
            try
            {
                font.Name = shapeData.FontName;
                font.Size = (float)shapeData.FontSize;
                font.Color.RGB = ColorTranslator.FromHtml(shapeData.FontColor).ToArgb();
                font.Bold = shapeData.Bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                font.Italic = shapeData.Italic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                font.Underline = shapeData.Underline ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            }
            catch
            {

            }
        }

        // Reverse mapping functions

        private Office.MsoAutoShapeType GetAutoShapeTypeFromString(string shapeType)
        {
            foreach (Office.MsoAutoShapeType type in Enum.GetValues(typeof(Office.MsoAutoShapeType)))
            {
                if (type.ToString() == shapeType)
                {
                    return type;
                }
            }
            return Office.MsoAutoShapeType.msoShapeMixed; // Indicates not found
        }

        private Office.MsoLineDashStyle GetLineDashStyleFromString(string dashStyle)
        {
            switch (dashStyle)
            {
                case "Dash":
                    return Office.MsoLineDashStyle.msoLineDash;
                case "DashDot":
                    return Office.MsoLineDashStyle.msoLineDashDot;
                case "DashDotDot":
                    return Office.MsoLineDashStyle.msoLineDashDotDot;
                case "LongDash":
                    return Office.MsoLineDashStyle.msoLineLongDash;
                case "LongDashDot":
                    return Office.MsoLineDashStyle.msoLineLongDashDot;
                case "LongDashDotDot":
                    return Office.MsoLineDashStyle.msoLineLongDashDotDot;
                case "SquareDot":
                    return Office.MsoLineDashStyle.msoLineSquareDot;
                case "Solid":
                    return Office.MsoLineDashStyle.msoLineSolid;
                default:
                    return Office.MsoLineDashStyle.msoLineSolid;
            }
        }

        private Office.MsoLineStyle GetLineStyleFromString(string style)
        {
            switch (style)
            {
                case "Single":
                    return Office.MsoLineStyle.msoLineSingle;
                case "ThinThin":
                    return Office.MsoLineStyle.msoLineThinThin;
                case "ThinThick":
                    return Office.MsoLineStyle.msoLineThinThick;
                case "ThickThin":
                    return Office.MsoLineStyle.msoLineThickThin;
                default:
                    return Office.MsoLineStyle.msoLineSingle;
            }
        }

        private Office.MsoArrowheadStyle GetArrowheadStyleFromString(string arrowStyle)
        {
            switch (arrowStyle)
            {
                case "None":
                    return Office.MsoArrowheadStyle.msoArrowheadNone;
                case "Triangle":
                    return Office.MsoArrowheadStyle.msoArrowheadTriangle;
                case "Stealth":
                    return Office.MsoArrowheadStyle.msoArrowheadStealth;
                case "Diamond":
                    return Office.MsoArrowheadStyle.msoArrowheadDiamond;
                case "Open":
                    return Office.MsoArrowheadStyle.msoArrowheadOpen;
                case "Oval":
                    return Office.MsoArrowheadStyle.msoArrowheadOval;
                default:
                    return Office.MsoArrowheadStyle.msoArrowheadNone;
            }
        }

        private PowerPoint.PpParagraphAlignment GetParagraphAlignment(string alignment)
        {
            switch (alignment)
            {
                case "ppAlignLeft":
                    return PowerPoint.PpParagraphAlignment.ppAlignLeft;
                case "ppAlignCenter":
                    return PowerPoint.PpParagraphAlignment.ppAlignCenter;
                case "ppAlignRight":
                    return PowerPoint.PpParagraphAlignment.ppAlignRight;
                case "ppAlignJustify":
                    return PowerPoint.PpParagraphAlignment.ppAlignJustify;
                default:
                    return PowerPoint.PpParagraphAlignment.ppAlignLeft;
            }
        }

        private void pictureBoxLoader_click(object sender,EventArgs e)
        {

        }


        private void myButton_Click(object sender, EventArgs e)
        {
            Logger.Log("started --");

            this.limit = 0;
            flowLayoutPanelSlides.Controls.Clear();
   
            try
            {
                
                this.loadTaskPaneSlides();
                
            }
            //}
            catch (Exception ex)
            {
                // Handle any unexpected errors
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private string Get_PPT_data()
        {
            try
            {
                // Get the active presentation
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

                if (presentation == null)
                {
                    MessageBox.Show("No active presentation found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return "None";
                }

                // Get the active slide
                PowerPoint.Slide activeSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                if (activeSlide == null)
                {
                    MessageBox.Show("No active slide found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return "None";
                }

                // Initialize a list to collect shape data
                List<ShapeData> shapesData = new List<ShapeData>();
                List<slideMaster> masterShapesData = new List<slideMaster>();


                // Fetch shapes from the active slide
                GetShapesFromSlide(activeSlide, shapesData);

                // Get the corresponding slide master and layout for the active slide
                PowerPoint.CustomLayout slideLayout = activeSlide.CustomLayout;
                PowerPoint.Master slideMaster = slideLayout.Design.SlideMaster;

                if (slideMaster != null)
                {
                    // Fetch shapes from the slide master if available
                    foreach (PowerPoint.CustomLayout layout in slideMaster.CustomLayouts)
                    {
                        List<slideMaster> layoutShapesData = new List<slideMaster>();

                        // Fetch shapes from the layout
                       slideMaster obj =  GetShapesColorsFromSlide(layout, layoutShapesData, isMaster: true);

                        if (obj!=null)
                        {
                            masterShapesData.AddRange(layoutShapesData);
                            break;
                        }

                        // Add the shapes from the layout to the master shapes data
                        //masterShapesData.AddRange(layoutShapesData);
                    }
                }




                // Fetch shapes from the slide master corresponding to the active slide's layout
                //GetShapesFromSlide(slideMaster, shapesData, isMaster: true);

                // Serialize the list to JSON with indentation for readability
                var shapesDictionary = new
                {
                    shapes = shapesData,
                    slidemaster = masterShapesData
                };

                // Serialize the dictionary to JSON with indentation for readability
                string jsonOutput = JsonConvert.SerializeObject(shapesDictionary, Formatting.Indented);

                // Define the file path where you want to save the JSON
                string filePath = base_dir_path + "\\inputData.json"; // Change this to your desired path

                // Ensure the directory exists before writing the file
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));

                File.WriteAllText(filePath, jsonOutput);

                //MessageBox.Show($"JSON data saved successfully to {filePath}.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return jsonOutput;
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "None";
            }
        }

        private slideMaster GetShapesColorsFromSlide(dynamic slide, List<slideMaster> shapesData, bool isMaster = false)
        {
            string fillColor = null;
            string fontColor = null;
            string Name = null;

            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                // Initialize variables for colors
                object shapeType = shape.Type;
                // Try to get the fill color (only for AutoShape or Freeform shapes)
                if (shape.Name.Contains("Rectangle"))
                {
                    if (shape.Fill != null && shape.Fill.ForeColor != null && fillColor == null)
                    {
                        int foreColor = shape.Fill.ForeColor.RGB;
                        fillColor = ColorTranslator.ToHtml(Color.FromArgb(foreColor)); // Convert RGB to HTML color
                    }
                }

                // Try to get the font color (only if the shape contains text)
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue &&
                    shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue && fontColor == null)
                {
                    int fontColorRgb = shape.TextFrame.TextRange.Font.Color.RGB;
                    fontColor = ColorTranslator.ToHtml(Color.FromArgb(fontColorRgb)); // Convert RGB to HTML color
                }

                // If either the fill color or the font color is found, return the shape data
                if (fillColor != null && fontColor != null)
                {
                    
                    slideMaster shapeData = new slideMaster
                    {
                        FillColor = fillColor ?? "No Fill",  // If fill color is not available, mark as "No Fill"
                        FontColor = fontColor ?? "No Text"  // If font color is not available, mark as "No Text"
                    };

                    // Add the shape data to the list
                    shapesData.Add(shapeData);

                    // Return the slideMaster object and exit the method
                    return shapeData;
                }
            }

            // If no valid shape is found, return null (or an empty slideMaster object as needed)
            return null;
        }



        // Helper method to retrieve shapes from both slide and slide master
        private void GetShapesFromSlide(dynamic slideOrMaster, List<ShapeData> shapesData, bool isMaster = false)
        {
            // Iterate through each shape in the slide or master
            foreach (PowerPoint.Shape shape in slideOrMaster.Shapes)
            {
                ShapeData shapeData = new ShapeData();

                // Shape Type
                shapeData.ShapeType = GetShapeType(shape);

                // ID and Name
                shapeData.Id = shape.Id;
                shapeData.Name = shape.Name;

                // Position and Size
                shapeData.Left = Math.Round(shape.Left, 4);
                shapeData.Top = Math.Round(shape.Top, 4);
                shapeData.Width = Math.Round(shape.Width, 4);
                shapeData.Height = Math.Round(shape.Height, 4);

                // Margins with Error Handling
                try
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        shapeData.MarginLeft = Math.Round(shape.TextFrame.MarginLeft, 4);
                        shapeData.MarginRight = Math.Round(shape.TextFrame.MarginRight, 4);
                        shapeData.MarginTop = Math.Round(shape.TextFrame.MarginTop, 4);
                        shapeData.MarginBottom = Math.Round(shape.TextFrame.MarginBottom, 4);
                    }
                }
                catch (Exception ex)
                {
                    shapeData.MarginLeft = null;
                    shapeData.MarginRight = null;
                    shapeData.MarginTop = null;
                    shapeData.MarginBottom = null;
                }

                // Rotation
                shapeData.Rotation = Math.Round(shape.Rotation, 4);

                // Fill Properties
                try
                {
                    if (shape.Fill.Visible == Office.MsoTriState.msoTrue)
                    {
                        shapeData.FillColor = ColorTranslator.ToHtml(Color.FromArgb(shape.Fill.ForeColor.RGB));
                        shapeData.FillType = GetFillType(shape.Fill.Type);
                    }
                    else
                    {
                        shapeData.FillColor = "none";
                        shapeData.FillType = "none";
                    }
                }
                catch (Exception ex)
                {
                    shapeData.FillColor = "none";
                    shapeData.FillType = "none";
                }

                // Line Properties with Error Handling
                try
                {
                    if (shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue)
                    {
                        shapeData.LineColor = ColorTranslator.ToHtml(Color.FromArgb(shape.Line.ForeColor.RGB));
                        shapeData.LineWidth = Math.Round(shape.Line.Weight, 4);
                        shapeData.LineDashStyle = GetLineDashStyle(shape.Line.DashStyle);
                        shapeData.LineStyle = GetLineStyle(shape.Line.Style);
                        shapeData.LineTransparency = Math.Round(shape.Line.Transparency, 4);
                        shapeData.LineBeginArrowheadStyle = GetArrowheadStyle(shape.Line.BeginArrowheadStyle);
                        shapeData.LineEndArrowheadStyle = GetArrowheadStyle(shape.Line.EndArrowheadStyle);
                    }
                    else
                    {
                        shapeData.LineColor = "none";
                        shapeData.LineWidth = 0;
                        shapeData.LineDashStyle = "";
                        shapeData.LineStyle = "";
                        shapeData.LineTransparency = 0;
                        shapeData.LineBeginArrowheadStyle = "";
                        shapeData.LineEndArrowheadStyle = "";
                    }
                }
                catch (Exception ex)
                {
                    shapeData.LineColor = "none";
                    shapeData.LineWidth = 0;
                    shapeData.LineDashStyle = "";
                    shapeData.LineStyle = "";
                    shapeData.LineTransparency = 0;
                    shapeData.LineBeginArrowheadStyle = "";
                    shapeData.LineEndArrowheadStyle = "";
                }

                // Font Properties
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                    shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    try
                    {
                        PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                        PowerPoint.Font font = textRange.Font;

                        shapeData.FontName = font.Name;
                        shapeData.FontSize = Math.Round(font.Size, 2);
                        shapeData.FontColor = ColorTranslator.ToHtml(Color.FromArgb(font.Color.RGB));
                        shapeData.Bold = font.Bold == Office.MsoTriState.msoTrue;
                        shapeData.Italic = font.Italic == Office.MsoTriState.msoTrue;
                        shapeData.Underline = font.Underline == Office.MsoTriState.msoTrue;

                        // Alignment
                        shapeData.allignment = textRange.ParagraphFormat.Alignment.ToString();

                        // Content
                        shapeData.Content = new List<ContentItem>();
                        foreach (PowerPoint.TextRange paragraph in textRange.Paragraphs())
                        {
                            ContentItem contentItem = new ContentItem();
                            if (paragraph.ParagraphFormat.Bullet.Type != PowerPoint.PpBulletType.ppBulletNone)
                            {
                                string bulletText = paragraph.Text.Trim();
                                if (!string.IsNullOrEmpty(bulletText))
                                {
                                    contentItem.Bullet = bulletText;
                                }
                            }
                            else
                            {
                                string regularText = paragraph.Text.Trim();
                                if (!string.IsNullOrEmpty(regularText))
                                {
                                    contentItem.Text = regularText;
                                }
                            }
                            if (!string.IsNullOrEmpty(contentItem.Bullet) || !string.IsNullOrEmpty(contentItem.Text))
                            {
                                shapeData.Content.Add(contentItem);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        shapeData.FontName = "";
                        shapeData.FontSize = 0;
                        shapeData.FontColor = "";
                        shapeData.Bold = false;
                        shapeData.Italic = false;
                        shapeData.Underline = false;
                        shapeData.allignment = "";
                        shapeData.Content = new List<ContentItem>();
                    }
                }
                else
                {
                    shapeData.FontName = "";
                    shapeData.FontSize = 0;
                    shapeData.FontColor = "";
                    shapeData.Bold = false;
                    shapeData.Italic = false;
                    shapeData.Underline = false;
                    shapeData.allignment = "";
                    shapeData.Content = new List<ContentItem>();
                }

                // Table Properties
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    try
                    {
                        PowerPoint.Table table = shape.Table;
                        TableData tableData = new TableData
                        {
                            Rows = table.Rows.Count,
                            Columns = table.Columns.Count,
                            TableRows = new List<TableRow>()
                        };

                        for (int i = 1; i <= table.Rows.Count; i++)
                        {
                            TableRow tableRow = new TableRow
                            {
                                Cells = new List<TableCell>()
                            };

                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                PowerPoint.Cell cell = table.Cell(i, j);
                                string cellText = "";

                                if (cell.Shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                {
                                    cellText = cell.Shape.TextFrame.TextRange.Text.Trim();
                                }

                                TableCell tableCell = new TableCell
                                {
                                    Text = cellText
                                };

                                tableRow.Cells.Add(tableCell);
                            }

                            tableData.TableRows.Add(tableRow);
                        }

                        shapeData.Table = tableData;
                    }
                    catch (Exception ex)
                    {
                        shapeData.Table = null;
                    }
                }
                else
                {
                    shapeData.Table = null;
                }

                // Add the shape data to the list
                shapesData.Add(shapeData);
            }
        }




        private Microsoft.Office.Core.MsoVerticalAnchor GetVerticalAlignment(string verticalAlignment)
        {
            switch (verticalAlignment)
            {
                case "msoAnchorTop":
                    return Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop;
                case "msoAnchorMiddle":
                    return Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                case "msoAnchorBottom":
                    return Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorBottom;
                default:
                    return Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle; // Default to top
            }
        }

        

        private async Task<string> getData(string inputSlide)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            // FastAPI endpoint URL
            string apiUrl = "https://3598-2407-d000-b-b41a-3df2-306d-a6c9-690f.ngrok-free.app/recommendation";
            // Initialize HttpClient
            var handler = new HttpClientHandler();
            handler.ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator;

            using (HttpClient client = new HttpClient(handler))

            {
                try
                {
                    // Convert the inputSlide string to a JSON content object
                    var content = new StringContent(inputSlide, Encoding.UTF8, "application/json");
                    Logger.Log("Calling API with data ---");
                    //Logger.Log(inputSlide); // Log the input string instead of the content object

                    // Make the POST request
                    HttpResponseMessage response = await client.PostAsync(apiUrl, content);
                    Logger.Log($"Response Status Code: {response.StatusCode}"); // Log the status code

                    // Check if the response is successful
                    if (response.IsSuccessStatusCode)
                    {
                        // Read the response content
                        string result = await response.Content.ReadAsStringAsync();
                        Logger.Log("API Response:");
                        Logger.Log(result);

                        return result; // Return the response (recommendations)
                    }
                    else
                    {
                        // Log or handle the error response
                        string errorContent = await response.Content.ReadAsStringAsync();
                        Logger.Log($"Error Content: {errorContent}");
                        return $"Error: {response.StatusCode} - {response.ReasonPhrase}";
                    }
                }
                catch (Exception ex)
                {
                    // Log or handle the exception
                    Logger.Log($"Exception: {ex.Message}");
                    if (ex.InnerException != null)
                    {
                        Logger.Log($"Inner Exception: {ex.InnerException.Message}");
                    }
                    Logger.Log($"Stack Trace: {ex.StackTrace}");
                    return $"Exception: {ex.Message}";
                }
            }
        }





        #region Helper Methods

        private string GetShapeType(PowerPoint.Shape shape)
        {
            switch (shape.Type)
            {
                case Office.MsoShapeType.msoAutoShape:
                    return shape.AutoShapeType.ToString();
                case Office.MsoShapeType.msoCallout:
                    return "Callout";
                case Office.MsoShapeType.msoChart:
                    return "Chart";
                case Office.MsoShapeType.msoComment:
                    return "Comment";
                case Office.MsoShapeType.msoFreeform:
                    return "Freeform";
                case Office.MsoShapeType.msoGroup:
                    return "Group";
                case Office.MsoShapeType.msoEmbeddedOLEObject:
                    return "EmbeddedOLEObject";
                case Office.MsoShapeType.msoFormControl:
                    return "FormControl";
                case Office.MsoShapeType.msoLine:
                    return "Line";
                case Office.MsoShapeType.msoLinkedOLEObject:
                    return "LinkedOLEObject";
                case Office.MsoShapeType.msoLinkedPicture:
                    return "LinkedPicture";
                case Office.MsoShapeType.msoMedia:
                    return "Media";
                case Office.MsoShapeType.msoOLEControlObject:
                    return "OLEControlObject";
                case Office.MsoShapeType.msoPicture:
                    return "Picture";
                case Office.MsoShapeType.msoPlaceholder:
                    return "Placeholder";
                case Office.MsoShapeType.msoTextBox:
                    return "TextBox";
                default:
                    return "Unknown";
            }
        }

        private string GetFillType(Office.MsoFillType fillType)
        {
            switch (fillType)
            {
                case Office.MsoFillType.msoFillBackground:
                    return "Background";
                case Office.MsoFillType.msoFillGradient:
                    return "Gradient";
                case Office.MsoFillType.msoFillPatterned:
                    return "Patterned";
                case Office.MsoFillType.msoFillPicture:
                    return "Picture";
                case Office.MsoFillType.msoFillSolid:
                    return "Solid";
                case Office.MsoFillType.msoFillTextured:
                    return "Textured";
                default:
                    return "None";
            }
        }

        private string GetLineDashStyle(Office.MsoLineDashStyle dashStyle)
        {
            switch (dashStyle)
            {
                case Office.MsoLineDashStyle.msoLineDash:
                    return "Dash";
                case Office.MsoLineDashStyle.msoLineDashDot:
                    return "DashDot";
                case Office.MsoLineDashStyle.msoLineDashDotDot:
                    return "DashDotDot";
                case Office.MsoLineDashStyle.msoLineLongDash:
                    return "LongDash";
                case Office.MsoLineDashStyle.msoLineLongDashDot:
                    return "LongDashDot";
                case Office.MsoLineDashStyle.msoLineLongDashDotDot:
                    return "LongDashDotDot";
                case Office.MsoLineDashStyle.msoLineSquareDot:
                    return "SquareDot";
                case Office.MsoLineDashStyle.msoLineSolid:
                    return "Solid";
                default:
                    return "Unknown";
            }
        }

        private string GetLineStyle(Office.MsoLineStyle style)
        {
            switch (style)
            {
                case Office.MsoLineStyle.msoLineSingle:
                    return "Single";
                case Office.MsoLineStyle.msoLineThinThin:
                    return "ThinThin";
                case Office.MsoLineStyle.msoLineThinThick:
                    return "ThinThick";
                case Office.MsoLineStyle.msoLineThickThin:
                    return "ThickThin";
                default:
                    return "Unknown";
            }
        }

        private string GetArrowheadStyle(Office.MsoArrowheadStyle arrowStyle)
        {
            switch (arrowStyle)
            {
                case Office.MsoArrowheadStyle.msoArrowheadNone:
                    return "None";
                case Office.MsoArrowheadStyle.msoArrowheadTriangle:
                    return "Triangle";
                case Office.MsoArrowheadStyle.msoArrowheadStealth:
                    return "Stealth";
                case Office.MsoArrowheadStyle.msoArrowheadDiamond:
                    return "Diamond";
                case Office.MsoArrowheadStyle.msoArrowheadOpen:
                    return "Open";
                case Office.MsoArrowheadStyle.msoArrowheadOval:
                    return "Oval";
                default:
                    return "Unknown";
            }
        }

        private void PreviewSlideFromJson22()
        {

            string jsonInput = this.json_api;

            //this.tempImagePaths.Clear();

            PowerPoint.Presentation tempPresentation = null;

            try
            {

                //var slidesList = JsonConvert.DeserializeObject<List<SlideData>>(jsonInput);
                //this.slides = slidesList;

                //// Check if we have any slides
                //if (slidesList == null || slidesList.Count == 0)
                //{
                //    MessageBox.Show("No slides found in the input data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    return;
                //}

                if (this.limit >= this.MAX_COUNT)
                {
                    MessageBox.Show("No More Slides Available", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Create a temporary presentation
                tempPresentation = Globals.ThisAddIn.Application.Presentations.Add(Office.MsoTriState.msoFalse);
                int slide_index_temp = 0;
                int imagesToDisplay = 8;
                // Iterate over each SlideData object in the list
                for (int slideIndex = this.limit; slideIndex < this.limit + imagesToDisplay; slideIndex++)
                {
                    SlideData slideData = this.slides[slideIndex];
                    List<ShapeData> shapesData = slideData.updated_template.Shapes;

                    // Add a blank slide for each SlideData object
                    PowerPoint.Slide tempSlide = tempPresentation.Slides.Add(slide_index_temp + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    slide_index_temp++;

                    // Iterate through each shapeData and create shapes accordingly
                    foreach (ShapeData shapeData in shapesData)
                    {
                        PowerPoint.Shape shape = CreateShapeFromType(tempSlide, shapeData);

                        if (shape != null)
                        {
                            // Set shape properties
                            shape.Left = (float)shapeData.Left;
                            shape.Top = (float)shapeData.Top;
                            shape.Width = (float)shapeData.Width;
                            shape.Height = (float)shapeData.Height;
                            shape.Rotation = (float)shapeData.Rotation;
                            shape.Name = shapeData.Name;

                            // Handle margins
                            try
                            {
                                if (shapeData.MarginLeft.HasValue)
                                    shape.TextFrame.MarginLeft = (float)shapeData.MarginLeft.Value;
                                if (shapeData.MarginRight.HasValue)
                                    shape.TextFrame.MarginRight = (float)shapeData.MarginRight.Value;
                                if (shapeData.MarginTop.HasValue)
                                    shape.TextFrame.MarginTop = (float)shapeData.MarginTop.Value;
                                if (shapeData.MarginBottom.HasValue)
                                    shape.TextFrame.MarginBottom = (float)shapeData.MarginBottom.Value;
                            }
                            catch { /* Ignore margin errors */ }

                            // Handle fill properties
                            try
                            {
                                if (shapeData.FillType != "none")
                                {
                                    shape.Fill.Visible = Office.MsoTriState.msoTrue;
                                    shape.Fill.ForeColor.RGB = ColorTranslator.FromHtml(shapeData.FillColor).ToArgb();
                                    shape.Fill.Solid();
                                }
                                else
                                {
                                    shape.Fill.Visible = Office.MsoTriState.msoFalse;
                                }
                            }
                            catch { /* Ignore fill errors */ }

                            // Handle line properties
                            try
                            {
                                if (shapeData.LineColor != "none")
                                {
                                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                                    shape.Line.ForeColor.RGB = ColorTranslator.FromHtml(shapeData.LineColor).ToArgb();
                                    shape.Line.Weight = (float)shapeData.LineWidth;
                                    shape.Line.DashStyle = GetLineDashStyleFromString(shapeData.LineDashStyle);
                                    shape.Line.Style = GetLineStyleFromString(shapeData.LineStyle);
                                    shape.Line.Transparency = (float)shapeData.LineTransparency;
                                    shape.Line.BeginArrowheadStyle = GetArrowheadStyleFromString(shapeData.LineBeginArrowheadStyle);
                                    shape.Line.EndArrowheadStyle = GetArrowheadStyleFromString(shapeData.LineEndArrowheadStyle);
                                }
                                else
                                {
                                    shape.Line.Visible = Office.MsoTriState.msoFalse;
                                }
                            }
                            catch { /* Ignore line errors */ }

                            // Handle text and content
                            if (shapeData.Content != null && shapeData.Content.Count > 0)
                            {
                                PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                                textRange.Text = String.Empty; // Clear existing text

                                foreach (ContentItem contentItem in shapeData.Content)
                                {
                                    if (!string.IsNullOrEmpty(contentItem.Text))
                                    {
                                        PowerPoint.TextRange newText = textRange.InsertAfter(contentItem.Text);
                                        SetFontProperties(newText.Font, shapeData);
                                    }
                                    else if (!string.IsNullOrEmpty(contentItem.Bullet))
                                    {
                                        PowerPoint.TextRange newText = textRange.InsertAfter("\r" + contentItem.Bullet);
                                        newText.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
                                        SetFontProperties(newText.Font, shapeData);
                                    }
                                }

                                // Set alignment
                                textRange.ParagraphFormat.Alignment = GetParagraphAlignment(shapeData.allignment);
                                shape.TextFrame2.VerticalAnchor = GetVerticalAlignment(shapeData.VerticalAlignment);
                            }

                            // Handle table shapes
                            if (shapeData.Table != null && shape.HasTable == Office.MsoTriState.msoTrue)
                            {
                                PowerPoint.Table table = shape.Table;
                                for (int i = 1; i <= table.Rows.Count; i++)
                                {
                                    for (int j = 1; j <= table.Columns.Count; j++)
                                    {
                                        PowerPoint.Cell cell = table.Cell(i, j);
                                        if (shapeData.Table.TableRows.Count >= i && shapeData.Table.TableRows[i - 1].Cells.Count >= j)
                                        {
                                            TableCell tableCellData = shapeData.Table.TableRows[i - 1].Cells[j - 1];
                                            cell.Shape.TextFrame.TextRange.Text = tableCellData.Text;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    string imagePath = GenerateUniqueFileName(".png");

                    // Export the slide as an image
                    tempSlide.Export(imagePath, "PNG", 485, 250);

                    if (!System.IO.File.Exists(imagePath))
                    {
                        MessageBox.Show($"Failed to export slide to image at {imagePath}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        continue;
                    }

                    // Store the image path for later display
                    this.tempImagePaths.Add(imagePath);

                }
                DisplaySlidesInFlowLayoutPanel();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during preview: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (tempPresentation != null)
                {
                    tempPresentation.Close();
                    Marshal.ReleaseComObject(tempPresentation);
                }
            }
        }


        private async void PreviewSlideFromJson1()
        {
            string jsonInput = this.json_api;
            

            PowerPoint.Presentation tempPresentation = null;

            try
            {
                var slidesList = JsonConvert.DeserializeObject<List<SlideData>>(jsonInput);
                this.slides = slidesList;
                this.MAX_COUNT = slidesList.Count;

                // Check if we have any slides
                if (slidesList == null || slidesList.Count == 0)
                {
                    MessageBox.Show("No slides found in the input data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (this.limit >= this.MAX_COUNT)
                {
                    MessageBox.Show("No More Slides Available", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Create a temporary presentation
                tempPresentation = Globals.ThisAddIn.Application.Presentations.Add(Office.MsoTriState.msoFalse);

                // Prepare tasks for exporting slides
                List<Task> exportTasks = new List<Task>();

                // Iterate over each SlideData object in the list
                for (int slideIndex = this.limit,i=0; slideIndex < this.limit+5 && slideIndex < this.MAX_COUNT; slideIndex++,i++)
                {
                    SlideData slideData = slidesList[slideIndex];
                    List<ShapeData> shapesData = slideData.updated_template.Shapes;

                    // Add a blank slide for each SlideData object
                    PowerPoint.Slide tempSlide = tempPresentation.Slides.Add(i + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    CreateShapesForSlide(tempSlide, shapesData);

                    // Prepare image export
                    string imagePath = GenerateUniqueFileName(".jpg");
                    // Capture slide index for later use
                    int currentIndex = slideIndex;

                    // Create a task for exporting the slide and displaying it
                    exportTasks.Add(Task.Run(() =>
                    {
                        ExportSlide(tempSlide, imagePath, currentIndex);
                    }));

 
                }

                // Await completion of all exports
                    await Task.WhenAll(exportTasks);
                    this.limit += 5;

                    // Update button visibility safely on the UI thread
                    if (this.limit <= this.MAX_COUNT)
                    {
                        this.Invoke((Action)(() =>
                        {
                            buttonLoadMore.Visible = true;
                        }));
                    }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred during preview: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (tempPresentation != null)
                {
                    //tempPresentation.Close();
                    Marshal.ReleaseComObject(tempPresentation);
                }
            }
        }


       

        private void CreateShapesForSlide(PowerPoint.Slide slide, List<ShapeData> shapesData)
        {
            foreach (ShapeData shapeData in shapesData)
            {
                PowerPoint.Shape shape = CreateShapeFromType(slide, shapeData);
                if (shape != null)
                {
                    // Set shape properties
                    shape.Left = (float)shapeData.Left;
                    shape.Top = (float)shapeData.Top;
                    shape.Width = (float)shapeData.Width;
                    shape.Height = (float)shapeData.Height;
                    shape.Rotation = (float)shapeData.Rotation;
                    shape.Name = shapeData.Name;

                    // Handle margins
                    try
                    {
                        if (shapeData.MarginLeft.HasValue)
                            shape.TextFrame.MarginLeft = (float)shapeData.MarginLeft.Value;
                        if (shapeData.MarginRight.HasValue)
                            shape.TextFrame.MarginRight = (float)shapeData.MarginRight.Value;
                        if (shapeData.MarginTop.HasValue)
                            shape.TextFrame.MarginTop = (float)shapeData.MarginTop.Value;
                        if (shapeData.MarginBottom.HasValue)
                            shape.TextFrame.MarginBottom = (float)shapeData.MarginBottom.Value;
                    }
                    catch { /* Ignore margin errors */ }

                    // Handle fill properties
                    try
                    {
                        if (shapeData.FillType != "none")
                        {
                            shape.Fill.Visible = Office.MsoTriState.msoTrue;
                            shape.Fill.ForeColor.RGB = ColorTranslator.FromHtml(shapeData.FillColor).ToArgb();
                            shape.Fill.Solid();
                        }
                        else
                        {
                            shape.Fill.Visible = Office.MsoTriState.msoFalse;
                        }
                    }
                    catch { /* Ignore fill errors */ }

                    // Handle line properties
                    try
                    {
                        if (shapeData.LineColor != "none")
                        {
                            shape.Line.Visible = Office.MsoTriState.msoTrue;
                            shape.Line.ForeColor.RGB = ColorTranslator.FromHtml(shapeData.LineColor).ToArgb();
                            shape.Line.Weight = (float)shapeData.LineWidth;
                            shape.Line.DashStyle = GetLineDashStyleFromString(shapeData.LineDashStyle);
                            shape.Line.Style = GetLineStyleFromString(shapeData.LineStyle);
                            shape.Line.Transparency = (float)shapeData.LineTransparency;
                            shape.Line.BeginArrowheadStyle = GetArrowheadStyleFromString(shapeData.LineBeginArrowheadStyle);
                            shape.Line.EndArrowheadStyle = GetArrowheadStyleFromString(shapeData.LineEndArrowheadStyle);
                        }
                        else
                        {
                            shape.Line.Visible = Office.MsoTriState.msoFalse;
                        }
                    }
                    catch { /* Ignore line errors */ }

                    // Handle text and content
                    if (shapeData.Content != null && shapeData.Content.Count > 0)
                    {
                        PowerPoint.TextRange textRange = shape.TextFrame.TextRange;
                        textRange.Text = String.Empty; // Clear existing text

                        foreach (ContentItem contentItem in shapeData.Content)
                        {
                            if (!string.IsNullOrEmpty(contentItem.Text))
                            {
                                PowerPoint.TextRange newText = textRange.InsertAfter(contentItem.Text);
                                SetFontProperties(newText.Font, shapeData);
                            }
                            else if (!string.IsNullOrEmpty(contentItem.Bullet))
                            {
                                PowerPoint.TextRange newText = textRange.InsertAfter("\r" + contentItem.Bullet);
                                newText.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
                                SetFontProperties(newText.Font, shapeData);
                            }
                        }

                        // Set alignment
                        textRange.ParagraphFormat.Alignment = GetParagraphAlignment(shapeData.allignment);
                        shape.TextFrame2.VerticalAnchor = GetVerticalAlignment(shapeData.VerticalAlignment);
                    }

                    // Handle table shapes
                    if (shapeData.Table != null && shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        PowerPoint.Table table = shape.Table;
                        for (int i = 1; i <= table.Rows.Count; i++)
                        {
                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                PowerPoint.Cell cell = table.Cell(i, j);
                                if (shapeData.Table.TableRows.Count >= i && shapeData.Table.TableRows[i - 1].Cells.Count >= j)
                                {
                                    TableCell tableCellData = shapeData.Table.TableRows[i - 1].Cells[j - 1];
                                    cell.Shape.TextFrame.TextRange.Text = tableCellData.Text;

                                    // Handle any additional table cell properties here if needed
                                }
                            }
                        }
                    }
                }
            }
        }

        private async Task ExportSlidesAsync(PowerPoint.Slide[] slides, string[] imagePaths, int[] slideIndices)
        {
            try
            {
                List<Task> exportTasks = new List<Task>();

                for (int i = 0; i < slides.Length; i++)
                {
                    PowerPoint.Slide slide = slides[i];
                    string imagePath = imagePaths[i];
                    int slideIndex = slideIndices[i];

                    if (slide == null)
                    {
                        MessageBox.Show($"Slide is null for index: {slideIndex}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        continue;
                    }

                    // Perform the export asynchronously to avoid blocking the main thread
                    Task exportTask = Task.Run(() =>
                    {
                        slide.Export(imagePath, "JPG", 280, 160);

                        if (System.IO.File.Exists(imagePath))
                        {
                            this.tempImagePaths.Add(imagePath);

                            // Since UI updates must be done on the UI thread, use Invoke
                            flowLayoutPanelSlides.Invoke(new Action(() =>
                            {
                                DisplaySlideImage(imagePath, slideIndex);
                            }));
                        }
                        else
                        {
                            MessageBox.Show($"Failed to export slide to image at {imagePath}. File does not exist after export.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    });

                    // Add the task to the list
                    exportTasks.Add(exportTask);
                }

                // Wait for all exports to complete
                await Task.WhenAll(exportTasks);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while exporting slides: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void ExportSlide(PowerPoint.Slide slide, string imagePath, int slideIndex)
        {
            try
            {
                // Debug: Check if the slide is valid
                if (slide == null)
                {
                    MessageBox.Show($"Slide is null for index: {slideIndex}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Debug: Show the image path being used for export
                //MessageBox.Show($"Exporting slide to: {imagePath}", "Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Attempt to export the slide
                slide.Export(imagePath, "JPG", 280, 160);

                // Check if the file was created successfully
                if (System.IO.File.Exists(imagePath))
                {
                    this.tempImagePaths.Add(imagePath);
                    // Display the slide immediately after export
                    DisplaySlideImage(imagePath, slideIndex);
                    if (firstPreview)
                    {
                        firstPreview = false;
                        HideLoader();
                    }
                }
                else
                {
                    MessageBox.Show($"Failed to export slide to image at {imagePath}. File does not exist after export.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                // Log the exception message for debugging
                MessageBox.Show($"Error exporting slide: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void DisplaySlideImage(string imagePath, int slideIndex)
        {
            // Use Invoke if required (for thread safety)
            if (flowLayoutPanelSlides.InvokeRequired)
            {
                flowLayoutPanelSlides.Invoke(new Action(() => DisplaySlideImage(imagePath, slideIndex)));
                return;
            }

            // Check if the file exists
            if (System.IO.File.Exists(imagePath))
            {
                // Create and configure the PictureBox
                PictureBox pictureBox = new PictureBox
                {
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Width = 280,  // Adjust width
                    Height = 160, // Adjust height
                    Margin = new Padding(10),
                    Tag = slideIndex,
                    Cursor = Cursors.Hand // Set the cursor to a hand for better interactivity
                };

                // Load the image asynchronously
                var imgTask = LoadImageAsync(imagePath, pictureBox.Width, pictureBox.Height);
                imgTask.ContinueWith(task =>
                {
                    if (task.Result != null)
                    {
                        pictureBox.Image = task.Result; // Set the image directly

                        // Create rounded corners for the PictureBox
                        var path = CreateRoundedRegion(pictureBox.Width, pictureBox.Height, 20);
                        pictureBox.Region = new Region(path);

                        // Add a tooltip
                        ToolTip toolTip = new ToolTip();
                        toolTip.SetToolTip(pictureBox, $"Slide Preview {slideIndex + 1}");

                        // Add hover effect
                        AddHoverEffect(pictureBox);

                        // Add click event to load slide
                        pictureBox.Click += (sender, e) => this.buttonLoadJSON_Click(sender, e, slideIndex);

                        // Add the PictureBox to the flowLayoutPanel on the UI thread
                        flowLayoutPanelSlides.Invoke(new Action(() =>
                        {
                            flowLayoutPanelSlides.Controls.Add(pictureBox);
                            flowLayoutPanelSlides.PerformLayout(); // Ensure the layout updates
                        }));
                    }
                    else
                    {
                        MessageBox.Show($"Could not load image: {imagePath}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                });
            }
            else
            {
                MessageBox.Show($"Image file not found: {imagePath}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }




        private string GenerateUniqueFileName(string extension)
        {
            string randomFileName = Guid.NewGuid().ToString();
            return Path.Combine(Path.GetTempPath(), randomFileName + extension);
        }

        // Helper method to display all slides in PictureBox controls or a UI container
        private async void DisplaySlidesInFlowLayoutPanel()
        {
            // Check if the limit has reached or exceeded the available images
            if (this.limit >= this.MAX_COUNT)
            {
                MessageBox.Show("No More Slides Available", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Use Invoke if required (for thread safety in case it's called from a non-UI thread)
            if (flowLayoutPanelSlides.InvokeRequired)
            {
                flowLayoutPanelSlides.Invoke(new Action(() => DisplaySlidesInFlowLayoutPanel()));
                return;
            }

            // Determine how many images can still be displayed without exceeding the bounds of the list
            int imagesToDisplay = 1; // Adjust this as per your requirement
            for (int i = this.limit; i < (this.limit + imagesToDisplay) && i < this.tempImagePaths.Count; i++)
            {
                string imagePath = this.tempImagePaths[i];

                // Check if the file exists
                if (System.IO.File.Exists(imagePath))
                {
                    // Create and configure the PictureBox
                    PictureBox pictureBox = new PictureBox
                    {
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Width = 280,  // Adjust width
                        Height = 160, // Adjust height
                        Margin = new Padding(10),
                        Tag = i,
                        Cursor = Cursors.Hand // Set the cursor to a hand for better interactivity
                    };

                    // Load the image asynchronously
                    Image img = await LoadImageAsync(imagePath, pictureBox.Width, pictureBox.Height);
                    if (img != null)
                    {
                        // Create rounded corners for the PictureBox
                        var path = CreateRoundedRegion(pictureBox.Width, pictureBox.Height, 20); // 20 is the corner radius
                        pictureBox.Region = new Region(path);

                        // Custom Paint event to manually draw the image and the border
                        pictureBox.Paint += (sender, e) =>
                        {
                            DrawRoundedImage(e.Graphics, img, pictureBox);
                            DrawBorder(e.Graphics, pictureBox);
                        };

                        // Add a tooltip
                        ToolTip toolTip = new ToolTip();
                        toolTip.SetToolTip(pictureBox, $"Slide Preview {i + 1}");

                        // Hover effect
                        AddHoverEffect(pictureBox);

                        // Add click event to load slide
                        pictureBox.Click += (sender, e) => this.buttonLoadJSON_Click(sender, e, (int)((PictureBox)sender).Tag);

                        // Add the PictureBox to the flowLayoutPanel
                        flowLayoutPanelSlides.Controls.Add(pictureBox);

                        // Allow the UI to refresh between adding images
                        await Task.Delay(100); // Increase delay to ensure the UI updates

                        // Force the layout to update
                        flowLayoutPanelSlides.PerformLayout();
                        flowLayoutPanelSlides.Refresh();
                    }
                    else
                    {
                        MessageBox.Show($"Could not load image: {imagePath}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show($"Image file not found: {imagePath}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Increment the limit to load the next set of slides in the future
            this.limit += imagesToDisplay;

            // Check if there are no more images left to display
            if (this.limit >= this.MAX_COUNT)
            {
                MessageBox.Show("No More Slides Available", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        // Asynchronously load an image and resize it
        private async Task<Image> LoadImageAsync(string path, int width, int height)
        {
            return await Task.Run(() =>
            {
                using (var originalImage = Image.FromFile(path))
                {
                    return ResizeImage(originalImage, width, height);
                }
            });
        }

        // Create rounded corners region
        private GraphicsPath CreateRoundedRegion(int width, int height, int cornerRadius)
        {
            var path = new GraphicsPath();
            path.AddArc(0, 0, cornerRadius, cornerRadius, 180, 90);
            path.AddArc(width - cornerRadius, 0, cornerRadius, cornerRadius, 270, 90);
            path.AddArc(width - cornerRadius, height - cornerRadius, cornerRadius, cornerRadius, 0, 90);
            path.AddArc(0, height - cornerRadius, cornerRadius, cornerRadius, 90, 90);
            path.CloseAllFigures();
            return path;
        }

        // Draw the rounded image and border
        private void DrawRoundedImage(Graphics g, Image img, PictureBox pictureBox)
        {
            var imagePath = CreateRoundedRegion(pictureBox.Width, pictureBox.Height, 20); // Ensure the corner radius matches your design
            g.SetClip(imagePath); // Set the clipping region
            g.DrawImage(img, pictureBox.ClientRectangle); // Draw the image within the clipped region
            g.ResetClip(); // Reset the clipping region to allow for other drawings
        }

        // Draw the border
        private void DrawBorder(Graphics g, PictureBox pictureBox)
        {
            using (Pen borderPen = new Pen(Color.White, 1))
            {
                var borderPath = CreateRoundedRegion(pictureBox.Width, pictureBox.Height, 20);
                g.DrawPath(borderPen, borderPath);
            }
        }

        // Add hover effect
        private void AddHoverEffect(PictureBox pictureBox)
        {
            bool isHovered = false;

            pictureBox.MouseEnter += (sender, e) =>
            {
                isHovered = true;
                pictureBox.Invalidate(); // Force repaint with hover state
            };

            pictureBox.MouseLeave += (sender, e) =>
            {
                isHovered = false;
                pictureBox.Invalidate(); // Force repaint to remove hover effect
            };

            // Custom Paint event with hover effect handling
            pictureBox.Paint += (sender, e) =>
            {
                // Draw the rounded image
                DrawRoundedImage(e.Graphics, pictureBox.Image, pictureBox);

                // Use hover effect variables to draw border
                DrawBorder(e.Graphics, pictureBox, isHovered ? 10 : 1, isHovered ? Color.FromArgb(228, 108, 92) : Color.White);
            };
        }


        private void DrawBorder(Graphics graphics, PictureBox pictureBox, int borderWidth, Color borderColor)
        {
            using (Pen pen = new Pen(borderColor, borderWidth))
            {
                Rectangle rect = new Rectangle(0, 0, pictureBox.Width - 1, pictureBox.Height - 1);
                graphics.DrawRectangle(pen, rect);
            }
        }






        // Function to resize the image to the desired dimensions (to improve performance)
        private Image ResizeImage(Image image, int width, int height)
        {
            var resized = new Bitmap(width, height);
            using (var g = Graphics.FromImage(resized))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(image, 0, 0, width, height);
            }
            return resized;
        }





        private void ApplyShapeProperties(PowerPoint.Shape shape, ShapeData shapeData)
        {
            // Set properties
            shape.Left = (float)shapeData.Left;
            shape.Top = (float)shapeData.Top;
            shape.Width = (float)shapeData.Width;
            shape.Height = (float)shapeData.Height;
            shape.Rotation = (float)shapeData.Rotation;
            shape.Name = shapeData.Name;

            // Margins
            try
            {
                if (shapeData.MarginLeft.HasValue)
                {
                    shape.TextFrame.MarginLeft = (float)shapeData.MarginLeft.Value;
                }
                if (shapeData.MarginRight.HasValue)
                {
                    shape.TextFrame.MarginRight = (float)shapeData.MarginRight.Value;
                }
                if (shapeData.MarginTop.HasValue)
                {
                    shape.TextFrame.MarginTop = (float)shapeData.MarginTop.Value;
                }
                if (shapeData.MarginBottom.HasValue)
                {
                    shape.TextFrame.MarginBottom = (float)shapeData.MarginBottom.Value;
                }
            }
            catch { /* Ignore margin errors */ }

            // Fill properties
            // ... (existing code)

            // Line properties
            // ... (existing code)

            // Font properties and content
            // ... (existing code)

            // Table properties
            // ... (existing code)
        }

        #endregion

       
        private void ShowLoader()
        {
            if (pictureBoxLoader.InvokeRequired)
            {
                pictureBoxLoader.Invoke(new Action(() => pictureBoxLoader.Visible = true));
                pictureBoxLoader.BringToFront();
            }
            if (buttonLoadMore.InvokeRequired)
            {
                buttonLoadMore.Invoke(new Action(() => buttonLoadMore.Visible = true));
            }
            else
            {
                pictureBoxLoader.Visible = true;
                buttonLoadMore.Visible = true;
                pictureBoxLoader.BringToFront();
            }
        }

        private void HideLoader()
        {
            if (pictureBoxLoader.InvokeRequired)
            {
                pictureBoxLoader.Invoke(new Action(() => pictureBoxLoader.Visible = false));
                buttonLoadMore.Invoke(new Action(() => buttonLoadMore.Visible = true));
            }
            else
            {
                pictureBoxLoader.Visible = false;
                buttonLoadMore.Visible = true;
            }
        }

        private void ClearFlowLayoutPanel()
        {
            // Ensure you're accessing the panel from the UI thread
            if (flowLayoutPanelSlides.InvokeRequired)
            {
                flowLayoutPanelSlides.Invoke(new Action(() => flowLayoutPanelSlides.Controls.Clear()));
                flowLayoutPanelSlides.Controls.Add(pictureBoxLoader);
            }
            else
            {
                flowLayoutPanelSlides.Controls.Clear();
                flowLayoutPanelSlides.Controls.Add(pictureBoxLoader);
            }
        }


        private async void loadTaskPaneSlides()
        {
            try
            {
                ClearFlowLayoutPanel();
                this.tempImagePaths.Clear();
                ShowLoader();
                this.firstPreview = true;
                this.limit = 0;

                string Input_slide = Get_PPT_data();
                string jsonOutput = await getData(Input_slide);
                this.json_api = jsonOutput;

                var slidesList = JsonConvert.DeserializeObject<List<SlideData>>(jsonOutput);

                if (slidesList == null || slidesList.Count == 0)
                {
                    MessageBox.Show("No slides found in the input data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                this.slides = slidesList;
                this.MAX_COUNT = slidesList.Count;

                Logger.Log("Preview Begin");
                PreviewSlideFromJson1();
                // HideLoader();
                Logger.Log("Preview End");

                await PreviewSlideFromJson1(0);
                Logger.Log("Preview End");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while loading JSON: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void pictureBoxPreview_Click(object sender, EventArgs e)
        {
           // buttonLoadJSON_Click(sender, e);
        }

        private void flowLayoutPanelSlides_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void loadMore (object sender, EventArgs e)
        {
            //this.DisplaySlidesInFlowLayoutPanel();
            this.PreviewSlideFromJson1();
            //this.Get_PPT_data();
        }
    }
}
