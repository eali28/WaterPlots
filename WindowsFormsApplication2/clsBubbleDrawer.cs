using System;

using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Windows.Forms.DataVisualization.Charting;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public class clsBubbleDrawer
    {
        /// <summary>
        /// Draws a Bubble Diagram on the main chart plotting area.
        /// </summary>
        public static void DrawBubbleDiagram(Graphics g)
        {
            // Remove existing event handlers
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxBubble_Click;
            int diagramWidth = (int)(0.5f * frmMainForm.mainChartPlotting.Width);
            int diagramHeight = (int)(0.7f * frmMainForm.mainChartPlotting.Height);
            // Define margins
            int leftMargin = (int)(0.1 * frmMainForm.mainChartPlotting.Width);
            int topMargin = (int)(0.01 * frmMainForm.mainChartPlotting.Height);
            int xOrigin = leftMargin + (int)(0.03f * frmMainForm.mainChartPlotting.Width);
            int yOrigin = topMargin + (frmMainForm.mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * frmMainForm.mainChartPlotting.Height);

            // Set the title of the diagram
            float fontSize = clsConstants.legendTextSize;
            Font titleFont = new Font("Times New Roman", 25, FontStyle.Bold);
            Brush titleBrush = Brushes.Black;
            g.DrawString("Bubble Diagram", titleFont, titleBrush, new PointF((int)(0.4 * frmMainForm.mainChartPlotting.Width), (float)(0.1*topMargin)));

            // Set the axis titles and labels
            string xAxisTitle = "metamorphic";
            string yAxisTitle = "desulphurization";
            List<Brush> Colors = new List<Brush>();

            if(frmImportSamples.WaterData.Count>0)
            {
                // Draw the axis labels
                Font axisFont = new Font("Times New Roman", fontSize, FontStyle.Bold);

                double maxX = double.MinValue;
                double maxY = double.MinValue;
                foreach (var data in frmImportSamples.WaterData)
                {
                    double xValue = (data.Cl - data.Na) / data.Mg; // X-axis value
                    double yValue = (data.So4 * 100) / data.Cl;    // Y-axis value
                    maxX = Math.Max(maxX, xValue - xValue % 10 + 10);
                    maxY = Math.Max(maxY, yValue - yValue % 10 + 10);
                }

                // Draw X-axis title
                g.DrawString(xAxisTitle, axisFont, titleBrush, new PointF(xOrigin + diagramWidth/2, yOrigin + diagramHeight + 30));
                
                // Draw Y-axis title (rotated)
                GraphicsState gstate = g.Save();
                g.TranslateTransform(xOrigin - (int)(0.5 * leftMargin), yOrigin + diagramHeight / 2);
                g.RotateTransform(-90);
                g.DrawString(yAxisTitle, axisFont, Brushes.Black, new PointF(0, 0));
                g.Restore(gstate);

                // Draw the axes (X and Y)
                Pen axisPen = new Pen(Color.Black, 1);
                
                // Draw X-axis
                g.DrawLine(axisPen, xOrigin, yOrigin + diagramHeight, xOrigin + diagramWidth, yOrigin + diagramHeight);
                
                // Draw Y-axis
                g.DrawLine(axisPen, xOrigin, yOrigin, xOrigin, yOrigin + diagramHeight);

                // Calculate scaling factors
                double fx = diagramWidth / maxX; // X-axis scale
                double fy = diagramHeight / maxY; // Y-axis scale

                // Draw grid lines and ticks
                int tickCountX = 6;
                double tickStepX = maxX / (tickCountX);
                for (int i = 0; i <= tickCountX; i++)
                {
                    double tickValueX = i * tickStepX;
                    float xPosition = xOrigin + (float)(tickValueX * fx);
                    
                    // Draw X-axis tick
                    g.DrawLine(axisPen, xPosition, yOrigin + diagramHeight, xPosition, yOrigin + diagramHeight + 5);
                    
                    // Draw X-axis label
                    g.DrawString(tickValueX.ToString("F0"), axisFont, Brushes.Black, xPosition - 10, yOrigin + diagramHeight + 10);
                    
                    // Draw vertical grid line
                    if (i > 0 && i<tickCountX)
                    {
                        Pen gridPen = new Pen(Color.LightGray, 1) { DashStyle = DashStyle.Dot };
                        g.DrawLine(gridPen, xPosition, yOrigin, xPosition, yOrigin + diagramHeight);
                    }
                    else
                    {
                        g.DrawLine(axisPen, xPosition, yOrigin, xPosition, yOrigin + diagramHeight);
                    }
                }

                int tickCountY = 6;
                double tickStepY = maxY / (tickCountY);
                for (int i = 0; i <= tickCountY; i++)
                {
                    double tickValueY = i * tickStepY;
                    double yPosition = yOrigin + diagramHeight - (double)(tickValueY * fy);
                    //tickValueY = Math.Ceiling(tickValueY);
                    // Draw Y-axis tick
                    g.DrawLine(axisPen, xOrigin - 5, (float)yPosition, xOrigin, (float)yPosition);
                    
                    // Draw Y-axis label
                    g.DrawString(tickValueY.ToString("F2"), axisFont, Brushes.Black, xOrigin - 40, (float)(yPosition - 10));
                    
                    // Draw horizontal grid line
                    if (i > 0 && i<tickCountY)
                    {
                        Pen gridPen = new Pen(Color.LightGray, 1) { DashStyle = DashStyle.Dot };
                        g.DrawLine(gridPen, xOrigin, (float)(yPosition), xOrigin + diagramWidth, (float)(yPosition));
                    }
                    else
                    {
                        g.DrawLine(axisPen, xOrigin, (float)(yPosition), xOrigin + diagramWidth, (float)(yPosition));
                    }

                }

                // Plot the data points
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    var data = frmImportSamples.WaterData[i];
                    double xValue = (data.Cl - data.Na) / data.Mg;
                    double yValue = (data.So4 * 100) / data.Cl;

                    // Scale the values to fit within the diagram
                    float scaledX = xOrigin + (float)(xValue * fx);
                    float scaledY = yOrigin + diagramHeight - (float)(yValue * fy);

                    // Determine the bubble size and color
                    int bubbleSize = (int)(0.015f * frmMainForm.mainChartPlotting.Width);
                    Color bubbleColor = GetColorByTDS(data.TDS);
                    Brush squareBrush = new SolidBrush(bubbleColor);
                    if (data.bubble)
                    {
                        squareBrush = new SolidBrush(data.color);
                    }

                    Pen bubbleBorderPen = new Pen(Color.Black, 1);
                    if (data.shape != null && data.shape != "Circle" && !float.IsNaN(scaledX) && !float.IsNaN(scaledY))
                    {
                        for (int j = 0; j < frmSymbolPicker.symbolNames.Count; j++)
                        {
                            if (data.shape == frmSymbolPicker.symbolNames.ElementAt(j))
                            {
                                frmSymbolPicker.DrawSymbol(g, j, (int)scaledX - 12, (int)scaledY - 12, 25, squareBrush);
                                break;
                            }
                        }
                    }
                    else
                    {
                        g.DrawEllipse(bubbleBorderPen, scaledX - bubbleSize / 2, scaledY - bubbleSize / 2, bubbleSize, bubbleSize);
                        g.FillEllipse(squareBrush, scaledX - bubbleSize / 2, scaledY - bubbleSize / 2, bubbleSize, bubbleSize);
                    }
                }
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    var data = frmImportSamples.WaterData[i];
                    double xValue = (data.Cl - data.Na) / data.Mg;
                    double yValue = (data.So4 * 100) / data.Cl;

                    // Scale the values to fit within the diagram
                    float scaledX = xOrigin + (float)(xValue * fx);
                    float scaledY = yOrigin + diagramHeight - (float)(yValue * fy);
                    int bubbleSize = (int)(0.015f * frmMainForm.mainChartPlotting.Width);
                    string label = "W" + (i + 1).ToString();
                    g.DrawString(label, new Font("Times New Roman", 8, FontStyle.Bold), titleBrush, new PointF(scaledX + bubbleSize / 2 + 5, scaledY - 10));
                }
            }
            if (frmImportSamples.WaterData.Count > 0)
            {
                DrawLegendBubble(g);
            }
            
        }
        /// <summary>
        /// Draws the legend for the Bubble Diagram, showing color ranges and sample metadata.
        /// </summary>
        public static void DrawLegendBubble(Graphics g)
        {

            // Define legend colors and labels
            var legendItems = new[]
            {
                new { Label = "20000-40000", Color = Color.Red },
                new { Label = "40000-60000", Color = Color.Orange },
                new { Label = "60000-80000", Color = Color.Gray },
                new { Label = "80000-100000", Color = Color.Yellow },
                new { Label = "100000-120000", Color = Color.LightGreen },
                new { Label = "120000-140000", Color = Color.Blue },
                new { Label = "140000-160000", Color = Color.Green }
            };

            #region Draw Legend

            if (frmImportSamples.WaterData.Count > 0)
            {
                int metaX = (int)(0.69f * frmMainForm.mainChartPlotting.Width);
                int metaY = clsConstants.metaY;

                int metaHeight = 0;
                int legendtextSize = clsConstants.legendTextSize;
                int metaWidth = (int)(0.2 * frmMainForm.mainChartPlotting.Width);

                using (Font font = new Font("Times New Roman", legendtextSize, FontStyle.Bold))
                {
                    StringFormat stringFormat = new StringFormat { FormatFlags = StringFormatFlags.NoClip, Trimming = StringTrimming.EllipsisCharacter };

                    for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                    {
                        var data = frmImportSamples.WaterData[i];
                        string fullText = "";
                        if (clsConstants.clickedHeaders.Count > 0)
                        {
                            int c = 0;
                            //fullText += "W" + (i + 1).ToString() + ", ";

                            foreach (var header in clsConstants.clickedHeaders)
                            {
                                if (header == "Job ID")
                                {
                                    fullText += data.jobID;
                                }
                                else if (header == "Sample ID")
                                {
                                    fullText += data.sampleID;
                                }
                                else if (header == "Client ID")
                                {
                                    fullText += data.ClientID;
                                }
                                else if (header == "Well Name")
                                {
                                    fullText += data.Well_Name;
                                }
                                else if (header == "Lat")
                                {
                                    fullText += data.latitude;
                                }
                                else if (header == "Long")
                                {
                                    fullText += data.longtude;
                                }
                                else if (header == "Sample Type")
                                {
                                    fullText += data.sampleType;
                                }
                                else if (header == "Formation Name")
                                {
                                    fullText += data.formName;
                                }
                                else if (header == "Depth")
                                {
                                    fullText += data.Depth;
                                }
                                else if (header == "Prep")
                                {
                                    fullText += data.prep;
                                }
                                if (c != clsConstants.clickedHeaders.Count - 1)
                                {
                                    fullText += ", ";
                                }
                                c++;
                            }
                        }
                        else
                        {
                            fullText += data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
                        }
                        SizeF textSize = g.MeasureString(fullText, font, metaWidth - 30); // Adjust for wrapping width
                        metaWidth = (int)Math.Max(metaWidth, textSize.Width); // Ensure metaWidth accounts for the largest text
                        metaHeight += (int)Math.Ceiling(textSize.Height); // Add spacing between lines

                    }
                }

                Bitmap metaBitmap = new Bitmap(metaWidth, metaHeight);
                PictureBox metaPictureBox = new PictureBox
                {
                    Size = new Size(metaWidth, metaHeight),
                    Image = metaBitmap
                };
                metaPictureBox.MouseDoubleClick += (_sender, e) =>
                    frmMainForm.pictureBoxCollinsPieStiffMeta_Click(_sender, e, "Stiff Legend");
                frmMainForm.metaPanel.Controls.Clear();
                frmMainForm.metaPanel.Controls.Add(metaPictureBox);
                frmMainForm.metaPanel.Size = new Size(metaWidth, metaHeight);
                frmMainForm.metaPanel.Visible = true;
                frmMainForm.metaPanel.BringToFront();

                g = Graphics.FromImage(metaBitmap);
                g.Clear(Color.White);
                g.DrawRectangle(new Pen(Color.Blue), metaX - 15.0f, metaY - 10.0f, metaWidth + 15.0f, metaHeight + 30.0f);

                int ysample = 0;
                g.FillRectangle(Brushes.White, 0, 0, metaWidth, metaHeight);
                g.DrawRectangle(new Pen(Color.Blue, 2), 0, 0, metaWidth, metaHeight);

                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    var data = frmImportSamples.WaterData[i];

                    string fullText = "";
                    if (clsConstants.clickedHeaders.Count > 0)
                    {
                        int c = 0;
                        //fullText += "W" + (i + 1).ToString() + ", ";
                        foreach (var header in clsConstants.clickedHeaders)
                        {
                            if (header == "Job ID")
                            {
                                fullText += data.jobID;
                            }
                            else if (header == "Sample ID")
                            {
                                fullText += data.sampleID;
                            }
                            else if (header == "Client ID")
                            {
                                fullText += data.ClientID;
                            }
                            else if (header == "Well Name")
                            {
                                fullText += data.Well_Name;
                            }
                            else if (header == "Lat")
                            {
                                fullText += data.latitude;
                            }
                            else if (header == "Long")
                            {
                                fullText += data.longtude;
                            }
                            else if (header == "Sample Type")
                            {
                                fullText += data.sampleType;
                            }
                            else if (header == "Formation Name")
                            {
                                fullText += data.formName;
                            }
                            else if (header == "Depth")
                            {
                                fullText += data.Depth;
                            }
                            else if (header == "Prep")
                            {
                                fullText += data.prep;
                            }
                            if (c != clsConstants.clickedHeaders.Count - 1)
                            {
                                fullText += ", ";
                            }
                            c++;
                        }
                    }
                    else
                    {
                        fullText += data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
                    }
                    RectangleF textRect = new RectangleF(30, ysample, metaWidth - 35, metaHeight);

                    Font font = new Font("Times New Roman", legendtextSize, FontStyle.Bold);
                    SizeF textSize = g.MeasureString(fullText, font, (int)textRect.Width); // Adjust for wrapping width
                    g.DrawString("W" + (i + 1).ToString() + ", ", font, Brushes.Black, 0, ysample);
                    g.DrawString(fullText,
                            font,
                            Brushes.Black,
                            textRect);
                    ysample += (int)Math.Ceiling(textSize.Height); // Move down based on wrapped height
                }

                frmMainForm.metaPanel.Location = new Point(metaX - 14, metaY - 9);
                frmMainForm.mainChartPlotting.Controls.Add(frmMainForm.metaPanel);
                int legendX = (int)(0.1f * frmMainForm.mainChartPlotting.Width);
                int legendY = (int)(0.1f * frmMainForm.mainChartPlotting.Height);
                int s = 0;
                for (int i = 0; i < legendItems.Length; i++)
                {

                    string fullText = legendItems[i].Label;
                    SizeF textSize = g.MeasureString(fullText, new Font("Times New Roman", clsConstants.legendTextSize));
                    s = (int)(s + textSize.Width + 40);
                }


                int legendBoxHeight = (int)(0.03f * frmMainForm.mainChartPlotting.Height);
                int legendBoxWidth = s;


                frmMainForm.legendPictureBox.Size = new Size(legendBoxWidth, legendBoxHeight);
                Bitmap bit = new Bitmap(legendBoxWidth, legendBoxHeight);
                g = Graphics.FromImage(bit);
                int xsample = legendX;


                using (Graphics legendGraphics = g)
                {
                    //legendGraphics.Clear(Color.White);  // Fill background
                    legendGraphics.FillRectangle(Brushes.White, 0, 0, legendBoxWidth - 5, legendBoxHeight - 5);
                    legendGraphics.DrawRectangle(new Pen(Color.Blue, 2), 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                    xsample = 0;
                    for (int i = 0; i < legendItems.Length; i++)
                    {
                        Brush myBrush = new SolidBrush(legendItems[i].Color);

                        if (frmCollinsLegend.IsUpdateClicked)
                        {
                            myBrush = new SolidBrush(clsCollinsDrawer.legendColors[i]);

                        }
                        legendGraphics.FillEllipse(myBrush, xsample + 5, 2, 18, 18);


                        // Draw text beside the shape
                        legendGraphics.DrawString(legendItems[i].Label, new Font("Times New Roman", clsConstants.legendTextSize), Brushes.Black, xsample + 25, 5);

                        string fullText = legendItems[i].Label;
                        SizeF textSize = g.MeasureString(fullText, new Font("Times New Roman", clsConstants.legendTextSize));
                        xsample += (int)textSize.Width + 40;
                    }
                }
                frmMainForm.legendPanel.Location = new Point(legendX - 14, legendY - 9);
                frmMainForm.legendPanel.Size = new System.Drawing.Size(legendBoxWidth, legendBoxHeight);
                frmMainForm.legendPictureBox.Image = bit;
                frmMainForm.legendPictureBox.Visible = true;
                frmMainForm.legendPanel.Controls.Add(frmMainForm.legendPictureBox);


                frmMainForm.legendPanel.Visible = true;

                frmMainForm.mainChartPlotting.Controls.Add(frmMainForm.legendPanel);
            }
            else
            {
                frmMainForm.legendPanel.AutoScroll = false;
            }
            frmMainForm.legendPanel.Show();
            frmMainForm.mainChartPlotting.Invalidate();
            #endregion


        }
        /// <summary>
        /// Exports the Bubble Diagram to a PowerPoint slide.
        /// </summary>
        public static void ExportBubbleDiagramToPowerPoint(PowerPoint.Slide slide, PowerPoint.Presentation presentation)
        {
            try
            {
                double maxX = double.MinValue;
                double maxY = double.MinValue;
                foreach (var data in frmImportSamples.WaterData)
                {
                    double xValue = (data.Cl - data.Na) / data.Mg; // X-axis value
                    double yValue = (data.So4 * 100) / data.Cl;    // Y-axis value
                    maxX = Math.Max(maxX, xValue - xValue % 10 + 10);
                    maxY = Math.Max(maxY, yValue - yValue % 10 + 10);
                }
                // Set chart area
                float chartX = 0.1f * presentation.PageSetup.SlideWidth, chartY = 100, chartWidth = 420, chartHeight = 0.7f * presentation.PageSetup.SlideHeight; ;
                // Add title
                PowerPoint.Shape title = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    (presentation.PageSetup.SlideWidth / 2) - 100, clsConstants.chartYPowerpoint, 200, 50);
                title.TextFrame.TextRange.Text = "Bubble Diagram";
                title.TextFrame.TextRange.Font.Size = 25;
                title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                title.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                title.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                title.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                // Draw grid lines
                int numHorizontalLines = 6, numVerticalLines = 6;
                float spacingX = (float)((maxX) / (numVerticalLines - 1)), spacingY = (float)((maxY) / (numHorizontalLines - 1));
                List<double> tickValues = new List<double>();
                maxY += spacingY;
                for (int i = 0; i <= numHorizontalLines; i++) // Horizontal lines
                {
                    double tickValueY = i * spacingY;
                    float yPosition = (float)((tickValueY) / (maxY) * (chartHeight));
                    var line = slide.Shapes.AddLine(chartX, (float)(chartY + yPosition), chartX + chartWidth, (float)(chartY + yPosition));
                    line.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                    line.Line.DashStyle = Office.MsoLineDashStyle.msoLineSolid;
                    if (i!=0 && i != numHorizontalLines)
                    {
                        line.Line.ForeColor.RGB = System.Drawing.Color.Gray.ToArgb();
                        line.Line.DashStyle = Office.MsoLineDashStyle.msoLineSquareDot;
                    }
                    
                    var n = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        chartX - 30,
                        (float)(chartY + yPosition),
                        150,
                        30
                    );
                    n.TextFrame.TextRange.Text = (maxY-tickValueY).ToString();
                    n.TextFrame2.TextRange.Font.Size = 8;
                }
                maxX += spacingX;
                for (int i = 0; i <= numVerticalLines; i++) // Vertical lines
                {
                    double tickValueX = i * spacingX;
                    float xPosition = (float)((tickValueX) / (maxX) * (chartWidth));
                    var line = slide.Shapes.AddLine(chartX + xPosition, chartY, chartX + xPosition, chartY + chartHeight);
                    line.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                    line.Line.DashStyle = Office.MsoLineDashStyle.msoLineSolid;
                    if (i != 0 && i!=numVerticalLines)
                    {
                        line.Line.ForeColor.RGB = System.Drawing.Color.Gray.ToArgb();
                        line.Line.DashStyle = Office.MsoLineDashStyle.msoLineSquareDot;
                    }
                    if (i != 0)
                    {
                        var n = slide.Shapes.AddTextbox(
                                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                chartX + xPosition-7,
                                chartY + chartHeight + 10,
                                150,
                                30
                            );
                        n.TextFrame.TextRange.Text = tickValueX.ToString();
                        n.TextFrame2.TextRange.Font.Size = 8;
                    }
                }

                // Add Axis Titles
                var xlabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX + chartWidth / 2 - 50, chartY + chartHeight + 30, 150, 30);
                    xlabel.TextFrame.TextRange.Text = "Metamorphic";
                    xlabel.TextFrame2.TextRange.Font.Size = 10;
                
                var yAxisLabel = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    chartX - 120,
                    chartY + chartHeight / 2 - 30,
                    150,
                    30
                );
                yAxisLabel.TextFrame.TextRange.Text = "Desulphurization";
                yAxisLabel.Rotation = -90; // Rotate the text 90 degrees counterclockwise
                yAxisLabel.TextFrame2.TextRange.Font.Size = 10;

                // Plot Points (as circles)
                for (int i = 0; i < frmImportSamples.WaterData.Count;i++)
                {
                    var data = frmImportSamples.WaterData[i];
                    double xValue = (frmImportSamples.WaterData[i].Cl - frmImportSamples.WaterData[i].Na) / frmImportSamples.WaterData[i].Mg;
                    double yValue = (frmImportSamples.WaterData[i].So4 * 100) / frmImportSamples.WaterData[i].Cl;

                    float xPos = (float)(chartX + (xValue * (chartWidth / 120))); // Map X value
                    float yPos = (float)(chartY + chartHeight - (yValue * (chartHeight / 20))); // Map Y value

                    System.Drawing.Color bubbleColor;
                    if (data.TDS >= 20000 && data.TDS < 40000) bubbleColor = Color.Red;
                    else if (data.TDS >= 40000 && data.TDS < 60000) bubbleColor = Color.Orange;
                    else if (data.TDS >= 60000 && data.TDS < 80000) bubbleColor = Color.Gray;
                    else if (data.TDS >= 80000 && data.TDS < 100000) bubbleColor = Color.Yellow;
                    else if (data.TDS >= 100000 && data.TDS < 120000) bubbleColor = Color.LightGreen;
                    else if (data.TDS >= 120000 && data.TDS < 140000) bubbleColor = Color.Blue;
                    else bubbleColor = Color.Green;
                    Office.MsoAutoShapeType bubbleType = Office.MsoAutoShapeType.msoShapeOval; // Default shape (rectangle)
                    Color brush = bubbleColor;
                    if (data.bubble)
                    {
                        brush = data.color;
                    }

                    switch (data.shape)
                    {
                        case "Circle":
                            bubbleType = Office.MsoAutoShapeType.msoShapeOval; // Perfect circle
                            break;
                        case "Diamond":
                            bubbleType = Office.MsoAutoShapeType.msoShapeDiamond; // Diamond shape
                            break;
                        case "Pentagon":
                            bubbleType = Office.MsoAutoShapeType.msoShapePentagon; // Pentagon shape
                            break;
                        case "Hexagon":
                            bubbleType = Office.MsoAutoShapeType.msoShapeHexagon; // Hexagon shape
                            break;
                        case "Octagon":
                            bubbleType = Office.MsoAutoShapeType.msoShapeOctagon; // Octagon shape
                            break;
                        case "Star (5-point)":
                            bubbleType = Office.MsoAutoShapeType.msoShape5pointStar; // 5-point star
                            break;
                        case "Star (6-point)":
                            bubbleType = Office.MsoAutoShapeType.msoShape6pointStar; // 6-point star
                            break;
                        case "Star (8-point)":
                            bubbleType = Office.MsoAutoShapeType.msoShape8pointStar; // 8-point star
                            break;
                        case "Rectangle":
                            bubbleType = Office.MsoAutoShapeType.msoShapeRectangle; // Rectangle shape
                            break;
                        case "Plus":
                            // For plus sign, we'll create two rectangles
                            var horizontalRectangle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos - 7, yPos - 3, 15, 7);
                            horizontalRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            horizontalRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            horizontalRectangle.Line.Weight = 1;

                            var verticalRectangle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos - 3, yPos - 7, 7, 15);
                            verticalRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            verticalRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            verticalRectangle.Line.Weight = 1;
                            break; // Exit since we've already created the plus sign
                        case "Trapezoid (up)":
                            var trapezoidUpPoints = new float[,] {
                        { xPos - 7, yPos + 7 },
                        { xPos + 7, yPos + 7 },
                        { xPos + 5, yPos - 7 },
                        { xPos - 5, yPos - 7 }
                    };
                            var trapezoidUp = slide.Shapes.AddPolyline(trapezoidUpPoints);
                            trapezoidUp.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            trapezoidUp.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            trapezoidUp.Line.Weight = 1;
                            break;
                        case "Trapezoid (right)":
                            var trapezoidRightPoints = new float[,] {
                        { xPos + 7, yPos - 5 },
                        { xPos - 7, yPos - 7 },
                        { xPos - 7, yPos + 7 },
                        { xPos + 7, yPos + 5 }
                    };
                            var trapezoidRight = slide.Shapes.AddPolyline(trapezoidRightPoints);
                            trapezoidRight.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            trapezoidRight.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            trapezoidRight.Line.Weight = 1;
                            break;
                        case "Trapezoid (down)":
                            var trapezoidDownPoints = new float[,] {
                        { xPos - 5, yPos + 7 },
                        { xPos + 5, yPos + 7 },
                        { xPos + 7, yPos - 7 },
                        { xPos - 7, yPos - 7 }
                    };
                            var trapezoidDown = slide.Shapes.AddPolyline(trapezoidDownPoints);
                            trapezoidDown.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            trapezoidDown.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            trapezoidDown.Line.Weight = 1;
                            break;
                        case "Trapezoid (left)":
                            var trapezoidLeftPoints = new float[,] {
                        { xPos + 7, yPos - 7 },
                        { xPos - 7, yPos - 5 },
                        { xPos - 7, yPos + 5 },
                        { xPos + 7, yPos + 7 }
                    };
                            var trapezoidLeft = slide.Shapes.AddPolyline(trapezoidLeftPoints);
                            trapezoidLeft.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            trapezoidLeft.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            trapezoidLeft.Line.Weight = 1;
                            break;
                        case "Vertical rectangle":
                            var vRect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos - 6, yPos - 7, 12, 15);
                            vRect.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            vRect.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            vRect.Line.Weight = 1;
                            break;
                        case "X":
                            var xPoints1 = new float[,] {
                        { xPos - 7, yPos - 7 },
                        { xPos - 3, yPos - 7 },
                        { xPos + 7, yPos + 7 },
                        { xPos + 3, yPos + 7 }
                    };
                            var xPoints2 = new float[,] {
                        { xPos + 7, yPos - 7 },
                        { xPos + 3, yPos - 7 },
                        { xPos - 7, yPos + 7 },
                        { xPos - 3, yPos + 7 }
                    };
                            var xShape1 = slide.Shapes.AddPolyline(xPoints1);
                            xShape1.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            xShape1.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            xShape1.Line.Weight = 1;
                            var xShape2 = slide.Shapes.AddPolyline(xPoints2);
                            xShape2.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            xShape2.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            xShape2.Line.Weight = 1;
                            break;
                        case "Horizontal bar":
                            var hBar = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos - 7, yPos - 6, 15, 12);
                            hBar.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            hBar.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            hBar.Line.Weight = 1;
                            break;
                        case "Up arrow":
                            var upArrowPoints = new float[,] {
                        { xPos, yPos - 7 },
                        { xPos + 7, yPos + 7 },
                        { xPos, yPos + 3 },
                        { xPos - 7, yPos + 7 }
                    };
                            var upArrow = slide.Shapes.AddPolyline(upArrowPoints);
                            upArrow.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            upArrow.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            upArrow.Line.Weight = 1;
                            break;
                        case "Right arrow":
                            var rightArrowPoints = new float[,] {
                        { xPos + 7, yPos },
                        { xPos - 7, yPos - 7 },
                        { xPos - 3, yPos },
                        { xPos - 7, yPos + 7 }
                    };
                            var rightArrow = slide.Shapes.AddPolyline(rightArrowPoints);
                            rightArrow.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            rightArrow.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            rightArrow.Line.Weight = 1;
                            break;
                        case "Down arrow":
                            var downArrowPoints = new float[,] {
                        { xPos, yPos + 7 },
                        { xPos - 7, yPos - 7 },
                        { xPos, yPos - 3 },
                        { xPos + 7, yPos - 7 }
                    };
                            var downArrow = slide.Shapes.AddPolyline(downArrowPoints);
                            downArrow.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            downArrow.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            downArrow.Line.Weight = 1;
                            break;
                        case "Left arrow":
                            var leftArrowPoints = new float[,] {
                        { xPos - 7, yPos },
                        { xPos + 7, yPos + 7 },
                        { xPos + 3, yPos },
                        { xPos + 7, yPos - 7 }
                    };
                            var leftArrow = slide.Shapes.AddPolyline(leftArrowPoints);
                            leftArrow.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            leftArrow.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            leftArrow.Line.Weight = 1;
                            break;
                        case "Arrow with tail (up)":
                            var upArrowTailPoints = new float[,] {
                        { xPos, yPos - 7 },
                        { xPos + 7, yPos + 7 },
                        { xPos, yPos + 3 },
                        { xPos - 7, yPos + 7 }
                    };
                            var upArrowTail = slide.Shapes.AddPolyline(upArrowTailPoints);
                            upArrowTail.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            upArrowTail.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            upArrowTail.Line.Weight = 1;
                            var upTail = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos - 5, yPos + 3, 10, 7);
                            upTail.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            upTail.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            upTail.Line.Weight = 1;
                            break;
                        case "Arrow with tail (right)":
                            var rightArrowTailPoints = new float[,] {
                        { xPos + 7, yPos },
                        { xPos - 7, yPos - 7 },
                        { xPos - 3, yPos },
                        { xPos - 7, yPos + 7 }
                    };
                            var rightArrowTail = slide.Shapes.AddPolyline(rightArrowTailPoints);
                            rightArrowTail.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            rightArrowTail.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            rightArrowTail.Line.Weight = 1;
                            var rightTail = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos - 3, yPos - 5, 7, 10);
                            rightTail.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            rightTail.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            rightTail.Line.Weight = 1;
                            break;
                        case "Arrow with tail (down)":
                            var downArrowTailPoints = new float[,] {
                        { xPos, yPos + 7 },
                        { xPos - 7, yPos - 7 },
                        { xPos, yPos - 3 },
                        { xPos + 7, yPos - 7 }
                    };
                            var downArrowTail = slide.Shapes.AddPolyline(downArrowTailPoints);
                            downArrowTail.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            downArrowTail.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            downArrowTail.Line.Weight = 1;
                            var downTail = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos - 5, yPos - 7, 10, 7);
                            downTail.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            downTail.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            downTail.Line.Weight = 1;
                            break;
                        case "Arrow with tail (left)":
                            var leftArrowTailPoints = new float[,] {
                        { xPos - 7, yPos },
                        { xPos + 7, yPos + 7 },
                        { xPos + 3, yPos },
                        { xPos + 7, yPos - 7 }
                    };
                            var leftArrowTail = slide.Shapes.AddPolyline(leftArrowTailPoints);
                            leftArrowTail.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            leftArrowTail.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            leftArrowTail.Line.Weight = 1;
                            var leftTail = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos + 3, yPos - 5, 7, 10);
                            leftTail.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            leftTail.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            leftTail.Line.Weight = 1;
                            break;
                        case "Upward fat arrow":
                            var fatArrowPoints = new float[,] {
                        { xPos, yPos - 7 },
                        { xPos + 7, yPos - 2 },
                        { xPos + 5, yPos - 2 },
                        { xPos + 5, yPos + 7 },
                        { xPos - 5, yPos + 7 },
                        { xPos - 5, yPos - 2 },
                        { xPos - 7, yPos - 2 }
                    };
                            var fatArrow = slide.Shapes.AddPolyline(fatArrowPoints);
                            fatArrow.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            fatArrow.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            fatArrow.Line.Weight = 1;
                            break;

                        case "Up triangle":
                            yPos -= 7;
                            var triangleUpPoints = new float[,] {
                                { xPos, yPos },
                                { xPos-8, yPos + 15 },
                                { xPos+8, yPos + 15 }
                            };
                            var triangleUp = slide.Shapes.AddPolyline(triangleUpPoints);
                            triangleUp.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            triangleUp.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            triangleUp.Line.Weight = 1;
                            break;
                        case "Down triangle":
                            yPos += 7;
                            var triangleDownPoints = new float[,] {
                                { xPos, yPos },
                                { xPos + 8, yPos - 15 },
                                { xPos-8, yPos - 15 }
                            };
                            var triangleDown = slide.Shapes.AddPolyline(triangleDownPoints);
                            triangleDown.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            triangleDown.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            triangleDown.Line.Weight = 1;
                            break;
                        case "Right triangle":
                            xPos += 7;
                            var triangleRightPoints = new float[,] {
                                { xPos, yPos },
                                { xPos - 15, yPos - 8 },
                                { xPos - 15, yPos +8 }
                            };
                            var triangleRight = slide.Shapes.AddPolyline(triangleRightPoints);
                            triangleRight.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            triangleRight.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            triangleRight.Line.Weight = 1;
                            break;
                        case "Left triangle":
                            xPos -= 7;
                            var triangleLeftPoints = new float[,] {
                                { xPos, yPos },
                                { xPos + 15, yPos - 8 },
                                { xPos + 15, yPos + 8 }
                            };
                            var triangleLeft = slide.Shapes.AddPolyline(triangleLeftPoints);
                            triangleLeft.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                            triangleLeft.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                            triangleLeft.Line.Weight = 1;
                            break;
                        //default:
                        //    // For any other shape, use a plus sign as default
                        //    var hRect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos - 7, yPos - 3, 15, 7);
                        //    hRect.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                        //    hRect.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                        //    hRect.Line.Weight = 1;

                        //    var vRectangle = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, xPos - 3, yPos - 7, 7, 15);
                        //    vRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                        //    vRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                        //    vRectangle.Line.Weight = 1;
                        //    return; // Exit since we've already created the plus sign
                    }

                    // Create the shape with the determined type
                    if(data.shape=="Circle" || data.shape=="Diamond" || data.shape== "Pentagon" || data.shape== "Hexagon" || data.shape== "Octagon" || data.shape== "Star (5-point)" || data.shape== "Star (6-point)" || data.shape== "Star (8-point)" || data.shape=="Rectangle" || data.shape==null)
                    {
                        var shapeObj = slide.Shapes.AddShape(bubbleType, xPos - 7, yPos - 7, 15, 15);
                        shapeObj.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                        shapeObj.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                        shapeObj.Line.Weight = 1;
                    }
                    


                    //bubbleType = Office.MsoAutoShapeType.msoShapeOval; // Perfect circle
                

                    // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
                    //var bubble = slide.Shapes.AddShape(bubbleType, xPos - 17, yPos - 17, 15, 15); // Adjust for your specific use case

                    //bubble.Fill.ForeColor.RGB = ColorTranslator.ToOle(bubbleColor);
                    //bubble.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

                    // Add Label

                    PowerPoint.Shape label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, xPos + 5, yPos - 20, 150, 15);
                        label.TextFrame2.TextRange.Font.Size = 8;
                        label.TextFrame.TextRange.Text = "W"+(i+1).ToString();
                        label.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        label.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                        label.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                }

                #region Draw Legend
                var legendColors = new[] { Color.Red, Color.Orange, Color.Gray, Color.Yellow, Color.LightGreen, Color.Blue, Color.Green };
                var legendLabels = new[] { "20000-40000", "40000-60000", "60000-80000", "80000-100000", "100000-120000", "120000-140000", "140000-160000" };
                int legendX = (int)(0.1f * presentation.PageSetup.SlideWidth);
                int legendY = 50;
                int xSample = legendX + 5;
                int fontSize = clsConstants.legendTextSize;
                int legendBoxWidth = 0;
                int legendBoxHeight = 0;
                // Add legend border


                for (int i = 0; i < legendLabels.Length; i++)
                {
                    // Create a temp shape just to measure
                    var temp = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        xSample + 50, legendY + 10, 100, 20);
                    temp.TextFrame.TextRange.Text = legendLabels[i];
                    temp.TextFrame.TextRange.Font.Size = fontSize;
                    temp.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    temp.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    temp.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

                    temp.TextFrame.TextRange.Text = legendLabels[i];
                    temp.TextFrame.TextRange.Font.Size = fontSize;
                    temp.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    int charCount = legendLabels[i].Length;
                    float approxWidth = fontSize * charCount * 0.4f;
                    temp.Width = approxWidth;
                    legendBoxWidth += (int)temp.Width+10;
                    legendBoxHeight = Math.Max(legendBoxHeight, (int)temp.Height+5);
                    temp.Delete(); // clean up

                }
                PowerPoint.Shape legendBorder = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                legendX, legendY, legendBoxWidth, legendBoxHeight);
                legendBorder.Fill.Transparency = 1.0f;
                legendBorder.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                legendBorder.Line.Weight = 2;
                for (int i = 0; i < legendLabels.Length; i++)
                {

                    PowerPoint.Shape legendBox = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeOval,
                        xSample, legendY + 5, 10, 10);
                    legendBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(legendColors[i]);


                    PowerPoint.Shape legendText = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        xSample + 10, legendY+2, 100, 20);
                    legendText.TextFrame.TextRange.Text = legendLabels[i];
                    legendText.TextFrame.TextRange.Font.Size = fontSize;
                    legendText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    legendText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    legendText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                    legendText.TextFrame.MarginLeft = 0;
                    legendText.TextFrame.MarginRight = 0;
                    legendText.TextFrame.MarginTop = 0;
                    legendText.TextFrame.MarginBottom = 0;
                    int charCount = legendLabels[i].Length;
                    float approxWidth = fontSize * charCount * 0.4f; // 0.6 is a rough factor

                    legendText.Width = approxWidth;

                    xSample += (int)legendText.Width+10;
                }



                // Add metadata
                float metadataX = 550;
                float metadataY = legendY;
                int metaWidth = 180; // Set a fixed width for the text box (enables wrapping)
                int metaHeight = 0;

                float ysample = metadataY;

                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    var data = frmImportSamples.WaterData[i];

                    // Prepare wrapped text
                    string fullText = "";
                    if (clsConstants.clickedHeaders.Count > 0)
                    {
                        int c = 0;
                        //fullText += "W" + (i + 1).ToString() + ", ";
                        foreach (var header in clsConstants.clickedHeaders)
                        {

                            if (header == "Job ID")
                            {
                                fullText += data.jobID;
                            }
                            else if (header == "Sample ID")
                            {
                                fullText += data.sampleID;
                            }
                            else if (header == "Client ID")
                            {
                                fullText += data.ClientID;
                            }
                            else if (header == "Well Name")
                            {
                                fullText += data.Well_Name;
                            }
                            else if (header == "Lat")
                            {
                                fullText += data.latitude;
                            }
                            else if (header == "Long")
                            {
                                fullText += data.longtude;
                            }
                            else if (header == "Sample Type")
                            {
                                fullText += data.sampleType;
                            }
                            else if (header == "Formation Name")
                            {
                                fullText += data.formName;
                            }
                            else if (header == "Depth")
                            {
                                fullText += data.Depth;
                            }
                            else if (header == "Prep")
                            {
                                fullText += data.prep;
                            }
                            if (c != clsConstants.clickedHeaders.Count - 1)
                            {
                                fullText += ", ";
                            }
                            c++;
                        }
                    }
                    else
                    {
                        fullText += data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
                    }

                    // Add textbox with wrapping and fixed width
                    PowerPoint.Shape metadataText = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        metadataX + 25, ysample, metaWidth, 20); // initial height, PowerPoint will auto-expand

                    metadataText.TextFrame.TextRange.Text = fullText;
                    metadataText.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                    metadataText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    metadataText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop;
                    metadataText.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;

                    // Remove margins to reduce waste of space
                    metadataText.TextFrame.MarginLeft = 0;
                    metadataText.TextFrame.MarginRight = 0;
                    metadataText.TextFrame.MarginTop = 0;
                    metadataText.TextFrame.MarginBottom = 0;

                    // Auto-resize height only
                    metadataText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    var temp = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        metadataX - 40, ysample - 5, 100, 20);
                    temp.TextFrame.TextRange.Text = "W" + (i + 1).ToString() + ", ";
                    temp.TextFrame.TextRange.Font.Size = 10;
                    temp.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    temp.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    temp.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                    temp.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
                    ysample += metadataText.Height;
                    metaHeight += (int)(metadataText.Height);
                }

                // Draw blue border box after content is drawn
                PowerPoint.Shape metaBorder = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    metadataX - 5, metadataY - 5, metaWidth + 35, metaHeight + 10);
                metaBorder.Fill.Transparency = 1.0f;
                metaBorder.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                metaBorder.Line.Weight = 1;
                #endregion

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

        }
        /// <summary>
        /// Returns a color based on the TDS value for a sample.
        /// </summary>
        private static Color GetColorByTDS(double tds)
        {
            if (tds >= 20000 && tds < 40000) return Color.Red;
            if (tds >= 40000 && tds < 60000) return Color.Orange;
            if (tds >= 60000 && tds < 80000) return Color.Gray;
            if (tds >= 80000 && tds < 100000) return Color.Yellow;
            if (tds >= 100000 && tds < 120000) return Color.LightGreen;
            if (tds >= 120000 && tds < 140000) return Color.Blue;
            return Color.Green;
        }

    }
}
