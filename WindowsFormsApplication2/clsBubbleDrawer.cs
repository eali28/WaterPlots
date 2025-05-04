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
        /// Draws a Bubble Diagram on chart2.
        /// </summary>
        public static void DrawBubbleDiagram(Graphics g)
        {
            int diagramWidth = (int)(0.5f * frmMainForm.mainChartPlotting.Width);
            int diagramHeight = (int)(0.7f * frmMainForm.mainChartPlotting.Height);
            // Define margins
            int leftMargin = (int)(0.1 * frmMainForm.mainChartPlotting.Width);
            int topMargin = (int)(0.01 * frmMainForm.mainChartPlotting.Height);
            int xOrigin = leftMargin + (int)(0.03f * frmMainForm.mainChartPlotting.Width);
            int yOrigin = topMargin + (frmMainForm.mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * frmMainForm.mainChartPlotting.Height);

            // Set the title of the diagram
            float fontSize = 12;
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
                    if (i > 0)
                    {
                        Pen gridPen = new Pen(Color.LightGray, 1) { DashStyle = DashStyle.Dot };
                        g.DrawLine(gridPen, xPosition, yOrigin, xPosition, yOrigin + diagramHeight);
                    }
                }

                int tickCountY = 6;
                double tickStepY = maxY / (tickCountY);
                for (int i = 0; i <= tickCountY; i++)
                {
                    double tickValueY = i * tickStepY;
                    double yPosition = yOrigin + diagramHeight - (double)(tickValueY * fy);
                    
                    // Draw Y-axis tick
                    g.DrawLine(axisPen, xOrigin - 5, (float)yPosition, xOrigin, (float)yPosition);
                    
                    // Draw Y-axis label
                    g.DrawString(tickValueY.ToString("F2"), axisFont, Brushes.Black, xOrigin - 40, (float)(yPosition - 10));
                    
                    // Draw horizontal grid line
                    if (i > 0)
                    {
                        Pen gridPen = new Pen(Color.LightGray, 1) { DashStyle = DashStyle.Dot };
                        g.DrawLine(gridPen, xOrigin, (float)(yPosition), xOrigin + diagramWidth, (float)(yPosition));
                    }
                }

                // Plot the data points
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    double xValue = (frmImportSamples.WaterData[i].Cl - frmImportSamples.WaterData[i].Na) / frmImportSamples.WaterData[i].Mg;
                    double yValue = (frmImportSamples.WaterData[i].So4 * 100) / frmImportSamples.WaterData[i].Cl;

                    // Scale the values to fit within the diagram
                    float scaledX = xOrigin + (float)(xValue * fx);
                    float scaledY = yOrigin + diagramHeight - (float)(yValue * fy);

                    // Determine the bubble size and color
                    int bubbleSize = (int)(0.015f * frmMainForm.mainChartPlotting.Width);
                    Color bubbleColor = GetColorByTDS(frmImportSamples.WaterData[i].TDS);

                    // Draw the bubble
                    Brush bubbleBrush = new SolidBrush(bubbleColor);
                    g.FillEllipse(bubbleBrush, scaledX - bubbleSize/2, scaledY - bubbleSize/2, bubbleSize, bubbleSize);
                    Colors.Add(bubbleBrush);

                    // Draw the bubble border
                    Pen bubbleBorderPen = new Pen(Color.Black, 1);
                    g.DrawEllipse(bubbleBorderPen, scaledX - bubbleSize/2, scaledY - bubbleSize/2, bubbleSize, bubbleSize);

                    //// Draw the label
                    //string label = "W" + (i + 1).ToString();
                    //graphics.DrawString(label, axisFont, titleBrush, new PointF(scaledX + bubbleSize/2 + 5, scaledY - 10));
                }
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    double xValue = (frmImportSamples.WaterData[i].Cl - frmImportSamples.WaterData[i].Na) / frmImportSamples.WaterData[i].Mg;
                    double yValue = (frmImportSamples.WaterData[i].So4 * 100) / frmImportSamples.WaterData[i].Cl;

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
                DrawLegendBubble(g, Colors);
            }
            
        }
        public static MarkerStyle GetMarkerStyle(string shape)
        {
            switch (shape)
            {
                case "circle":
                    return MarkerStyle.Circle;
                case "cube":
                    return MarkerStyle.Square;
                case "hexagon":
                    return MarkerStyle.Diamond; // Closest to hexagon
                case "merkaba":
                    return MarkerStyle.Star5; // 5-point star for Merkaba
                case "triangle":
                    return MarkerStyle.Triangle;
                default:
                    return MarkerStyle.Circle;
            }
        }

        public static Office.MsoAutoShapeType ConvertShapeToPowerPoint(string shape)
        {
            switch (shape)
            {
                case "circle":
                    return Office.MsoAutoShapeType.msoShapeOval; // Perfect circle
                case "cube":
                    return Office.MsoAutoShapeType.msoShapeRectangle; 
                case "hexagon":
                    return Office.MsoAutoShapeType.msoShapeHexagon; // Hexagon shape
                case "merkaba":
                    return Office.MsoAutoShapeType.msoShape5pointStar; // Star shape for Merkaba
                case "triangle":
                    return Office.MsoAutoShapeType.msoShapeIsoscelesTriangle;
                default:
                    return Office.MsoAutoShapeType.msoShapeOval; // Default fallback
            }
        }


        public static void DrawLegendBubble(Graphics g, List<Brush> BubbleColors)
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
                int metaY = (int)(0.13f * frmMainForm.mainChartPlotting.Height);
                int size = 0;
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    if (frmImportSamples.WaterData[i].Well_Name.Length + frmImportSamples.WaterData[i].ClientID.Length + frmImportSamples.WaterData[i].Depth.Length > size)
                    {
                        size = frmImportSamples.WaterData[i].Well_Name.Length + frmImportSamples.WaterData[i].ClientID.Length + frmImportSamples.WaterData[i].Depth.Length;
                    }
                }


                double metaHeight = 0;
                int legendtextSize = clsConstants.legendTextSize;
                int metaWidth = 0;

                using (Font font = new Font("Times New Roman", legendtextSize, FontStyle.Bold))
                {
                    for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                    {
                        string fullText = "W" + (i + 1).ToString() + ", " + frmImportSamples.WaterData[i].Well_Name + ", " + frmImportSamples.WaterData[i].ClientID + ", " + frmImportSamples.WaterData[i].Depth;
                        SizeF textSize = g.MeasureString(fullText, font);
                        if (textSize.Width > metaWidth)
                        {
                            metaWidth = (int)Math.Round(textSize.Width, 0);

                        }
                        metaHeight += Math.Round(textSize.Height, 0);
                    }
                }
                Bitmap metaBitmap = new Bitmap(metaWidth, (int)metaHeight);
                //Form1.pic.Visible = true;
                PictureBox metaPictureBox = new PictureBox();
                metaPictureBox.Size = new Size(metaWidth, (int)metaHeight);
                metaPictureBox.Image = metaBitmap;

                frmMainForm.metaPanel.Controls.Add(metaPictureBox);
                frmMainForm.metaPanel.Size = new Size(metaWidth, (int)metaHeight);
                frmMainForm.metaPanel.Visible = true;
                frmMainForm.metaPanel.BringToFront();

                g = Graphics.FromImage(metaBitmap);
                g.Clear(Color.White);
                g.DrawRectangle(new Pen(Color.Blue), metaX - 15.0f, metaY - 10.0f, metaWidth + 15.0f, (int)metaHeight + 30.0f);
                int ysample = metaY;
                //legendGraphics.Clear(Color.White);  // Fill background
                g.FillRectangle(Brushes.White, 0, 0, metaWidth - 1, (int)metaHeight - 1);
                g.DrawRectangle(new Pen(Color.Blue, 2), 0, 0, metaWidth - 1, (int)metaHeight - 1);
                ysample = 0;
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {



                    // Draw text beside the shape
                    g.DrawString("W" + (i + 1).ToString() + ", " +
                        frmImportSamples.WaterData[i].Well_Name + ", " + frmImportSamples.WaterData[i].ClientID + ", " + frmImportSamples.WaterData[i].Depth,
                        new Font("Times New Roman", legendtextSize, FontStyle.Bold),
                        Brushes.Black, 0, ysample
                    );
                    string fullText = "W" + (i + 1).ToString() + ", " + frmImportSamples.WaterData[i].Well_Name + ", " + frmImportSamples.WaterData[i].ClientID + ", " + frmImportSamples.WaterData[i].Depth;
                    SizeF textSize = g.MeasureString(fullText, new Font("Times New Roman", legendtextSize, FontStyle.Bold));
                    ysample += (int)(Math.Round(textSize.Height, 0));
                }

                //Form1.legendPanel.BackColor = Color.Transparent;
                frmMainForm.metaPanel.Location = new Point(metaX - 14, metaY - 9);
                frmMainForm.metaPanel.Size = new System.Drawing.Size(metaWidth, metaWidth);
                frmMainForm.legendPictureBox.Image = metaBitmap;
                //Form1.pic.Location = new Point(0, 0);
                //Form1.pic.Visible = true;
                frmMainForm.metaPanel.Controls.Add(frmMainForm.legendPictureBox);


                frmMainForm.metaPanel.Visible = true;

                frmMainForm.mainChartPlotting.Controls.Add(frmMainForm.metaPanel);
                int legendX = (int)(0.1 * frmMainForm.mainChartPlotting.Width);
                int legendY = (int)(0.1f * frmMainForm.mainChartPlotting.Height);
                int s = 0;
                for (int i = 0; i < legendItems.Length; i++)
                {
                    
                    string fullText = legendItems[i].Label;
                    SizeF textSize = g.MeasureString(fullText, new Font("Times New Roman", 10));
                    s = (int)(s + textSize.Width + 20);
                }


                int legendBoxHeight = (int)(0.03f * frmMainForm.mainChartPlotting.Height);
                float fontSize = clsConstants.legendTextSize; // Make font size relative
                int legendBoxWidth = s;


                //Form1.pic.Visible = true;
                frmMainForm.legendPictureBox.Size = new Size(legendBoxWidth, legendBoxHeight);
                //frmMainForm.legendPictureBox.Location = new Point(legendX, legendY);
                Bitmap bit = new Bitmap(legendBoxWidth, legendBoxHeight);
                
                g = Graphics.FromImage(bit);
                //g.DrawRectangle(new Pen(Color.Blue), legendX - 15.0f, legendY - 10.0f, legendBoxWidth + 15.0f, legendBoxHeight + 30.0f);
                int xsample = 0;


                using (Graphics legendGraphics = g)
                {
                    //legendGraphics.Clear(Color.White);  // Fill background
                    legendGraphics.FillRectangle(Brushes.White, 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                    legendGraphics.DrawRectangle(new Pen(Color.Blue, 2), 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                    xsample = 0;
                    for (int i = 0; i < legendItems.Length; i++)
                    {

                        Brush legendBrush = new SolidBrush(legendItems[i].Color);
                        legendGraphics.FillEllipse(legendBrush, xsample, 0, 20, 20);
 

                        // Draw text beside the shape
                        legendGraphics.DrawString(legendItems[i].Label, new Font("Times New Roman", fontSize), Brushes.Black, xsample + 20, 5);
                        string fullText = legendItems[i].Label;
                        SizeF textSize = g.MeasureString(fullText, new Font("Times New Roman", fontSize));
                        xsample += (int)textSize.Width+20;
                    }
                }
                //Form1.legendPanel.BackColor = Color.Transparent;
                frmMainForm.legendPanel.Location = new Point(legendX - 14, legendY - 9);
                frmMainForm.legendPanel.Size = new System.Drawing.Size(legendBoxWidth, legendBoxHeight);
                frmMainForm.legendPictureBox.Image = bit;
                //Form1.pic.Location = new Point(0, 0);
                frmMainForm.legendPictureBox.Visible = true;
                //frmMainForm.legendPictureBox.MouseDoubleClick += frmMainForm.pictureBoxPie_Click;
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
        public static void ExportBubbleDiagramToPowerPoint(PowerPoint.Slide slide, PowerPoint.Presentation presentation)
        {
            try
            {
                List<Color> BubbleColors = new List<Color>();
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
                float chartX = 100, chartY = 100, chartWidth = 820, chartHeight = 800;
                // Add title
                PowerPoint.Shape title = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    450, 20, 600, 50);
                title.TextFrame.TextRange.Text = "Bubble Diagram";
                title.TextFrame.TextRange.Font.Size = 55;
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
                        line.Line.ForeColor.RGB = System.Drawing.Color.LightGray.ToArgb();
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
                        line.Line.ForeColor.RGB = System.Drawing.Color.LightGray.ToArgb();
                        line.Line.DashStyle = Office.MsoLineDashStyle.msoLineSquareDot;
                    }
                        var n = slide.Shapes.AddTextbox(
                            Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            chartX + xPosition,
                            chartY + chartHeight + 10,
                            150,
                            30
                        );
                        n.TextFrame.TextRange.Text = tickValueX.ToString();
                    

                }

                // Add Axis Titles
                slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX + chartWidth / 2 - 50, chartY + chartHeight + 30, 150, 30)
                    .TextFrame.TextRange.Text = "Metamorphic";
                var yAxisLabel = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    chartX - 150,
                    chartY + chartHeight / 2 - 30,
                    150,
                    30
                );
                yAxisLabel.TextFrame.TextRange.Text = "Desulphurization";
                yAxisLabel.Rotation = -90; // Rotate the text 90 degrees counterclockwise

                // Plot Points (as circles)
                for (int i = 0; i < frmImportSamples.WaterData.Count;i++)
                {

                    double xValue = (frmImportSamples.WaterData[i].Cl - frmImportSamples.WaterData[i].Na) / frmImportSamples.WaterData[i].Mg;
                    double yValue = (frmImportSamples.WaterData[i].So4 * 100) / frmImportSamples.WaterData[i].Cl;

                    float xPos = (float)(chartX + (xValue * (chartWidth / 120))); // Map X value
                    float yPos = (float)(chartY + chartHeight - (yValue * (chartHeight / 20))); // Map Y value

                    System.Drawing.Color bubbleColor;
                    if (frmImportSamples.WaterData[i].TDS >= 20000 && frmImportSamples.WaterData[i].TDS < 40000) bubbleColor = Color.Red;
                    else if (frmImportSamples.WaterData[i].TDS >= 40000 && frmImportSamples.WaterData[i].TDS < 60000) bubbleColor = Color.Orange;
                    else if (frmImportSamples.WaterData[i].TDS >= 60000 && frmImportSamples.WaterData[i].TDS < 80000) bubbleColor = Color.Gray;
                    else if (frmImportSamples.WaterData[i].TDS >= 80000 && frmImportSamples.WaterData[i].TDS < 100000) bubbleColor = Color.Yellow;
                    else if (frmImportSamples.WaterData[i].TDS >= 100000 && frmImportSamples.WaterData[i].TDS < 120000) bubbleColor = Color.LightGreen;
                    else if (frmImportSamples.WaterData[i].TDS >= 120000 && frmImportSamples.WaterData[i].TDS < 140000) bubbleColor = Color.Blue;
                    else bubbleColor = Color.Green;
                    Office.MsoAutoShapeType bubbleType = Office.MsoAutoShapeType.msoShapeOval; // Default shape (rectangle)

                    BubbleColors.Add(bubbleColor);
                    switch (frmImportSamples.WaterData[i].shape)
                    {
                        case "circle":
                            bubbleType = Office.MsoAutoShapeType.msoShapeOval; // Perfect circle
                            break;
                        case "cube":
                            bubbleType = Office.MsoAutoShapeType.msoShapeRectangle; // Cube shape as a rectangle
                            break;
                        case "hexagon":
                            bubbleType = Office.MsoAutoShapeType.msoShapeHexagon; // Hexagon shape
                            break;
                        case "merkaba":
                            bubbleType = Office.MsoAutoShapeType.msoShape5pointStar; // Star shape for Merkaba
                            break;
                        case "triangle":
                            bubbleType = Office.MsoAutoShapeType.msoShapeIsoscelesTriangle; // Triangle shape
                            break;
                    }

                    // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
                    var bubble = slide.Shapes.AddShape(bubbleType, xPos - 17, yPos - 17, 35, 35); // Adjust for your specific use case

                    bubble.Fill.ForeColor.RGB = ColorTranslator.ToOle(bubbleColor);
                    bubble.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

                    // Add Label

                    PowerPoint.Shape label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, xPos + 30, yPos - 20, 150, 15);
                        label.TextFrame.TextRange.Text = "W"+(i+1).ToString();
                        label.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        label.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                        label.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                }

                // Draw Legend
                float legendX = chartX + chartWidth + 80, legendY = chartY;
                var legendColors = new[] { Color.Red, Color.Orange, Color.Gray, Color.Yellow, Color.LightGreen, Color.Blue, Color.Green };
                var legendLabels = new[] { "20000-40000", "40000-60000", "60000-80000", "80000-100000", "100000-120000", "120000-140000", "140000-160000" };

                float legendItemHeight = 40; // Height of each legend item (circle + spacing)
                float legendWidth = 200;    // Width of the legend box (adjust as needed)
                float legendHeight = legendColors.Length * legendItemHeight; // Total height of legend box

                // Draw the rectangle around the legend
                PowerPoint.Shape legendRectangle = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    legendX - 10,       // Add padding on the left
                    legendY - 10,       // Add padding on the top
                    legendWidth,        // Legend box width
                    legendHeight + 60   // Legend box height with padding
                );
                legendRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb(); // Border color
                legendRectangle.Fill.Transparency = 1.0f; // Make the rectangle transparent
                PowerPoint.Shape legendTitle = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    legendX + 40,
                    legendY + 10,
                    150,
                    15
                );

                legendTitle.TextFrame.TextRange.Text = "TDS (mg/L)";
                legendTitle.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue; // Make text bold
                legendTitle.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                legendTitle.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                legendTitle.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                // Add legend items inside the rectangle
                for (int i = 0; i < legendColors.Length; i++)
                {
                    // Add legend circle
                    PowerPoint.Shape legendBox = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, legendX, legendY + 40, 35, 35);
                    legendBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(legendColors[i]);
                    legendBox.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

                    // Add legend label

                    PowerPoint.Shape legendLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, legendX + 40, legendY + 40, 150, 15);
                    legendLabel.TextFrame.TextRange.Text = legendLabels[i];
                    legendLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    legendLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    legendLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

                    legendY += legendItemHeight; // Move down for the next legend item
                }
                int s = 0;
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    if ("W".Length + ",".Length + frmImportSamples.WaterData[i].Well_Name.Length + ",".Length + frmImportSamples.WaterData[i].ClientID.Length + ",".Length + frmImportSamples.WaterData[i].Depth.Length > s)
                    {
                        s = "W".Length + ",".Length + frmImportSamples.WaterData[i].Well_Name.Length + ",".Length + frmImportSamples.WaterData[i].ClientID.Length + ",".Length + frmImportSamples.WaterData[i].Depth.Length;
                    }
                }
                // Add metadata
                float metadataX = legendX - 20;
                float metadataY = legendY + 80;
                int metaHeight = (frmImportSamples.WaterData.Count * 33);
                int fontSize = 17;
                int metaWidth = s*(fontSize-7);
                PowerPoint.Shape metaRectangle = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        metadataX - 10,       // Add padding on the left
                        metadataY - 10,       // Add padding on the top
                        metaWidth,        // Legend box width
                        metaHeight   // Legend box height with padding
                    );
                metaRectangle.Fill.Transparency = 1.0f;
                metaRectangle.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                metaRectangle.Line.Weight = 2;
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    PowerPoint.Shape metadataText = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        metadataX+40, metadataY + (i * 32), 500, 20);
                    metadataText.TextFrame.TextRange.Text = "W" + (i + 1).ToString() + "," + (frmImportSamples.WaterData[i].Well_Name) + "," + (frmImportSamples.WaterData[i].ClientID) + "," + (frmImportSamples.WaterData[i].Depth);
                    metadataText.TextFrame.TextRange.Font.Size = fontSize;
                    metadataText.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    metadataText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    metadataText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                    Office.MsoAutoShapeType bubbleType = Office.MsoAutoShapeType.msoShapeOval; // Default shape (rectangle)
                    switch (frmImportSamples.WaterData[i].shape)
                    {
                        case "circle":
                            bubbleType = Office.MsoAutoShapeType.msoShapeOval; // Perfect circle
                            break;
                        case "cube":
                            bubbleType = Office.MsoAutoShapeType.msoShapeRectangle; // Cube shape as a rectangle
                            break;
                        case "hexagon":
                            bubbleType = Office.MsoAutoShapeType.msoShapeHexagon; // Hexagon shape
                            break;
                        case "merkaba":
                            bubbleType = Office.MsoAutoShapeType.msoShape5pointStar; // Star shape for Merkaba
                            break;
                        case "triangle":
                            bubbleType = Office.MsoAutoShapeType.msoShapeIsoscelesTriangle; // Triangle shape
                            break;
                    }
                    // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
                    var bubble = slide.Shapes.AddShape(bubbleType, metadataX - 7, metadataY + (i * 32), 28, 28); // Adjust for your specific use case

                    bubble.Fill.ForeColor.RGB = ColorTranslator.ToOle(BubbleColors[i]);
                    bubble.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

        }
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
