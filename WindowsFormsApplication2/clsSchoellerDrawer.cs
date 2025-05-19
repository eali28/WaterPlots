using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Office = Microsoft.Office.Core;
using System.Windows.Forms.DataVisualization.Charting;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WindowsFormsApplication2
{
    public class clsSchoellerDrawer
    {
        /// <summary>
        /// Draws the Schoeller Diagram, plotting ion concentrations for each sample.
        /// </summary>
        public static void DrawSchoellerDiagram(Graphics g)
        {
            // Remove existing event handlers
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxBubble_Click;

            // Set up fonts and constants
            Font labelFont = new Font("Times New Roman", 12, FontStyle.Bold);
            Font legendFont = new Font("Times New Roman", 8, FontStyle.Bold);
            
            // Define starting coordinates

            
            // Define margins
            int leftMargin = (int)(0.1 * frmMainForm.mainChartPlotting.Width);
            int topMargin = (int)(0.01 * frmMainForm.mainChartPlotting.Height);

            // Calculate diagram dimensions based on starting point
            int diagramWidth = (int)(0.5f * frmMainForm.mainChartPlotting.Width);
            int diagramHeight = (int)(0.7f * frmMainForm.mainChartPlotting.Height);
            int x1 = leftMargin + (int)(0.03f * frmMainForm.mainChartPlotting.Width);
            int y1 = topMargin + (frmMainForm.mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * frmMainForm.mainChartPlotting.Height);
            // Draw the title
            g.DrawString("SCHOELLER logarithmic diagram of major ions in meq/L demonstrate different water types on the same diagram.", 
                labelFont, Brushes.Black, x1, y1+diagramHeight-(int)(0.1*diagramHeight));

            // Schoeller X-axis components
            string[] XaxisItems = { "K", "Mg", "Ca", "Na", "Cl", "HCO3", "SO4" };

            // Factor conversions
            double Nafac = 22.99, Kfac = 39.0983, Cafac = 20.039, Mgfac = 12.1525;
            double Clfac = 35.453, HCO3fac = 61.01684, CO3fac = 30.004, SO4fac = 48.0313;

            // Draw X-axis labels
            int xSpacing = diagramWidth / (XaxisItems.Length + 1);
            for (int i = 0; i < XaxisItems.Length; i++)
            {
                int xPos = x1 + (i + 1) * xSpacing;
                g.DrawString(XaxisItems[i], labelFont, Brushes.Black, xPos - 10, y1 + diagramHeight + 20);
            }

            // Draw Y-axis label (rotated)
            GraphicsState gstate = g.Save();
            g.TranslateTransform(x1 - (int)(0.5 * leftMargin), y1 + diagramHeight / 2);
            g.RotateTransform(-90);
            g.DrawString("Concentration (meq/L)", labelFont, Brushes.Black, new PointF(0, 0));
            g.Restore(gstate);

            // Logarithmic scaling on Y-axis
            double yMax = 10000; // Max concentration
            double yMin = 0.1;  // Min concentration
            
            // Draw Y-axis and grid lines
            g.DrawLine(Pens.Black, x1, y1, x1, y1 + diagramHeight); // Y-axis line
            g.DrawLine(Pens.Black, x1 + diagramWidth, y1, x1 + diagramWidth, y1 + diagramHeight);
            
            for (double yValue = yMin; yValue <= yMax; yValue *= 10)
            {
                int yPos = y1 + diagramHeight - (int)((Math.Log10(yValue) - Math.Log10(yMin)) / Math.Log10(yMax / yMin) * diagramHeight);
                
                // Draw horizontal grid line
                Pen linePen = new Pen(Color.LightGray, 1);
                linePen.DashStyle = DashStyle.Dot;
                g.DrawLine(linePen, x1, yPos, x1 + diagramWidth, yPos);
                
                // Draw Y-axis labels
                g.DrawString(yValue.ToString(), labelFont, Brushes.Black, x1 - (int)(0.3*leftMargin), yPos - 10);
            }

            // Draw X-axis line
            g.DrawLine(Pens.Black, x1, y1 + diagramHeight, x1 + diagramWidth, y1 + diagramHeight);
            g.DrawLine(Pens.Black, x1, y1, x1 + diagramWidth, y1);
            // Extract data and plot the points
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                double Na = frmImportSamples.WaterData[i].Na / Nafac;
                double K = frmImportSamples.WaterData[i].K / Kfac;
                double Ca = frmImportSamples.WaterData[i].Ca / Cafac;
                double Mg = frmImportSamples.WaterData[i].Mg / Mgfac;
                double Cl = frmImportSamples.WaterData[i].Cl / Clfac;
                double HCO3 = frmImportSamples.WaterData[i].HCO3 / HCO3fac;
                double SO4 = frmImportSamples.WaterData[i].So4 / SO4fac;

                double[] values = { K, Mg, Ca, Na, Cl, HCO3, SO4 };
                Brush sampleBrush = new SolidBrush(frmImportSamples.WaterData[i].color);
                List<Point> points = new List<Point>();

                // Plot the points
                for (int j = 0; j < XaxisItems.Length; j++)
                {
                    double value = values[j];
                    int yPos;
                    if (value != 0)
                    {
                        yPos = y1 + diagramHeight - (int)((Math.Log10(value) - Math.Log10(yMin)) / Math.Log10(yMax / yMin) * diagramHeight);
                    }
                    else
                    {
                        yPos=y1+diagramHeight;
                    }
                    int xPos = x1 + (j + 1) * xSpacing;

                    points.Add(new Point(xPos, yPos));
                    if (points.Count >= 2)
                    {
                        using (Pen linePen = new Pen(frmImportSamples.WaterData[i].color, 2))
                        {
                            linePen.Width = frmImportSamples.WaterData[i].lineWidth;
                            linePen.DashStyle = frmImportSamples.WaterData[i].selectedStyle;
                            g.DrawLine(linePen, points[j - 1].X, points[j - 1].Y, points[j].X, points[j].Y);
                        }
                    }
                }
            }

            #region Draw Legend

            if (frmImportSamples.WaterData.Count > 0)
            {
                int xsample = (int)(0.69f * frmMainForm.mainChartPlotting.Width);
                int ysample = clsConstants.metaY;
                int legendX = xsample;
                int legendY = ysample;

                int legendBoxHeight = 0;
                int legendtextSize = clsConstants.legendTextSize;
                int legendBoxWidth = (int)(0.2 * frmMainForm.mainChartPlotting.Width); // Set fixed width for wrapping area

                using (Font font = new Font("Times New Roman", legendtextSize, FontStyle.Bold))
                {
                    for (int i= 0;i < frmImportSamples.WaterData.Count;i++)
                    {
                        var data = frmImportSamples.WaterData[i];
                        string fullText = "";
                        if (clsConstants.clickedHeaders.Count > 0)
                        {
                            int c = 0;

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
                            fullText += "W" + (i + 1).ToString() + ", " + data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
                        }
                        SizeF textSize = g.MeasureString(fullText, font, legendBoxWidth - 30); // limit width for wrapping
                        legendBoxWidth = (int)Math.Max(legendBoxWidth, textSize.Width);
                        legendBoxHeight += (int)Math.Ceiling(textSize.Height); // add spacing between lines
                    }
                }

                frmMainForm.legendPictureBox.Size = new Size(legendBoxWidth, legendBoxHeight);
                Bitmap bit = new Bitmap(legendBoxWidth, legendBoxHeight);
                g = Graphics.FromImage(bit);

                //g.DrawRectangle(new Pen(Color.Black), legendX - 15, legendY - 10, legendBoxWidth + 15, legendBoxHeight + 30);

                using (Graphics legendGraphics = g)
                {
                    legendGraphics.FillRectangle(Brushes.White, 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                    legendGraphics.DrawRectangle(new Pen(Color.Blue, 2), 0, 0, legendBoxWidth, legendBoxHeight - 1);
                    ysample = 0;

                    for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                    {
                        var data = frmImportSamples.WaterData[i];
                        Brush squareBrush = new SolidBrush(data.color);
                        Pen axisPen = new Pen(data.color, data.lineWidth)
                        {
                            DashStyle = data.selectedStyle
                        };

                        g.DrawLine(axisPen, 5, ysample + 10, 25, ysample + 10);

                        string fullText = "";
                        if (clsConstants.clickedHeaders.Count > 0)
                        {
                            int c = 0;
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
                            fullText +=data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
                        }
                        RectangleF textRect = new RectangleF(30, ysample, legendBoxWidth - 35, legendBoxHeight); // large height to wrap

                        Font font = new Font("Times New Roman", legendtextSize, FontStyle.Bold);
                        SizeF textSize = legendGraphics.MeasureString(fullText, font, (int)textRect.Width);

                        legendGraphics.DrawString(
                            fullText,
                            font,
                            Brushes.Black,
                            textRect
                        );

                        ysample += (int)Math.Ceiling(textSize.Height); // Move down based on wrapped height
                    }
                }

                frmMainForm.legendPanel.Location = new Point(legendX - 14, legendY - 9);
                frmMainForm.legendPanel.Size = new Size(legendBoxWidth, legendBoxHeight);
                frmMainForm.legendPictureBox.Image = bit;

                frmMainForm.legendPictureBox.MouseDoubleClick += frmMainForm.pictureBoxSchoeller_Click;
                frmMainForm.legendPanel.Controls.Add(frmMainForm.legendPictureBox);
                frmMainForm.legendPanel.Visible = true;
                frmMainForm.mainChartPlotting.Controls.Add(frmMainForm.legendPanel);
            }
            else
            {
                frmMainForm.legendPanel.AutoScroll = false;
            }

            frmMainForm.legendPanel.Show();

            #endregion
        }
        /// <summary>
        /// Exports the Schoeller Diagram to a PowerPoint slide.
        /// </summary>
        public static void ExportSchoellerDiagramToPowerPoint(PowerPoint.Slide slide, PowerPoint.Presentation presentation)
        {

            // Chart area dimensions
            float chartX = (int)(0.1f * presentation.PageSetup.SlideWidth), chartY = 100, chartWidth = 420, chartHeight = 0.6f * (int)presentation.PageSetup.SlideHeight;

            // Add chart title
            PowerPoint.Shape title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, (presentation.PageSetup.SlideWidth / 2) - 100, clsConstants.chartYPowerpoint, 200, 50);
            title.TextFrame.TextRange.Text = "Schoeller Diagram";
            title.TextFrame.TextRange.Font.Size = 25;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            title.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            title.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            // Remove margins to reduce waste of space
            title.TextFrame.MarginLeft = 0;
            title.TextFrame.MarginRight = 0;
            title.TextFrame.MarginTop = 0;
            title.TextFrame.MarginBottom = 0;

            // Add description text
            PowerPoint.Shape descriptionText = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX, chartY + chartHeight + 60, 750, 100);
            descriptionText.TextFrame.TextRange.Text = "SCHOELLER logarithmic diagram of major ions in meq/L demonstrates different water types on the same diagram.";
            descriptionText.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
            descriptionText.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            descriptionText.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            descriptionText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            descriptionText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            descriptionText.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;
            // Remove margins to reduce waste of space
            descriptionText.TextFrame.MarginLeft = 0;
            descriptionText.TextFrame.MarginRight = 0;
            descriptionText.TextFrame.MarginTop = 0;
            descriptionText.TextFrame.MarginBottom = 0;
            var yAxisLabel = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                chartX - 100,
                chartY + chartHeight / 2 - 30,
                100, 30
            );
            yAxisLabel.TextFrame.TextRange.Text = "Concentration (meq/L)";
            yAxisLabel.Rotation = -90;
            yAxisLabel.TextFrame.TextRange.Font.Size = 12;
            yAxisLabel.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            yAxisLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            yAxisLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            yAxisLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            yAxisLabel.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;
            // Remove margins to reduce waste of space
            yAxisLabel.TextFrame.MarginLeft = 0;
            yAxisLabel.TextFrame.MarginRight = 0;
            yAxisLabel.TextFrame.MarginTop = 0;
            yAxisLabel.TextFrame.MarginBottom = 0;

            // X-axis labels
            string[] XaxisItems = { "K", "Mg", "Ca", "Na", "Cl", "HCO3", "SO4" };
            double[] YaxisItems = { 0.1,1, 10, 100, 1000, 10000 };

            // Draw X and Y axes
            slide.Shapes.AddLine(chartX, chartY + chartHeight, chartX + chartWidth, chartY + chartHeight) // X axis
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
            slide.Shapes.AddLine(chartX, chartY, chartX, chartY + chartHeight) // Y axis
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            // Y-axis grid lines and labels (log scale)
            for (int i = 0; i < YaxisItems.Length; i++)
            {
                // Logarithmic Y-axis position calculation
                float yPos = chartY +chartHeight - (int)((Math.Log10(YaxisItems[i])-Math.Log10(YaxisItems[0]))/Math.Log10(10000/YaxisItems[0])*chartHeight);
                // Y-axis labels
                slide.Shapes.AddLine(chartX, yPos, chartX - 10, yPos)
                    .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                PowerPoint.Shape labelY = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 70, yPos - 10, 100, 30);
                labelY.TextFrame.TextRange.Text = YaxisItems[i].ToString();
                labelY.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                labelY.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                labelY.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                labelY.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                labelY.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                labelY.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;

                // Remove margins to reduce waste of space
                labelY.TextFrame.MarginLeft = 0;
                labelY.TextFrame.MarginRight = 0;
                labelY.TextFrame.MarginTop = 0;
                labelY.TextFrame.MarginBottom = 0;
            }
            List<PointF> Positions = new List<PointF>();
            for (int j = 1; j <= XaxisItems.Length; j++)
            {
                // Logarithmic scale Y position
                float xPos = chartX + (j * (chartWidth / XaxisItems.Length));
                Positions.Add(new PointF(xPos, chartY + chartHeight));
                slide.Shapes.AddLine(xPos, chartY + chartHeight, xPos, chartY + chartHeight + 10).Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                PowerPoint.Shape labelX = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, xPos - 5, chartY + chartHeight + 20, 100, 50);
                labelX.TextFrame.TextRange.Text = XaxisItems[j - 1];
                labelX.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                labelX.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                labelX.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                labelX.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                labelX.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                labelX.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;

                // Remove margins to reduce waste of space
                labelX.TextFrame.MarginLeft = 0;
                labelX.TextFrame.MarginRight = 0;
                labelX.TextFrame.MarginTop = 0;
                labelX.TextFrame.MarginBottom = 0;
            }

            // Add lines and data points for each water sample

            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                // Extract data values and normalize
                double Na = frmImportSamples.WaterData[i].Na / 22.99;
                double K = frmImportSamples.WaterData[i].K / 39.0983;
                double Ca = frmImportSamples.WaterData[i].Ca / 20.039;
                double Mg = frmImportSamples.WaterData[i].Mg / 12.1525;
                double Cl = frmImportSamples.WaterData[i].Cl / 35.453;
                double HCO3 = frmImportSamples.WaterData[i].HCO3 / 61.01684;
                double CO3 = frmImportSamples.WaterData[i].CO3 / 30.004;
                double SO4 = frmImportSamples.WaterData[i].So4 / 48.0313;

                double[] values = { K, Mg, Ca, Na, Cl, HCO3 + CO3, SO4 };
                PointF[] points = new PointF[values.Length];
                // Create a new series for each sample
                for (int j = 1; j <= XaxisItems.Length; j++)
                {
                    // Logarithmic scale Y position
                    float xPos = chartX + (j * (chartWidth / XaxisItems.Length));
                    float yPos = chartY + chartHeight - (int)((Math.Log10(values[j - 1]) - Math.Log10(YaxisItems[0])) / Math.Log10(10000 / YaxisItems[0]) * chartHeight);

                    points[j - 1] = new PointF(xPos - 7.5f, yPos - 7.5f);
                }

                // Flatten the points array into a float array for AddPolyline
                float[] polylinePoints = new float[points.Length * 2];
                for (int j = 0; j < points.Length; j++)
                {
                    polylinePoints[j * 2] = points[j].X;
                    polylinePoints[j * 2 + 1] = points[j].Y;
                    
                    string str = points[j].Y.ToString();
                    if (float.IsInfinity(polylinePoints[j*2+1]) || str.Contains("E"))
                    {
                        polylinePoints[j * 2 + 1] = chartY+chartHeight;
                    }
                }

                // Add the polyline to the slide
                PowerPoint.Shape setLine = slide.Shapes.AddPolyline(new float[,]
                {
                    { polylinePoints[0], polylinePoints[1] },
                    { polylinePoints[2], polylinePoints[3] },
                    { polylinePoints[4], polylinePoints[5] },
                    { polylinePoints[6], polylinePoints[7] },
                    { polylinePoints[8], polylinePoints[9] },
                    { polylinePoints[10],polylinePoints[11]},
                    { polylinePoints[12],polylinePoints[13]},
                });
                setLine.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color); // Set line color
                setLine.Line.Weight = 2; // Set line width


                //// Loop Through Water Data Samples and Add Text
                //var line = slide.Shapes.AddLine(legendX, ysample + 10, legendX + 30, ysample + 10);
                //line.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color);
                //string sampleText = (frmImportSamples.WaterData[i].Well_Name) + "," + (frmImportSamples.WaterData[i].ClientID) + "," + (frmImportSamples.WaterData[i].Depth);

                //PowerPoint.Shape sampleTextShape = slide.Shapes.AddTextbox(
                //    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                //    legendX + 50, ysample, 700, 20
                //);

                //sampleTextShape.TextFrame.TextRange.Text = sampleText;
                //sampleTextShape.TextFrame.TextRange.Font.Size = 15;
                //sampleTextShape.TextFrame.TextRange.Font.Name = "Times New Roman";
                //sampleTextShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                //sampleTextShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);

                //ysample += 30;

            }
            //int s = 0;
            //for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            //{
            //    if (frmImportSamples.WaterData[i].Well_Name.Length + frmImportSamples.WaterData[i].ClientID.Length + frmImportSamples.WaterData[i].Depth.Length + 5 > s)
            //    {
            //        s = frmImportSamples.WaterData[i].Well_Name.Length + frmImportSamples.WaterData[i].ClientID.Length + frmImportSamples.WaterData[i].Depth.Length + 5;
            //    }
            //}


            //int legendBoxHeight = (frmImportSamples.WaterData.Count * 30) + 5;
            //int fontSize = Math.Max(8, Math.Min(12, legendBoxHeight / frmImportSamples.WaterData.Count));
            //int legendBoxWidth = s * (fontSize - 1);
            //// Add border around legend
            //PowerPoint.Shape borderShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, legendX, legendY, legendBoxWidth, legendBoxHeight);
            //borderShape.Fill.Transparency = 1.0f;
            //borderShape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue);
            //borderShape.Line.Weight = 2;
            #region Draw Legend
            if (frmImportSamples.WaterData.Count > 0)
            {
                int legendY = 50;

                float metadataX = 550;
                float metadataY = legendY;
                int metaWidth = 180; // Set a fixed width for the text box (enables wrapping)
                int metaHeight = 0;

                float ysample = metadataY;

                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    var data = frmImportSamples.WaterData[i];

                    // Draw the colored line
                    var line = slide.Shapes.AddLine(metadataX, ysample + 10, metadataX + 20, ysample + 10);
                    line.Line.ForeColor.RGB = ColorTranslator.ToOle(data.color);
                    line.Line.Weight = data.lineWidth;
                    line.Line.DashStyle = clsRadarDrawer.ConvertDashStyle(data.selectedStyle);

                    // Prepare wrapped text
                    string fullText = data.Well_Name + ", " + data.ClientID + ", " + data.Depth;

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

                    ysample += metadataText.Height + 5;
                    metaHeight += (int)(metadataText.Height + 5);
                }

                // Draw blue border box after content is drawn
                PowerPoint.Shape metaBorder = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    metadataX - 5, metadataY - 5, metaWidth + 35, metaHeight + 10);
                metaBorder.Fill.Transparency = 1.0f;
                metaBorder.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                metaBorder.Line.Weight = 1;
            }
            #endregion

        }
    }
}
