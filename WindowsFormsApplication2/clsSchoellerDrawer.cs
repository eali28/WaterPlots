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
        public static void DrawSchoellerDiagram(Graphics g)
        {
            // Remove existing event handlers
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;

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
                int ysample = (int)(0.13f * frmMainForm.mainChartPlotting.Height);
                int legendX = xsample;
                int legendY = ysample;


                int legendBoxHeight = 0;
                int legendtextSize = clsConstants.legendTextSize;

                int legendBoxWidth = 0;

                using (Font font = new Font("Times New Roman", legendtextSize,FontStyle.Bold))
                {
                    foreach (var data in frmImportSamples.WaterData)
                    {
                        string fullText = data.Well_Name+", "+ data.ClientID+", " +data.Depth;
                        SizeF textSize = g.MeasureString(fullText, font);
                        if (textSize.Width+30 > legendBoxWidth)
                        {
                            legendBoxWidth = (int)Math.Round(textSize.Width, 0)+30;
                        }
                        legendBoxHeight += (int)Math.Round(textSize.Height, 0);
                    }
                }
                
                //Form1.pic.Visible = true;
                frmMainForm.legendPictureBox.Size = new Size(legendBoxWidth, legendBoxHeight);
                Bitmap bit = new Bitmap(legendBoxWidth, legendBoxHeight);
                g = Graphics.FromImage(bit);
                g.DrawRectangle(new Pen(Color.Black), legendX - 15, legendY - 10, legendBoxWidth + 15, legendBoxHeight + 30);



                using (Graphics legendGraphics = g)
                {
                    //legendGraphics.Clear(Color.White);  // Fill background
                    legendGraphics.FillRectangle(Brushes.White, 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                    legendGraphics.DrawRectangle(new Pen(Color.Blue, 2), 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                    ysample = 0;
                    for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                    {
                        Brush squareBrush = new SolidBrush(frmImportSamples.WaterData[i].color);

                        Pen axisPen = new Pen(frmImportSamples.WaterData[i].color, 2);
                        axisPen.Width = frmImportSamples.WaterData[i].lineWidth;
                        axisPen.DashStyle = frmImportSamples.WaterData[i].selectedStyle;
                        g.DrawLine(axisPen, 5, ysample + 5, 25, ysample + 5);
                        string fullText = frmImportSamples.WaterData[i].Well_Name + ", " + frmImportSamples.WaterData[i].ClientID + ", " + frmImportSamples.WaterData[i].Depth;
                        SizeF textSize = g.MeasureString(fullText, new Font("Times New Roman", legendtextSize, FontStyle.Bold));
                        // Draw text beside the shape
                        legendGraphics.DrawString(
                            fullText,
                            new Font("Times New Roman", legendtextSize, FontStyle.Bold),
                            Brushes.Black, 30, ysample
                        );

                        ysample += (int)(textSize.Height);
                    }
                }
                //Form1.legendPanel.BackColor = Color.Transparent;
                frmMainForm.legendPanel.Location = new Point(legendX - 14, legendY - 9);
                frmMainForm.legendPanel.Size = new System.Drawing.Size(legendBoxWidth, legendBoxHeight);
                frmMainForm.legendPictureBox.Image = bit;
                //Form1.pic.Location = new Point(0, 0);
                //Form1.pic.Visible = true;
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
        public static void ExportSchoellerDiagramToPowerPoint(PowerPoint.Slide slide, PowerPoint.Presentation presentation)
        {

            // Chart area dimensions
            float chartX = 160, chartY = 110, chartWidth = 700, chartHeight = 700;

            // Add chart title
            PowerPoint.Shape chartTitle = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, presentation.PageSetup.SlideWidth / 2 - 20, chartY - 100, 400, 50);
            chartTitle.TextFrame.TextRange.Text = "Schoeller Diagram";
            chartTitle.TextFrame.TextRange.Font.Size = 27;
            chartTitle.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;

            // Add description text
            PowerPoint.Shape descriptionText = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX, chartY + chartHeight + 60, 750, 100);
            descriptionText.TextFrame.TextRange.Text = "SCHOELLER logarithmic diagram of major ions in meq/L demonstrates different water types on the same diagram.";
            descriptionText.TextFrame.TextRange.Font.Size = 17;
            descriptionText.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            var yAxisLabel = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                chartX - 200,
                chartY + chartHeight / 2 - 30,
                200,
                30
            );
            yAxisLabel.TextFrame.TextRange.Text = "Concentration (meq/L)";
            yAxisLabel.Rotation = -90;

            // X-axis labels
            string[] XaxisItems = { "K", "Mg", "Ca", "Na", "Cl", "HCO3", "SO4" };
            double[] YaxisItems = { 1, 10, 100, 1000, 10000 };

            // Draw X and Y axes
            slide.Shapes.AddLine(chartX, chartY + chartHeight, chartX + chartWidth, chartY + chartHeight) // X axis
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
            slide.Shapes.AddLine(chartX, chartY, chartX, chartY + chartHeight) // Y axis
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            // Y-axis grid lines and labels (log scale)
            for (int i = 0; i < YaxisItems.Length; i++)
            {
                // Logarithmic Y-axis position calculation
                float yPos = chartY + chartHeight - (float)(Math.Log10(YaxisItems[i]) * (chartHeight / Math.Log10(10000)));
                // Y-axis labels
                slide.Shapes.AddLine(chartX, yPos, chartX - 10, yPos)
                    .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                PowerPoint.Shape label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 70, yPos - 10, 100, 30);
                label.TextFrame.TextRange.Text = YaxisItems[i].ToString();
                label.TextFrame.TextRange.Font.Size = 17;
                label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            }

            for (int j = 1; j <= XaxisItems.Length; j++)
            {
                // Logarithmic scale Y position
                float xPos = chartX + (j * (chartWidth / XaxisItems.Length));
                slide.Shapes.AddLine(xPos, chartY + chartHeight, xPos, chartY + chartHeight + 10).Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
                PowerPoint.Shape labelX = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, xPos - 5, chartY + chartHeight + 20, 100, 50);
                labelX.TextFrame.TextRange.Text = XaxisItems[j - 1];
                labelX.TextFrame.TextRange.Font.Size = 20;
                labelX.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            }

            // Add lines and data points for each water sample
            int legendY = (int)chartY;
            int ysample = legendY;

            int legendX = (int)(chartX + chartWidth + 60);
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
                    float yPos = chartY + chartHeight - (float)(Math.Log10(values[j - 1]) * (chartHeight / Math.Log10(10000)));

                    points[j - 1] = new PointF(xPos - 7.5f, yPos - 7.5f);
                }

                // Flatten the points array into a float array for AddPolyline
                float[] polylinePoints = new float[points.Length * 2];
                for (int j = 0; j < points.Length; j++)
                {
                    polylinePoints[j * 2] = points[j].X;
                    polylinePoints[j * 2 + 1] = points[j].Y;
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


                // Loop Through Water Data Samples and Add Text
                var line = slide.Shapes.AddLine(legendX, ysample + 10, legendX + 30, ysample + 10);
                line.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color);
                string sampleText = (frmImportSamples.WaterData[i].Well_Name) + "," + (frmImportSamples.WaterData[i].ClientID) + "," + (frmImportSamples.WaterData[i].Depth);

                PowerPoint.Shape sampleTextShape = slide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    legendX + 50, ysample, 700, 20
                );

                sampleTextShape.TextFrame.TextRange.Text = sampleText;
                sampleTextShape.TextFrame.TextRange.Font.Size = 15;
                sampleTextShape.TextFrame.TextRange.Font.Name = "Times New Roman";
                sampleTextShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                sampleTextShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);

                ysample += 30;

            }
            int s = 0;
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                if (frmImportSamples.WaterData[i].Well_Name.Length + frmImportSamples.WaterData[i].ClientID.Length + frmImportSamples.WaterData[i].Depth.Length + 5 > s)
                {
                    s = frmImportSamples.WaterData[i].Well_Name.Length + frmImportSamples.WaterData[i].ClientID.Length + frmImportSamples.WaterData[i].Depth.Length + 5;
                }
            }


            int legendBoxHeight = (frmImportSamples.WaterData.Count * 30) + 5;
            int fontSize = Math.Max(8, Math.Min(12, legendBoxHeight / frmImportSamples.WaterData.Count));
            int legendBoxWidth = s * (fontSize - 1);
            // Add border around legend
            PowerPoint.Shape borderShape = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, legendX, legendY, legendBoxWidth, legendBoxHeight);
            borderShape.Fill.Transparency = 1.0f;
            borderShape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue);
            borderShape.Line.Weight = 2;

        }
    }
}
