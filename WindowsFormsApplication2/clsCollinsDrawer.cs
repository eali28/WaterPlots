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
using Microsoft.Win32;
using System.Diagnostics;

using System.Management;
using System.Windows.Forms.DataVisualization.Charting;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication2
{
    public class clsCollinsDrawer
    {
        public static string[] labels = { "Na+K", "Ca", "Mg", "Cl", "SO4", "HCO3 + CO3" };
        public static Color[] legendColors = { Color.Cyan, Color.Purple, Color.DarkSlateBlue, Color.Yellow, Color.Magenta, Color.Green };
        public static Rectangle chartBounds = frmMainForm.mainChartPlotting.ClientRectangle;
        public static int margin = (int)(0.02 * chartBounds.Width); // Make margin relative to width

        // Calculate triangle and diamond dimensions within the chart area
        public static int availableWidth = chartBounds.Width - 4 * margin;
        public static int availableHeight = chartBounds.Height - 4 * margin;
        /// <summary>
        /// Draws the Collins Diagram, plotting stacked bars for cations and anions for each sample.
        /// </summary>
        public static void DrawCollinsDiagram(Graphics g, int chartWidth, int chartHeight)
        {
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;

            frmMainForm.mainChartPlotting.Invalidate();

            // Calculate center position
            int leftMargin = (int)(0.1 * chartWidth);
            int topMargin = (int)(0.01 * chartHeight);
            float diagramWidth = (float)(frmImportSamples.WaterData.Count* 0.03f*chartWidth); // Make width relative
            int diagramHeight = (int)(0.7 * chartHeight);
            int xOrigin = leftMargin+(int)(0.03f*chartWidth);
            int yOrigin = topMargin + (chartHeight - diagramHeight) / 2 - (int)(0.02 * chartHeight);
            // factors
            double Nafac = 22.99, Kfac = 39.0983, Cafac = 20.039, Mgfac = 12.1525, Clfac = 35.453, HCO3fac = 61.01684, CO3fac = 30.004, SO4fac = 48.0313;

            // Draw diagram border
            Pen borderPen = new Pen(Color.Black, 2);

            g.DrawRectangle(borderPen, xOrigin, yOrigin, diagramWidth, diagramHeight);
            float fontSize = 25; // Make font size relative
            // Add diagram title
            string title = "COLLINS DIAGRAM";
            Font titleFont = new Font("Times New Roman", fontSize, FontStyle.Bold);
            SizeF titleSize = g.MeasureString(title, titleFont);
            int titleX = (int)(xOrigin + (diagramWidth - (int)titleSize.Width) / 2);
            int titleY = (int)(0.01 * chartHeight);
            g.DrawString(title, titleFont, Brushes.Black, 0.4f * frmMainForm.mainChartPlotting.Width, 0.01f * frmMainForm.mainChartPlotting.Height);
            fontSize = 0.01f * frmMainForm.mainChartPlotting.Height;
            List<string> samples = new List<string>();
            for (int i = 1; i <= frmImportSamples.WaterData.Count; i++)
            {
                samples.Add("W" + i.ToString());

            }

            double[] Na = new double[frmImportSamples.WaterData.Count];
            double[] K = new double[frmImportSamples.WaterData.Count];
            double[] Ca = new double[frmImportSamples.WaterData.Count];
            double[] Mg = new double[frmImportSamples.WaterData.Count];
            double[] Cl = new double[frmImportSamples.WaterData.Count];
            double[] HCO3 = new double[frmImportSamples.WaterData.Count];
            double[] CO3 = new double[frmImportSamples.WaterData.Count];
            double[] SO4 = new double[frmImportSamples.WaterData.Count];
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                Na[i] += frmImportSamples.WaterData[i].Na;
                K[i] += frmImportSamples.WaterData[i].K;
                Ca[i] += frmImportSamples.WaterData[i].Ca;
                Mg[i] += frmImportSamples.WaterData[i].Mg;
                Cl[i] += frmImportSamples.WaterData[i].Cl;
                SO4[i] += frmImportSamples.WaterData[i].So4;
                HCO3[i] += frmImportSamples.WaterData[i].HCO3;
                CO3[i] += frmImportSamples.WaterData[i].CO3;
            }
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                Na[i] /= Nafac;
                K[i] /= Kfac;
                Ca[i] /= Cafac;
                Mg[i] /= Mgfac;
                Cl[i] /= Clfac;
                HCO3[i] /= HCO3fac;
                CO3[i] /= CO3fac;
                SO4[i] /= SO4fac;
            }
            double Max_Value = 0;
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                Max_Value = Math.Max(Max_Value, (Na[i] + K[i] + Ca[i] + Mg[i]));
                Max_Value = Math.Max(Max_Value, (Cl[i] + HCO3[i] + CO3[i] + SO4[i]));
            }
            Max_Value = Max_Value * 1.1;
            // Axis scaling
            double fx = diagramWidth / (samples.Count + 1); // X-axis scale
            double fy = diagramHeight / 3000.0; // Y-axis scale (max value = 3000)

            // Draw X-axis
            for (int i = 0; i < samples.Count; i++)
            {
                int x = xOrigin + (int)(fx * (i + 1));
                g.DrawLine(borderPen, x, yOrigin + diagramHeight, x, yOrigin + diagramHeight + 10); // Tick marks
                g.DrawString(samples[i], new Font("Times New Roman", fontSize), Brushes.Black, x - 10, yOrigin + diagramHeight + topMargin);
            }

            // Draw Y-axis
            for (int i = 0; i <= 3000; i += 500)
            {
                int y = yOrigin + diagramHeight - (int)(fy * i);
                g.DrawLine(borderPen, xOrigin - 10, y, xOrigin, y); // Tick marks
                g.DrawString(i.ToString(), new Font("Times New Roman", fontSize), Brushes.Black, xOrigin - 40, y - 10);
            }


            Pen axisPen = new Pen(Color.Black, 2);
            double fy_F = (double)(diagramHeight / (double)Max_Value);
            // Draw stacked bars
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                int z = 0;
                double NaK = Na[i] + K[i], HCO3CO3 = HCO3[i] + CO3[i];
                int x = xOrigin + (int)(fx * (i + 1)) - (int)(0.01 * diagramWidth);
                double currentY = yOrigin + diagramHeight;
                List<double> Items = new List<double>();
                Items.Add(NaK * fy_F);
                Items.Add(Ca[i] * fy_F);
                Items.Add(Mg[i] * fy_F);
                Items.Add(Cl[i] * fy_F);
                Items.Add(SO4[i] * fy_F);
                Items.Add(HCO3CO3 * fy_F);
                // First bar: Na+K, Ca, Mg
                for (int j = 0; j < Items.Count/2; j++)
                {
                    g.DrawRectangle(axisPen, x, (float)(currentY - Items[j]), (int)(0.02 * diagramWidth), (float)Items[j]);
                    if (!frmCollinsLegend.IsUpdateClicked)
                    {
                        g.FillRectangle(new SolidBrush(legendColors[z]), x, (float)(currentY - Items[j]), (int)(0.02 * diagramWidth), (float)Items[j]);
                    }
                    else
                    {
                        g.FillRectangle(new SolidBrush(frmCollinsLegend.CollinsColor[z]), x, (float)(currentY - Items[j]), (int)(0.02 * diagramWidth), (float)Items[j]);
                    }
                    currentY -= Items[j];
                    z++;

                }



                // Second bar: Cl, SO4, HCO3
                x += (int)(0.02 * diagramWidth); // Shift for second bar

                
                currentY = yOrigin + diagramHeight;
                for (int j = 3; j < Items.Count; j++)
                {
                    g.DrawRectangle(axisPen, x, (float)(currentY - Items[j]), (int)(0.02 * diagramWidth), (float)Items[j]);
                    if (!frmCollinsLegend.IsUpdateClicked)
                    {
                        g.FillRectangle(new SolidBrush(legendColors[z]), x, (float)(currentY - Items[j]), (int)(0.02 * diagramWidth), (float)Items[j]);
                    }
                    else
                    {
                        g.FillRectangle(new SolidBrush(frmCollinsLegend.CollinsColor[z]), x, (float)(currentY - Items[j]), (int)(0.01 * diagramWidth), (float)Items[j]);
                    }
                    currentY -= Items[j];
                    z++;

                }

                
            }
            fontSize = chartHeight * 0.01f;
            // Draw the legend

            fontSize = clsConstants.legendTextSize;


            // Draw legend background (blue rectangle)
            
            
            g.DrawString("COLLINS diagram display of concentrations (meq/L) ( not ratios )\n for individual samples  in a cumulative chart, Total height reflects\n the difference in TDS", new Font("Times New Roman", fontSize, FontStyle.Bold), Brushes.Black, xOrigin + 30, yOrigin);

            #region Draw Legend

            if (frmImportSamples.WaterData.Count > 0)
            {
                int metaX = (int)(0.69f * frmMainForm.mainChartPlotting.Width);
                int metaY = (int)(0.13f * frmMainForm.mainChartPlotting.Height);

                int metaHeight = 0;
                int legendtextSize = clsConstants.legendTextSize;
                int metaWidth = (int)(0.2 * frmMainForm.mainChartPlotting.Width);

                using (Font font = new Font("Times New Roman", legendtextSize, FontStyle.Bold))
                {
                    StringFormat stringFormat = new StringFormat { FormatFlags = StringFormatFlags.NoClip, Trimming = StringTrimming.EllipsisCharacter };

                    for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                    {
                        string fullText = "W" + (i + 1).ToString() + ", " + frmImportSamples.WaterData[i].Well_Name + ", " + frmImportSamples.WaterData[i].ClientID + ", " + frmImportSamples.WaterData[i].Depth;
                        SizeF textSize = g.MeasureString(fullText, font, metaWidth, stringFormat); // Adjust for wrapping width
                        metaWidth = (int)Math.Max(metaWidth, textSize.Width); // Ensure metaWidth accounts for the largest text
                        metaHeight += (int)Math.Ceiling(textSize.Height + 10); // Add spacing between lines
                    }
                }

                Bitmap metaBitmap = new Bitmap(metaWidth, metaHeight);
                PictureBox metaPictureBox = new PictureBox
                {
                    Size = new Size(metaWidth, metaHeight),
                    Image = metaBitmap
                };

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
                    //Brush squareBrush = new SolidBrush(data.color);
                    //Pen axisPen = new Pen(data.color, data.lineWidth)
                    //{
                    //    DashStyle = data.selectedStyle
                    //};

                    //g.DrawLine(axisPen, 5, ysample + 10, 25, ysample + 10);

                    string fullText = "W" + (i + 1).ToString() + ", " + data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
                    RectangleF textRect = new RectangleF(0, ysample, metaWidth, metaHeight);

                    Font font = new Font("Times New Roman", legendtextSize, FontStyle.Bold);
                    SizeF textSize = g.MeasureString(fullText, font, metaWidth); // Adjust for wrapping width

                    g.DrawString(fullText, font, Brushes.Black, textRect);
                    ysample += (int)Math.Ceiling(textSize.Height + 10); // Move down based on wrapped height
                }

                frmMainForm.metaPanel.Location = new Point(metaX - 14, metaY - 9);
                frmMainForm.mainChartPlotting.Controls.Add(frmMainForm.metaPanel);
                int legendX = (int)(0.1f * frmMainForm.mainChartPlotting.Width);
                int legendY = (int)(0.1f * frmMainForm.mainChartPlotting.Height);
                int s = 0;
                for (int i = 0; i < labels.Length; i++)
                {

                    string fullText = labels[i];
                    SizeF textSize = g.MeasureString(fullText, new Font("Times New Roman", clsConstants.legendTextSize));
                    s = (int)(s + textSize.Width + 40);
                }


                int legendBoxHeight = (int)(0.03f * frmMainForm.mainChartPlotting.Height);
                int legendBoxWidth = s;


                //Form1.pic.Visible = true;
                frmMainForm.legendPictureBox.Size = new Size(legendBoxWidth, legendBoxHeight);
                Bitmap bit = new Bitmap(legendBoxWidth, legendBoxHeight);
                g = Graphics.FromImage(bit);
                //g.DrawRectangle(new Pen(Color.Blue), legendX - 15.0f, legendY - 10.0f, legendBoxWidth + 15.0f, legendBoxHeight + 30.0f);
                int xsample = legendX;


                using (Graphics legendGraphics = g)
                {
                    //legendGraphics.Clear(Color.White);  // Fill background
                    legendGraphics.FillRectangle(Brushes.White, 0, 0, legendBoxWidth - 5, legendBoxHeight - 5);
                    legendGraphics.DrawRectangle(new Pen(Color.Blue, 2), 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                    xsample = 0;
                    for (int i = 0; i < labels.Length; i++)
                    {
                        Brush myBrush = new SolidBrush(legendColors[i]);
                        
                        if(frmCollinsLegend.IsUpdateClicked)
                        {
                            myBrush = new SolidBrush(clsCollinsDrawer.legendColors[i]);
                            
                        }
                        legendGraphics.FillRectangle(myBrush, xsample+5, 2, 18, 18);


                        // Draw text beside the shape
                        legendGraphics.DrawString(labels[i], new Font("Times New Roman", fontSize), Brushes.Black, xsample + 25, 5);

                        string fullText = labels[i];
                        SizeF textSize = g.MeasureString(fullText, new Font("Times New Roman", fontSize));
                        xsample += (int)textSize.Width + 40;
                    }
                }
                //Form1.legendPanel.BackColor = Color.Transparent;
                frmMainForm.legendPanel.Location = new Point(legendX - 14, legendY - 9);
                frmMainForm.legendPanel.Size = new System.Drawing.Size(legendBoxWidth, legendBoxHeight);
                frmMainForm.legendPictureBox.Image = bit;
                //Form1.pic.Location = new Point(0, 0);
                frmMainForm.legendPictureBox.Visible = true;
                frmMainForm.legendPictureBox.MouseDoubleClick += frmMainForm.pictureBoxCollins_Click;
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
        /// Exports the Collins Diagram to a PowerPoint slide.
        /// </summary>
        public static void ExportCollinsToPowerPoint(PowerPoint.Slide slide, PowerPoint.Presentation presentation)
        {


            int chartWidth = (int)presentation.PageSetup.SlideWidth;
            int chartHeight = (int)presentation.PageSetup.SlideHeight;

            // Calculate center position
            int diagramWidth = 450; // Fixed width for the diagram
            int diagramHeight = (int)(0.7 * chartHeight); // Fixed height for the diagram
            int x1 = (int)(0.1f * chartWidth); // Center horizontally
            int y1 = 100; // Center vertically

            // factors
            double Nafac = 22.99, Kfac = 39.0983, Cafac = 20.039, Mgfac = 12.1525, Clfac = 35.453, HCO3fac = 61.01684, CO3fac = 30.004, SO4fac = 48.0313;
            // Dummy data for Collins diagram
            double totaltds = 0.0;

            List<string> samples = new List<string>();
            for (int i = 1; i <= frmImportSamples.WaterData.Count; i++)
            {
                samples.Add("W" + i.ToString());
            }

            double[] Na = new double[frmImportSamples.WaterData.Count];
            double[] K = new double[frmImportSamples.WaterData.Count];
            double[] Ca = new double[frmImportSamples.WaterData.Count];
            double[] Mg = new double[frmImportSamples.WaterData.Count];
            double[] Cl = new double[frmImportSamples.WaterData.Count];
            double[] HCO3 = new double[frmImportSamples.WaterData.Count];
            double[] CO3 = new double[frmImportSamples.WaterData.Count];
            double[] SO4 = new double[frmImportSamples.WaterData.Count];

            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                Na[i] += frmImportSamples.WaterData[i].Na;
                totaltds += frmImportSamples.WaterData[i].TDS;
                K[i] += frmImportSamples.WaterData[i].K;
                Ca[i] += frmImportSamples.WaterData[i].Ca;
                Mg[i] += frmImportSamples.WaterData[i].Mg;
                Cl[i] += frmImportSamples.WaterData[i].Cl;
                SO4[i] += frmImportSamples.WaterData[i].So4;
                HCO3[i] += frmImportSamples.WaterData[i].HCO3;
                CO3[i] += frmImportSamples.WaterData[i].CO3;
            }
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                Na[i] /= Nafac;
                K[i] /= Kfac;
                Ca[i] /= Cafac;
                Mg[i] /= Mgfac;
                Cl[i] /= Clfac;
                HCO3[i] /= HCO3fac;
                CO3[i] /= CO3fac;
                SO4[i] /= SO4fac;
            }
            double Max_Value = 0;
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                Max_Value = Math.Max(Max_Value, (Na[i] + K[i] + Ca[i] + Mg[i]));
                Max_Value = Math.Max(Max_Value, (Cl[i] + HCO3[i] + CO3[i] + SO4[i]));
            }
            Max_Value = Max_Value * 1.1;
            // Axis scaling
            double fx = diagramWidth / (samples.Count + 1);
            double fy = diagramHeight / 3000.0;
            double maxtds = frmImportSamples.WaterData.Max(w => w.TDS);

            // Add title
            PowerPoint.Shape title = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                (presentation.PageSetup.SlideWidth / 2) - 100, clsConstants.chartYPowerpoint, 200, 50);
            title.TextFrame.TextRange.Text = "COLLINS DIAGRAM";
            title.TextFrame.TextRange.Font.Name = "Times New Roman";
            title.TextFrame.TextRange.Font.Size = 25;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            title.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            //title.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            title.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;

            // Draw X-axis labels
            for (int i = 0; i < samples.Count; i++)
            {
                float x = x1 + (float)(fx * (i + 1));
                var xTick = slide.Shapes.AddLine(x, y1 + (int)diagramHeight, x, y1 + (int)diagramHeight + 10);
                xTick.Line.ForeColor.RGB = Color.Black.ToArgb();
                slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                        x - 10, y1 + (int)diagramHeight + 15, 50, 15)
                            .TextFrame.TextRange.Text = samples[i];
                slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                slide.Shapes[slide.Shapes.Count].TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                slide.Shapes[slide.Shapes.Count].TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                slide.Shapes[slide.Shapes.Count].TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;

                // Remove margins to reduce waste of space
                slide.Shapes[slide.Shapes.Count].TextFrame.MarginLeft = 0;
                slide.Shapes[slide.Shapes.Count].TextFrame.MarginRight = 0;
                slide.Shapes[slide.Shapes.Count].TextFrame.MarginTop = 0;
                slide.Shapes[slide.Shapes.Count].TextFrame.MarginBottom = 0;
            }

            // Draw Y-axis labels
            for (int i = 0; i <= 3000; i += 500)
            {
                float y = y1 + (int)diagramHeight - (float)(fy * i);
                var yTick = slide.Shapes.AddLine(x1 - 10, y, x1, y);
                yTick.Line.ForeColor.RGB = Color.Black.ToArgb();
                slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                        x1 - 50, y - 10, 60, 15)
                            .TextFrame.TextRange.Text = i.ToString();
                slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                slide.Shapes[slide.Shapes.Count].TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                slide.Shapes[slide.Shapes.Count].TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                slide.Shapes[slide.Shapes.Count].TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue;

                // Remove margins to reduce waste of space
                slide.Shapes[slide.Shapes.Count].TextFrame.MarginLeft = 0;
                slide.Shapes[slide.Shapes.Count].TextFrame.MarginRight = 0;
                slide.Shapes[slide.Shapes.Count].TextFrame.MarginTop = 0;
                slide.Shapes[slide.Shapes.Count].TextFrame.MarginBottom = 0;
            }

            var verticalAxis = slide.Shapes.AddLine(x1, y1, x1, y1 + (int)diagramHeight);
            verticalAxis.Line.ForeColor.RGB = Color.Black.ToArgb();  // Set color to black
            verticalAxis.Line.Weight = 1;
            var rightAxis = slide.Shapes.AddLine(x1 + (int)diagramWidth, y1, x1 + (int)diagramWidth, y1 + (int)diagramHeight);
            rightAxis.Line.ForeColor.RGB = Color.Black.ToArgb();
            rightAxis.Line.Weight = 1;
            // Add horizontal axis (bottom)
            var horizontalAxis = slide.Shapes.AddLine(x1, y1 + (int)diagramHeight, x1 + diagramWidth, y1 + (int)diagramHeight);
            horizontalAxis.Line.ForeColor.RGB = Color.Black.ToArgb();  // Set color to black
            horizontalAxis.Line.Weight = 1;
            var topAxis = slide.Shapes.AddLine(x1, y1, x1 + (int)diagramWidth, y1);
            topAxis.Line.ForeColor.RGB = Color.Black.ToArgb();
            topAxis.Line.Weight = 1;
            // Draw stacked bars for each sample
            Color[] colors = { Color.Cyan, Color.Purple, Color.DarkSlateBlue, Color.Yellow, Color.Magenta, Color.Green };
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {


                float width = 10;

                double NaK = Na[i] + K[i], HCO3CO3 = HCO3[i] + CO3[i];
                int x = x1 + (int)(fx * (i + 1)) - 10;
                double currentY = y1 + diagramHeight;
                // First bar: Na+K, Ca, Mg
                double temp = (double)diagramHeight;
                double fy_F = (double)(diagramHeight / (double)Max_Value);
                double heightPart = (NaK * fy_F);

                // Na+K
                PowerPoint.Shape rectangle = slide.Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    x,                          // Left
                    (float)(currentY - heightPart), // Top
                    width,                      // Width
                    (float)heightPart           // Height
                );

                // Fill the rectangle with orange color
                if (!frmCollinsLegend.IsUpdateClicked)
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(legendColors[0]);
                }
                else 
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(frmCollinsLegend.CollinsColor[0]);
                }
                currentY -= heightPart;
                //temp -= heightPart;

                heightPart = (Ca[i] * fy_F);
                rectangle = slide.Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    x,                          // Left
                    (float)(currentY - heightPart), // Top
                    width,                      // Width
                    (float)heightPart           // Height
                );
                if (!frmCollinsLegend.IsUpdateClicked)
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(legendColors[1]);
                }
                else
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(frmCollinsLegend.CollinsColor[1]);
                }
                currentY -= heightPart;
                //temp -= heightPart;

                heightPart = Mg[i] * fy_F;
                rectangle = slide.Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    x,                          // Left
                    (float)(currentY - heightPart), // Top
                    width,                      // Width
                    (float)heightPart           // Height
                );
                if (!frmCollinsLegend.IsUpdateClicked)
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(legendColors[2]);
                }
                else
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(frmCollinsLegend.CollinsColor[2]);
                }
                //temp = (double)diagramHeight;
                x += 10; // Shift for second bar
                currentY = y1 + diagramHeight;

                heightPart = (Cl[i] * fy_F);
                rectangle = slide.Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    x,                          // Left
                    (float)(currentY - heightPart), // Top
                    width,                      // Width
                    (float)heightPart           // Height
                );
                if (!frmCollinsLegend.IsUpdateClicked)
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(legendColors[3]);
                }
                else
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(frmCollinsLegend.CollinsColor[3]);
                }
                currentY -= heightPart;
                //temp -= heightPart;

                heightPart = (HCO3CO3 * fy_F);
                rectangle = slide.Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    x,                          // Left
                    (float)(currentY - heightPart), // Top
                    width,                      // Width
                    (float)heightPart           // Height
                );
                if (!frmCollinsLegend.IsUpdateClicked)
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(legendColors[4]);
                }
                else
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(frmCollinsLegend.CollinsColor[4]);
                }
                currentY -= heightPart;
                //temp -= heightPart;

                heightPart = (SO4[i] * fy_F);
                rectangle = slide.Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                    x,                          // Left
                    (float)(currentY - heightPart), // Top
                    width,                      // Width
                    (float)heightPart           // Height
                );
                if (!frmCollinsLegend.IsUpdateClicked)
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(legendColors[5]);
                }
                else
                {
                    rectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(frmCollinsLegend.CollinsColor[5]);
                }
            }
            string[] legendItems = { "Na+K", "Ca", "Mg", "Cl", "SO4", "HCO3 + CO3" };
            #region Draw Legend
            int legendX = x1;
            int legendY = 50;
            int xSample = legendX + 5;
            int fontSize = clsConstants.legendTextSize;
            int legendBoxWidth = 0;
            int legendBoxHeight = 0;
            // Add legend border


            for (int i = 0; i < labels.Length; i++)
            {
                // Create a temp shape just to measure
                var temp = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    xSample + 50, legendY + 10, 100, 20);
                temp.TextFrame.TextRange.Text = labels[i];
                temp.TextFrame.TextRange.Font.Size = fontSize;
                temp.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                temp.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                temp.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

                temp.TextFrame.TextRange.Text = labels[i];
                temp.TextFrame.TextRange.Font.Size = fontSize;
                temp.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                int charCount = labels[i].Length;
                float approxWidth = fontSize * charCount * 0.9f;
                temp.Width = approxWidth;
                legendBoxWidth += (int)temp.Width + 10;
                legendBoxHeight = Math.Max(legendBoxHeight, (int)temp.Height);
                temp.Delete(); // clean up

            }
            PowerPoint.Shape legendBorder = slide.Shapes.AddShape(
            Office.MsoAutoShapeType.msoShapeRectangle,
            legendX, legendY, legendBoxWidth, legendBoxHeight);
            legendBorder.Fill.Transparency = 1.0f;
            legendBorder.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue);
            legendBorder.Line.Weight = 2;
            for (int i = 0; i < labels.Length; i++)
            {
                if (!frmPieLegend.IsUpdateClicked)
                {
                    PowerPoint.Shape legendBox = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        xSample, legendY + 5, 10, 10);
                    legendBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(legendColors[i]);
                }
                else
                {
                    PowerPoint.Shape legendBox = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        xSample, legendY + 5, 10, 10);
                    legendBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(frmPieLegend.PieColor[i]);
                }

                PowerPoint.Shape legendText = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    xSample + 10, legendY, 100, 20);
                legendText.TextFrame.TextRange.Text = labels[i];
                legendText.TextFrame.TextRange.Font.Size = fontSize;
                legendText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                legendText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                legendText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                legendText.TextFrame.MarginLeft = 0;
                legendText.TextFrame.MarginRight = 0;
                legendText.TextFrame.MarginTop = 0;
                legendText.TextFrame.MarginBottom = 0;
                int charCount = labels[i].Length;
                float approxWidth = fontSize * charCount * 0.9f; // 0.6 is a rough factor

                legendText.Width = approxWidth;

                xSample += (int)legendText.Width + 10;
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
                string fullText = "W" + (i + 1).ToString() + ", " + data.Well_Name + ", " + data.ClientID + ", " + data.Depth;

                // Add textbox with wrapping and fixed width
                PowerPoint.Shape metadataText = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    metadataX + 5, ysample, metaWidth, 20); // initial height, PowerPoint will auto-expand

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
            #endregion



        }
    }
}
