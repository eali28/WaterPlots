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


namespace WaterPlots
{
    public class clsStiffDrawer
    {
        public static string[] labels = { "Na+K", "Ca", "Mg", "Cl", "SO4", "HCO3 + CO3" };
        /// <summary>
        /// Draws the Stiff Diagram, plotting cation and anion concentrations for each sample.
        /// </summary>
        public static void DrawStiffDiagram(Graphics g)
        {
            // Calculate center position
            int leftMargin = (int)(0.1 * frmMainForm.mainChartPlotting.Width);
            int topMargin = (int)(0.01 * frmMainForm.mainChartPlotting.Height);

            // Diagram Dimensions
            int diagramWidth = (int)(0.5f * frmMainForm.mainChartPlotting.Width); // Make width relative
            float diagramHeight = (float)(frmImportSamples.WaterData.Count * 0.04f * frmMainForm.mainChartPlotting.Height);
            
            int yOrigin = (int)(topMargin + (frmMainForm.mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * frmMainForm.mainChartPlotting.Height)+diagramHeight);
            float fontSize = clsConstants.legendTextSize;
            Font titleFont = new Font("Times New Roman", 25, FontStyle.Bold);
            // Fonts and Pens
            Font labelFont = new Font("Times New Roman", fontSize, FontStyle.Regular);
            Pen axisPen = new Pen(Color.Black, 1f);
            Pen linePen = new Pen(Color.Black, 1f);
            Brush[] ionColors = { Brushes.Cyan, Brushes.Orange, Brushes.Purple, Brushes.Blue, Brushes.Magenta, Brushes.Green };

            // Draw Title
            string title = "STIFF DIAGRAM";
            g.DrawString(title, titleFont, Brushes.Black, 0.4f*frmMainForm.mainChartPlotting.Width, 0.01f*frmMainForm.mainChartPlotting.Height);

            // Draw Axes
            float axisHalfLength = diagramWidth / 2f;
            int xOrigin = (int)(leftMargin + 0.03f * frmMainForm.mainChartPlotting.Width+axisHalfLength);
            g.DrawLine(axisPen, xOrigin-axisHalfLength, yOrigin, xOrigin+axisHalfLength, yOrigin);
            g.DrawLine(axisPen, xOrigin - axisHalfLength, yOrigin-diagramHeight, xOrigin + axisHalfLength, yOrigin-diagramHeight);
            g.DrawLine(axisPen, xOrigin, yOrigin, xOrigin, yOrigin - diagramHeight);
            g.DrawLine(axisPen, xOrigin+axisHalfLength, yOrigin, xOrigin+axisHalfLength, yOrigin - diagramHeight);
            g.DrawLine(axisPen, xOrigin-axisHalfLength, yOrigin, xOrigin-axisHalfLength, yOrigin - diagramHeight);
            // Draw X-axis ticks (10% increments)
            int numTicks = 10;
            float tickSpacing = axisHalfLength / numTicks;

            g.DrawLine(axisPen, xOrigin, yOrigin - 5, xOrigin, yOrigin + 5);
            g.DrawString((0).ToString(), labelFont, Brushes.Black, xOrigin - 5, yOrigin + 8);

            for (int i = 1; i <= numTicks; i++)
            {
                float offset = i * tickSpacing;

                // Right side (positive)
                g.DrawLine(axisPen, xOrigin + offset, yOrigin - 5, xOrigin + offset, yOrigin + 5);
                g.DrawString((i*10).ToString(), labelFont, Brushes.Black, xOrigin + offset - 5, yOrigin + 8);

                // Left side (negative)
                g.DrawLine(axisPen, xOrigin - offset, yOrigin - 5, xOrigin - offset, yOrigin + 5);
                g.DrawString((i * 10).ToString(), labelFont, Brushes.Black, xOrigin - offset - 10, yOrigin + 8);
            }


            List<PointF> Points = new List<PointF>();
            double Nafac = 22.99, Kfac = 39.0983, Cafac = 20.039, Mgfac = 12.1525, Clfac = 35.453, HCO3fac = 61.01684, CO3fac = 30.004, SO4fac = 48.0313;
            int totalSamples = frmImportSamples.WaterData.Count;
            int sampleSpacing = (int)(totalSamples > 0 ? (diagramHeight) / totalSamples : 0); // Adjusted spacing
            int offsetY = yOrigin - (int)(0.02f*diagramHeight); // Start from top with some margin
            // Plot each sample
            int maxPoint=0;
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                var existingTuple = frmImportSamples.WaterData[i];
                double Nab, Kb, Mgb, Cab, Clb, SO4b, HCO3b, CO3b;

                Nab = existingTuple.Na / Nafac;
                Kb = existingTuple.K / Kfac;
                Mgb = existingTuple.Mg / Mgfac;
                Cab = existingTuple.Ca / Cafac;
                Clb = existingTuple.Cl / Clfac;
                SO4b = existingTuple.So4 / SO4fac;
                HCO3b = existingTuple.HCO3 / HCO3fac;
                CO3b = existingTuple.CO3 / CO3fac;
                double total = Nab + Kb + Mgb + Cab + Clb + SO4b + HCO3b + CO3b;

                // Calculate percentages for each component
                double NaK = (Nab + Kb) / total;
                double Mg = Mgb / total;
                double Ca = Cab / total;
                double Cl = Clb / total;
                double So4 = SO4b / total;
                double HCO3CO3 = (HCO3b + CO3b) / total;

                // Left points (cations)
                List<PointF> leftPoints = new List<PointF>
                    {
                        new PointF(xOrigin - (float)(NaK * axisHalfLength), offsetY - 10), // Na+K
                        new PointF(xOrigin - (float)(Mg  * axisHalfLength), offsetY),      // Mg
                        new PointF(xOrigin - (float)(Ca  * axisHalfLength), offsetY + 10) // Ca
                    };

                // Right points (anions)
                List<PointF> rightPoints = new List<PointF>
                    {
                        new PointF(xOrigin + (float)(Cl  * axisHalfLength), offsetY - 10),  // Cl
                        new PointF(xOrigin + (float)(So4* axisHalfLength), offsetY),     // SO4
                        new PointF(xOrigin + (float)(HCO3CO3 * axisHalfLength), offsetY + 10) // HCO3+CO3
                    };
                Points.Add(new PointF((xOrigin + (float)(Cl * axisHalfLength)), offsetY - 20));
                if ((xOrigin + (float)(Cl * axisHalfLength)) > maxPoint)
                {
                    maxPoint = (int)(xOrigin + (float)(Cl * axisHalfLength));
                }

                // Draw connecting lines between cations and anions within the same sample
                g.DrawLine(linePen, leftPoints[0], rightPoints[0]);
                g.DrawLine(linePen, leftPoints[0], leftPoints[1]);
                g.DrawLine(linePen, leftPoints[1], leftPoints[2]);
                g.DrawLine(linePen, leftPoints[2], rightPoints[2]);
                g.DrawLine(linePen, rightPoints[2], rightPoints[1]);
                g.DrawLine(linePen, rightPoints[1], rightPoints[0]);

                // Draw individual points with colors
                for (int j = 0; j < leftPoints.Count; j++)
                {
                    g.FillEllipse(ionColors[j], leftPoints[j].X - 4, leftPoints[j].Y - 4, 8, 8);
                }

                for (int j = 0; j < rightPoints.Count; j++)
                {
                    g.FillEllipse(ionColors[j + 3], rightPoints[j].X - 4, rightPoints[j].Y - 4, 8, 8);
                }

                offsetY -= sampleSpacing; // Move down for the next sample
            }
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                g.DrawString("W" + (i + 1).ToString(), labelFont, Brushes.Black, maxPoint+10,Points[i].Y);
            }
            g.DrawString("meq/L", new Font("Times New Roman", 15, FontStyle.Bold), Brushes.Black, xOrigin - 0.01f*diagramWidth, yOrigin + 30);
            g.DrawString("Cations", new Font("Times New Roman", 15, FontStyle.Bold), Brushes.Black, xOrigin - 0.25f * diagramWidth, yOrigin + 30);
            g.DrawString("Anions", new Font("Times New Roman", 15, FontStyle.Bold), Brushes.Black, xOrigin + 0.25f*diagramWidth, yOrigin + 30);
            string subtitle = "STIFF diagram displaying concentration ratios (meq/L) for individual samples.";
            g.DrawString(subtitle, labelFont, Brushes.Black, 0.2f * frmMainForm.mainChartPlotting.Width, 0.9f * frmMainForm.mainChartPlotting.Height);
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
                            fullText +=  data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
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
                    RectangleF textRect = new RectangleF(30, ysample, metaWidth-35, metaHeight);

                    Font font = new Font("Times New Roman", legendtextSize, FontStyle.Bold);
                    SizeF textSize = g.MeasureString(fullText, font, (int)textRect.Width); // Adjust for wrapping width
                    g.DrawString("W" + (i + 1).ToString()+", ", font, Brushes.Black, 0, ysample);
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
                        
                        Brush myBrush = ionColors[i];
                        legendGraphics.FillEllipse(myBrush, xsample + 5, 2, 18, 18);
                        

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


        /// <summary>
        /// Exports the Stiff Diagram to a PowerPoint slide.
        /// </summary>
        public static void ExportStiffDiagramToPowerPoint(PowerPoint.Slide slide, PowerPoint.Presentation presentation)
        {

            // Define Diagram Position
            int diagramWidth = 450;
            int diagramHeight = (int)(0.6*(int)presentation.PageSetup.SlideHeight);
            int xOrigin = (int)(0.6f * diagramWidth);
            int yOrigin = 100;
            float axisHalfLength = diagramWidth / 2f;

            // Add Title
            PowerPoint.Shape title = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                (presentation.PageSetup.SlideWidth / 2) - 100, clsConstants.chartYPowerpoint, 200, 50);
            title.TextFrame.TextRange.Text = "Stiff Diagram";
            title.TextFrame.TextRange.Font.Size = 25;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            title.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            title.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Draw Axes
            PowerPoint.Shape CenterVerticalLine = slide.Shapes.AddLine(xOrigin, yOrigin, xOrigin, yOrigin + diagramHeight);
            CenterVerticalLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
            PowerPoint.Shape RightVerticalLine = slide.Shapes.AddLine(xOrigin+axisHalfLength, yOrigin, xOrigin+axisHalfLength, yOrigin + diagramHeight);
            RightVerticalLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
            PowerPoint.Shape LeftVerticalLine = slide.Shapes.AddLine(xOrigin-axisHalfLength, yOrigin, xOrigin-axisHalfLength, yOrigin + diagramHeight);
            LeftVerticalLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
            // Draw Horizontal Axis
            PowerPoint.Shape bottomHorizontalLine = slide.Shapes.AddLine(
                xOrigin - axisHalfLength,
                yOrigin + diagramHeight,
                xOrigin + axisHalfLength,
                yOrigin + diagramHeight
            );
            bottomHorizontalLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black); // Set color to black
            PowerPoint.Shape topHorizontalLine = slide.Shapes.AddLine(
                xOrigin - axisHalfLength,
                yOrigin,
                xOrigin + axisHalfLength,
                yOrigin
            );
            topHorizontalLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black); // Set color to black

            // Draw X-axis Ticks
            int numTicks = 10;
            float tickSpacing = axisHalfLength / numTicks;
            double fx = axisHalfLength / numTicks; // X-axis scale
            for (int i = 0; i <= numTicks; i++)
            {
                float x = xOrigin + (float)(fx * i);

                slide.Shapes.AddLine(x, yOrigin + (int)diagramHeight, x, yOrigin + (int)diagramHeight + 10).Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);

                var textBox = slide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    x - 10, yOrigin + (int)diagramHeight + 15, 50, 15);

                textBox.TextFrame.TextRange.Text = (i * 10).ToString();
                textBox.TextFrame.TextRange.Font.Size = 8;
            }

            for (int i = 1; i <= numTicks; i++)
            {
                float x = xOrigin - i * tickSpacing;

                slide.Shapes.AddLine(x, yOrigin + (int)diagramHeight, x, yOrigin + (int)diagramHeight + 10).Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);

                var textBox = slide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    x - 10, yOrigin + (int)diagramHeight + 15, 50, 15);

                textBox.TextFrame.TextRange.Text = (i * 10).ToString();
                textBox.TextFrame.TextRange.Font.Size = 8;
            }


            // Sample Plotting
            int totalSamples = frmImportSamples.WaterData.Count;
            int sampleSpacing = totalSamples > 0 ? diagramHeight / totalSamples : 0;
            int offsetY = yOrigin + diagramHeight-10;
            Color[] ionColors = { Color.Cyan, Color.Orange, Color.Purple, Color.Blue, Color.Magenta, Color.Green };
            double Nafac = 22.99, Kfac = 39.0983, Cafac = 20.039, Mgfac = 12.1525, Clfac = 35.453, HCO3fac = 61.01684, CO3fac = 30.004, SO4fac = 48.0313;
            List<PointF> Points = new List<PointF>();
            float maxX = 0;
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                var existingTuple = frmImportSamples.WaterData[i];
                double Nab, Kb, Mgb, Cab, Clb, SO4b, HCO3b, CO3b;

                Nab = existingTuple.Na / Nafac;
                Kb = existingTuple.K / Kfac;
                Mgb = existingTuple.Mg / Mgfac;
                Cab = existingTuple.Ca / Cafac;
                Clb = existingTuple.Cl / Clfac;
                SO4b = existingTuple.So4 / SO4fac;
                HCO3b = existingTuple.HCO3 / HCO3fac;
                CO3b = existingTuple.CO3 / CO3fac;
                double total = Nab + Kb + Mgb + Cab + Clb + SO4b + HCO3b + CO3b;

                // Calculate percentages for each component
                double NaK = (Nab + Kb) / total;
                double Mg = Mgb / total;
                double Ca = Cab / total;
                double Cl = Clb / total;
                double So4 = SO4b / total;
                double HCO3CO3 = (HCO3b + CO3b) / total;

                // Left Points (Cations)
                float[] leftX = { xOrigin - (float)(NaK * axisHalfLength), xOrigin - (float)(Mg * axisHalfLength), xOrigin - (float)(Ca * axisHalfLength) };
                float[] leftY = { offsetY - 10, offsetY, offsetY + 10 };

                // Right Points (Anions)
                float[] rightX = { xOrigin + (float)(Cl * axisHalfLength), xOrigin + (float)(So4 * axisHalfLength), xOrigin + (float)(HCO3CO3 * axisHalfLength) };
                float[] rightY = { offsetY - 10, offsetY, offsetY + 10 };

                // Connect Cations and Anions with Black Lines
                PowerPoint.Shape line1 = slide.Shapes.AddLine(leftX[0], leftY[0], rightX[0], rightY[0]);
                line1.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);

                PowerPoint.Shape line2 = slide.Shapes.AddLine(leftX[0], leftY[0], leftX[1], leftY[1]);
                line2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);

                PowerPoint.Shape line3 = slide.Shapes.AddLine(leftX[1], leftY[1], leftX[2], leftY[2]);
                line3.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);

                PowerPoint.Shape line4 = slide.Shapes.AddLine(leftX[2], leftY[2], rightX[2], rightY[2]);
                line4.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);

                PowerPoint.Shape line5 = slide.Shapes.AddLine(rightX[2], rightY[2], rightX[1], rightY[1]);
                line5.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);

                PowerPoint.Shape line6 = slide.Shapes.AddLine(rightX[1], rightY[1], rightX[0], rightY[0]);
                line6.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                maxX = Math.Max(maxX, rightX[0] - 15);
                




                // Draw Circles for Each Ion Group
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, leftX[0] - 4, leftY[0] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[0]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, leftX[1] - 4, leftY[1] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[1]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, leftX[2] - 4, leftY[2] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[2]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, rightX[0] - 4, rightY[0] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[3]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, rightX[1] - 4, rightY[1] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[4]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, rightX[2] - 4, rightY[2] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[5]);
                Points.Add(new PointF(rightX[0], rightY[0]));
                maxX = Math.Max(maxX, rightX[0]);
                offsetY -= sampleSpacing; // Move Down for Next Sample
            }
            offsetY = yOrigin + diagramHeight - 10;
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                float[] rightY = { offsetY - 10, offsetY, offsetY + 10 };

                var temp = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    maxX, rightY[0] - 10, 100, 20);
                temp.TextFrame.TextRange.Text = "W" + (i + 1).ToString();
                temp.TextFrame.TextRange.Font.Size = 10;
                temp.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                temp.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                temp.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                temp.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
                offsetY -= sampleSpacing; // Move Down for Next Sample
            }


                #region Draw Legend
                int legendX = (int)(0.1f * presentation.PageSetup.SlideWidth);
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
                
                    PowerPoint.Shape legendBox = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeOval,
                        xSample, legendY + 5, 10, 10);
                    legendBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(ionColors[i]);
                

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
                    metadataX-40, ysample-5, 100, 20);
                temp.TextFrame.TextRange.Text = "W" + (i + 1).ToString()+", ";
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
            string cation = "Cations";
            PowerPoint.Shape cationShape = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                xOrigin - axisHalfLength/2-10, yOrigin + diagramHeight + 30, 600, 20
            );
            cationShape.TextFrame.TextRange.Text = cation;
            cationShape.TextFrame.TextRange.Font.Size = 15;
            cationShape.TextFrame.TextRange.Font.Name = "Times New Roman";
            cationShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Black text
            cationShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            cationShape.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            string meqL = "meq/L";
            PowerPoint.Shape meqLShape = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                xOrigin-10, yOrigin + diagramHeight + 30, 600, 20
            );
            meqLShape.TextFrame.TextRange.Text = meqL;
            meqLShape.TextFrame.TextRange.Font.Size = 15;
            meqLShape.TextFrame.TextRange.Font.Name = "Times New Roman";
            meqLShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Black text
            meqLShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            meqLShape.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            string anions = "Anions";
            PowerPoint.Shape anionsShape = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                xOrigin + axisHalfLength/2+10, yOrigin + diagramHeight + 30, 600, 20
            );
            anionsShape.TextFrame.TextRange.Text = anions;
            anionsShape.TextFrame.TextRange.Font.Size = 15;
            anionsShape.TextFrame.TextRange.Font.Name = "Times New Roman";
            anionsShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Black text
            anionsShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            anionsShape.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            // Add Subtitle at Bottom
            string subtitle = "STIFF diagram displaying concentration ratios (meq/L) for individual samples.";

            PowerPoint.Shape subtitleShape = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                xOrigin - 200, yOrigin + diagramHeight + 80, 600, 20
            );
            subtitleShape.TextFrame.TextRange.Text = subtitle;
            subtitleShape.TextFrame.TextRange.Font.Size = 8;
            subtitleShape.TextFrame.TextRange.Font.Name = "Times New Roman";
            subtitleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Black text
        }
    }
}
