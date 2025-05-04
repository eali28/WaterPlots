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
    public class clsStiffDrawer
    {
        public static string[] labels = { "Na+K", "Ca", "Mg", "Cl", "SO4", "HCO3 + CO3" };
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
            Pen axisPen = new Pen(Color.Black, 2f);
            Pen linePen = new Pen(Color.Black, 1f);
            Brush[] ionColors = { Brushes.Cyan, Brushes.Orange, Brushes.Purple, Brushes.Blue, Brushes.Magenta, Brushes.Green };

            // Draw Title
            string title = "STIFF DIAGRAM";
            g.DrawString(title, titleFont, Brushes.Black, 0.4f*frmMainForm.mainChartPlotting.Width, 0.01f*frmMainForm.mainChartPlotting.Height);

            // Draw Axes
            float axisHalfLength = diagramWidth / 2f;
            int xOrigin = (int)(leftMargin + 0.03f * frmMainForm.mainChartPlotting.Width+axisHalfLength);
            g.DrawLine(axisPen, xOrigin, yOrigin, xOrigin+axisHalfLength, yOrigin);
            g.DrawLine(axisPen, xOrigin, yOrigin, xOrigin - axisHalfLength, yOrigin);
            g.DrawLine(axisPen, xOrigin, yOrigin, xOrigin, yOrigin - diagramHeight);

            // Draw X-axis ticks (10% increments)
            int numTicks = 10;
            float tickSpacing = axisHalfLength / numTicks;

            for (int i = 1; i <= numTicks; i++)
            {
                float offset = i * tickSpacing;

                // Right side (positive)
                g.DrawLine(axisPen, xOrigin + offset, yOrigin - 5, xOrigin + offset, yOrigin + 5);
                g.DrawString(i.ToString(), labelFont, Brushes.Black, xOrigin + offset - 5, yOrigin + 8);

                // Left side (negative)
                g.DrawLine(axisPen, xOrigin - offset, yOrigin - 5, xOrigin - offset, yOrigin + 5);
                g.DrawString(i.ToString(), labelFont, Brushes.Black, xOrigin - offset - 10, yOrigin + 8);
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
            g.DrawString("meq/L", titleFont, Brushes.Black, xOrigin - 0.01f*diagramWidth, yOrigin + diagramHeight + 30);
            g.DrawString("Cations", titleFont, Brushes.Black, xOrigin - 0.25f * diagramWidth, yOrigin + diagramHeight + 30);
            g.DrawString("Anions", titleFont, Brushes.Black, xOrigin + 0.25f*diagramWidth, yOrigin + diagramHeight + 30);
            string subtitle = "STIFF diagram displaying concentration ratios (meq/L) for individual samples.";
            g.DrawString(subtitle, labelFont, Brushes.Black, xOrigin - 200, yOrigin + 60);
            #region Draw Legend

            if (frmImportSamples.WaterData.Count > 0)
            {

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

                int metaX = (int)(0.69f * frmMainForm.mainChartPlotting.Width);
                int metaY = (int)(0.13f * frmMainForm.mainChartPlotting.Height);
               


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


                //Form1.pic.Visible = true;
                frmMainForm.legendPictureBox.Size = new Size(legendBoxWidth, legendBoxHeight);
                Bitmap bit = new Bitmap(legendBoxWidth, legendBoxHeight);
                g = Graphics.FromImage(bit);
                g.DrawRectangle(new Pen(Color.Blue), legendX - 15.0f, legendY - 10.0f, legendBoxWidth + 15.0f, legendBoxHeight + 30.0f);
                int xsample = legendX;


                using (Graphics legendGraphics = g)
                {
                    //legendGraphics.Clear(Color.White);  // Fill background
                    legendGraphics.FillRectangle(Brushes.White, 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                    legendGraphics.DrawRectangle(new Pen(Color.Blue, 2), 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                    xsample = 0;
                    for (int i = 0; i < labels.Length; i++)
                    {

                        legendGraphics.FillEllipse(ionColors[i], xsample, 0, 20, 20);


                        // Draw text beside the shape
                        legendGraphics.DrawString(labels[i], new Font("Times New Roman", fontSize), Brushes.Black, xsample + 20, 5);

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
                // Draw Subtitle

            }
            else
            {
                frmMainForm.legendPanel.AutoScroll = false;
            }
            frmMainForm.legendPanel.Show();
            frmMainForm.mainChartPlotting.Invalidate();
            #endregion
            
        }


        public static void ExportStiffDiagramToPowerPoint(PowerPoint.Slide slide, PowerPoint.Presentation presentation)
        {

            // Define Diagram Position
            int chartWidth = 700;
            int chartHeight = (int)(frmMainForm.mainChartPlotting.Height / 1.3f);
            int xOrigin = 620;
            int yOrigin = 100;
            float axisHalfLength = chartWidth / 1.5f;

            // Add Title
            PowerPoint.Shape title = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                450, 20, 600, 50);
            title.TextFrame.TextRange.Text = "Stiff Diagram";
            title.TextFrame.TextRange.Font.Size = 55;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            title.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            title.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Draw Axes
            PowerPoint.Shape verticalLine = slide.Shapes.AddLine(xOrigin, yOrigin, xOrigin, yOrigin + chartHeight);
            verticalLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);

            // Draw Horizontal Axis
            PowerPoint.Shape horizontalLine = slide.Shapes.AddLine(
                xOrigin - axisHalfLength,
                yOrigin + chartHeight,
                xOrigin + axisHalfLength,
                yOrigin + chartHeight
            );
            horizontalLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black); // Set color to black


            // Draw X-axis Ticks
            int numTicks = 10;
            float tickSpacing = axisHalfLength / numTicks;
            double fx = axisHalfLength / numTicks; // X-axis scale
            for (int i = 0; i <= numTicks; i++)
            {
                float x = xOrigin + (float)(fx * (i));
                slide.Shapes.AddLine(x, yOrigin + (int)chartHeight, x, yOrigin + (int)chartHeight + 10);
                slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                        x - 10, yOrigin + (int)chartHeight + 15, 50, 15)
                            .TextFrame.TextRange.Text = (i * 10).ToString();
            }
            for (int i = 1; i <= numTicks; i++)
            {
                float x = xOrigin - i * tickSpacing;
                slide.Shapes.AddLine(x, yOrigin + (int)chartHeight, x, yOrigin + (int)chartHeight + 10);
                slide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                        x - 10, yOrigin + (int)chartHeight + 15, 50, 15)
                            .TextFrame.TextRange.Text = (i * 10).ToString();
            }

            // Sample Plotting
            int totalSamples = frmImportSamples.WaterData.Count;
            int sampleSpacing = totalSamples > 0 ? (chartHeight - 20) / totalSamples : 0;
            int offsetY = yOrigin + chartHeight - sampleSpacing;
            Color[] ionColors = { Color.Cyan, Color.Orange, Color.Purple, Color.Blue, Color.Gray, Color.Green };
            double Nafac = 22.99, Kfac = 39.0983, Cafac = 20.039, Mgfac = 12.1525, Clfac = 35.453, HCO3fac = 61.01684, CO3fac = 30.004, SO4fac = 48.0313;
            List<PointF> Points = new List<PointF>();
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
                float[] leftY = { offsetY - 20, offsetY, offsetY + 20 };

                // Right Points (Anions)
                float[] rightX = { xOrigin + (float)(Cl * axisHalfLength), xOrigin + (float)(So4 * axisHalfLength), xOrigin + (float)(HCO3CO3 * axisHalfLength) };
                float[] rightY = { offsetY - 20, offsetY, offsetY + 20 };

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



                // Draw Circles for Each Ion Group
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, leftX[0] - 4, leftY[0] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[0]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, leftX[1] - 4, leftY[1] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[1]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, leftX[2] - 4, leftY[2] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[2]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, rightX[0] - 4, rightY[0] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[3]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, rightX[1] - 4, rightY[1] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[4]);
                slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval, rightX[2] - 4, rightY[2] - 4, 8, 8).Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[5]);
                Points.Add(new PointF(rightX[0], rightY[0]));
                offsetY -= sampleSpacing; // Move Down for Next Sample
            }



            #region Draw Legend
            // Define legend position
            int legendX = (int)(xOrigin + chartWidth / 2 + 50);
            int legendY = (int)yOrigin;
            // Define labels and colors for ions
            string[] ionLabels = { "Na+K", "Mg", "Ca", "Cl", "SO4", "HCO3+CO3" };

            // Loop to create legend items
            for (int i = 0; i < ionLabels.Length; i++)
            {
                // Add Circle for Legend
                PowerPoint.Shape legendCircle = slide.Shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                    legendX, legendY, 20, 20
                );
                legendCircle.Fill.ForeColor.RGB = ColorTranslator.ToOle(ionColors[i]); // Set circle color
                legendCircle.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // Remove border

                // Add Text Label next to the Circle
                PowerPoint.Shape legendText = slide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    legendX + 20, legendY - 2, 70, 15
                );
                legendText.TextFrame.TextRange.Text = ionLabels[i];
                legendText.TextFrame.TextRange.Font.Size = 15;
                legendText.TextFrame.TextRange.Font.Name = "Times New Roman";
                legendText.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Set text color

                // Move to next legend item
                legendY += 27;
            }
            int legendBoxX = legendX - 15;
            int legendBoxY = legendY - 170;
            int legendBoxWidth = 150;
            int legendBoxHeight = 200;

            // Draw Transparent Legend Box with Blue Border
            PowerPoint.Shape legendBox = slide.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
                legendBoxX, legendBoxY, legendBoxWidth, legendBoxHeight
            );
            legendBox.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // Make it transparent
            legendBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue); // Blue border
            // Add metadata
            float metadataX = legendX - 20;
            float metadataY = legendY + 50;

            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                PowerPoint.Shape metadataText = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    metadataX, metadataY + (i * 20), 500, 20);
                metadataText.TextFrame.TextRange.Text = "W" + (i + 1).ToString() + "," + (frmImportSamples.WaterData[i].Well_Name) + "," + (frmImportSamples.WaterData[i].ClientID) + "," + (frmImportSamples.WaterData[i].Depth);
                metadataText.TextFrame.TextRange.Font.Size = 15;
            }
            #endregion
            // Position for Water Sample List
            int ysample = legendY + 50;

            // Loop Through Water Data Samples and Add Text
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                string sampleText = "W" + (i + 1).ToString();

                PowerPoint.Shape sampleTextShape = slide.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    Points[i].X + 15, Points[i].Y - 10, 700, 20
                );

                sampleTextShape.TextFrame.TextRange.Text = sampleText;
                sampleTextShape.TextFrame.TextRange.Font.Size = 15;
                sampleTextShape.TextFrame.TextRange.Font.Name = "Times New Roman";
                sampleTextShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                sampleTextShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Set text color to black

                ysample += 30; // Move down for the next sample
            }
            string cation = "Cations";
            PowerPoint.Shape cationShape = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                xOrigin - 200, yOrigin + chartHeight + 50, 600, 20
            );
            cationShape.TextFrame.TextRange.Text = cation;
            cationShape.TextFrame.TextRange.Font.Size = 20;
            cationShape.TextFrame.TextRange.Font.Name = "Times New Roman";
            cationShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Black text
            cationShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            string meqL = "meq/L";
            PowerPoint.Shape meqLShape = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                xOrigin, yOrigin + chartHeight + 50, 600, 20
            );
            meqLShape.TextFrame.TextRange.Text = meqL;
            meqLShape.TextFrame.TextRange.Font.Size = 20;
            meqLShape.TextFrame.TextRange.Font.Name = "Times New Roman";
            meqLShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Black text
            meqLShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            string anions = "Anions";
            PowerPoint.Shape anionsShape = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                xOrigin + 200, yOrigin + chartHeight + 50, 600, 20
            );
            anionsShape.TextFrame.TextRange.Text = anions;
            anionsShape.TextFrame.TextRange.Font.Size = 20;
            anionsShape.TextFrame.TextRange.Font.Name = "Times New Roman";
            anionsShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Black text
            anionsShape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            // Add Subtitle at Bottom
            string subtitle = "STIFF diagram displaying concentration ratios (meq/L) for individual samples.";

            PowerPoint.Shape subtitleShape = slide.Shapes.AddTextbox(
                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                xOrigin - 200, yOrigin + chartHeight + 80, 600, 20
            );
            subtitleShape.TextFrame.TextRange.Text = subtitle;
            subtitleShape.TextFrame.TextRange.Font.Size = 15;
            subtitleShape.TextFrame.TextRange.Font.Name = "Times New Roman";
            subtitleShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black); // Black text
        }
    }
}
