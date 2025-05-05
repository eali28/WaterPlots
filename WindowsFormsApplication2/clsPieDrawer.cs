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

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WindowsFormsApplication2
{
    public class clsPieDrawer
    {
        public static Color[] pieColors = { Color.Cyan, Color.Purple, Color.Orange, Color.Blue, Color.Gray, Color.Green };
        public static string[] labels = { "Na+K", "Ca", "Mg", "Cl", "SO4", "HCO3 + CO3" };
        public static void DrawPieChart(Graphics g, int chartWidth, int chartHeight)
        {
            // Detach the event handler if it is attached
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;

            int diagramWidth;
            if (frmImportSamples.WaterData.Count > 0)
            {
                // Calculate the total width needed for all pies in a row (8 pies per row)
                int piesPerRow = 8;
                int totalPieSpacing =(int)(0.4f * chartWidth); // Spacing between pies
                diagramWidth = (totalPieSpacing) / piesPerRow; // Width for each pie
            }
            else
            {
                diagramWidth = (int)(chartWidth);
            }
            

            int x1 = (int)(0.1f*chartWidth); // Center horizontally
            int y1 = (int)(0.17f*chartHeight); // Center vertically
            float titleSize = 0.04f * chartHeight;
            // Draw diagram title
            string title = "PIE CHART";
            Font titleFont = new Font("Times New Roman", titleSize, FontStyle.Bold);
            
            int titleX = (int)(frmMainForm.mainChartPlotting.Width*0.4f);
            int titleY = (int)(0.01f*frmMainForm.mainChartPlotting.Height);
            g.DrawString(title, titleFont, Brushes.Black, titleX, titleY);

            // Dummy data extraction
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
                Na[i] = frmImportSamples.WaterData[i].Na;
                K[i] = frmImportSamples.WaterData[i].K;
                Ca[i] = frmImportSamples.WaterData[i].Ca;
                Mg[i] = frmImportSamples.WaterData[i].Mg;
                Cl[i] = frmImportSamples.WaterData[i].Cl;
                HCO3[i] = frmImportSamples.WaterData[i].HCO3;
                CO3[i] = frmImportSamples.WaterData[i].CO3;
                SO4[i] = frmImportSamples.WaterData[i].So4;
            }

            // Pie chart parameters
            int pieDiameter = diagramWidth; // Size of each pie chart
            int pieSpacing = (int)(0.3f * diagramWidth);   // Space between pie charts
            int startAngle = 0;    // Starting angle for pie segments

            // Colors for each component

            
            // Draw pie charts for each sample
            for (int i = 0; i < samples.Count; i++)
            {
                float fontSize = 12; // Make font size relative
                // Calculate total for normalization
                double total = Na[i] + K[i] + Ca[i] + Mg[i] + Cl[i] + SO4[i] + HCO3[i] + CO3[i];
                double[] values = { Na[i] + K[i], Ca[i], Mg[i], Cl[i], SO4[i], HCO3[i] + CO3[i] };

                // Calculate position of the pie chart
                int pieX = x1 + (i % 8) * (pieDiameter + pieSpacing);
                int pieY = y1 + (i / 8) * (pieDiameter + pieSpacing);

                // Draw pie segments
                startAngle = 0;
                for (int j = 0; j < values.Length; j++)
                {
                    float sweepAngle = (float)(values[j] / total * 360.0);
                    if (!frmPieLegend.IsUpdateClicked)
                    {
                        Brush myBrush = new SolidBrush(pieColors[j]);

                        g.FillPie(myBrush, pieX, pieY, pieDiameter, pieDiameter, startAngle, sweepAngle);

                    }
                    else 
                    {
                        Brush myBrush = new SolidBrush(frmPieLegend.PieColor[j]);
                        g.FillPie(myBrush, pieX, pieY, pieDiameter, pieDiameter, startAngle, sweepAngle);
                    }
                    startAngle += (int)sweepAngle;
                }

                // Draw sample label
                g.DrawString(samples[i], new Font("Times New Roman", fontSize), Brushes.Black, pieX + pieDiameter / 2, pieY + pieDiameter + 5);
            }
            
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
                    for (int i = 0; i < frmImportSamples.WaterData.Count;i++ )
                    {
                        string fullText = "W" + (i+1).ToString() + ", " + frmImportSamples.WaterData[i].Well_Name + ", " + frmImportSamples.WaterData[i].ClientID + ", " + frmImportSamples.WaterData[i].Depth;
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
                    g.DrawString("W"+(i+1).ToString()+", "+
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
                int legendX = x1;
                int legendY = (int)(0.1f*chartHeight);
                int s = 0;
                for (int i = 0; i < labels.Length; i++)
                {

                    string fullText = labels[i];
                    SizeF textSize = g.MeasureString(fullText, new Font("Times New Roman", clsConstants.legendTextSize));
                    s = (int)(s + textSize.Width + 40);
                }


                int legendBoxHeight = (int)(0.03f * frmMainForm.mainChartPlotting.Height);
                float fontSize = clsConstants.legendTextSize; // Make font size relative
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
                        if (!frmPieLegend.IsUpdateClicked)
                        {
                            Brush myBrush = new SolidBrush(pieColors[i]);
                            legendGraphics.FillRectangle(myBrush, xsample+5, 2, 18, 18);
                        }
                        else 
                        {
                            Brush myBrush = new SolidBrush(frmPieLegend.PieColor[i]);
                            legendGraphics.FillRectangle(myBrush, xsample+5, 2, 18, 18);
                        }

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
                frmMainForm.legendPictureBox.MouseDoubleClick += frmMainForm.pictureBoxPie_Click;
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
        public static void ExportPieChartToPowerPoint(PowerPoint.Slide slide, PowerPoint.Presentation presentation)
        {
            // Get the chart dimensions from the main form
            int chartWidth = (int)presentation.PageSetup.SlideWidth;
            int chartHeight = (int)presentation.PageSetup.SlideHeight;

            // Calculate the diagram width based on the number of samples
            int diagramWidth;
            if (frmImportSamples.WaterData.Count > 0)
            {
                int piesPerRow = 8;
                int totalPieSpacing = (int)(0.4f * chartWidth);
                diagramWidth = totalPieSpacing / piesPerRow;
            }
            else
            {
                diagramWidth = chartWidth;
            }

            // Calculate the position to center the diagram on the slide

            int x1 = (int)(0.1f * chartWidth); // Center horizontally
            int y1 = 150; // Center vertically

            // Title
            PowerPoint.Shape title = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                50, 20, 600, 50);
            title.TextFrame.TextRange.Text = "PIE CHART";
            title.TextFrame.TextRange.Font.Size = 40;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            title.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            title.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Pie chart parameters - maintain same proportions as in the application
            int pieDiameter = diagramWidth;
            int pieSpacing = (int)(0.3f * diagramWidth);

            // Process each sample
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                // Create a temporary bitmap for individual pie chart
                Bitmap bitmap = new Bitmap(pieDiameter + 10, pieDiameter + 10);
                Graphics graphics = Graphics.FromImage(bitmap);
                graphics.Clear(Color.White);

                // Calculate total for normalization
                double total = frmImportSamples.WaterData[i].Na + frmImportSamples.WaterData[i].K + 
                             frmImportSamples.WaterData[i].Ca + frmImportSamples.WaterData[i].Mg + 
                             frmImportSamples.WaterData[i].Cl + frmImportSamples.WaterData[i].So4 + 
                             frmImportSamples.WaterData[i].HCO3 + frmImportSamples.WaterData[i].CO3;

                double[] values = {
                    frmImportSamples.WaterData[i].Na + frmImportSamples.WaterData[i].K,
                    frmImportSamples.WaterData[i].Ca,
                    frmImportSamples.WaterData[i].Mg,
                    frmImportSamples.WaterData[i].Cl,
                    frmImportSamples.WaterData[i].So4,
                    frmImportSamples.WaterData[i].HCO3 + frmImportSamples.WaterData[i].CO3
                };

                // Draw pie segments
                int startAngle = 0;
                for (int j = 0; j < values.Length; j++)
                {
                    float sweepAngle = (float)(values[j] / total * 360.0);
                    if (!frmPieLegend.IsUpdateClicked)
                    {
                        graphics.FillPie(new SolidBrush(pieColors[j]), 5, 5, pieDiameter, pieDiameter, startAngle, sweepAngle);
                    }
                    else
                    {
                        graphics.FillPie(new SolidBrush(frmPieLegend.PieColor[j]), 5, 5, pieDiameter, pieDiameter, startAngle, sweepAngle);
                    }
                    startAngle += (int)sweepAngle;
                }

                // Save the bitmap to a temporary file
                string tempImagePath = Path.GetTempFileName() + ".png";
                bitmap.Save(tempImagePath, System.Drawing.Imaging.ImageFormat.Png);

                // Calculate position on PowerPoint slide - maintain same grid layout as in the application
                int pieX = x1 + (i % 8) * (pieDiameter + pieSpacing);
                int pieY = y1 + (i / 8) * (pieDiameter + pieSpacing);

                // Insert image into PowerPoint
                slide.Shapes.AddPicture(tempImagePath,
                    Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, pieX, pieY, pieDiameter, pieDiameter);

                // Sample Label
                PowerPoint.Shape sampleLabel = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    pieX + pieDiameter / 2 - 20, pieY + pieDiameter, 50, 15);
                sampleLabel.TextFrame.TextRange.Text = "W" + (i + 1).ToString();
                sampleLabel.TextFrame.TextRange.Font.Size = 8;
                sampleLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                sampleLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                sampleLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

                // Cleanup
                graphics.Dispose();
                bitmap.Dispose();
                File.Delete(tempImagePath);
            }

            #region Draw Legend
            int legendX = x1;
            int legendY = 100;
            int xSample = legendX+5;
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
                legendBoxWidth += (int)temp.Width+10;
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
                        xSample, legendY+5, 10, 10);
                    legendBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(pieColors[i]);
                }
                else
                {
                    PowerPoint.Shape legendBox = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        xSample, legendY+5, 10, 10);
                    legendBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(frmPieLegend.PieColor[i]);
                }

                PowerPoint.Shape legendText = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    xSample+10, legendY, 100, 20);
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

                xSample += (int)legendText.Width+10;
            }



            // Add metadata
            float metadataX = 500;
            float metadataY = legendY;
            int metaWidth = 0;
            int metaHeight = 0;
            
            
            float ySample = metadataY;
            //List<PowerPoint.Shape> addedTexts = new List<PowerPoint.Shape>();

            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                string fullText = "W" + (i + 1).ToString() + "," +
                    frmImportSamples.WaterData[i].Well_Name + "," +
                    frmImportSamples.WaterData[i].ClientID + "," +
                    frmImportSamples.WaterData[i].Depth;

                PowerPoint.Shape metadataText = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    metadataX+2, ySample, 500, 20);

                metadataText.TextFrame.TextRange.Text = fullText;
                metadataText.TextFrame.TextRange.Font.Size = fontSize;
                metadataText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                metadataText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                metadataText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                metadataText.TextFrame.MarginLeft = 0;
                metadataText.TextFrame.MarginRight = 0;
                metadataText.TextFrame.MarginTop = 0;
                metadataText.TextFrame.MarginBottom = 0;
                metadataText.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;

                

                metaWidth = Math.Max(metaWidth, (int)metadataText.Width);
                ySample += metadataText.Height;
                metaHeight += (int)metadataText.Height+1;
                //addedTexts.Add(metadataText);
            }

            // Now draw the border box *after*
            PowerPoint.Shape metaBorder = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                metadataX, metadataY, metaWidth, metaHeight);
            metaBorder.Fill.Transparency = 1.0f;
            metaBorder.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue);
            metaBorder.Line.Weight = 2;
            #endregion
        }
    }
}
