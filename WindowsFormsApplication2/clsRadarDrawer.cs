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
    public class clsRadarDrawer
    {
        
        public static double maxCl = 0, maxNa = 0, maxK = 0, maxCa = 0, maxMg = 0, maxBa = 0, maxSr = 0;
        public static double maxNaCl = 0, maxClCa = 0, maxHCO3Cl = 0, maxClSr = 0, maxNaCa = 0, maxKNa = 0, maxSrMg = 0, maxMgCl = 0, maxSrCl = 0, maxSrK = 0, maxMgK = 0, maxCaK = 0, maxtK = 0, maxBCl = 0, maxBNa = 0, maxBMg = 0;
        public static double maxAl = 0, maxCo = 0, maxCu = 0, maxMn = 0, maxNi = 0, maxZn = 0, maxPb = 0, maxFe = 0, maxCd = 0, maxCr = 0, maxTl = 0, maxBe = 0, maxSe = 0, maxLi=0,maxB=0;
        public static double Bm = 35453, Bn = 22989.7, Bo = 39098.3, Bp = 40078, Bq = 24305, Br = 137327, Bs = 87620;
        public static TextBox txt;
        public static string[] Radar1Scales=new string[7];
        public static string[] Radar2Scales = new string[16];
        public static string[] Radar3Scales = new string[21];
        /// <summary>
        /// Draw Radar legend
        /// </summary>
        /// <param name="g">graphics</param>
        /// <param name="bounds">bounds of the radar</param>
        public static void Radar_legend(Graphics g, Rectangle bounds)
        {

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

                using (Font font = new Font("Times New Roman", legendtextSize, FontStyle.Bold))
                {
                    foreach (var data in frmImportSamples.WaterData)
                    {
                        string fullText = data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
                        SizeF textSize = g.MeasureString(fullText, font);
                        if (textSize.Width + 30 > legendBoxWidth)
                        {
                            legendBoxWidth = (int)Math.Round(textSize.Width, 0) + 30;
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
                frmMainForm.legendPictureBox.MouseDoubleClick += frmMainForm.legendPictureBoxRadar;
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
        /// 
        /// </summary>
        /// <param name="g"></param>
        /// <param name="bounds"></param>
        /// <param name="flag"></param>
        public static void DrawRadarChart1(Graphics g, Rectangle bounds,bool flag)
        {
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.legendPictureBoxRadar;
            frmMainForm.mainChartPlotting.Invalidate();
            // Data labels and values
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];



            float fontSize = 12; // Make font size relative
            PrecomputeMaxValues(flag);
            Radar1Scales = new string[] { maxCl.ToString("F5"), maxNa.ToString("F5"), maxK.ToString("F5"), maxCa.ToString("F5"), maxMg.ToString("F5"), maxBa.ToString("F5"), maxSr.ToString("F5") };
            Font AxisFont = new Font("Times New Roman", fontSize, FontStyle.Bold);
            List<string> scales = new List<string>();
            for (int i = 0; i < Radar1Scales.Count(); i++)
            {
                string s = Radar1Scales[i];
                string temp = "";
                bool checking = false;
                bool found = false;
                for (int j = 0; j < s.Length; j++)
                {
                    if (s[j] == '.')
                    {
                        checking = true;
                    }
                    else if (s[j] != '0' && checking)
                    {
                        found = true;
                        temp += s[j];
                        break;
                    }
                    temp += s[j];
                }
                if (!found)
                {
                    int dotIndex = temp.IndexOf('.');
                    if (dotIndex != -1)
                    {
                        temp = temp.Substring(0, dotIndex);
                    }
                }
                scales.Add(temp);
                Radar1Scales[i] =temp;
            }
            string[] labels =
            {
                "K (mol/L)\n"+ scales[2],
                "Ca (mol/L)\n"+ scales[3],
                "Mg (mol/L)\n"+ scales[4],
                "Ba (mol/L)\n"+ scales[5],
                "Sr (mol/L)\n"+ scales[6],
                "Cl (mol/L)\n"+ scales[0],
                "Na (mol/L)\n"+ scales[1]
            };
            
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                sampleData[i] = new clsRadarScale[]
                {
                new clsRadarScale { Item = frmImportSamples.WaterData[i].K / Bo, Scale = maxK},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Ca / Bp, Scale = maxCa},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Mg / Bq, Scale = maxMg},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Ba / Br, Scale = maxBa},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / Bs, Scale = maxSr},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Cl / Bm, Scale = maxCl},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Na / Bn, Scale = maxNa}
                };
            }
            // Colors for the samples
            Color[] colors = new Color[frmImportSamples.WaterData.Count];
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                colors[i] = frmImportSamples.WaterData[i].color;
            }

            // Center of the diagram
            float centerX = 0.3f * frmMainForm.mainChartPlotting.Width;
            float centerY = 0.4f * frmMainForm.mainChartPlotting.Height;

            // Radius of the radar diagram
            float radius = Math.Min(bounds.Width, bounds.Height) / 3;

            // Number of axes
            int numAxes = labels.Length;

            // Angle between each axis (in radians)
            double angleIncrement = 2 * Math.PI / numAxes;

            // Draw the axes

            Pen axisPen = new Pen(Color.LightGray, 1);
            axisPen.DashStyle = DashStyle.Dot;
            PointF[] quarterList = new PointF[numAxes];
            PointF[] halfList = new PointF[numAxes];
            PointF[] thirdQuarterList = new PointF[numAxes];
            PointF[] allList = new PointF[numAxes];
            for (int i = 0; i < numAxes; i++)
            {
                double angle = i * angleIncrement;
                float allX = centerX + (float)(radius * Math.Cos(angle));
                float allY = centerY + (float)(radius * Math.Sin(angle));
                g.DrawLine(axisPen, centerX, centerY, allX, allY);
                float quarterX = centerX + (float)((0.25*radius) * Math.Cos(angle));
                float quarterY = centerY + (float)((0.25*radius) * Math.Sin(angle));
                float halfX = centerX + (float)((0.5*radius) * Math.Cos(angle));
                float halfY = centerY + (float)((0.5*radius) * Math.Sin(angle));
                float thirdQuarterX = centerX + (float)((0.75*radius) * Math.Cos(angle));
                float thirdQuarterY = centerY + (float)((0.75*radius) * Math.Sin(angle));
                allList[i]=new PointF(allX,allY);
                quarterList[i]=new PointF(quarterX, quarterY);
                halfList[i]=new PointF(halfX, halfY);
                thirdQuarterList[i]=new PointF(thirdQuarterX, thirdQuarterY);
                // Draw axis labels
                string label = labels[i];
                SizeF labelSize = g.MeasureString(label, AxisFont);
                float labelX = (float)(allX + (0.3 * radius) * Math.Cos(angle));
                float labelY = (float)(allY + (0.3 * radius) * Math.Sin(angle));
                labelX -= 0.02f * frmMainForm.mainChartPlotting.Width;
                g.DrawString(label, AxisFont, Brushes.Black, labelX, labelY);

                
            }
            g.DrawPolygon(axisPen, quarterList);
            g.DrawPolygon(axisPen, allList);
            g.DrawPolygon(axisPen, halfList);
            g.DrawPolygon(axisPen, thirdQuarterList);

            g.DrawString("Radar diagram showing the molar concentrations for major ions", new Font("Times New Roman", fontSize, FontStyle.Bold), Brushes.Black, 0.2f*frmMainForm.mainChartPlotting.Width, 0.9f*frmMainForm.mainChartPlotting.Height);

            #region Draw Radar legend
            for (int s = 0; s < sampleData.Length; s++)
            {
                PointF[] points = new PointF[numAxes];

                for (int i = 0; i < numAxes; i++)
                {
                    clsRadarScale value = sampleData[s][i];

                    // Normalize value to be within 0 and its max scale
                    double normalizedValue = Math.Min(value.Item / value.Scale, 1.0);

                    // Scale the normalized value according to its axis' maximum
                    float scaledRadius = (float)(normalizedValue * radius);

                    double angle = i * angleIncrement;
                    float x = centerX + (float)(scaledRadius * Math.Cos(angle));
                    float y = centerY + (float)(scaledRadius * Math.Sin(angle));
                    if (double.IsNaN(x))
                    {
                        x = centerX;
                    }
                    if (double.IsNaN(y))
                    {
                        y = centerY;
                    }
                    points[i] = new PointF(x, y);
                }

                // Draw the polygon by connecting points
                using (Pen linePen = new Pen(colors[s], 2))
                {
                    linePen.Color = frmImportSamples.WaterData[s].color;
                    linePen.DashStyle = frmImportSamples.WaterData[s].selectedStyle;
                    linePen.Width = frmImportSamples.WaterData[s].lineWidth;
                    g.DrawPolygon(linePen, points);
                }
            }
            Radar_legend(g, bounds);
            #endregion

            flag = false;
        }
        public static void ExportRadar1ToPowerpoint(Rectangle bounds, PowerPoint.Slide slide, PowerPoint.Presentation presentation,bool flag)
        {

            // Data labels and values
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];
            double Bm = 35453, Bn = 22989.7, Bo = 39098.3, Bp = 40078, Bq = 24305, Br = 137327, Bs = 87620;
            PrecomputeMaxValues(flag);

            string[] labels =
            {
                "K (mol/L)"+ (maxK).ToString("F5"),
                "Ca (mol/L)"+ (maxCa).ToString("F5"),
                "Mg (mol/L)"+ (maxMg).ToString("F5"),
                "Ba (mol/L)"+ (maxBa).ToString("F5"),
                "Sr (mol/L)"+ (maxSr).ToString("F5"),
                "Cl (mol/L)"+ (maxCl).ToString("F5"),
                "Na (mol/L)"+ (maxNa).ToString("F5")
            };
            // Initialize jagged array
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                sampleData[i] = new clsRadarScale[]
                {
                new clsRadarScale { Item = frmImportSamples.WaterData[i].K / Bo, Scale = maxK},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Ca / Bp, Scale = maxCa},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Mg / Bq, Scale = maxMg},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Ba / Br, Scale = maxBa},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / Bs, Scale = maxSr},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Cl / Bm, Scale = maxCl},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Na / Bn, Scale = maxNa}
                };
            }

            // Colors for the samples
            Color[] colors = new Color[frmImportSamples.WaterData.Count];
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                colors[i] = frmImportSamples.WaterData[i].color;
            }

            // Center of the diagram
            float centerX = (bounds.Width / 2);
            float centerY = bounds.Y + bounds.Height / 2 - 50;


            // Radius of the radar diagram
            float radius = Math.Min(bounds.Width, bounds.Height) / 3;
            // Number of axes
            int numAxes = labels.Length;

            // Angle between each axis (in radians)
            PointF[] quarterList = new PointF[numAxes];
            PointF[] halfList = new PointF[numAxes];
            PointF[] thirdQuarterList = new PointF[numAxes];
            PointF[] allList = new PointF[numAxes];
            double angleIncrement = 2 * Math.PI / numAxes;
            for (int i = 0; i < numAxes; i++)
            {
                double angle = i * angleIncrement;
                float allX = centerX + (float)(radius * Math.Cos(angle));
                float allY = centerY + (float)(radius * Math.Sin(angle));
                float quarterX = centerX + (float)((0.25 * radius) * Math.Cos(angle));
                float quarterY = centerY + (float)((0.25 * radius) * Math.Sin(angle));
                float halfX = centerX + (float)((0.5 * radius) * Math.Cos(angle));
                float halfY = centerY + (float)((0.5 * radius) * Math.Sin(angle));
                float thirdQuarterX = centerX + (float)((0.75 * radius) * Math.Cos(angle));
                float thirdQuarterY = centerY + (float)((0.75 * radius) * Math.Sin(angle));
                allList[i] = new PointF(allX, allY);
                quarterList[i] = new PointF(quarterX, quarterY);
                halfList[i] = new PointF(halfX, halfY);
                thirdQuarterList[i] = new PointF(thirdQuarterX, thirdQuarterY);
                var radiusLine=slide.Shapes.AddLine(centerX, centerY, allX, allY);
                radiusLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
                radiusLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                string label = labels[i];
                SizeF labelSize = TextRenderer.MeasureText(label, SystemFonts.DefaultFont);
                float labelX = allX + (float)((radius*0.3) * Math.Cos(angle)) - labelSize.Width / 2;
                float labelY = allY + (float)((radius*0.3) * Math.Sin(angle)) - labelSize.Height / 2;
                
                PowerPoint.Shape itemText = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, labelX, labelY, 200, 30);
                itemText.TextFrame.TextRange.Text = label;
                itemText.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                itemText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                itemText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                itemText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                itemText.TextFrame.MarginLeft = 0;
                itemText.TextFrame.MarginRight = 0;
                itemText.TextFrame.MarginTop = 0;
                itemText.TextFrame.MarginBottom = 0;
                itemText.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            }

            for (int i = 0; i < allList.Length - 1; i++)
            {
                float[,] points1 = new float[,]
                {
                    { allList[i].X, allList[i].Y },
                    { allList[i + 1].X, allList[i + 1].Y}
                };

                PowerPoint.Shape allPolygon= slide.Shapes.AddPolyline(points1);
                allPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                allPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                float[,] points2 = new float[,]
                {
                    { thirdQuarterList[i].X, thirdQuarterList[i].Y },
                    { thirdQuarterList[i + 1].X, thirdQuarterList[i + 1].Y}
                };

                PowerPoint.Shape thirdQuarterPolygon = slide.Shapes.AddPolyline(points2);
                thirdQuarterPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                thirdQuarterPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                float[,] points3 = new float[,]
                {
                    { halfList[i].X, halfList[i].Y },
                    { halfList[i + 1].X, halfList[i + 1].Y}
                };

                PowerPoint.Shape halfPolygon = slide.Shapes.AddPolyline(points3);
                halfPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                halfPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                float[,] points4 = new float[,]
                {
                    { quarterList[i].X, quarterList[i].Y },
                    { quarterList[i + 1].X, quarterList[i + 1].Y}
                };

                PowerPoint.Shape quarterPolygon = slide.Shapes.AddPolyline(points4);
                quarterPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                quarterPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            }
            float[,] allPoints = new float[,]
                {
                    { allList[0].X, allList[0].Y },
                    { allList[6].X, allList[6].Y}
                };

            PowerPoint.Shape allPolygon2 = slide.Shapes.AddPolyline(allPoints);
            allPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            allPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            float[,] thirdQuarterPoints = new float[,]
                {
                    { thirdQuarterList[0].X, thirdQuarterList[0].Y },
                    { thirdQuarterList[6].X, thirdQuarterList[6].Y}
                };

            PowerPoint.Shape thirdQuarterPolygon2 = slide.Shapes.AddPolyline(thirdQuarterPoints);
            thirdQuarterPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            thirdQuarterPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            float[,] halfPoints = new float[,]
                {
                    { halfList[0].X, halfList[0].Y },
                    { halfList[6].X, halfList[6].Y}
                };

            PowerPoint.Shape halfPolygon2 = slide.Shapes.AddPolyline(halfPoints);
            halfPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            halfPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            float[,] quarterPoints = new float[,]
                {
                    { quarterList[0].X, quarterList[0].Y },
                    { quarterList[6].X, quarterList[6].Y}
                };

            PowerPoint.Shape quarterPolygon2 = slide.Shapes.AddPolyline(quarterPoints);
            quarterPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            quarterPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 5, centerY + radius + 70, 1000, 30)
                .TextFrame.TextRange.Text = "Radar diagram showing the molar concentrations for major ions";
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.Shapes[slide.Shapes.Count].TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
            slide.Shapes[slide.Shapes.Count].TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginLeft = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginRight = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginTop = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginBottom = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            for (int i = 0; i < sampleData.Length; i++)
            {
                PointF[] points = new PointF[numAxes];

                for (int j = 0; j < numAxes; j++)
                {
                    clsRadarScale value = sampleData[i][j];

                    // Normalize value to be within 0 and its max scale
                    double normalizedValue = Math.Min(value.Item / value.Scale, 1.0); // Ensures it doesn't exceed 1

                    // Scale the normalized value according to its axis' maximum
                    float scaledRadius = (float)(normalizedValue * radius);

                    double angle = j * angleIncrement;
                    float x = centerX + (float)(scaledRadius * Math.Cos(angle));
                    float y = centerY + (float)(scaledRadius * Math.Sin(angle));
                    if (double.IsNaN(x))
                    {
                        x = centerX;
                    }
                    if (double.IsNaN(y))
                    {
                        y = centerY;
                    }
                    points[j] = new PointF(x, y);
                }
                // Flatten the points array into a float array for AddPolyline
                float[] polylinePoints = new float[points.Length * 2];
                for (int j = 0; j < points.Length; j++)
                {
                    polylinePoints[j * 2] = points[j].X;
                    polylinePoints[j * 2 + 1] = points[j].Y;
                }

                // Add the polyline to the slide
                PowerPoint.Shape samplePolygon = slide.Shapes.AddPolyline(new float[,]
                {
                    { polylinePoints[0], polylinePoints[1] },
                    { polylinePoints[2], polylinePoints[3] },
                    { polylinePoints[4], polylinePoints[5] },
                    { polylinePoints[6], polylinePoints[7] },
                    { polylinePoints[8], polylinePoints[9] },
                    { polylinePoints[10],polylinePoints[11]},
                    { polylinePoints[12],polylinePoints[13]},
                });
                samplePolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color); // Set line color
                samplePolygon.Line.Weight = frmImportSamples.WaterData[i].lineWidth; // Set line width
                samplePolygon.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
                PowerPoint.Shape sampleLastLine = slide.Shapes.AddPolyline(new float[,]
                    {
                      { polylinePoints[0], polylinePoints[1] },
                      { polylinePoints[12],polylinePoints[13]},
                    }
                );
                sampleLastLine.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color); // Set line color
                sampleLastLine.Line.Weight = frmImportSamples.WaterData[i].lineWidth; // Set line width
                sampleLastLine.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
            }
            #region Draw Legend
            if (frmImportSamples.WaterData.Count > 0)
            {
                int legendY = 50;
                
                // Add metadata
                float metadataX = 500;
                float metadataY = legendY;
                int metaWidth = 0;
                int metaHeight = 0;


                float ysample = metadataY;
                //List<PowerPoint.Shape> addedTexts = new List<PowerPoint.Shape>();

                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    var line = slide.Shapes.AddLine(metadataX, ysample + 10, metadataX + 20, ysample + 10);
                    line.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color);
                    line.Line.Weight = frmImportSamples.WaterData[i].lineWidth;
                    line.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
                    string fullText = "W" + (i + 1).ToString() + "," +
                        frmImportSamples.WaterData[i].Well_Name + "," +
                        frmImportSamples.WaterData[i].ClientID + "," +
                        frmImportSamples.WaterData[i].Depth;

                    PowerPoint.Shape metadataText = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        metadataX + 22, ysample, 500, 20);

                    metadataText.TextFrame.TextRange.Text = fullText;
                    metadataText.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                    metadataText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    metadataText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    metadataText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                    metadataText.TextFrame.MarginLeft = 0;
                    metadataText.TextFrame.MarginRight = 0;
                    metadataText.TextFrame.MarginTop = 0;
                    metadataText.TextFrame.MarginBottom = 0;
                    metadataText.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;



                    metaWidth = Math.Max(metaWidth, (int)metadataText.Width+30);
                    ysample += metadataText.Height;
                    metaHeight += (int)metadataText.Height + 1;
                    //addedTexts.Add(metadataText);
                }

                // Now draw the border box *after*
                PowerPoint.Shape metaBorder = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    metadataX, metadataY, metaWidth, metaHeight);
                metaBorder.Fill.Transparency = 1.0f;
                metaBorder.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                metaBorder.Line.Weight = 2;
                // Refresh PowerPoint slide
                //pptApplication.ActiveWindow.View.GotoSlide(presentation.Slides.Count);
            }
            #endregion
            
        }
        public static void DrawRadarChart2(Graphics g, Rectangle bounds, bool flag)
        {
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.legendPictureBoxRadar;
            frmMainForm.mainChartPlotting.Invalidate();
            float fontSize = 12; // Make font size relative
            // Data labels and values
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];
            PrecomputeMaxValues(flag);

            Radar2Scales = new string[] { maxNaCl.ToString("F5"), maxClCa.ToString("F5"), maxHCO3Cl.ToString("F5"), maxClSr.ToString("F5"), maxNaCa.ToString("F5"), maxKNa.ToString("F5"), maxSrMg.ToString("F5"), maxMgCl.ToString("F5"), maxSrCl.ToString("F5"), maxSrK.ToString("F5"), maxMgK.ToString("F5"), maxCaK.ToString("F5"), maxtK.ToString("F5"), maxBCl.ToString("F5"), maxBNa.ToString("F5"), maxBMg.ToString("F5") };
            List<string> scales = new List<string>();
            for (int i = 0; i < Radar2Scales.Count(); i++)
            {
                string s = Radar2Scales[i];
                string temp = "";
                bool checking = false;
                bool found = false;
                for (int j = 0; j < s.Length; j++)
                {
                    if (s[j] == '.')
                    {
                        checking = true;
                    }
                    else if (s[j] != '0' && checking)
                    {
                        found = true;
                        temp += s[j];
                        break;
                    }
                    temp += s[j];
                }
                if (!found)
                {
                    int dotIndex = temp.IndexOf('.');
                    if (dotIndex != -1)
                    {
                        temp = temp.Substring(0, dotIndex);
                    }
                }
                scales.Add(temp);
                Radar2Scales[i] = temp;
            }
            string[] labels =
            {
            "EV_Na-Ca \n"+scales[4],
            "GT_K-Na \n"+ scales[5],
            "SS_Sr-Mg \n"+scales[6],
            "SS_Mg-Cl \n"+scales[7],
            "SS_Sr-Cl \n"+ scales[8],
            "Lith_Sr-K \n"+ scales[9],
            "Lith_Mg-K \n"+ scales[10],
            "Lith_Ca-K \n"+ scales[11],
            "Wt%K \n"+ scales[12],
            "OM_B-Cl \n"+ scales[13],
            "OM_B-Na \n"+ scales[14],
            "OM_B-Mg \n"+ scales[15],
            "EV_Na-Cl \n"+ scales[0],
            "EV_Cl-Ca \n"+ scales[1],
            "EV_HCO3-Cl \n"+scales[2],
            "EV_Cl-Sr \n"+ scales[3]

        };
            // Initialize jagged array
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                sampleData[i] = new clsRadarScale[]
                {
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Na / frmImportSamples.WaterData[i].Ca, Scale = maxNaCa},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].K / frmImportSamples.WaterData[i].Na, Scale = maxKNa},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / frmImportSamples.WaterData[i].Mg, Scale = maxSrMg},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Mg / frmImportSamples.WaterData[i].Cl, Scale = maxMgCl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / frmImportSamples.WaterData[i].Cl, Scale = maxSrCl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / frmImportSamples.WaterData[i].K, Scale = maxSrK},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Mg / frmImportSamples.WaterData[i].K, Scale = maxMgK},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Ca / frmImportSamples.WaterData[i].K, Scale = maxCaK},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].K / 10000, Scale = maxtK},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].B / frmImportSamples.WaterData[i].Cl, Scale = maxBCl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].B / frmImportSamples.WaterData[i].Na, Scale = maxBNa},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].B / frmImportSamples.WaterData[i].Mg, Scale = maxBMg},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Na / frmImportSamples.WaterData[i].Cl, Scale = maxNaCl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Cl / frmImportSamples.WaterData[i].Ca, Scale = maxClCa},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].HCO3 / frmImportSamples.WaterData[i].Cl, Scale = maxHCO3Cl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Cl / frmImportSamples.WaterData[i].Sr, Scale = maxClSr},

                };
            }

            // Colors for the samples
            Color[] colors = new Color[frmImportSamples.WaterData.Count];
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                colors[i] = frmImportSamples.WaterData[i].color;
            }

            // Center of the diagram
            float centerX = 0.3f*frmMainForm.mainChartPlotting.Width;
            float centerY = 0.4f*frmMainForm.mainChartPlotting.Height;

            // Radius of the radar diagram
            float radius = Math.Min(bounds.Width, bounds.Height) *0.32f;

            // Number of axes
            int numAxes = labels.Length;

            // Angle between each axis (in radians)
            double angleIncrement = 2 * Math.PI / numAxes;

            #region Draw Axis
            Font AxisFont = new Font("Times New Roman", fontSize, FontStyle.Bold);
            Pen axisPen = new Pen(Color.LightGray, 1);
            axisPen.DashStyle = DashStyle.Dot;
            PointF[] quarterList = new PointF[numAxes];
            PointF[] halfList = new PointF[numAxes];
            PointF[] thirdQuarterList = new PointF[numAxes];
            PointF[] allList = new PointF[numAxes];
            for (int i = 0; i < numAxes; i++)
            {
                double angle = i * angleIncrement;
                float allX = centerX + (float)(radius * Math.Cos(angle));
                float allY = centerY + (float)(radius * Math.Sin(angle));
                g.DrawLine(axisPen, centerX, centerY, allX, allY);
                float quarterX = centerX + (float)((0.25 * radius) * Math.Cos(angle));
                float quarterY = centerY + (float)((0.25 * radius) * Math.Sin(angle));
                float halfX = centerX + (float)((0.5 * radius) * Math.Cos(angle));
                float halfY = centerY + (float)((0.5 * radius) * Math.Sin(angle));
                float thirdQuarterX = centerX + (float)((0.75 * radius) * Math.Cos(angle));
                float thirdQuarterY = centerY + (float)((0.75 * radius) * Math.Sin(angle));
                allList[i] = new PointF(allX, allY);
                quarterList[i] = new PointF(quarterX, quarterY);
                halfList[i] = new PointF(halfX, halfY);
                thirdQuarterList[i] = new PointF(thirdQuarterX, thirdQuarterY);
                // Draw axis labels
                string label = labels[i];
                SizeF labelSize = g.MeasureString(label, AxisFont);
                float labelX = (float)(allX + (0.3 * radius) * Math.Cos(angle));
                float labelY = (float)(allY + (0.3 * radius) * Math.Sin(angle));
                
                labelX -= 0.02f * frmMainForm.mainChartPlotting.Width;
                g.DrawString(label, AxisFont, Brushes.Black, labelX, labelY);

            }
            g.DrawPolygon(axisPen, quarterList);
            g.DrawPolygon(axisPen, allList);
            g.DrawPolygon(axisPen, halfList);
            g.DrawPolygon(axisPen, thirdQuarterList);
            #endregion
            g.DrawString("Genetic Origin and Alteration Tool Radar Plot for study waters. Ratio categories: \nwater evolution(EV), geothermometers(GT), lithology(Lith), salinity source(SS), and organic matter related(OM) \nare listed in front of each axis label.", new Font("Times New Roman", fontSize, FontStyle.Bold), Brushes.Black, centerX - 350, centerY + radius + 120);


            #region Draw Radar legend

            for (int s = 0; s < sampleData.Length; s++)
            {
                PointF[] points = new PointF[numAxes];

                for (int i = 0; i < numAxes; i++)
                {
                    clsRadarScale value = sampleData[s][i];

                    // Normalize value to be within 0 and its max scale
                    double normalizedValue = Math.Min(value.Item / value.Scale, 1.0); // Ensures it doesn't exceed 1

                    // Scale the normalized value according to its axis' maximum
                    float scaledRadius = (float)(normalizedValue * radius);

                    double angle = i * angleIncrement;
                    float x = centerX + (float)(scaledRadius * Math.Cos(angle));
                    float y = centerY + (float)(scaledRadius * Math.Sin(angle));
                    if (double.IsNaN(x))
                    {
                        x = centerX;
                    }
                    if (double.IsNaN(y))
                    {
                        y = centerY;
                    }
                    points[i] = new PointF(x, y);
                }

                // Draw the polygon by connecting points
                using (Pen linePen = new Pen(colors[s], 2))
                {
                    linePen.Color = frmImportSamples.WaterData[s].color;
                    linePen.DashStyle = frmImportSamples.WaterData[s].selectedStyle;

                    linePen.Width = frmImportSamples.WaterData[s].lineWidth;
                    g.DrawPolygon(linePen, points);
                }


            }
            Radar_legend(g, bounds);
            #endregion
            flag = false;
        }

        public static void ExportRadar2ToPowerpoint(Rectangle bounds, PowerPoint.Slide slide, PowerPoint.Presentation presentation,bool flag)
        {
            



            // Data labels and values
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];
            PrecomputeMaxValues(flag);
            string[] labels =
            {
            "EV_Na-Ca \n"+(maxNaCa).ToString("F5"),
            "GT_K-Na \n"+ (maxKNa).ToString("F5"),
            "SS_Sr-Mg \n"+(maxSrMg).ToString("F5"),
            "SS_Mg-Cl \n"+(maxMgCl).ToString("F5"),
            "SS_Sr-Cl \n"+ (maxSrCl).ToString("F5"),
            "Lith_Sr-K \n"+ (maxSrK).ToString("F5"),
            "Lith_Mg-K \n"+ (maxMgK).ToString("F5"),
            "Lith_Ca-K \n"+ (maxCaK).ToString("F5"),
            "Wt%K \n"+ (maxtK).ToString("F5"),
            "OM_B-Cl \n"+ (maxBCl).ToString("F5"),
            "OM_B-Na \n"+ (maxBNa).ToString("F5"),
            "OM_B-Mg \n"+ (maxBMg).ToString("F5"),
            "EV_Na-Cl \n"+ (maxNaCl).ToString("F5"),
            "EV_Cl-Ca \n"+ (maxClCa).ToString("F5"),
            "EV_HCO3-Cl \n"+(maxHCO3Cl).ToString("F5"),
            "EV_Cl-Sr \n"+ (maxClSr).ToString("F5")

        };

            // Initialize jagged array
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                sampleData[i] = new clsRadarScale[]
                {
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Na / frmImportSamples.WaterData[i].Ca, Scale = maxNaCa},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].K / frmImportSamples.WaterData[i].Na, Scale = maxKNa},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / frmImportSamples.WaterData[i].Mg, Scale = maxSrMg},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Mg / frmImportSamples.WaterData[i].Cl, Scale = maxMgCl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / frmImportSamples.WaterData[i].Cl, Scale = maxSrCl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / frmImportSamples.WaterData[i].K, Scale = maxSrK},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Mg / frmImportSamples.WaterData[i].K, Scale = maxMgK},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Ca / frmImportSamples.WaterData[i].K, Scale = maxCaK},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].K / 10000, Scale = maxtK},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].B / frmImportSamples.WaterData[i].Cl, Scale = maxBCl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].B / frmImportSamples.WaterData[i].Na, Scale = maxBNa},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].B / frmImportSamples.WaterData[i].Mg, Scale = maxBMg},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Na / frmImportSamples.WaterData[i].Cl, Scale = maxNaCl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Cl / frmImportSamples.WaterData[i].Ca, Scale = maxClCa},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].HCO3 / frmImportSamples.WaterData[i].Cl, Scale = maxHCO3Cl},
                    new clsRadarScale { Item = frmImportSamples.WaterData[i].Cl / frmImportSamples.WaterData[i].Sr, Scale = maxClSr}
                };
            }

            // Colors for the samples
            Color[] colors = new Color[frmImportSamples.WaterData.Count];
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                colors[i] = frmImportSamples.WaterData[i].color;
            }

            // Center of the diagram
            float centerX = (bounds.Width / 2);
            float centerY = bounds.Y + bounds.Height / 2 - 50;


            // Radius of the radar diagram
            float radius = Math.Min(bounds.Width, bounds.Height) / 3;
            // Number of axes
            int numAxes = labels.Length;

            // Angle between each axis (in radians)
            double angleIncrement = 2 * Math.PI / numAxes;
            // Angle between each axis (in radians)
            PointF[] quarterList = new PointF[numAxes];
            PointF[] halfList = new PointF[numAxes];
            PointF[] thirdQuarterList = new PointF[numAxes];
            PointF[] allList = new PointF[numAxes];
            for (int i = 0; i < numAxes; i++)
            {
                double angle = i * angleIncrement;
                float allX = centerX + (float)(radius * Math.Cos(angle));
                float allY = centerY + (float)(radius * Math.Sin(angle));
                float quarterX = centerX + (float)((0.25 * radius) * Math.Cos(angle));
                float quarterY = centerY + (float)((0.25 * radius) * Math.Sin(angle));
                float halfX = centerX + (float)((0.5 * radius) * Math.Cos(angle));
                float halfY = centerY + (float)((0.5 * radius) * Math.Sin(angle));
                float thirdQuarterX = centerX + (float)((0.75 * radius) * Math.Cos(angle));
                float thirdQuarterY = centerY + (float)((0.75 * radius) * Math.Sin(angle));
                allList[i] = new PointF(allX, allY);
                quarterList[i] = new PointF(quarterX, quarterY);
                halfList[i] = new PointF(halfX, halfY);
                thirdQuarterList[i] = new PointF(thirdQuarterX, thirdQuarterY);
                var radiusLine = slide.Shapes.AddLine(centerX, centerY, allX, allY);
                radiusLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
                radiusLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                string label = labels[i];
                SizeF labelSize = TextRenderer.MeasureText(label, SystemFonts.DefaultFont);
                float labelX = allX + (float)((radius * 0.3) * Math.Cos(angle)) - labelSize.Width / 2;
                float labelY = allY + (float)((radius * 0.3) * Math.Sin(angle)) - labelSize.Height / 2;

                PowerPoint.Shape itemText = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, labelX, labelY, 200, 30);
                itemText.TextFrame.TextRange.Text = label;
                itemText.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                itemText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                itemText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                itemText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                itemText.TextFrame.MarginLeft = 0;
                itemText.TextFrame.MarginRight = 0;
                itemText.TextFrame.MarginTop = 0;
                itemText.TextFrame.MarginBottom = 0;
                itemText.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            }

            for (int i = 0; i < allList.Length - 1; i++)
            {
                float[,] points1 = new float[,]
                {
                    { allList[i].X, allList[i].Y },
                    { allList[i + 1].X, allList[i + 1].Y}
                };

                PowerPoint.Shape allPolygon = slide.Shapes.AddPolyline(points1);
                allPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                allPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                float[,] points2 = new float[,]
                {
                    { thirdQuarterList[i].X, thirdQuarterList[i].Y },
                    { thirdQuarterList[i + 1].X, thirdQuarterList[i + 1].Y}
                };

                PowerPoint.Shape thirdQuarterPolygon = slide.Shapes.AddPolyline(points2);
                thirdQuarterPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                thirdQuarterPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                float[,] points3 = new float[,]
                {
                    { halfList[i].X, halfList[i].Y },
                    { halfList[i + 1].X, halfList[i + 1].Y}
                };

                PowerPoint.Shape halfPolygon = slide.Shapes.AddPolyline(points3);
                halfPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                halfPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                float[,] points4 = new float[,]
                {
                    { quarterList[i].X, quarterList[i].Y },
                    { quarterList[i + 1].X, quarterList[i + 1].Y}
                };

                PowerPoint.Shape quarterPolygon = slide.Shapes.AddPolyline(points4);
                quarterPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                quarterPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            }
            float[,] allPoints = new float[,]
                {
                    { allList[0].X, allList[0].Y },
                    { allList[15].X, allList[15].Y}
                };

            PowerPoint.Shape allPolygon2 = slide.Shapes.AddPolyline(allPoints);
            allPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            allPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            float[,] thirdQuarterPoints = new float[,]
                {
                    { thirdQuarterList[0].X, thirdQuarterList[0].Y },
                    { thirdQuarterList[15].X, thirdQuarterList[15].Y}
                };

            PowerPoint.Shape thirdQuarterPolygon2 = slide.Shapes.AddPolyline(thirdQuarterPoints);
            thirdQuarterPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            thirdQuarterPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            float[,] halfPoints = new float[,]
                {
                    { halfList[0].X, halfList[0].Y },
                    { halfList[15].X, halfList[15].Y}
                };

            PowerPoint.Shape halfPolygon2 = slide.Shapes.AddPolyline(halfPoints);
            halfPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            halfPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            float[,] quarterPoints = new float[,]
                {
                    { quarterList[0].X, quarterList[0].Y },
                    { quarterList[15].X, quarterList[15].Y}
                };

            PowerPoint.Shape quarterPolygon2 = slide.Shapes.AddPolyline(quarterPoints);
            quarterPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            quarterPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 5, centerY + radius + 70, 1000, 30)
                .TextFrame.TextRange.Text = "Genetic Origin and Alteration Tool Radar Plot for study waters. Ratio categories: \nwater evolution(EV), geothermometers(GT), lithology(Lith), salinity source(SS), and organic matter related(OM) \nare listed in front of each axis label.";
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.Shapes[slide.Shapes.Count].TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
            slide.Shapes[slide.Shapes.Count].TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginLeft = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginRight = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginTop = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginBottom = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            for (int i = 0; i < sampleData.Length; i++)
            {
                PointF[] points = new PointF[numAxes];

                for (int j = 0; j < numAxes; j++)
                {
                    clsRadarScale value = sampleData[i][j];

                    // Normalize value to be within 0 and its max scale
                    double normalizedValue = Math.Min(value.Item / value.Scale, 1.0); // Ensures it doesn't exceed 1

                    // Scale the normalized value according to its axis' maximum
                    float scaledRadius = (float)(normalizedValue * radius);

                    double angle = j * angleIncrement;
                    float x = centerX + (float)(scaledRadius * Math.Cos(angle));
                    float y = centerY + (float)(scaledRadius * Math.Sin(angle));
                    if (double.IsNaN(x))
                    {
                        x = centerX;
                    }
                    if (double.IsNaN(y))
                    {
                        y = centerY;
                    }
                    points[j] = new PointF(x, y);
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
                    {polylinePoints[14],polylinePoints[15] },
                    {polylinePoints[16],polylinePoints[17] },
                    {polylinePoints[18],polylinePoints[19] },
                    {polylinePoints[20],polylinePoints[21] },
                    {polylinePoints[22],polylinePoints[23] },
                    {polylinePoints[24],polylinePoints[25] },
                    {polylinePoints[26],polylinePoints[27] },
                    {polylinePoints[28],polylinePoints[29] },
                    {polylinePoints[30],polylinePoints[31] },

                });
                setLine.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color); // Set line color
                setLine.Line.Weight = frmImportSamples.WaterData[i].lineWidth; // Set line width
                setLine.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
                PowerPoint.Shape setLine2 = slide.Shapes.AddPolyline(new float[,]
                    {
                      { polylinePoints[0], polylinePoints[1] },
                      { polylinePoints[30],polylinePoints[31]},
                    }
                );
                setLine2.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color); // Set line color
                setLine2.Line.Weight = frmImportSamples.WaterData[i].lineWidth; // Set line width
                setLine2.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
            }
            #region Draw Legend
            if (frmImportSamples.WaterData.Count > 0)
            {
                int legendY = 50;

                // Add metadata
                float metadataX = 500;
                float metadataY = legendY;
                int metaWidth = 0;
                int metaHeight = 0;


                float ysample = metadataY;
                //List<PowerPoint.Shape> addedTexts = new List<PowerPoint.Shape>();

                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    var line = slide.Shapes.AddLine(metadataX, ysample + 10, metadataX + 20, ysample + 10);
                    line.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color);
                    line.Line.Weight = frmImportSamples.WaterData[i].lineWidth;
                    line.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
                    string fullText = "W" + (i + 1).ToString() + "," +
                        frmImportSamples.WaterData[i].Well_Name + "," +
                        frmImportSamples.WaterData[i].ClientID + "," +
                        frmImportSamples.WaterData[i].Depth;

                    PowerPoint.Shape metadataText = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        metadataX + 22, ysample, 500, 20);

                    metadataText.TextFrame.TextRange.Text = fullText;
                    metadataText.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                    metadataText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    metadataText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    metadataText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                    metadataText.TextFrame.MarginLeft = 0;
                    metadataText.TextFrame.MarginRight = 0;
                    metadataText.TextFrame.MarginTop = 0;
                    metadataText.TextFrame.MarginBottom = 0;
                    metadataText.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;



                    metaWidth = Math.Max(metaWidth, (int)metadataText.Width + 30);
                    ysample += metadataText.Height;
                    metaHeight += (int)metadataText.Height + 1;
                    //addedTexts.Add(metadataText);
                }

                // Now draw the border box *after*
                PowerPoint.Shape metaBorder = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    metadataX, metadataY, metaWidth, metaHeight);
                metaBorder.Fill.Transparency = 1.0f;
                metaBorder.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                metaBorder.Line.Weight = 2;
                // Refresh PowerPoint slide
                //pptApplication.ActiveWindow.View.GotoSlide(presentation.Slides.Count);
            }
            #endregion

        }
        public static void DrawRadarChart3(Graphics g, Rectangle bounds, bool flag)
        {
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.legendPictureBoxRadar;
            frmMainForm.mainChartPlotting.Invalidate();
            // Data labels and values
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];



            float fontSize = 12; // Make font size relative
            // Data labels and values
            PrecomputeMaxValues(flag);

            Radar3Scales = new string[] { 
            maxNa.ToString("F5"),
            maxK.ToString("F5"),
            maxCa.ToString("F5"),
            maxMg.ToString("F5"),
            maxAl.ToString("F5"),
            maxCo.ToString("F5"),
            maxCu.ToString("F5"),
            maxMn.ToString("F5"),
            maxNi.ToString("F5"),
            maxSr.ToString("F5"),
            maxZn.ToString("F5"),
            maxBa.ToString("F5"),
            maxPb.ToString("F5"),
            maxFe.ToString("F5"),
            maxCd.ToString("F5"),
            maxCr.ToString("F5"),
            maxTl.ToString("F5"),
            maxBe.ToString("F5"),
            maxSe.ToString("F5"),
            maxB.ToString("F5"),
            maxLi.ToString("F5") };
            List<string> scales = new List<string>();
            for (int i = 0; i < Radar3Scales.Count(); i++)
            {
                string s = Radar3Scales[i];
                string temp="";
                bool checking=false;
                bool found = false;
                for (int j = 0; j < s.Length; j++)
                {
                    if (s[j] == '.')
                    {
                        checking = true;
                    }
                    else if (s[j] != '0' && checking)
                    {
                        found = true;
                        temp += s[j];
                        break;
                    }
                    temp += s[j];
                }
                if (!found)
                {
                    int dotIndex = temp.IndexOf('.');
                    if (dotIndex != -1)
                    {
                        temp = temp.Substring(0, dotIndex);
                    }
                }
                scales.Add(temp);
                Radar3Scales[i] = temp;
            }

            string[] labels =
            {
            "Co \n"+scales[5],
            "Cu \n"+ scales[6],
            "Mn \n"+scales[7],
            "Ni \n"+scales[8],
            "Sr \n"+ scales[9],
            "Zn \n"+ scales[10],
            "Ba \n"+ scales[11],
            "Pb \n"+ scales[12],
            "Fe \n"+ scales[13],
            "Cd \n"+ scales[14],
            "Cr \n"+ scales[15],
            "Tl \n"+ scales[16],
            "Be \n"+ scales[17],
            "Se \n"+ scales[18],
            "B \n"+scales[19],
            "Li \n"+ scales[20],
            "Na \n"+scales[0],
            "K \n"+scales[1],
            "Ca \n"+scales[2],
            "Mg \n"+scales[3],
            "Al \n"+scales[4]

        };

            // Initialize jagged array
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                sampleData[i] = new clsRadarScale[]
                {
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Co), Scale = maxCo },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cu), Scale = maxCu },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Mn), Scale = maxMn },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ni), Scale = maxNi },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Sr), Scale = maxSr },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Zn), Scale = maxZn },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ba), Scale = maxBa },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Pb), Scale = maxPb },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Fe), Scale = maxFe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cd), Scale = maxCd },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cr), Scale = maxCr },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Tl), Scale = maxTl },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Be), Scale = maxBe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Se), Scale = maxSe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].B),  Scale = maxB },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Li), Scale = maxLi },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Na), Scale = maxNa },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].K),  Scale = maxK },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ca), Scale = maxCa },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Mg), Scale = maxMg },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Al), Scale = maxAl },
                };
            }
            // Colors for the samples
            Color[] colors = new Color[frmImportSamples.WaterData.Count];
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                colors[i] = frmImportSamples.WaterData[i].color;
            }

            // Center of the diagram
            float centerX = 0.3f * frmMainForm.mainChartPlotting.Width;
            float centerY = 0.4f * frmMainForm.mainChartPlotting.Height;

            // Radius of the radar diagram
            float radius = Math.Min(bounds.Width, bounds.Height) / 3;

            // Number of axes
            int numAxes = labels.Length;

            // Angle between each axis (in radians)
            double angleIncrement = 2 * Math.PI / numAxes;

            // Draw the axes

            Pen axisPen = new Pen(Color.LightGray, 1);
            Font AxisFont = new Font("Times New Roman", fontSize, FontStyle.Bold);
            axisPen.DashStyle = DashStyle.Dot;
            PointF[] quarterList = new PointF[numAxes];
            PointF[] halfList = new PointF[numAxes];
            PointF[] thirdQuarterList = new PointF[numAxes];
            PointF[] allList = new PointF[numAxes];
            for (int i = 0; i < numAxes; i++)
            {
                double angle = i * angleIncrement;
                float allX = centerX + (float)(radius * Math.Cos(angle));
                float allY = centerY + (float)(radius * Math.Sin(angle));
                g.DrawLine(axisPen, centerX, centerY, allX, allY);
                float quarterX = centerX + (float)((0.25 * radius) * Math.Cos(angle));
                float quarterY = centerY + (float)((0.25 * radius) * Math.Sin(angle));
                float halfX = centerX + (float)((0.5 * radius) * Math.Cos(angle));
                float halfY = centerY + (float)((0.5 * radius) * Math.Sin(angle));
                float thirdQuarterX = centerX + (float)((0.75 * radius) * Math.Cos(angle));
                float thirdQuarterY = centerY + (float)((0.75 * radius) * Math.Sin(angle));
                allList[i] = new PointF(allX, allY);
                quarterList[i] = new PointF(quarterX, quarterY);
                halfList[i] = new PointF(halfX, halfY);
                thirdQuarterList[i] = new PointF(thirdQuarterX, thirdQuarterY);
                // Draw axis labels
                string label = labels[i];
                SizeF labelSize = g.MeasureString(label, AxisFont);
                float labelX = (float)(allX + (0.3 * radius) * Math.Cos(angle));
                float labelY = (float)(allY + (0.3 * radius) * Math.Sin(angle));
                labelX -= 0.02f * frmMainForm.mainChartPlotting.Width;
                g.DrawString(label, AxisFont, Brushes.Black, labelX, labelY);
            }
            g.DrawPolygon(axisPen, quarterList);
            g.DrawPolygon(axisPen, allList);
            g.DrawPolygon(axisPen, halfList);
            g.DrawPolygon(axisPen, thirdQuarterList);

            g.DrawString("ICP Reproducibility", new Font("Times New Roman", fontSize, FontStyle.Bold), Brushes.Black, 0.2f * frmMainForm.mainChartPlotting.Width, 0.9f * frmMainForm.mainChartPlotting.Height);
            
            g.SmoothingMode = SmoothingMode.AntiAlias;
            for (int s = 0; s < sampleData.Length; s++)
            {
                PointF[] points = new PointF[numAxes];

                for (int i = 0; i < numAxes; i++)
                {
                    clsRadarScale value = sampleData[s][i];

                    // Normalize value to be within 0 and its max scale
                    double normalizedValue = Math.Min(value.Item / value.Scale, 1.0); // Ensures it doesn't exceed 1

                    // Scale the normalized value according to its axis' maximum
                    float scaledRadius = (float)(normalizedValue * radius);

                    double angle = i * angleIncrement;
                    float x = centerX + (float)(scaledRadius * Math.Cos(angle));
                    float y = centerY + (float)(scaledRadius * Math.Sin(angle));
                    if(double.IsNaN(x))
                    {
                        x = centerX;
                    }
                    if(double.IsNaN(y))
                    {
                        y = centerY;
                    }
                    points[i] = new PointF(x, y);
                    
                    
                }
                Pen polygonPen = new Pen(colors[s], 2f);
                polygonPen.Width = frmImportSamples.WaterData[s].lineWidth;
                polygonPen.DashStyle = frmImportSamples.WaterData[s].selectedStyle;

                g.DrawPolygon(polygonPen, points);
                


            }
            #region Draw Radar legend
            Radar_legend(g, bounds);
            #endregion
            flag = false;
        }
        public static void ExportRadar3ToPowerpoint(Rectangle bounds, PowerPoint.Slide slide, PowerPoint.Presentation presentation, bool flag)
        {
            #region Setup

            // Process scales to be cleaner strings
            List<string> scales = new List<string>();
            foreach (var value in Radar3Scales)
            {
                string s = value;
                string temp = "";
                bool checking = false, found = false;

                for (int j = 0; j < s.Length; j++)
                {
                    if (s[j] == '.')
                        checking = true;
                    else if (s[j] != '0' && checking)
                    {
                        found = true;
                        temp += s[j];
                        break;
                    }
                    temp += s[j];
                }

                if (!found)
                {
                    int dotIndex = temp.IndexOf('.');
                    if (dotIndex != -1)
                        temp = temp.Substring(0, dotIndex);
                }

                scales.Add(temp);
            }

            // Axis labels based on elements
            string[] elements = { 
            "Co \n"+scales[5],
            "Cu \n"+ scales[6],
            "Mn \n"+scales[7],
            "Ni \n"+scales[8],
            "Sr \n"+ scales[9],
            "Zn \n"+ scales[10],
            "Ba \n"+ scales[11],
            "Pb \n"+ scales[12],
            "Fe \n"+ scales[13],
            "Cd \n"+ scales[14],
            "Cr \n"+ scales[15],
            "Tl \n"+ scales[16],
            "Be \n"+ scales[17],
            "Se \n"+ scales[18],
            "B \n"+scales[19],
            "Li \n"+ scales[20],
            "Na \n"+scales[0],
            "K \n"+scales[1],
            "Ca \n"+scales[2],
            "Mg \n"+scales[3],
            "Al \n"+scales[4] };

            int numAxes = elements.Length;
            double angleIncrement = 2 * Math.PI / numAxes;

            // Radar center and radius
            // Center of the diagram
            

            #endregion

            #region Build Sample Data

            var data = frmImportSamples.WaterData;
            clsRadarScale[][] sampleData = new clsRadarScale[data.Count][];

            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                sampleData[i] = new clsRadarScale[]
                {
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Co), Scale = maxCo },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cu), Scale = maxCu },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Mn), Scale = maxMn },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ni), Scale = maxNi },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Sr), Scale = maxSr },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Zn), Scale = maxZn },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ba), Scale = maxBa },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Pb), Scale = maxPb },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Fe), Scale = maxFe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cd), Scale = maxCd },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cr), Scale = maxCr },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Tl), Scale = maxTl },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Be), Scale = maxBe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Se), Scale = maxSe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].B),  Scale = maxB },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Li), Scale = maxLi },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Na), Scale = maxNa },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].K),  Scale = maxK },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ca), Scale = maxCa },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Mg), Scale = maxMg },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Al), Scale = maxAl },
                };
            }

            #endregion

            // Colors for the samples
            Color[] colors = new Color[frmImportSamples.WaterData.Count];
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                colors[i] = frmImportSamples.WaterData[i].color;
            }

            // Center of the diagram
            float centerX = (bounds.Width / 2);
            float centerY = bounds.Y + bounds.Height / 2 - 50;


            // Radius of the radar diagram
            float radius = Math.Min(bounds.Width, bounds.Height) / 3;
            // Number of axes

            PointF[] quarterList = new PointF[numAxes];
            PointF[] halfList = new PointF[numAxes];
            PointF[] thirdQuarterList = new PointF[numAxes];
            PointF[] allList = new PointF[numAxes];
            for (int i = 0; i < numAxes; i++)
            {
                double angle = i * angleIncrement;
                float allX = centerX + (float)(radius * Math.Cos(angle));
                float allY = centerY + (float)(radius * Math.Sin(angle));
                float quarterX = centerX + (float)((0.25 * radius) * Math.Cos(angle));
                float quarterY = centerY + (float)((0.25 * radius) * Math.Sin(angle));
                float halfX = centerX + (float)((0.5 * radius) * Math.Cos(angle));
                float halfY = centerY + (float)((0.5 * radius) * Math.Sin(angle));
                float thirdQuarterX = centerX + (float)((0.75 * radius) * Math.Cos(angle));
                float thirdQuarterY = centerY + (float)((0.75 * radius) * Math.Sin(angle));
                allList[i] = new PointF(allX, allY);
                quarterList[i] = new PointF(quarterX, quarterY);
                halfList[i] = new PointF(halfX, halfY);
                thirdQuarterList[i] = new PointF(thirdQuarterX, thirdQuarterY);
                var radiusLine = slide.Shapes.AddLine(centerX, centerY, allX, allY);
                radiusLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
                radiusLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                string label = elements[i];
                SizeF labelSize = TextRenderer.MeasureText(label, SystemFonts.DefaultFont);
                float labelX = allX + (float)((radius * 0.3) * Math.Cos(angle)) - labelSize.Width / 2;
                float labelY = allY + (float)((radius * 0.3) * Math.Sin(angle)) - labelSize.Height / 2;

                PowerPoint.Shape itemText = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, labelX, labelY, 200, 30);
                itemText.TextFrame.TextRange.Text = label;
                itemText.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                itemText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                itemText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                itemText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                itemText.TextFrame.MarginLeft = 0;
                itemText.TextFrame.MarginRight = 0;
                itemText.TextFrame.MarginTop = 0;
                itemText.TextFrame.MarginBottom = 0;
                itemText.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            }

            for (int i = 0; i < allList.Length - 1; i++)
            {
                float[,] points1 = new float[,]
                {
                    { allList[i].X, allList[i].Y },
                    { allList[i + 1].X, allList[i + 1].Y}
                };

                PowerPoint.Shape allPolygon = slide.Shapes.AddPolyline(points1);
                allPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                allPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                float[,] points2 = new float[,]
                {
                    { thirdQuarterList[i].X, thirdQuarterList[i].Y },
                    { thirdQuarterList[i + 1].X, thirdQuarterList[i + 1].Y}
                };

                PowerPoint.Shape thirdQuarterPolygon = slide.Shapes.AddPolyline(points2);
                thirdQuarterPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                thirdQuarterPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                float[,] points3 = new float[,]
                {
                    { halfList[i].X, halfList[i].Y },
                    { halfList[i + 1].X, halfList[i + 1].Y}
                };

                PowerPoint.Shape halfPolygon = slide.Shapes.AddPolyline(points3);
                halfPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                halfPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                float[,] points4 = new float[,]
                {
                    { quarterList[i].X, quarterList[i].Y },
                    { quarterList[i + 1].X, quarterList[i + 1].Y}
                };

                PowerPoint.Shape quarterPolygon = slide.Shapes.AddPolyline(points4);
                quarterPolygon.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
                quarterPolygon.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            }
            float[,] allPoints = new float[,]
                {
                    { allList[0].X, allList[0].Y },
                    { allList[20].X, allList[20].Y}
                };

            PowerPoint.Shape allPolygon2 = slide.Shapes.AddPolyline(allPoints);
            allPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            allPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            float[,] thirdQuarterPoints = new float[,]
                {
                    { thirdQuarterList[0].X, thirdQuarterList[0].Y },
                    { thirdQuarterList[20].X, thirdQuarterList[20].Y}
                };

            PowerPoint.Shape thirdQuarterPolygon2 = slide.Shapes.AddPolyline(thirdQuarterPoints);
            thirdQuarterPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            thirdQuarterPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            float[,] halfPoints = new float[,]
                {
                    { halfList[0].X, halfList[0].Y },
                    { halfList[20].X, halfList[20].Y}
                };

            PowerPoint.Shape halfPolygon2 = slide.Shapes.AddPolyline(halfPoints);
            halfPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            halfPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            float[,] quarterPoints = new float[,]
                {
                    { quarterList[0].X, quarterList[0].Y },
                    { quarterList[20].X, quarterList[20].Y}
                };

            PowerPoint.Shape quarterPolygon2 = slide.Shapes.AddPolyline(quarterPoints);
            quarterPolygon2.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray); // Set line color
            quarterPolygon2.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
            slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 5, centerY + radius + 70, 1000, 30)
                .TextFrame.TextRange.Text = "ICP Reproducibility";
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.Shapes[slide.Shapes.Count].TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
            slide.Shapes[slide.Shapes.Count].TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginLeft = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginRight = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginTop = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame.MarginBottom = 0;
            slide.Shapes[slide.Shapes.Count].TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            for (int i = 0; i < sampleData.Length; i++)
            {
                PointF[] points = new PointF[numAxes];

                for (int j = 0; j < numAxes; j++)
                {
                    clsRadarScale value = sampleData[i][j];

                    // Normalize value to be within 0 and its max scale
                    double normalizedValue = Math.Min(value.Item / value.Scale, 1.0); // Ensures it doesn't exceed 1

                    // Scale the normalized value according to its axis' maximum
                    float scaledRadius = (float)(normalizedValue * radius);

                    double angle = j * angleIncrement;
                    float x = centerX + (float)(scaledRadius * Math.Cos(angle));
                    float y = centerY + (float)(scaledRadius * Math.Sin(angle));
                    if (double.IsNaN(x))
                    {
                        x = centerX;
                    }
                    if (double.IsNaN(y))
                    {
                        y = centerY;
                    }
                    points[j] = new PointF(x, y);
                }
                for(int j=0;j<points.Length-1;j++)
                {
                    
                    PowerPoint.Shape polygon= slide.Shapes.AddPolyline(new float[,] { { points[j].X, points[j].Y }, { points[j + 1].X, points[j + 1].Y } });
                    polygon.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color); // Set line color
                    polygon.Line.Weight = frmImportSamples.WaterData[i].lineWidth; // Set line width
                    polygon.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
                }
                
                PowerPoint.Shape polygonLastline = slide.Shapes.AddPolyline(new float[,] { { points[0].X, points[0].Y }, { points[20].X, points[20].Y } });
                polygonLastline.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color); // Set line color
                polygonLastline.Line.Weight = frmImportSamples.WaterData[i].lineWidth; // Set line width
                polygonLastline.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
                
                
               
            }
            #region Draw Legend
            if (frmImportSamples.WaterData.Count > 0)
            {
                int legendY = 50;

                // Add metadata
                float metadataX = 500;
                float metadataY = legendY;
                int metaWidth = 0;
                int metaHeight = 0;


                float ysample = metadataY;
                //List<PowerPoint.Shape> addedTexts = new List<PowerPoint.Shape>();

                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    var line = slide.Shapes.AddLine(metadataX, ysample + 10, metadataX + 20, ysample + 10);
                    line.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color);
                    line.Line.Weight = frmImportSamples.WaterData[i].lineWidth;
                    line.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
                    string fullText = "W" + (i + 1).ToString() + "," +
                        frmImportSamples.WaterData[i].Well_Name + "," +
                        frmImportSamples.WaterData[i].ClientID + "," +
                        frmImportSamples.WaterData[i].Depth;

                    PowerPoint.Shape metadataText = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        metadataX + 22, ysample, 500, 20);

                    metadataText.TextFrame.TextRange.Text = fullText;
                    metadataText.TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
                    metadataText.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    metadataText.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    metadataText.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
                    metadataText.TextFrame.MarginLeft = 0;
                    metadataText.TextFrame.MarginRight = 0;
                    metadataText.TextFrame.MarginTop = 0;
                    metadataText.TextFrame.MarginBottom = 0;
                    metadataText.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;



                    metaWidth = Math.Max(metaWidth, (int)metadataText.Width + 30);
                    ysample += metadataText.Height;
                    metaHeight += (int)metadataText.Height + 1;
                    //addedTexts.Add(metadataText);
                }

                // Now draw the border box *after*
                PowerPoint.Shape metaBorder = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    metadataX, metadataY, metaWidth, metaHeight);
                metaBorder.Fill.Transparency = 1.0f;
                metaBorder.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.Blue);
                metaBorder.Line.Weight = 2;
                // Refresh PowerPoint slide
                //pptApplication.ActiveWindow.View.GotoSlide(presentation.Slides.Count);
            }
            #endregion
        }

        private static void PrecomputeMaxValues(bool flag)
        {
            if (flag) return;
            maxCl = maxNa = maxK = maxCa = maxMg = maxBa = maxSr = 0;
            foreach (var data in frmImportSamples.WaterData)
            {
                maxCl = Math.Max(maxCl, Math.Abs(data.Cl) / Bm);
                maxNa = Math.Max(maxNa, Math.Abs(data.Na)/Bn);
                maxK = Math.Max(maxK, Math.Abs(data.K)/Bo);
                maxCa = Math.Max(maxCa, Math.Abs(data.Ca)/Bp);
                maxMg = Math.Max(maxMg, Math.Abs(data.Mg)/Bq);
                maxBa = Math.Max(maxBa, Math.Abs(data.Ba)/Br);
                maxSr = Math.Max(maxSr, Math.Abs(data.Sr)/Bs);
                maxNaCl = Math.Max(maxNaCl, Math.Abs(data.Na) / Math.Abs(data.Cl));
                maxClCa = Math.Max(maxClCa, Math.Abs(data.Cl) / Math.Abs(data.Ca));
                maxHCO3Cl = Math.Max(maxHCO3Cl, Math.Abs(data.HCO3) / Math.Abs(data.Cl));
                maxClSr = Math.Max(maxClSr, Math.Abs(data.Cl) / Math.Abs(data.Sr));
                maxNaCa = Math.Max(maxNaCa, Math.Abs(data.Na) / Math.Abs(data.Ca));
                maxKNa = Math.Max(maxKNa, Math.Abs(data.K) / Math.Abs(data.Na));
                maxSrMg = Math.Max(maxSrMg, Math.Abs(data.Sr) / Math.Abs(data.Mg));
                maxMgCl = Math.Max(maxMgCl, Math.Abs(data.Mg) / Math.Abs(data.Cl));
                maxSrCl = Math.Max(maxSrCl, Math.Abs(data.Sr) / Math.Abs(data.Cl));
                maxSrK = Math.Max(maxSrK, Math.Abs(data.Sr / data.K));
                maxMgK = Math.Max(maxMgK, Math.Abs(data.Mg / data.K));
                maxCaK = Math.Max(maxCaK, Math.Abs(data.Ca / data.K));
                maxtK = Math.Max(maxtK, Math.Abs(data.K) / 10000);
                maxBCl = Math.Max(maxBCl, Math.Abs(data.B / data.Cl));
                maxBNa = Math.Max(maxBNa, Math.Abs(data.B / data.Na));
                maxBMg = Math.Max(maxBMg, Math.Abs(data.B / data.Mg));
                maxAl = Math.Max(maxAl, Math.Abs(data.Al));
                maxCo = Math.Max(maxCo, Math.Abs(data.Co));
                maxCu = Math.Max(maxCu, Math.Abs(data.Cu));
                maxMn = Math.Max(maxMn, Math.Abs(data.Mn));
                maxNi = Math.Max(maxNi, Math.Abs(data.Ni));
                maxZn = Math.Max(maxZn, Math.Abs(data.Zn));
                maxPb = Math.Max(maxPb, Math.Abs(data.Pb));
                maxFe = Math.Max(maxFe, Math.Abs(data.Fe));
                maxCd = Math.Max(maxCd, Math.Abs(data.Cd));
                maxCr = Math.Max(maxCr, Math.Abs(data.Cr));
                maxTl = Math.Max(maxTl, Math.Abs(data.Tl));
                maxBe = Math.Max(maxBe, Math.Abs(data.Be));
                maxSe = Math.Max(maxSe, Math.Abs(data.Se));
                maxB = Math.Max(maxB, Math.Abs(data.B));
                maxLi = Math.Max(maxLi, Math.Abs(data.Li));
            }
            

        }

        public static Office.MsoLineDashStyle ConvertDashStyle(DashStyle dashStyle)
        {
            switch (dashStyle)
            {
                case DashStyle.Solid:
                    return Office.MsoLineDashStyle.msoLineSolid;
                case DashStyle.Dash:
                    return Office.MsoLineDashStyle.msoLineDash;
                case DashStyle.DashDot:
                    return Office.MsoLineDashStyle.msoLineDashDot;
                case DashStyle.DashDotDot:
                    return Office.MsoLineDashStyle.msoLineDashDotDot;
                case DashStyle.Dot:
                    return Office.MsoLineDashStyle.msoLineSquareDot;
                default:
                    return Office.MsoLineDashStyle.msoLineSolid;
            }
        }


    }
}
