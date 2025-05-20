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
        /// <summary>
        /// Scale values for First Radar Diagram
        /// </summary>
        public static double maxCl = 0, maxNa1 = 0, maxK1 = 0, maxCa1 = 0, maxMg1 = 0, maxBa1 = 0, maxSr1 = 0;
        /// <summary>
        /// Scale values for Second Radar Diagram
        /// </summary>
        public static double maxNaCl = 0, maxClCa = 0, maxHCO3Cl = 0, maxClSr = 0, maxNaCa = 0, maxKNa = 0, maxSrMg = 0, maxMgCl = 0, maxSrCl = 0, maxSrK = 0, maxMgK = 0, maxCaK = 0, maxtK = 0, maxBCl = 0, maxBNa = 0, maxBMg = 0;
        /// <summary>
        /// Scale values for Third Radar Diagram
        /// </summary>
        public static double maxNa3 = 0, maxK3 = 0, maxCa3 = 0, maxMg3 = 0, maxBa3 = 0, maxSr3 = 0,maxAl = 0, maxCo = 0, maxCu = 0, maxMn = 0, maxNi = 0, maxZn = 0, maxPb = 0, maxFe = 0, maxCd = 0, maxCr = 0, maxTl = 0, maxBe = 0, maxSe = 0, maxLi=0,maxB=0;
        /// <summary>
        /// Factors to convert into (mol/L)
        /// </summary>
        public static double Bm = 35453, Bn = 22989.7, Bo = 39098.3, Bp = 40078, Bq = 24305, Br = 137327, Bs = 87620;
        public static TextBox txt;
        public static string[] Radar1Scales=new string[7];
        public static string[] Radar2Scales = new string[16];
        public static string[] Radar3Scales = new string[21];
        /// <summary>
        /// Draws a legend for the radar chart showing sample information including well name, client ID, and depth.
        /// The legend is positioned on the right side of the chart and includes colored lines matching each sample's style.
        /// </summary>
        /// <param name="g">Graphics object used for drawing</param>
        /// <param name="bounds">Bounds of the radar chart area</param>
        public static void RadarLegend(Graphics g, Rectangle bounds)
        {
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
                    for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
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
                            fullText += data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
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
                            fullText += data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
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
        /// Draws a legend for the radar chart in PowerPoint format, showing sample information including well name, client ID, and depth.
        /// The legend is positioned on the right side of the slide and includes colored lines matching each sample's style.
        /// </summary>
        /// <param name="slide">PowerPoint slide to draw the legend on</param>
        public static void RadarLegendPowerpoint(PowerPoint.Slide slide)
        {
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
                    var line = slide.Shapes.AddLine(metadataX, ysample+5, metadataX + 20, ysample+5);
                    line.Line.ForeColor.RGB = ColorTranslator.ToOle(data.color);
                    line.Line.Weight = data.lineWidth;
                    line.Line.DashStyle = ConvertDashStyle(data.selectedStyle);

                    // Prepare wrapped text
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
            }
            #endregion
        }

        /// <summary>
        /// Draws the first radar chart showing molar concentrations of major ions (Cl, Na, K, Ca, Mg, Ba, Sr).
        /// Each axis represents a different ion and shows its maximum concentration value.
        /// </summary>
        /// <param name="g">Graphics object used for drawing</param>
        /// <param name="bounds">Bounds of the radar chart area</param>
        /// <param name="flag">Flag to control whether to recompute maximum values</param>
        public static void DrawRadarChart1(Graphics g, Rectangle bounds, bool flag)
        {
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.legendPictureBoxRadar;
            frmMainForm.mainChartPlotting.Invalidate();
            
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];

            string title = "Elements Molar concentration";
            Font titleFont = new Font("Times New Roman", 25, FontStyle.Bold);
            int titleX = (int)(frmMainForm.mainChartPlotting.Width * 0.4f);
            int titleY = (int)(0.01f * frmMainForm.mainChartPlotting.Height);
            g.DrawString(title, titleFont, Brushes.Black, titleX, titleY);

            float fontSize = 12; // Make font size relative
            PrecomputeMaxValues(flag);


            Radar1Scales = new string[] { maxCl.ToString(), maxNa1.ToString(), maxK1.ToString(), maxCa1.ToString(), maxMg1.ToString(), maxBa1.ToString(), maxSr1.ToString() };

            Font AxisFont = new Font("Times New Roman", fontSize, FontStyle.Bold);
            List<string> scales = new List<string>();
            for (int i = 0; i < Radar1Scales.Count(); i++)
            {
                string s = Radar1Scales[i];
                decimal parsedValue;

                //Try to parse the number and use the parsed value for more accurate formatting
                if (decimal.TryParse(s, out parsedValue))
                    {
                        //Format the number to remove unnecessary trailing zeros and retain precision
                    string formattedValue = parsedValue.ToString("0.#######");  // Adjust the number of '#' based on desired precision
                        scales.Add(formattedValue);
                        Radar1Scales[i] = formattedValue;
                    }
                    else
                    {
                        //If parsing fails, just add the original string(error handling)
                        scales.Add(s);
                        Radar1Scales[i] = s;
                    }
            }
            for (int i = 0; i < Radar1Scales.Length; i++)
            {
                string fullString = Radar1Scales[i];
                if (fullString.Contains("E-"))
                {
                    int eIndex = fullString.IndexOf("E");
                    string temp = "";
                    for (int j = 0; j < fullString.Length; j++)
                    {
                        if (fullString[j] != '.')
                        {
                            temp += fullString[j];
                        }
                        else
                        {
                            temp += fullString[j];
                            temp += fullString[j + 1];
                            break;
                        }

                    }
                    for (int j = eIndex; j < fullString.Length; j++)
                    {
                        temp += fullString[j];
                    }
                    fullString = temp;
                    Radar1Scales[i] = fullString;

                }
            }
            string[] labels =
            {
                "K (mol/L)\n"+ Radar1Scales[2],
                "Ca (mol/L)\n"+ Radar1Scales[3],
                "Mg (mol/L)\n"+ Radar1Scales[4],
                "Ba (mol/L)\n"+ Radar1Scales[5],
                "Sr (mol/L)\n"+ Radar1Scales[6],
                "Cl (mol/L)\n"+ Radar1Scales[0],
                "Na (mol/L)\n"+ Radar1Scales[1]
            };

            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                sampleData[i] = new clsRadarScale[]
                {
                new clsRadarScale { Item = frmImportSamples.WaterData[i].K / Bo, Scale = maxK1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Ca / Bp, Scale = maxCa1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Mg / Bq, Scale = maxMg1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Ba / Br, Scale = maxBa1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / Bs, Scale = maxSr1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Cl / Bm, Scale = maxCl},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Na / Bn, Scale = maxNa1}
                };
            }
            //Colors for the samples

           Color[] colors = new Color[frmImportSamples.WaterData.Count];
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                colors[i] = frmImportSamples.WaterData[i].color;
            }

            //Center of the diagram
            float centerX = 0.3f * frmMainForm.mainChartPlotting.Width;
            float centerY = 0.4f * frmMainForm.mainChartPlotting.Height;

            //Radius of the radar diagram
            float radius = Math.Min(bounds.Width, bounds.Height) / 3;

            //Number of axes
            int numAxes = labels.Length;

            //Angle between each axis(in radians)
            double angleIncrement = 2 * Math.PI / numAxes;

            //Draw the axes

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
                //Draw axis labels
                string label = labels[i];
                SizeF labelSize = g.MeasureString(label, AxisFont);
                float labelX = (float)(allX + (0.2 * radius) * Math.Cos(angle));
                float labelY = (float)(allY + (0.2 * radius) * Math.Sin(angle));
                labelX -= 0.02f * frmMainForm.mainChartPlotting.Width;
                labelY -= 0.02f * frmMainForm.mainChartPlotting.Height;
                g.DrawString(label, AxisFont, Brushes.Black, labelX, labelY);


            }
            g.DrawPolygon(axisPen, quarterList);
            g.DrawPolygon(axisPen, allList);
            g.DrawPolygon(axisPen, halfList);
            g.DrawPolygon(axisPen, thirdQuarterList);

            g.DrawString("Radar diagram showing the molar concentrations for major ions", new Font("Times New Roman", fontSize, FontStyle.Bold), Brushes.Black, 0.2f * frmMainForm.mainChartPlotting.Width, 0.9f * frmMainForm.mainChartPlotting.Height);

            #region Draw Radar legend
            for (int s = 0; s < sampleData.Length; s++)
            {
                PointF[] points = new PointF[numAxes];

                for (int i = 0; i < numAxes; i++)
                {
                    clsRadarScale value = sampleData[s][i];

                    //Normalize value to be within 0 and its max scale
                    double normalizedValue = Math.Min(value.Item / value.Scale, 1.0);

                    //Scale the normalized value according to its axis' maximum
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

                //Draw the polygon by connecting points
                using (Pen linePen = new Pen(colors[s], 2))
                {
                    linePen.Color = frmImportSamples.WaterData[s].color;
                    linePen.DashStyle = frmImportSamples.WaterData[s].selectedStyle;
                    linePen.Width = frmImportSamples.WaterData[s].lineWidth;
                    g.DrawPolygon(linePen, points);
                }
            }
            RadarLegend(g, bounds);
            #endregion

            flag = false;
        }
        /// <summary>
        /// Exports the first radar chart to PowerPoint, showing molar concentrations of major ions.
        /// Creates a new slide with the chart and legend in PowerPoint format.
        /// </summary>
        /// <param name="bounds">Bounds of the radar chart area</param>
        /// <param name="slide">PowerPoint slide to draw on</param>
        /// <param name="presentation">PowerPoint presentation object</param>
        /// <param name="flag">Flag to control whether to recompute maximum values</param>
        public static void ExportRadar1ToPowerpoint(Rectangle bounds, PowerPoint.Slide slide, PowerPoint.Presentation presentation, bool flag)
        {
            PowerPoint.Shape title = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                (presentation.PageSetup.SlideWidth / 2) - 100, clsConstants.chartYPowerpoint, 200, 50);
            title.TextFrame.TextRange.Text = "Elements Molar Concentration";
            title.TextFrame.TextRange.Font.Size = 25;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            title.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            title.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            
            //Data labels and values
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];
            double Bm = 35453, Bn = 22989.7, Bo = 39098.3, Bp = 40078, Bq = 24305, Br = 137327, Bs = 87620;
            PrecomputeMaxValues(flag);

            string[] labels =
            {
                "K (mol/L)\n"+ Radar1Scales[2],
                "Ca (mol/L)\n"+ Radar1Scales[3],
                "Mg (mol/L)\n"+ Radar1Scales[4],
                "Ba (mol/L)\n"+ Radar1Scales[5],
                "Sr (mol/L)\n"+ Radar1Scales[6],
                "Cl (mol/L)\n"+ Radar1Scales[0],
                "Na (mol/L)\n"+ Radar1Scales[1]
            };
            //Initialize jagged array
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                sampleData[i] = new clsRadarScale[]
                {
                new clsRadarScale { Item = frmImportSamples.WaterData[i].K / Bo, Scale = maxK1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Ca / Bp, Scale = maxCa1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Mg / Bq, Scale = maxMg1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Ba / Br, Scale = maxBa1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Sr / Bs, Scale = maxSr1},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Cl / Bm, Scale = maxCl},
                new clsRadarScale { Item = frmImportSamples.WaterData[i].Na / Bn, Scale = maxNa1}
                };
            }

            //Colors for the samples

           Color[] colors = new Color[frmImportSamples.WaterData.Count];
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                colors[i] = frmImportSamples.WaterData[i].color;
            }

            //Center of the diagram
            float centerX = (bounds.Width / 2);
            float centerY = bounds.Y + bounds.Height / 2 - 50;


            //Radius of the radar diagram
            float radius = (float)Math.Min(bounds.Width / 1.5, bounds.Height / 1.5) / 3;
            //Number of axes
            int numAxes = labels.Length;

            //Angle between each axis(in radians)
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
                var radiusLine = slide.Shapes.AddLine(centerX, centerY, allX, allY);
                radiusLine.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Gray);
                radiusLine.Line.DashStyle = Office.MsoLineDashStyle.msoLineRoundDot;
                string label = labels[i];
                SizeF labelSize = TextRenderer.MeasureText(label, SystemFonts.DefaultFont);
                float labelX = allX + (float)((radius * 0.2) * Math.Cos(angle)) - labelSize.Width / 2;
                float labelY = allY + (float)((radius * 0.2) * Math.Sin(angle)) - labelSize.Height / 2;
                labelX += 8;
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
            slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, -200, centerY + radius + 70, 1000, 30)
                .TextFrame.TextRange.Text = "Radar diagram showing the molar concentrations for major ions";
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.Shapes[slide.Shapes.Count].TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
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

                    //Normalize value to be within 0 and its max scale
                    double normalizedValue = Math.Min(value.Item / value.Scale, 1.0); // Ensures it doesn't exceed 1

                    //Scale the normalized value according to its axis' maximum
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
                //Flatten the points array into a float array for AddPolyline

               float[] polylinePoints = new float[points.Length * 2];
                for (int j = 0; j < points.Length; j++)
                    {
                        polylinePoints[j * 2] = points[j].X;
                        polylinePoints[j * 2 + 1] = points[j].Y;
                    }

                //Add the polyline to the slide
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
            RadarLegendPowerpoint(slide);


        }
        /// <summary>
        /// Draws the second radar chart showing genetic origin and alteration ratios.
        /// Includes ratios for water evolution, geothermometers, lithology, salinity source, and organic matter.
        /// </summary>
        /// <param name="g">Graphics object used for drawing</param>
        /// <param name="bounds">Bounds of the radar chart area</param>
        /// <param name="flag">Flag to control whether to recompute maximum values</param>
        public static void DrawRadarChart2(Graphics g, Rectangle bounds, bool flag)
        {
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.legendPictureBoxRadar;
            frmMainForm.mainChartPlotting.Invalidate();
            float fontSize = 12; // Make font size relative
            string title = "Genetic Origin and Alteration\n                  Plot";
            Font titleFont = new Font("Times New Roman", 25, FontStyle.Bold);
            int titleX = (int)(frmMainForm.mainChartPlotting.Width * 0.4f);
            int titleY = (int)(0.01f * frmMainForm.mainChartPlotting.Height);
            g.DrawString(title, titleFont, Brushes.Black, titleX, titleY);
            // Data labels and values
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];
            PrecomputeMaxValues(flag);

            Radar2Scales = new string[] { maxNaCl.ToString("F5"), maxClCa.ToString("F5"), maxHCO3Cl.ToString("F5"), maxClSr.ToString("F5"), maxNaCa.ToString("F5"), maxKNa.ToString("F5"), maxSrMg.ToString("F5"), maxMgCl.ToString("F5"), maxSrCl.ToString("F5"), maxSrK.ToString("F5"), maxMgK.ToString("F5"), maxCaK.ToString("F5"), maxtK.ToString("F5"), maxBCl.ToString("F5"), maxBNa.ToString("F5"), maxBMg.ToString("F5") };
            List<string> scales = new List<string>();
            for (int i = 0; i < Radar2Scales.Count(); i++)
            {
                string s = Radar2Scales[i];
                decimal parsedValue;

                // Try to parse the number and use the parsed value for more accurate formatting
                if (decimal.TryParse(s, out parsedValue))
                {
                    // Format the number to remove unnecessary trailing zeros and retain precision
                    string formattedValue = parsedValue.ToString("0.#######");  // Adjust the number of '#' based on desired precision
                    scales.Add(formattedValue);
                    Radar2Scales[i] = formattedValue;
                }
                else
                {
                    // If parsing fails, just add the original string (error handling)
                    scales.Add(s);
                    Radar2Scales[i] = s;
                }
            }
            for (int i = 0; i < Radar2Scales.Length; i++)
            {
                string fullString = Radar2Scales[i];
                if (fullString.Contains("E-"))
                {
                    int eIndex = fullString.IndexOf("E");
                    string temp = "";
                    for (int j = 0; j < fullString.Length; j++)
                    {
                        if (fullString[j] != '.')
                        {
                            temp += fullString[j];
                        }
                        else
                        {
                            temp += fullString[j];
                            temp += fullString[j + 1];
                            break;
                        }

                    }
                    for (int j = eIndex; j < fullString.Length; j++)
                    {
                        temp += fullString[j];
                    }
                    fullString = temp;
                    Radar2Scales[i] = fullString;

                }
            }
            string[] labels =
            {
            "EV_Na-Ca \n"+Radar2Scales[4],
            "GT_K-Na \n"+ Radar2Scales[5],
            "SS_Sr-Mg \n"+Radar2Scales[6],
            "SS_Mg-Cl \n"+Radar2Scales[7],
            "SS_Sr-Cl \n"+ Radar2Scales[8],
            "Lith_Sr-K \n"+ Radar2Scales[9],
            "Lith_Mg-K \n"+ Radar2Scales[10],
            "Lith_Ca-K \n"+ Radar2Scales[11],
            "Wt%K \n"+ Radar2Scales[12],
            "OM_B-Cl \n"+ Radar2Scales[13],
            "OM_B-Na \n"+ Radar2Scales[14],
            "OM_B-Mg \n"+ Radar2Scales[15],
            "EV_Na-Cl \n"+ Radar2Scales[0],
            "EV_Cl-Ca \n"+ Radar2Scales[1],
            "EV_HCO3-Cl \n"+Radar2Scales[2],
            "EV_Cl-Sr \n"+ Radar2Scales[3]

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
            float centerX = 0.3f * frmMainForm.mainChartPlotting.Width;
            float centerY = 0.4f * frmMainForm.mainChartPlotting.Height;

            // Radius of the radar diagram
            float radius = Math.Min(bounds.Width, bounds.Height) * 0.32f;

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
            RadarLegend(g, bounds);
            #endregion
            flag = false;
        }

        /// <summary>
        /// Exports the second radar chart to PowerPoint, showing genetic origin and alteration ratios.
        /// Creates a new slide with the chart and legend in PowerPoint format.
        /// </summary>
        /// <param name="bounds">Bounds of the radar chart area</param>
        /// <param name="slide">PowerPoint slide to draw on</param>
        /// <param name="presentation">PowerPoint presentation object</param>
        /// <param name="flag">Flag to control whether to recompute maximum values</param>
        public static void ExportRadar2ToPowerpoint(Rectangle bounds, PowerPoint.Slide slide, PowerPoint.Presentation presentation, bool flag)
        {



            PowerPoint.Shape title = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                (presentation.PageSetup.SlideWidth / 2) - 100, clsConstants.chartYPowerpoint, 200, 50);
            title.TextFrame.TextRange.Text = "Genetic Origin and Alteration Plot";
            title.TextFrame.TextRange.Font.Size = 25;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            title.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            //title.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            title.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            // Data labels and values
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];
            PrecomputeMaxValues(flag);

            string[] labels =
            {
            "EV_Na-Ca \n"+Radar2Scales[4],
            "GT_K-Na \n"+ Radar2Scales[5],
            "SS_Sr-Mg \n"+Radar2Scales[6],
            "SS_Mg-Cl \n"+Radar2Scales[7],
            "SS_Sr-Cl \n"+ Radar2Scales[8],
            "Lith_Sr-K \n"+ Radar2Scales[9],
            "Lith_Mg-K \n"+ Radar2Scales[10],
            "Lith_Ca-K \n"+ Radar2Scales[11],
            "Wt%K \n"+ Radar2Scales[12],
            "OM_B-Cl \n"+ Radar2Scales[13],
            "OM_B-Na \n"+ Radar2Scales[14],
            "OM_B-Mg \n"+ Radar2Scales[15],
            "EV_Na-Cl \n"+ Radar2Scales[0],
            "EV_Cl-Ca \n"+ Radar2Scales[1],
            "EV_HCO3-Cl \n"+Radar2Scales[2],
            "EV_Cl-Sr \n"+ Radar2Scales[3]

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
            float radius = (float)Math.Min(bounds.Width / 1.5, bounds.Height / 1.5) / 3;
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
            slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, -200, centerY + radius + 70, 1000, 30)
                .TextFrame.TextRange.Text = "Genetic Origin and Alteration Tool Radar Plot for study waters. Ratio categories: \nwater evolution(EV), geothermometers(GT), lithology(Lith), salinity source(SS), and organic matter related(OM) \nare listed in front of each axis label.";
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            slide.Shapes[slide.Shapes.Count].TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
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
            RadarLegendPowerpoint(slide);

        }
        /// <summary>
        /// Draws the third radar chart showing ICP reproducibility data for various elements.
        /// Includes major ions and trace elements with their concentrations.
        /// </summary>
        /// <param name="g">Graphics object used for drawing</param>
        /// <param name="bounds">Bounds of the radar chart area</param>
        /// <param name="flag">Flag to control whether to recompute maximum values</param>
        public static void DrawRadarChart3(Graphics g, Rectangle bounds, bool flag)
        {
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.legendPictureBoxRadar;
            // Data labels and values
            clsRadarScale[][] sampleData = new clsRadarScale[frmImportSamples.WaterData.Count][];


            string title = "ICP Reproducibility";
            Font titleFont = new Font("Times New Roman", 25, FontStyle.Bold);
            int titleX = (int)(frmMainForm.mainChartPlotting.Width * 0.4f);
            int titleY = (int)(0.01f * frmMainForm.mainChartPlotting.Height);
            g.DrawString(title, titleFont, Brushes.Black, titleX, titleY);
            float fontSize = 12; // Make font size relative
            // Data labels and values
            PrecomputeMaxValues(flag);
            //maxNa = 0;
            //maxK = 0;
            //maxCa = 0;
            //maxMg = 0;
            //foreach (var data in frmImportSamples.WaterData)
            //{
            //    maxNa = Math.Max(maxNa, Math.Abs(data.Na));
            //    maxK = Math.Max(maxK, Math.Abs(data.K));
            //    maxCa = Math.Max(maxCa, Math.Abs(data.Ca));
            //    maxMg = Math.Max(maxMg, Math.Abs(data.Mg));
            //    maxBa = Math.Max(maxBa, Math.Abs(data.Ba));
            //    maxSr = Math.Max(maxSr, Math.Abs(data.Sr));
            //}

            Radar3Scales = new string[] {
            maxNa3.ToString("F5"),
            maxK3.ToString("F5"),
            maxCa3.ToString("F5"),
            maxMg3.ToString("F5"),
            maxAl.ToString("F5"),
            maxCo.ToString("F5"),
            maxCu.ToString("F5"),
            maxMn.ToString("F5"),
            maxNi.ToString("F5"),
            maxSr3.ToString("F5"),
            maxZn.ToString("F5"),
            maxBa3.ToString("F5"),
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
                decimal parsedValue;

                // Try to parse the number and use the parsed value for more accurate formatting
                if (decimal.TryParse(s, out parsedValue))
                {
                    // Format the number to remove unnecessary trailing zeros and retain precision
                    string formattedValue = parsedValue.ToString("0.#######");  // Adjust the number of '#' based on desired precision
                    scales.Add(formattedValue);
                    Radar3Scales[i] = formattedValue;
                }
                else
                {
                    // If parsing fails, just add the original string (error handling)
                    scales.Add(s);
                    Radar3Scales[i] = s;
                }
            }
            for (int i = 0; i < Radar3Scales.Length; i++)
            {
                string fullString = Radar3Scales[i];
                if (fullString.Contains("E-"))
                {
                    int eIndex = fullString.IndexOf("E");
                    string temp = "";
                    for (int j = 0; j < fullString.Length; j++)
                    {
                        if (fullString[j] != '.')
                        {
                            temp += fullString[j];
                        }
                        else
                        {
                            temp += fullString[j];
                            temp += fullString[j + 1];
                            break;
                        }

                    }
                    for (int j = eIndex; j < fullString.Length; j++)
                    {
                        temp += fullString[j];
                    }
                    fullString = temp;
                    Radar3Scales[i] = fullString;

                }
            }

            string[] labels =
            {
            "Co \n"+Radar3Scales[5],
            "Cu \n"+ Radar3Scales[6],
            "Mn \n"+Radar3Scales[7],
            "Ni \n"+Radar3Scales[8],
            "Sr \n"+ Radar3Scales[9],
            "Zn \n"+ Radar3Scales[10],
            "Ba \n"+ Radar3Scales[11],
            "Pb \n"+ Radar3Scales[12],
            "Fe \n"+ Radar3Scales[13],
            "Cd \n"+ Radar3Scales[14],
            "Cr \n"+ Radar3Scales[15],
            "Tl \n"+ Radar3Scales[16],
            "Be \n"+ Radar3Scales[17],
            "Se \n"+ Radar3Scales[18],
            "B \n"+Radar3Scales[19],
            "Li \n"+ Radar3Scales[20],
            "Na \n"+Radar3Scales[0],
            "K \n"+Radar3Scales[1],
            "Ca \n"+Radar3Scales[2],
            "Mg \n"+Radar3Scales[3],
            "Al \n"+Radar3Scales[4]
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
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Sr), Scale = maxSr3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Zn), Scale = maxZn },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ba), Scale = maxBa3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Pb), Scale = maxPb },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Fe), Scale = maxFe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cd), Scale = maxCd },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cr), Scale = maxCr },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Tl), Scale = maxTl },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Be), Scale = maxBe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Se), Scale = maxSe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].B),  Scale = maxB },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Li), Scale = maxLi },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Na), Scale = maxNa3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].K),  Scale = maxK3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ca), Scale = maxCa3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Mg), Scale = maxMg3 },
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
                float labelX = (float)(allX + (0.2 * radius) * Math.Cos(angle));
                float labelY = (float)(allY + (0.2 * radius) * Math.Sin(angle));
                labelY -= 0.02f * frmMainForm.mainChartPlotting.Height;
                labelX -= 0.01f * frmMainForm.mainChartPlotting.Width;
                g.DrawString(label, AxisFont, Brushes.Black, labelX, labelY);
            }
            g.DrawPolygon(axisPen, quarterList);
            g.DrawPolygon(axisPen, allList);
            g.DrawPolygon(axisPen, halfList);
            g.DrawPolygon(axisPen, thirdQuarterList);

            //g.DrawString("ICP Reproducibility", new Font("Times New Roman", fontSize, FontStyle.Bold), Brushes.Black, 0.2f * frmMainForm.mainChartPlotting.Width, 0.9f * frmMainForm.mainChartPlotting.Height);

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
                Pen polygonPen = new Pen(colors[s], 2f);
                polygonPen.Width = frmImportSamples.WaterData[s].lineWidth;
                polygonPen.DashStyle = frmImportSamples.WaterData[s].selectedStyle;

                g.DrawPolygon(polygonPen, points);



            }
            #region Draw Radar legend
            RadarLegend(g, bounds);
            #endregion
            flag = false;
        }
        /// <summary>
        /// Exports the third radar chart to PowerPoint, showing ICP reproducibility data.
        /// Creates a new slide with the chart and legend in PowerPoint format.
        /// </summary>
        /// <param name="bounds">Bounds of the radar chart area</param>
        /// <param name="slide">PowerPoint slide to draw on</param>
        /// <param name="presentation">PowerPoint presentation object</param>
        /// <param name="flag">Flag to control whether to recompute maximum values</param>
        public static void ExportRadar3ToPowerpoint(Rectangle bounds, PowerPoint.Slide slide, PowerPoint.Presentation presentation, bool flag)
        {
            #region Setup
            PowerPoint.Shape title = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                (presentation.PageSetup.SlideWidth / 2) - 100, clsConstants.chartYPowerpoint, 200, 50);
            title.TextFrame.TextRange.Text = "ICP Reproducibility";
            title.TextFrame.TextRange.Font.Size = 25;
            title.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            title.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            title.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            //title.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            title.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
            PrecomputeMaxValues(flag);
            //maxNa = 0;
            //maxK = 0;
            //maxCa = 0;
            //maxMg = 0;
            //foreach (var wdata in frmImportSamples.WaterData)
            //{
            //    maxNa = Math.Max(maxNa, Math.Abs(wdata.Na));
            //    maxK = Math.Max(maxK, Math.Abs(wdata.K));
            //    maxCa = Math.Max(maxCa, Math.Abs(wdata.Ca));
            //    maxMg = Math.Max(maxMg, Math.Abs(wdata.Mg));
            //    maxBa = Math.Max(maxBa, Math.Abs(wdata.Ba));
            //    maxSr = Math.Max(maxSr, Math.Abs(wdata.Sr));
            //}
            
            // Axis labels based on elements
            string[] elements = {
            "Co \n"+Radar3Scales[5],
            "Cu \n"+ Radar3Scales[6],
            "Mn \n"+Radar3Scales[7],
            "Ni \n"+Radar3Scales[8],
            "Sr \n"+ Radar3Scales[9],
            "Zn \n"+ Radar3Scales[10],
            "Ba \n"+ Radar3Scales[11],
            "Pb \n"+ Radar3Scales[12],
            "Fe \n"+ Radar3Scales[13],
            "Cd \n"+ Radar3Scales[14],
            "Cr \n"+ Radar3Scales[15],
            "Tl \n"+ Radar3Scales[16],
            "Be \n"+ Radar3Scales[17],
            "Se \n"+ Radar3Scales[18],
            "B \n"+Radar3Scales[19],
            "Li \n"+ Radar3Scales[20],
            "Na \n"+Radar3Scales[0],
            "K \n"+Radar3Scales[1],
            "Ca \n"+Radar3Scales[2],
            "Mg \n"+Radar3Scales[3],
            "Al \n"+Radar3Scales[4] };

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
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Sr), Scale = maxSr3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Zn), Scale = maxZn },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ba), Scale = maxBa3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Pb), Scale = maxPb },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Fe), Scale = maxFe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cd), Scale = maxCd },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Cr), Scale = maxCr },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Tl), Scale = maxTl },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Be), Scale = maxBe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Se), Scale = maxSe },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].B),  Scale = maxB },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Li), Scale = maxLi },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Na), Scale = maxNa3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].K),  Scale = maxK3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Ca), Scale = maxCa3 },
                    new clsRadarScale { Item = Math.Abs(frmImportSamples.WaterData[i].Mg), Scale = maxMg3 },
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
            float radius = (float)Math.Min(bounds.Width / 1.5, bounds.Height / 1.5) / 3;
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
                float labelX = allX + (float)((radius * 0.2) * Math.Cos(angle)) - labelSize.Width / 2;
                float labelY = allY + (float)((radius * 0.2) * Math.Sin(angle)) - labelSize.Height / 2;
                labelX += 8;

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
            //slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 5, centerY + radius + 70, 1000, 30)
            //    .TextFrame.TextRange.Text = "ICP Reproducibility";
            //slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Size = clsConstants.legendTextSize;
            //slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            //slide.Shapes[slide.Shapes.Count].TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            //slide.Shapes[slide.Shapes.Count].TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignLeft;
            //slide.Shapes[slide.Shapes.Count].TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            //slide.Shapes[slide.Shapes.Count].TextFrame.MarginLeft = 0;
            //slide.Shapes[slide.Shapes.Count].TextFrame.MarginRight = 0;
            //slide.Shapes[slide.Shapes.Count].TextFrame.MarginTop = 0;
            //slide.Shapes[slide.Shapes.Count].TextFrame.MarginBottom = 0;
            //slide.Shapes[slide.Shapes.Count].TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;
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
                for (int j = 0; j < points.Length - 1; j++)
                {

                    PowerPoint.Shape polygon = slide.Shapes.AddPolyline(new float[,] { { points[j].X, points[j].Y }, { points[j + 1].X, points[j + 1].Y } });
                    polygon.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color); // Set line color
                    polygon.Line.Weight = frmImportSamples.WaterData[i].lineWidth; // Set line width
                    polygon.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);
                }

                PowerPoint.Shape polygonLastline = slide.Shapes.AddPolyline(new float[,] { { points[0].X, points[0].Y }, { points[20].X, points[20].Y } });
                polygonLastline.Line.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color); // Set line color
                polygonLastline.Line.Weight = frmImportSamples.WaterData[i].lineWidth; // Set line width
                polygonLastline.Line.DashStyle = ConvertDashStyle(frmImportSamples.WaterData[i].selectedStyle);



            }
            RadarLegendPowerpoint(slide);
        }

        /// <summary>
        /// Precomputes maximum values for all elements and ratios used in the radar charts.
        /// This is used to scale the radar chart axes appropriately.
        /// </summary>
        /// <param name="flag">Flag to control whether to recompute values</param>
        private static void PrecomputeMaxValues(bool flag)
        {
            if (flag) return;
            maxAl = 0; maxCo = 0; maxCu = 0; maxMn = 0; maxNi = 0; maxZn = 0; maxPb = 0; maxFe = 0; maxCd = 0; maxCr = 0; maxTl = 0; maxBe = 0; maxSe = 0; maxLi = 0; maxB = 0;
            maxNaCl = 0; maxClCa = 0; maxHCO3Cl = 0; maxClSr = 0; maxNaCa = 0; maxKNa = 0; maxSrMg = 0; maxMgCl = 0; maxSrCl = 0; maxSrK = 0; maxMgK = 0; maxCaK = 0; maxtK = 0; maxBCl = 0; maxBNa = 0; maxBMg = 0;
            maxCl = 0; maxNa1 = 0; maxK1 = 0; maxCa1 = 0; maxMg1 = 0; maxBa1 = 0; maxSr1 = 0;
            maxNa3 = 0; maxK3 = 0; maxCa3 = 0; maxMg3 = 0; maxBa3 = 0; maxSr3 = 0;
            foreach (var data in frmImportSamples.WaterData)
            {
                maxCl = Math.Max(maxCl, Math.Abs(data.Cl) / Bm);
                
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
                
                // Always update both sets of scales
                maxNa3 = Math.Max(maxNa3, Math.Abs(data.Na));
                maxK3 = Math.Max(maxK3, Math.Abs(data.K));
                maxCa3 = Math.Max(maxCa3, Math.Abs(data.Ca));
                maxMg3 = Math.Max(maxMg3, Math.Abs(data.Mg));
                maxBa3 = Math.Max(maxBa3, Math.Abs(data.Ba));
                maxSr3 = Math.Max(maxSr3, Math.Abs(data.Sr));
                
                maxNa1 = Math.Max(maxNa1, Math.Abs(data.Na) / Bn);
                maxK1 = Math.Max(maxK1, Math.Abs(data.K) / Bo);
                maxCa1 = Math.Max(maxCa1, Math.Abs(data.Ca) / Bp);
                maxMg1 = Math.Max(maxMg1, Math.Abs(data.Mg) / Bq);
                maxBa1 = Math.Max(maxBa1, Math.Abs(data.Ba) / Br);
                maxSr1 = Math.Max(maxSr1, Math.Abs(data.Sr) / Bs);
            }
            

        }

        /// <summary>
        /// Converts a System.Drawing.DashStyle to PowerPoint's MsoLineDashStyle.
        /// Used for maintaining consistent line styles between Windows Forms and PowerPoint.
        /// </summary>
        /// <param name="dashStyle">System.Drawing.DashStyle to convert</param>
        /// <returns>Equivalent PowerPoint MsoLineDashStyle</returns>
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
