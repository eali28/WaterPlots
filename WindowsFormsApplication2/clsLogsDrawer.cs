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
    public class clsLogsDrawer
    {
        public static Rectangle chartBounds = frmMainForm.mainChartPlotting.ClientRectangle;
        public static int margin = (int)(0.02 * chartBounds.Width); // Make margin relative to width
        public static int leftMargin = (int)(0.1 * frmMainForm.mainChartPlotting.Width);
        public static int topMargin = (int)(0.01 * frmMainForm.mainChartPlotting.Height);
        public static void DrawlogNa_VS_logCl(Graphics g,int diagramWidth,int diagramHeight,int x,int y)
        {
            y += topMargin;
            x += leftMargin;
            float labelSize = 12; // Make font size relative
            float titleSize = 25;
            // Set up fonts
            Font labelFont = new Font("Times New Roman", labelSize, FontStyle.Bold);
            Font titleFont = new Font("Times New Roman", titleSize, FontStyle.Bold);


            if (frmMainForm.listBoxCharts.SelectedItem.ToString() == "log Na Vs log Cl")
            {
                g.DrawString("Log Na Vs Log Cl", titleFont, Brushes.Black, diagramWidth / 2, 0.01f * frmMainForm.mainChartPlotting.Height);
            }
            else if (frmMainForm.listBoxCharts.SelectedItem.ToString() == "Major Element Logs")
            {
                g.DrawString("Log Na Vs Log Cl", titleFont, Brushes.Black, x, 0);
            }

            // Define chart area limits for X and Y
            int xAxisMin = 0;
            int xAxisMax = 6;
            int yAxisMin = 0;
            int yAxisMax = 6;

            // X-axis label and grid
            g.DrawLine(Pens.Black, x, y + diagramHeight,x+ diagramWidth , y + diagramHeight); // X-axis line
            g.DrawString("Log Na", labelFont, Brushes.Black, x + diagramWidth / 2, (int)(diagramHeight +6*topMargin+y));

            // Y-axis label and grid
            g.DrawLine(Pens.Black, x,y,x,y+diagramHeight); // Y-axis line
            GraphicsState gstate = g.Save();

            gstate = g.Save();

            g.TranslateTransform((int)(x-0.5f*leftMargin), y+ diagramHeight / 3);

            // Rotate counterclockwise by 90 degrees
            g.RotateTransform(-90);
            g.DrawString("Log Cl", labelFont, Brushes.Black, new PointF(0, 0));
            g.Restore(gstate);

            // Draw grid lines for better readability
            for (int i = xAxisMin; i <= xAxisMax; i++)
            {
                int xPos = (int)((i - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)) + x;
                if (i != xAxisMin)
                {
                    g.DrawLine(Pens.LightGray, xPos, y, xPos,y +diagramHeight); // Vertical grid lines
                    g.DrawString(i.ToString(), labelFont, Brushes.Black, xPos - 10, diagramHeight + y + 10);
                }
                else
                {
                    g.DrawString(i.ToString(), labelFont, Brushes.Black, xPos - 10, diagramHeight + y + 10);
                }
            }

            for (int i = yAxisMin + 1; i <= yAxisMax; i++)
            {
                int yPos = diagramHeight - (int)((i - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight));
                g.DrawLine(Pens.LightGray, x, yPos + y, x+diagramWidth, yPos + y); // Horizontal grid lines
                g.DrawString(i.ToString(), labelFont, Brushes.Black, x - 40, yPos + y - 10);
            }

            // Plot red line (SERT)
            PointF sertStart = new PointF(x +(int)((0.5 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                         y + diagramHeight - (int)((0.5 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            PointF sertEnd = new PointF(x  + (int)((4.1 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                       y + diagramHeight - (int)((4.3 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            g.DrawLine(new Pen(Color.Red, 3), sertStart, sertEnd);

            // Plot blue line (SET)
            PointF setStart = new PointF(x + (int)((4.1 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                        diagramHeight + y - (int)((4.3 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            PointF setEnd = new PointF( x + (int)((4.9 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                      diagramHeight + y - (int)((5.2 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);

            setStart = new PointF(x + (int)((4.9 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                        diagramHeight + y - (int)((5.2 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            setEnd = new PointF(x + (int)((4.0 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                        diagramHeight + y - (int)((5.5 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));

            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);
            setStart = new PointF(x + (int)((4.0 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                        diagramHeight  + y - (int)((5.5 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            setEnd = new PointF( x + (int)((3.7 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                        diagramHeight  + y - (int)((5.3 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);
            setStart = new PointF(x + (int)((3.7 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                            diagramHeight + y - (int)((5.3 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            setEnd = new PointF( x + (int)((3.2 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                        diagramHeight + y - (int)((5.5 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);
            // Plot data points
            foreach (var waterData in frmImportSamples.WaterData)
            {
                double logNa = Math.Log10(waterData.Na);
                double logCl = Math.Log10(waterData.Cl);

                int xPos = (int)((logNa - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)) + x;
                int yPos = diagramHeight + y - (int)((logCl - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight ));

                // Draw a circle at each data point
                g.DrawEllipse(new Pen(Color.Red,1), xPos - 7, yPos - 7, 15, 15);
            }
            float setSertSize = frmMainForm.mainChartPlotting.Height * 0.018f;
            // Add annotations for "SERT" and "SET"
            StringFormat drawFormat = new StringFormat();
            drawFormat.FormatFlags = StringFormatFlags.NoWrap;
            g.DrawString("SERT", new Font("Times New Roman", setSertSize, FontStyle.Bold), Brushes.Red, sertStart.X + (int)(0.1 * (diagramHeight - margin - sertStart.X)), sertStart.Y, drawFormat);
            g.DrawString("SET", new Font("Times New Roman", setSertSize, FontStyle.Bold), Brushes.Blue, setEnd.X - (int)(0.1 * (setEnd.X - margin)), setEnd.Y, drawFormat);
        }


        public static void ExportLogNaVsLogClChartToPowerPoint(PowerPoint.Slide slide, int titleX, int titleY, PowerPoint.Presentation presentation, float slideWidth, float slideHeight, int x, int y)
        {
            // Chart area dimensions
            float chartX = x, chartY = y, diagramWidth = slideWidth, diagramHeight = slideHeight;

            // Add chart title
            PowerPoint.Shape chartTitle = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, titleX, titleY, 200, 50);
            chartTitle.TextFrame.TextRange.Text = "Log Na vs. Log Cl";
            chartTitle.TextFrame.TextRange.Font.Size = 25;
            chartTitle.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            chartTitle.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            chartTitle.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            chartTitle.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Draw grid and axes
            int numGridLinesX = 6, numGridLinesY = 6;
            float xInterval = diagramWidth / numGridLinesX, yInterval = diagramHeight / numGridLinesY;

            // X-axis
            slide.Shapes.AddLine(chartX, chartY + diagramHeight, chartX + diagramWidth, chartY + diagramHeight)
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            // Y-axis
            slide.Shapes.AddLine(chartX, chartY, chartX, chartY + diagramHeight)
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            // Grid lines
            for (int i = 1; i <= numGridLinesX; i++) // Vertical grid lines
            {
                var gridLine = slide.Shapes.AddLine(chartX + i * xInterval, chartY, chartX + i * xInterval, chartY + diagramHeight);
                gridLine.Line.ForeColor.RGB = System.Drawing.Color.Gray.ToArgb();
                gridLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineRoundDot; // Make line dotted

                var lineLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX + (i - 1) * xInterval, chartY + diagramHeight + 5, 100, 30);
                lineLabel.TextFrame.TextRange.Text = i.ToString();
                lineLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                lineLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                lineLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            }

            for (int i = 0; i < numGridLinesY; i++) // Horizontal grid lines
            {
                var gridLine = slide.Shapes.AddLine(chartX, chartY + i * yInterval, chartX + diagramWidth, chartY + i * yInterval);
                gridLine.Line.ForeColor.RGB = System.Drawing.Color.Gray.ToArgb();
                gridLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineRoundDot; // Make line dotted

                var lineLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 60, chartY + (i) * yInterval - 10, 100, 30);
                lineLabel.TextFrame.TextRange.Text = (6 - i).ToString();
                lineLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                lineLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            }
            var textbox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 60, chartY + 6 * yInterval - 10, 100, 30);
            textbox.TextFrame.TextRange.Text = (0).ToString();
            textbox.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            textbox.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;

            // Add axis titles
            var xAxisLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX + diagramWidth / 2 - 50, chartY + diagramHeight + 30, 100, 30);
            xAxisLabel.TextFrame.TextRange.Text = "Log Na";
            xAxisLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            xAxisLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            xAxisLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            var yAxisLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 100, chartY + diagramHeight / 2 - 40, 100, 30);
            yAxisLabel.TextFrame.TextRange.Text = "Log Cl";
            yAxisLabel.Rotation = -90;
            yAxisLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            yAxisLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            yAxisLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Add SERT (red line)
            PowerPoint.Shape sertLine = slide.Shapes.AddLine(
                chartX + (float)(0.5 / 6 * diagramWidth), chartY + diagramHeight - (float)(0.5 / 6 * diagramHeight),
                chartX + (float)(4.1 / 6 * diagramWidth), chartY + diagramHeight - (float)(4.3 / 6 * diagramHeight)
            );
            sertLine.Line.ForeColor.RGB = System.Drawing.Color.Blue.ToArgb();
            sertLine.Line.Weight = 3;

            // Add annotation for SERT (start of the line)
            var sertAnnotation = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                chartX + (float)(0.5 / 6 * diagramWidth) + 10, chartY + diagramHeight - (float)(0.5 / 6 * diagramHeight)-10,
                100,
                30
            );
            sertAnnotation.TextFrame.TextRange.Text = "SERT";
            sertAnnotation.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.Blue.ToArgb();
            sertAnnotation.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            sertAnnotation.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            sertAnnotation.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            sertAnnotation.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Add SET (blue line)
            PowerPoint.Shape setLine = slide.Shapes.AddPolyline(new float[,]
            {
        { chartX + (float)(4.1 / 6 * diagramWidth), chartY + diagramHeight - (float)(4.3 / 6 * diagramHeight) },
        { chartX + (float)(4.9 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.2 / 6 * diagramHeight) },
        { chartX + (float)(4.0 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.5 / 6 * diagramHeight) },
        { chartX + (float)(3.7 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.3 / 6 * diagramHeight) },
        { chartX + (float)(3.2 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.5 / 6 * diagramHeight) }
            });
            setLine.Line.ForeColor.RGB = System.Drawing.Color.Red.ToArgb();
            setLine.Line.Weight = 3;

            // Add annotation for SET (end of the line)
            var setAnnotation = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                chartX + (float)(3.2 / 6 * diagramWidth) - 70,
                chartY + diagramHeight - (float)(5.5 / 6 * diagramHeight) - 15,
                100,
                30
            );
            setAnnotation.TextFrame.TextRange.Text = "SET";
            setAnnotation.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.Red.ToArgb();
            setAnnotation.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            setAnnotation.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            setAnnotation.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            setAnnotation.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Add data points (hollow circles)
            foreach (var data in frmImportSamples.WaterData)
            {
                double logNa = Math.Log10(data.Na);
                double logCl = Math.Log10(data.Cl);

                // Normalize data points to chart area
                float xPos = chartX + (float)((logNa / 6) * diagramWidth);
                float yPos = chartY + diagramHeight - (float)((logCl / 6) * diagramHeight);

                // Draw hollow circle
                PowerPoint.Shape dataPoint = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, xPos - 7.5f, yPos - 7.5f, 15, 15);
                dataPoint.Fill.Transparency = 1.0f; // Hollow circle
                dataPoint.Line.ForeColor.RGB = System.Drawing.Color.Blue.ToArgb();
                dataPoint.Line.Weight = 2;
            }
        }

        public static void DrawlogMg_VS_logCl(Graphics g, int diagramWidth, int diagramHeight, int x, int y)
        {
            x += leftMargin;
            y += topMargin;
            float labelSize = 12; // Make font size relative
            float titleSize = 25;
            // Set up fonts
            Font labelFont = new Font("Times New Roman", labelSize, FontStyle.Bold);
            Font titleFont = new Font("Times New Roman", titleSize, FontStyle.Bold);

            // Draw the title
            if (frmMainForm.listBoxCharts.SelectedItem.ToString() == "log Mg Vs log Cl")
            {
                g.DrawString("Log Mg Vs Log Cl", titleFont, Brushes.Black, diagramWidth / 2, 0.01f * frmMainForm.mainChartPlotting.Height);
            }
            else if (frmMainForm.listBoxCharts.SelectedItem.ToString() == "Major Element Logs")
            {
                g.DrawString("Log Mg Vs Log Cl", titleFont, Brushes.Black, x, 0);
            }
            

            // Define chart area limits for X and Y
            int xAxisMin = 0;
            int xAxisMax = 6;
            int yAxisMin = 0;
            int yAxisMax = 6;

            // X-axis label and grid
            g.DrawLine(Pens.Black, x, diagramHeight +y, diagramWidth +x, diagramHeight +y); // X-axis line
            g.DrawString("Log Mg", labelFont, Brushes.Black, x + diagramWidth / 2, (int)(diagramHeight + 6 * topMargin + y));

            // Y-axis label and grid
            g.DrawLine(Pens.Black, x, y, x, y+diagramHeight); // Y-axis line
            GraphicsState gstate = g.Save();





            gstate = g.Save();

            g.TranslateTransform((int)(x - 0.5f * leftMargin), y + diagramHeight / 3);

            // Rotate counterclockwise by 90 degrees
            g.RotateTransform(-90);
            g.DrawString("Log Cl", labelFont, Brushes.Black, new PointF(0, 0));
            g.Restore(gstate);

            // Draw grid lines for better readability
            for (int i = xAxisMin; i <= xAxisMax; i++)
            {
                int xPos = (int)((i - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )) + x;
                if (i != xAxisMin)
                {
                    g.DrawLine(Pens.LightGray, xPos, y, xPos, diagramHeight + y); // Vertical grid lines
                }
                else
                {
                    g.DrawLine(Pens.Black, xPos, y, xPos, diagramHeight + y); // Vertical grid lines
                }
                g.DrawString(i.ToString(), labelFont, Brushes.Black, xPos - 10, diagramHeight + y + 10);
            }

            for (int i = yAxisMin + 1; i <= yAxisMax; i++)
            {
                int yPos = diagramHeight - (int)((i - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight));
                g.DrawLine(Pens.LightGray, x, yPos + y, x + diagramWidth, yPos + y); // Horizontal grid lines
                g.DrawString(i.ToString(), labelFont, Brushes.Black, x - 40, yPos + y - 10);
            }

            // Plot red line (SERT)
            PointF sertStart = new PointF(x + (int)((0.5 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                         diagramHeight +y - (int)((0.5 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            PointF sertEnd = new PointF(x + (int)((3.1 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                       diagramHeight +y - (int)((4.1 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            g.DrawLine(new Pen(Color.Red, 3), sertStart, sertEnd);

            // Plot blue line (SET)
            PointF setStart = new PointF(x + (int)((3.1 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                        diagramHeight +y - (int)((4.1 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            PointF setEnd = new PointF(x + (int)((4.0 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                      diagramHeight+y - (int)((5.1 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);

            setStart = new PointF(x + (int)((4.0 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                        diagramHeight+y - (int)((5.1 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            setEnd = new PointF(x + (int)((4.6 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                        diagramHeight +y - (int)((5.1 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));

            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);
            setStart = new PointF(x + (int)((4.6 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                        diagramHeight +y - (int)((5.1 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            setEnd = new PointF(x + (int)((4.8 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                        diagramHeight +y - (int)((5.3 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);
            setStart = new PointF(x + (int)((4.8 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                            diagramHeight +y - (int)((5.3 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            setEnd = new PointF(x + (int)((5.0 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                        diagramHeight +y - (int)((5.5 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);
            // Plot data points
            foreach (var waterData in frmImportSamples.WaterData)
            {
                double logMg = Math.Log10(waterData.Mg);
                double logCl = Math.Log10(waterData.Cl);

                int xPos = (int)((logMg - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth ))+x;
                int yPos = diagramHeight +y - (int)((logCl - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight));

                // Draw a circle at each data point
                g.DrawEllipse(new Pen(Color.Red, 1), xPos - 5, yPos - 5, 15, 15);
            }
            float setSertSize = frmMainForm.mainChartPlotting.Height * 0.018f; // Make font size relative
            // Add annotations for "SERT" and "SET"
            StringFormat drawFormat = new StringFormat();
            drawFormat.FormatFlags = StringFormatFlags.NoWrap;
            g.DrawString("SERT", new Font("Times New Roman", setSertSize, FontStyle.Bold), Brushes.Red, sertStart.X, sertStart.Y, drawFormat);
            g.DrawString("SET", new Font("Times New Roman", setSertSize, FontStyle.Bold), Brushes.Blue, setEnd.X - (int)(0.1 * (setEnd.X - margin)), setEnd.Y, drawFormat);
        }

        public static void ExportlogMgVslogCltoPowerpoint(PowerPoint.Slide slide, int titleX,int titleY, PowerPoint.Presentation presentation, float slideWidth, float slideHeight, int x, int y)
        {

            // Chart area dimensions
            float chartX = x, chartY = y, diagramWidth = slideWidth, diagramHeight = slideHeight;

            // Add chart title
            PowerPoint.Shape chartTitle = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, titleX, titleY, 200, 50);
            chartTitle.TextFrame.TextRange.Text = "Log Mg vs. Log Cl";
            chartTitle.TextFrame.TextRange.Font.Size = 25;
            chartTitle.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            chartTitle.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            chartTitle.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            chartTitle.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Draw grid and axes
            int numGridLinesX = 6, numGridLinesY = 6;
            float xInterval = diagramWidth / numGridLinesX, yInterval = diagramHeight / numGridLinesY;

            // X-axis
            slide.Shapes.AddLine(chartX, chartY + diagramHeight, chartX + diagramWidth, chartY + diagramHeight)
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            // Y-axis
            slide.Shapes.AddLine(chartX, chartY, chartX, chartY + diagramHeight)
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            // Grid lines
            for (int i = 1; i <= numGridLinesX; i++) // Vertical grid lines
            {
                var gridLine = slide.Shapes.AddLine(chartX + i * xInterval, chartY, chartX + i * xInterval, chartY + diagramHeight);
                gridLine.Line.ForeColor.RGB = System.Drawing.Color.Gray.ToArgb();
                gridLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineRoundDot; // Make line dotted

                var lineLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX + (i - 1) * xInterval, chartY + diagramHeight + 5, 100, 30);
                lineLabel.TextFrame.TextRange.Text = i.ToString();
                lineLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                lineLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                lineLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            }

            for (int i = 0; i < numGridLinesY; i++) // Horizontal grid lines
            {
                var gridLine = slide.Shapes.AddLine(chartX, chartY + i * yInterval, chartX + diagramWidth, chartY + i * yInterval);
                gridLine.Line.ForeColor.RGB = System.Drawing.Color.Gray.ToArgb();
                gridLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineRoundDot; // Make line dotted

                var lineLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 60, chartY + (i) * yInterval - 10, 100, 30);
                lineLabel.TextFrame.TextRange.Text = (6 - i).ToString();
                lineLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                // Horizontally center the text
                lineLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;

                // Vertically center the text
                lineLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            }

            var textBox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 60, chartY + 6 * yInterval - 10, 100, 30);
            textBox.TextFrame.TextRange.Text = (0).ToString();
            textBox.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            textBox.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            textBox.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            // Add axis titles
            var xAxisLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX + diagramWidth / 2 - 50, chartY + diagramHeight + 30, 100, 30);
            xAxisLabel.TextFrame.TextRange.Text = "Log Mg";
            xAxisLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            xAxisLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            xAxisLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            var yAxisLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 100, chartY + diagramHeight / 2 - 40, 100, 30);
            yAxisLabel.TextFrame.TextRange.Text = "Log Cl";
            yAxisLabel.Rotation = -90;
            yAxisLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            yAxisLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            yAxisLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            // Add SERT (red line)
            PowerPoint.Shape sertLine = slide.Shapes.AddLine(
                chartX + (float)(0.5 / 6 * diagramWidth), chartY + diagramHeight - (float)(0.5 / 6 * diagramHeight),
                chartX + (float)(3.1 / 6 * diagramWidth), chartY + diagramHeight - (float)(4.1 / 6 * diagramHeight)
            );
            sertLine.Line.ForeColor.RGB = System.Drawing.Color.Blue.ToArgb();
            sertLine.Line.Weight = 3;

            // Add annotation for SERT
            var sertAnnotation = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                chartX + (float)(0.5 / 6 * diagramWidth) + 10, chartY + diagramHeight - (float)(0.5 / 6 * diagramHeight)-10,
                100,
                30
            );

            sertAnnotation.TextFrame.TextRange.Text = "SERT";
            sertAnnotation.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.Blue.ToArgb();
            sertAnnotation.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            sertAnnotation.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            sertAnnotation.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            sertAnnotation.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Add SET (blue line)
            PowerPoint.Shape setLine = slide.Shapes.AddPolyline(new float[,]
            {
                { chartX + (float)(3.1 / 6 * diagramWidth), chartY + diagramHeight - (float)(4.1 / 6 * diagramHeight) },
                { chartX + (float)(4.0 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.1 / 6 * diagramHeight) },
                { chartX + (float)(4.6 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.1 / 6 * diagramHeight) },
                { chartX + (float)(4.8 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.3 / 6 * diagramHeight) },
                { chartX + (float)(5.0 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.5 / 6 * diagramHeight) }
            });
            setLine.Line.ForeColor.RGB = System.Drawing.Color.Red.ToArgb();
            setLine.Line.Weight = 3;

            // Add annotation for SET
            var setAnnotation = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                chartX + (float)(3.2 / 6 * diagramWidth) + 50,
                chartY + diagramHeight - (float)(5.5 / 6 * diagramHeight) -20,
                100,
                30
            );
            setAnnotation.TextFrame.TextRange.Text = "SET";
            setAnnotation.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.Red.ToArgb();
            setAnnotation.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            setAnnotation.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            setAnnotation.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            setAnnotation.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            // Add data points (hollow circles)
            foreach (var data in frmImportSamples.WaterData)
            {
                double logMg = Math.Log10(data.Mg);
                double logCl = Math.Log10(data.Cl);

                // Normalize data points to chart area
                float xPos = chartX + (float)((logMg / 6) * diagramWidth);
                float yPos = chartY + diagramHeight - (float)((logCl / 6) * diagramHeight);

                // Draw hollow circle
                PowerPoint.Shape dataPoint = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, xPos - 7.5f, yPos - 7.5f, 15, 15);
                dataPoint.Fill.Transparency = 1.0f; // Hollow circle
                dataPoint.Line.ForeColor.RGB = System.Drawing.Color.Blue.ToArgb();
                dataPoint.Line.Weight = 2;
            }
        }
        public static void DrawlogCa_VS_logCl(Graphics g, int diagramWidth, int diagramHeight, int x, int y)
        {
            y += topMargin;
            x += leftMargin;
            float labelSize = 12; // Make font size relative
            float titleSize = 25;
            // Set up fonts
            Font labelFont = new Font("Times New Roman", labelSize, FontStyle.Bold);
            Font titleFont = new Font("Times New Roman", titleSize, FontStyle.Bold);


            if (frmMainForm.listBoxCharts.SelectedItem.ToString() == "log Ca Vs log Cl")
            {
                g.DrawString("Log Ca Vs Log Cl", titleFont, Brushes.Black, diagramWidth / 2, 0.01f * frmMainForm.mainChartPlotting.Height);
            }
            else if (frmMainForm.listBoxCharts.SelectedItem.ToString() == "Major Element Logs")
            {
                g.DrawString("Log Ca Vs Log Cl", titleFont, Brushes.Black, x, y-6*topMargin);
            }

            // Define chart area limits for X and Y
            int xAxisMin = 0;
            int xAxisMax = 6;
            int yAxisMin = 0;
            int yAxisMax = 6;

            // X-axis label and grid
            g.DrawLine(Pens.Black, x, y + diagramHeight, x + diagramWidth, y + diagramHeight); // X-axis line
            g.DrawString("Log Ca", labelFont, Brushes.Black, x + diagramWidth / 2, (int)(diagramHeight + 6 * topMargin + y));

            // Y-axis label and grid
            g.DrawLine(Pens.Black, x, y, x, y + diagramHeight); // Y-axis line
            GraphicsState gstate = g.Save();

            gstate = g.Save();

            g.TranslateTransform((int)(x - 0.5f * leftMargin), y + diagramHeight / 3);

            // Rotate counterclockwise by 90 degrees
            g.RotateTransform(-90);
            g.DrawString("Log Cl", labelFont, Brushes.Black, new PointF(0, 0));
            g.Restore(gstate);

            // Draw grid lines for better readability
            for (int i = xAxisMin; i <= xAxisMax; i++)
            {
                int xPos = (int)((i - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)) + x;
                if (i != xAxisMin)
                {
                    g.DrawLine(Pens.LightGray, xPos, y, xPos, y + diagramHeight); // Vertical grid lines
                    g.DrawString(i.ToString(), labelFont, Brushes.Black, xPos - 10, diagramHeight + y + 10);
                }
                else
                {
                    g.DrawString(i.ToString(), labelFont, Brushes.Black, xPos - 10, diagramHeight + y + 10);
                }
            }

            for (int i = yAxisMin + 1; i <= yAxisMax; i++)
            {
                int yPos = diagramHeight - (int)((i - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight));
                g.DrawLine(Pens.LightGray, x, yPos + y, x + diagramWidth, yPos + y); // Horizontal grid lines
                g.DrawString(i.ToString(), labelFont, Brushes.Black, x - 40, yPos + y - 10);
            }

            // Plot red line (SERT)
            PointF sertStart = new PointF(x + (int)((0.8 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                         diagramHeight +y - (int)((0.6 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            PointF sertEnd = new PointF(x + (int)((2.5 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                       diagramHeight +y - (int)((4.2 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            g.DrawLine(new Pen(Color.Red, 3), sertStart, sertEnd);

            // Plot blue line (SET)
            PointF setStart = new PointF(x + (int)((2.5 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                        diagramHeight +y - (int)((4.2 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            PointF setEnd = new PointF(x + (int)((3.2 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                      diagramHeight +y - (int)((4.9 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);

            setStart = new PointF(x + (int)((3.2 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                        diagramHeight+y - (int)((4.9 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));
            setEnd = new PointF(x + (int)((2.8 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)),
                                        diagramHeight+y - (int)((5.2 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight)));

            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);
            setStart = new PointF(x + (int)((2.8 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                        diagramHeight +y - (int)((5.2 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            setEnd = new PointF(x + (int)((2.5 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                        diagramHeight +y - (int)((5.1 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);
            setStart = new PointF(x + (int)((2.5 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                            diagramHeight +y - (int)((5.1 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            setEnd = new PointF(x + (int)((2.0 - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth )),
                                        diagramHeight +y - (int)((5.6 - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight )));
            g.DrawLine(new Pen(Color.Blue, 3), setStart, setEnd);
            // Plot data points
            foreach (var waterData in frmImportSamples.WaterData)
            {
                double logCa = Math.Log10(waterData.Ca);
                double logCl = Math.Log10(waterData.Cl);

                int xPos = (int)((logCa - xAxisMin) / (double)(xAxisMax - xAxisMin) * (diagramWidth)) + x;
                int yPos = diagramHeight + y - (int)((logCl - yAxisMin) / (double)(yAxisMax - yAxisMin) * (diagramHeight));

                // Draw a circle at each data point
                g.DrawEllipse(new Pen(Color.Red, 1), xPos - 7, yPos - 7, 15, 15);
            }
            float setSertSize = frmMainForm.mainChartPlotting.Height * 0.018f; // Make font size relative
            // Add annotations for "SERT" and "SET"
            StringFormat drawFormat = new StringFormat();
            drawFormat.FormatFlags = StringFormatFlags.NoWrap;
            g.DrawString("SERT", new Font("Times New Roman", setSertSize, FontStyle.Bold), Brushes.Red, sertStart.X, sertStart.Y, drawFormat);
            g.DrawString("SET", new Font("Times New Roman", setSertSize, FontStyle.Bold), Brushes.Blue, setEnd.X - (int)(0.1 * (setEnd.X - margin)), setEnd.Y, drawFormat);
        }

        public static void ExportlogCaVslogCltoPowerPoint(PowerPoint.Slide slide,int titleX, int titleY, PowerPoint.Presentation presentation, float slideWidth, float slideHeight, int x, int y)
        {
            // Chart area dimensions
            float chartX = x, chartY = y, diagramWidth = slideWidth, diagramHeight = slideHeight;

            // Add chart title
            PowerPoint.Shape chartTitle = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, titleX, titleY, 200, 50);
            chartTitle.TextFrame.TextRange.Text = "Log Ca vs. Log Cl";
            chartTitle.TextFrame.TextRange.Font.Size = 25;
            chartTitle.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            chartTitle.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            chartTitle.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            chartTitle.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            
            // Draw grid and axes
            int numGridLinesX = 6, numGridLinesY = 6;
            float xInterval = diagramWidth / numGridLinesX, yInterval = diagramHeight / numGridLinesY;

            // X-axis
            slide.Shapes.AddLine(chartX, chartY + diagramHeight, chartX + diagramWidth, chartY + diagramHeight)
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            // Y-axis
            slide.Shapes.AddLine(chartX, chartY, chartX, chartY + diagramHeight)
                .Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            // Grid lines
            for (int i = 1; i <= numGridLinesX; i++) // Vertical grid lines
            {
                var gridLine = slide.Shapes.AddLine(chartX + i * xInterval, chartY, chartX + i * xInterval, chartY + diagramHeight);
                gridLine.Line.ForeColor.RGB = System.Drawing.Color.Gray.ToArgb();
                gridLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineRoundDot; // Make line dotted

                var lineLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX + (i - 1) * xInterval, chartY + diagramHeight + 5, 100, 30);
                lineLabel.TextFrame.TextRange.Text = i.ToString();
                lineLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                lineLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
                lineLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            }

            for (int i = 0; i < numGridLinesY; i++) // Horizontal grid lines
            {
                var gridLine = slide.Shapes.AddLine(chartX, chartY + i * yInterval, chartX + diagramWidth, chartY + i * yInterval);
                gridLine.Line.ForeColor.RGB = System.Drawing.Color.Gray.ToArgb();
                gridLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineRoundDot; // Make line dotted

                var lineLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 60, chartY + (i) * yInterval - 10, 100, 30);
                lineLabel.TextFrame.TextRange.Text = (6 - i).ToString();
                lineLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                // Horizontally center the text
                lineLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;

                // Vertically center the text
                lineLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            }
            var textBox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 60, chartY + 6 * yInterval - 10, 100, 30);
            textBox.TextFrame.TextRange.Text = (0).ToString();
            textBox.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            textBox.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            textBox.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            // Add axis titles
            var xAxisLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX + diagramWidth / 2 - 50, chartY + diagramHeight + 30, 100, 30);
            xAxisLabel.TextFrame.TextRange.Text = "Log Ca";
            xAxisLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            xAxisLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            xAxisLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            var yAxisLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, chartX - 100, chartY + diagramHeight / 2 - 40, 100, 30);
            yAxisLabel.TextFrame.TextRange.Text = "Log Cl";
            yAxisLabel.Rotation = -90;
            yAxisLabel.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            yAxisLabel.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            yAxisLabel.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;


            // Add SERT (red line)
            PowerPoint.Shape sertLine = slide.Shapes.AddLine(
                chartX + (float)(0.8 / 6 * diagramWidth), chartY + diagramHeight - (float)(0.6 / 6 * diagramHeight),
                chartX + (float)(2.5 / 6 * diagramWidth), chartY + diagramHeight - (float)(4.2 / 6 * diagramHeight)
            );
            sertLine.Line.ForeColor.RGB = System.Drawing.Color.Blue.ToArgb();
            sertLine.Line.Weight = 3;

            // Add annotation for SERT
            var sertAnnotation = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                chartX + (float)(0.8 / 6 * diagramWidth) + 10, chartY + diagramHeight - (float)(0.6 / 6 * diagramHeight)-10,
                100,
                30
            );

            sertAnnotation.TextFrame.TextRange.Text = "SERT";
            sertAnnotation.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.Blue.ToArgb();
            sertAnnotation.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            sertAnnotation.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            sertAnnotation.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            sertAnnotation.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;

            // Add SET (blue line)
            PowerPoint.Shape setLine = slide.Shapes.AddPolyline(new float[,]
            {
                { chartX + (float)(2.5 / 6 * diagramWidth), chartY + diagramHeight - (float)(4.2 / 6 * diagramHeight) },
                { chartX + (float)(3.2 / 6 * diagramWidth), chartY + diagramHeight - (float)(4.9 / 6 * diagramHeight) },
                { chartX + (float)(2.8 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.2 / 6 * diagramHeight) },
                { chartX + (float)(2.5 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.1 / 6 * diagramHeight) },
                { chartX + (float)(2.0 / 6 * diagramWidth), chartY + diagramHeight - (float)(5.6 / 6 * diagramHeight) }
            });
            setLine.Line.ForeColor.RGB = System.Drawing.Color.Red.ToArgb();
            setLine.Line.Weight = 3;


            // Add annotation for SET
            var setAnnotation = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                chartX + (float)(1.7 / 6 * diagramWidth) + 50,
                chartY + diagramHeight - (float)(5.5 / 6 * diagramHeight) - 15,
                100,
                30
            );
            setAnnotation.TextFrame.TextRange.Text = "SET";
            setAnnotation.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.Red.ToArgb();
            setAnnotation.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            setAnnotation.TextFrame.AutoSize = Microsoft.Office.Interop.PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            setAnnotation.TextFrame.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment.ppAlignCenter;
            setAnnotation.TextFrame2.VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorMiddle;
            // Add data points (hollow circles)
            foreach (var data in frmImportSamples.WaterData)
            {
                double logCa = Math.Log10(data.Ca);
                double logCl = Math.Log10(data.Cl);

                // Normalize data points to chart area
                float xPos = chartX + (float)((logCa / 6) * diagramWidth);
                float yPos = chartY + diagramHeight - (float)((logCl / 6) * diagramHeight);

                // Draw hollow circle
                PowerPoint.Shape dataPoint = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, xPos - 7.5f, yPos - 7.5f, 15, 15);
                dataPoint.Fill.Transparency = 1.0f; // Hollow circle
                dataPoint.Line.ForeColor.RGB = System.Drawing.Color.Blue.ToArgb();
                dataPoint.Line.Weight = 2;
            }
        }
    }
}
