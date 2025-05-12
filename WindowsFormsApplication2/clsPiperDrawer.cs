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
    public class clsPiperDrawer
    {
        public static List<PointF> cationVertices = new List<PointF>();
        public static List<PointF> anionVertices = new List<PointF>();
        public static Rectangle chartBounds = frmMainForm.mainChartPlotting.ClientRectangle;
        public static int margin = (int)(0.02 * chartBounds.Width); // Make margin relative to width

        // Calculate triangle and diamond dimensions within the chart area
        public static int availableWidth = chartBounds.Width - 4 * margin;
        public static int availableHeight = chartBounds.Height - 4 * margin;
        // Store clickable legend items

        public static void DrawPiperDiagram(Graphics g)
        {
            // Detach the event handler if it is attached

            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPiper_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxPie_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxSchoeller_Click;
            frmMainForm.legendPictureBox.MouseDoubleClick -= frmMainForm.pictureBoxCollins_Click;
            //frmMainForm.mainChartPlotting.Invalidate();

            cationVertices.Clear();
            anionVertices.Clear();
            #region Variables

            // Get chart drawing area (ClientRectangle)

            int triangleHeight = (int)(frmMainForm.mainChartPlotting.Height * 0.28);
            int triangleWidth = triangleHeight;

            int diamondHeight = triangleWidth * 2;
            int diamondWidth = triangleWidth;
            int xOrigin = (int)(0.15 * chartBounds.Width); // Make origin relative
            int yOrigin = (int)(0.01 * chartBounds.Height); // Make origin relative
            #endregion
            #region Text and tittle
            // Ensure all elements fit within the chart area
            float fontSize = 25;  // Make font size relative
            
            Font titleFont = new Font("Times New Roman", fontSize, FontStyle.Bold);
            string title = "PIPER DIAGRAM";
            SizeF titleSize = g.MeasureString(title, titleFont);
            g.DrawString(title, titleFont, Brushes.Black, (1200 - titleSize.Width) / 2 + xOrigin, yOrigin);
            Font font = new Font("Times New Roman", 16, FontStyle.Underline);

            // Define the brush and position
            Brush brush = Brushes.Black;
            string text = "Classification of water";
            PointF position = new PointF(50, 50);

            g.DrawString(text, font, brush, position);
            text = "Ca-SO4 waters";
            position = new PointF(50, 80);
            font = new Font("Times New Roman", 16, FontStyle.Bold);

            g.DrawString(text, font, Brushes.Red, position);

            text = " - typical of gypsum";
            position = new PointF(190, 80);
            font = new Font("Times New Roman", 16, FontStyle.Regular);

            g.DrawString(text, font, Brushes.Black, position);
            text = "ground waters and mine drainage.";
            position = new PointF(50, 110);
            font = new Font("Times New Roman", 16, FontStyle.Regular);
            g.DrawString(text, font, Brushes.Black, position);

            text = "Ca-HCO3 waters";
            position = new PointF(50, 140);
            font = new Font("Times New Roman", 16, FontStyle.Bold);
            g.DrawString(text, font, Brushes.Blue, position);
            text = "- typical of shallow, ";
            position = new PointF(210, 140);
            font = new Font("Times New Roman", 16, FontStyle.Regular);
            g.DrawString(text, font, Brushes.Black, position);
            text = "fresh ground waters.";
            position = new PointF(50, 170);
            font = new Font("Times New Roman", 16, FontStyle.Regular);
            g.DrawString(text, font, Brushes.Black, position);
            text = "Na-Cl waters";
            position = new PointF(50, 200);
            font = new Font("Times New Roman", 16, FontStyle.Bold);
            g.DrawString(text, font, Brushes.Green, position);
            text = " - typical of marine and";
            position = new PointF(170, 200);
            font = new Font("Times New Roman", 16, FontStyle.Regular);
            g.DrawString(text, font, Brushes.Black, position);
            text = "deep ancient ground waters.";
            position = new PointF(50, 230);
            font = new Font("Times New Roman", 16, FontStyle.Regular);
            g.DrawString(text, font, Brushes.Black, position);
            text = "Na-HCO3 waters";
            position = new PointF(50, 260);
            font = new Font("Times New Roman", 16, FontStyle.Bold);
            g.DrawString(text, font, Brushes.Black, position);
            text = " - typical of deeper ";
            position = new PointF(210, 260);
            font = new Font("Times New Roman", 16, FontStyle.Regular);
            g.DrawString(text, font, Brushes.Black, position);
            text = "ground waters influenced by ion exchange";
            position = new PointF(50, 290);
            font = new Font("Times New Roman", 16, FontStyle.Regular);
            g.DrawString(text, font, Brushes.Black, position);
            #endregion
            #region Define triangle and diamond bounds
            // Scale the triangle and diamond coordinates and sizes
            Rectangle cationTriangleBounds = new Rectangle(
                (int)(0.32f*frmMainForm.mainChartPlotting.Width),
                (int)(0.5f * frmMainForm.mainChartPlotting.Height),
                triangleWidth,
                triangleHeight);

            Rectangle diamondBounds = new Rectangle(
                (int)(0.4f * frmMainForm.mainChartPlotting.Width),
                (int)(0.4f * frmMainForm.mainChartPlotting.Height)-(int)(0.5*diamondHeight),
                diamondWidth,
                diamondHeight);

            Rectangle anionTriangleBounds = new Rectangle(
                (int)(0.48f * frmMainForm.mainChartPlotting.Width),
                (int)(0.5f * frmMainForm.mainChartPlotting.Height),
                triangleWidth,
                triangleHeight);

            #endregion

            string[] cations = { "Mg", "Ca", "Sodium (Na) + Potassium (K)" };
            string[] anions = { "SO4", "Carbonate (CO3) + Bicarbonate (HCO3)", "CL" };
            // Draw cation triangle
            DrawTriangle(g, cationTriangleBounds, "Cations", Pens.Black, cations);

            // Draw anion triangle
            DrawTriangle(g, anionTriangleBounds, "Anions", Pens.Black, anions);

            // Draw diamond
            DrawDiamond(g, diamondBounds, cationTriangleBounds, anionTriangleBounds);

            #region Draw Legend

            if (frmImportSamples.WaterData.Count > 0)
            {
                int xsample = (int)(0.69f * frmMainForm.mainChartPlotting.Width);
                int legendY = (int)(0.13f * frmMainForm.mainChartPlotting.Height);
                int legendX = xsample;

                int legendBoxHeight = 0;
                int legendtextSize = clsConstants.legendTextSize;
                int legendBoxWidth = (int)(0.2 * frmMainForm.mainChartPlotting.Width); // Set fixed width for wrapping area

                using (Font fontStyle = new Font("Times New Roman", legendtextSize, FontStyle.Bold))
                {
                    foreach (var data in frmImportSamples.WaterData)
                    {
                        string fullText = data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
                        SizeF textSize = g.MeasureString(fullText, fontStyle, legendBoxWidth - 30); // limit width for wrapping
                        legendBoxWidth = (int)Math.Max(legendBoxWidth, textSize.Width);
                        legendBoxHeight += (int)Math.Ceiling(textSize.Height + 10); // add spacing between lines
                    }
                }

                frmMainForm.legendPictureBox.Size = new Size(legendBoxWidth, legendBoxHeight);
                Bitmap bit = new Bitmap(legendBoxWidth, legendBoxHeight);
                g = Graphics.FromImage(bit);
                g.DrawRectangle(new Pen(Color.Blue), legendX - 15.0f, legendY - 10.0f, legendBoxWidth + 15.0f, legendBoxHeight);
                //int ysample = legendY;
                //legendGraphics.Clear(Color.White);  // Fill background
                g.FillRectangle(Brushes.White, 0, 0, legendBoxWidth - 1, legendBoxHeight - 1);
                g.DrawRectangle(new Pen(Color.Blue, 2), 0,0, legendBoxWidth - 1, legendBoxHeight - 1);
                 int ysample = 0;
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    Brush squareBrush = new SolidBrush(frmImportSamples.WaterData[i].color);
                    if (frmImportSamples.WaterData[i].shape != null && frmImportSamples.WaterData[i].shape!="Plus")
                    {
                        for (int j = 0; j < frmSymbolPicker.symbolNames.Count; j++)
                        {
                            if (frmImportSamples.WaterData[i].shape == frmSymbolPicker.symbolNames.ElementAt(j))
                            {
                                frmSymbolPicker.DrawSymbol(g, j, 2, ysample-3, 25, squareBrush);
                                break;
                            }

                        }

                    }
                    else
                    {
                        g.FillRectangle(squareBrush, 8, ysample, 0.005f * frmMainForm.mainChartPlotting.Width, 0.02f * frmMainForm.mainChartPlotting.Height);
                        g.FillRectangle(squareBrush, 5, ysample + 3, 0.02f * frmMainForm.mainChartPlotting.Height, 0.005f * frmMainForm.mainChartPlotting.Width);
                    }

                    // Draw text beside the shape
                    var data = frmImportSamples.WaterData[i];
                    string fullText = data.Well_Name + ", " + data.ClientID + ", " + data.Depth;
                    RectangleF textRect = new RectangleF(30, ysample, legendBoxWidth - 35, legendBoxHeight); // large height to wrap

                    Font fontStyle = new Font("Times New Roman", legendtextSize, FontStyle.Bold);
                    SizeF textSize = g.MeasureString(fullText, fontStyle, (int)textRect.Width);

                    g.DrawString(
                        fullText,
                        fontStyle,
                        Brushes.Black,
                        textRect
                    );

                    ysample += (int)Math.Ceiling(textSize.Height + 10); // Move down based on wrapped height
                }
                
                //Form1.legendPanel.BackColor = Color.Transparent;
                frmMainForm.legendPanel.Location = new Point(legendX - 14, legendY - 9);
                frmMainForm.legendPanel.Size = new System.Drawing.Size(legendBoxWidth, legendBoxHeight);
                frmMainForm.legendPictureBox.Image = bit;
                //Form1.pic.Location = new Point(0, 0);
                //Form1.pic.Visible = true;
                frmMainForm.legendPictureBox.MouseDoubleClick += frmMainForm.pictureBoxPiper_Click;
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
        public static void DrawTriangle(Graphics g, Rectangle bounds, string label, Pen pen, string[] data)
        {
            // Define triangle vertices
            PointF[] vertices = new PointF[]
            {
                new PointF(bounds.Left, bounds.Bottom), // Bottom-left
                new PointF(bounds.Right, bounds.Bottom), // Bottom-right
                new PointF(bounds.Left + bounds.Width / 2, bounds.Top), // Top
            };



            #region Draw triangle outline
            g.DrawPolygon(pen, vertices);
            float fontSize = Math.Min((int)(availableWidth*0.4), (int)(availableHeight * 0.4)) / 40;  // Adjust the divisor for your desired text size scale
            Font font = new Font("Times New Roman", fontSize, FontStyle.Bold);
            // Label vertices
            if (data[1] == "Carbonate (CO3) + Bicarbonate (HCO3)")
            {

                // Save the current graphics state
                GraphicsState gstate = g.Save();

                // Translate to the position of the text
                g.TranslateTransform(vertices[2].X - 0.08f*frmMainForm.mainChartPlotting.Width, vertices[2].Y + 0.2f*frmMainForm.mainChartPlotting.Height);

                // Rotate counterclockwise by 62 degrees
                g.RotateTransform(-64);

                // Draw the rotated text
                g.DrawString(data[1], font, Brushes.Black, new PointF(0, 0));

                // Restore the graphics state
                g.Restore(gstate);
            }
            else
            {
                g.DrawString(data[1], font, Brushes.Black, new PointF(vertices[2].X, vertices[0].Y + 15)); // Bottom-left
            }
            if (data[2] == "CL")
            {
                g.DrawString(data[2], font, Brushes.Black, new PointF(vertices[2].X, vertices[0].Y + 15));
            }
            else
            {
                GraphicsState gstate = g.Save();

                // Translate to the position of the text
                g.TranslateTransform(vertices[2].X + 0.04f*frmMainForm.mainChartPlotting.Width, vertices[2].Y + 50);

                // Rotate clockwise by 62 degrees
                g.RotateTransform(62);

                // Draw the rotated text
                g.DrawString(data[2], font, Brushes.Black, new PointF(0, 0));

                // Restore the graphics state
                g.Restore(gstate);
            }

            if (data[0] == "Mg")
            {
                g.DrawString(data[0], font, Brushes.Black, new PointF(vertices[2].X - 150, vertices[2].Y + 130)); // Top
            }
            else if (data[0] == "SO4")
            {
                g.DrawString(data[0], font, Brushes.Black, new PointF(vertices[2].X + 80, vertices[2].Y + 130)); // Top
            }


            #region Draw grid lines and numbered ranges
            int gridLines = 10; // Number of divisions
            Pen gridPen = new Pen(Color.LightGray, 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash };
            Font rangeFont = new Font("Times New Roman", 8);
            float ii = (float)8;
            float c = (float)10;
            for (int i = 0; i < gridLines; i += 2)
            {
                // Fraction for positioning
                float fraction = i / (float)gridLines;


                // Bottom-left to top
                PointF leftToTop = new PointF(
                    vertices[0].X + fraction * (vertices[2].X - vertices[0].X),
                    vertices[0].Y - fraction * (vertices[0].Y - vertices[2].Y)
                );

                // Bottom-right to top
                PointF rightToTop = new PointF(
                    vertices[1].X - fraction * (vertices[1].X - vertices[2].X),
                    vertices[1].Y - fraction * (vertices[1].Y - vertices[2].Y)
                );

                // Bottom-left to bottom-right
                PointF leftToRight = new PointF(
                    vertices[0].X + fraction * (vertices[1].X - vertices[0].X),
                    vertices[0].Y
                );

                // Draw grid lines
                g.DrawLine(gridPen, leftToTop, leftToRight); // Correct diagonal
                //g.DrawLine(gridPen, rightToTop, leftToRight);
                g.DrawLine(gridPen, leftToTop, rightToTop);


                // Labels for sides
                g.DrawString((i * 10).ToString("0"), rangeFont, Brushes.Black, leftToTop.X - 20, leftToTop.Y - 10);
                g.DrawString(((c) * 10).ToString("0"), rangeFont, Brushes.Black, rightToTop.X + 5, rightToTop.Y - 10);
                g.DrawString(((c) * 10).ToString("0"), rangeFont, Brushes.Black, leftToRight.X - 5, leftToRight.Y + 5);
                c -= 2;
                if (i != 0)
                {
                    ii /= 10;
                    rightToTop = new PointF(
                    vertices[1].X - ii * (vertices[1].X - vertices[2].X),
                    vertices[1].Y - ii * (vertices[1].Y - vertices[2].Y)
                    );

                    g.DrawLine(gridPen, leftToRight, rightToTop);
                    ii *= 10;
                    ii -= 2;
                }


            }


            // Bottom-left to top
            PointF topToLeft = new PointF(
                vertices[0].X + (vertices[2].X - vertices[0].X),
                vertices[0].Y - (vertices[0].Y - vertices[2].Y)
            );

            // Bottom-right to top
            PointF TopToRight = new PointF(
                vertices[1].X - (vertices[1].X - vertices[2].X),
                vertices[1].Y - (vertices[1].Y - vertices[2].Y)
            );

            // Bottom-left to bottom-right
            PointF RightToLeft = new PointF(
                vertices[0].X + (vertices[1].X - vertices[0].X),
                vertices[0].Y
            );
            g.DrawString((100).ToString("0"), rangeFont, Brushes.Black, topToLeft.X - 20, topToLeft.Y - 10);
            g.DrawString((0).ToString("0"), rangeFont, Brushes.Black, TopToRight.X + 5, TopToRight.Y - 10);
            g.DrawString((0).ToString("0"), rangeFont, Brushes.Black, RightToLeft.X - 5, RightToLeft.Y + 5);
            #endregion

            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                Color brush = frmImportSamples.WaterData[i].color;
                if (label == "Cations")
                {

                    PlotPointInTriangle(g, bounds, frmImportSamples.WaterData[i].Mg, frmImportSamples.WaterData[i].Na + frmImportSamples.WaterData[i].K, frmImportSamples.WaterData[i].Ca, brush, label,frmImportSamples.WaterData[i].shape);

                }
                else
                {
                    PlotPointInTriangle(g, bounds, frmImportSamples.WaterData[i].So4, frmImportSamples.WaterData[i].Cl, frmImportSamples.WaterData[i].HCO3 + frmImportSamples.WaterData[i].CO3, brush, label, frmImportSamples.WaterData[i].shape);
                }
                
            }

            float size = Math.Min((int)(availableWidth * 0.4), (int)(availableHeight * 0.4)) / 40;  // Adjust the divisor for your desired text size scale
            Font style = new Font("Times New Roman", size, FontStyle.Italic);
            float labelSize = Math.Min((int)(availableWidth * 0.7), (int)(availableHeight * 0.7)) / 40;  // Adjust the divisor for your desired text size scale
            Font labelStyle = new Font("Times New Roman", labelSize, FontStyle.Bold);
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Near;      // Horizontal center
            //format.LineAlignment = StringAlignment.Center;  // Vertical center (optional)
            if (label == "Cations")
            {
                PointF[] magnesiumVertices = new PointF[]
                {
                    new PointF(bounds.Left + bounds.Width / 2, bounds.Top), // Top
                    new PointF((vertices[0].X+vertices[2].X)/2,(vertices[0].Y+vertices[2].Y)/2), //bottom left
                    new PointF((vertices[1].X+vertices[2].X)/2,(vertices[1].Y+vertices[2].Y)/2)//bottom right
                };
                g.FillPolygon(new SolidBrush(Color.FromArgb(100, Color.Green)), magnesiumVertices);
                PointF[] calciumVertices = new PointF[]
                {
                    new PointF((vertices[0].X+vertices[2].X)/2,(vertices[0].Y+vertices[2].Y)/2), // Top
                    new PointF(bounds.Left, bounds.Bottom),//bottom left
                    new PointF((vertices[0].X+vertices[1].X)/2,(vertices[0].Y+vertices[1].Y)/2)//bottom right
                };
                g.FillPolygon(new SolidBrush(Color.FromArgb(100, Color.Gray)), calciumVertices);
                PointF[] sodiumPotassiumVertices = new PointF[]
                {
                    new PointF((vertices[1].X+vertices[2].X)/2,(vertices[1].Y+vertices[2].Y)/2),//top
                    new PointF(bounds.Right, bounds.Bottom),//bottom right
                    new PointF((vertices[0].X+vertices[1].X)/2,(vertices[0].Y+vertices[1].Y)/2)//bottom left
                };
                g.FillPolygon(new SolidBrush(Color.FromArgb(100, Color.Cyan)), sodiumPotassiumVertices);
                g.DrawString("Magnesium", style, Brushes.Black, magnesiumVertices[1].X + 40, magnesiumVertices[1].Y - 50,format);
                g.DrawString("Calcium\ntype", style, Brushes.Black, calciumVertices[1].X + 40, calciumVertices[1].Y - 70,format);
                g.DrawString("Sodium\nand\nPotassium", style, Brushes.Black, sodiumPotassiumVertices[2].X + 40, sodiumPotassiumVertices[2].Y - 70,format);
                g.DrawString("No\ndominant\ntype", style, Brushes.Black, calciumVertices[0].X + 40, calciumVertices[0].Y+20,format);

            }
            else if (label == "Anions")
            {
                PointF[] sulphate = new PointF[]
                {
                    new PointF(bounds.Left + bounds.Width / 2, bounds.Top), // Top
                    new PointF((vertices[0].X+vertices[2].X)/2,(vertices[0].Y+vertices[2].Y)/2), //bottom left
                    new PointF((vertices[1].X+vertices[2].X)/2,(vertices[1].Y+vertices[2].Y)/2)//bottom right
                };
                g.FillPolygon(new SolidBrush(Color.FromArgb(100, Color.Pink)), sulphate);
                PointF[] Bicarbonate = new PointF[]
                {
                    new PointF((vertices[0].X+vertices[2].X)/2,(vertices[0].Y+vertices[2].Y)/2), // Top
                    new PointF(bounds.Left, bounds.Bottom),//bottom left
                    new PointF((vertices[0].X+vertices[1].X)/2,(vertices[0].Y+vertices[1].Y)/2)//bottom right
                };
                g.FillPolygon(new SolidBrush(Color.FromArgb(100, Color.Magenta)), Bicarbonate);
                PointF[] chloride = new PointF[]
                {
                    new PointF((vertices[1].X+vertices[2].X)/2,(vertices[1].Y+vertices[2].Y)/2),//top
                    new PointF(bounds.Right, bounds.Bottom),//bottom right
                    new PointF((vertices[0].X+vertices[1].X)/2,(vertices[0].Y+vertices[1].Y)/2)//bottom left
                };

                g.FillPolygon(new SolidBrush(Color.FromArgb(100, Color.DarkOrange)), chloride);
                g.DrawString("Sulphate\ntype", style, Brushes.Black, sulphate[1].X + 40, sulphate[1].Y - 70,format);
                g.DrawString("Bicarbonate\ntype", style, Brushes.Black, Bicarbonate[1].X + 20, Bicarbonate[1].Y - 50,format);
                g.DrawString("Chloride\ntype", style, Brushes.Black, chloride[2].X + 40, chloride[2].Y - 70,format);
                g.DrawString("No\ndominant\ntype", style, Brushes.Black, Bicarbonate[0].X + 40, Bicarbonate[0].Y + 20,format);
            }
            g.DrawString(label, labelStyle, Brushes.Black, vertices[2].X - 20, vertices[0].Y + 30);
            #endregion
        }
        public static void DrawDiamond(Graphics g, Rectangle bounds, Rectangle cationTriangleBounds, Rectangle anionTriangleBounds)
        {
            float fontSize = Math.Min(0.4f*availableWidth,0.4f*availableHeight)/40;  // Adjust the divisor for your desired text size scale
            Font rangeFont = new Font("Times New Roman", fontSize);
            float size = Math.Min((int)(availableWidth * 0.6), (int)(availableHeight * 0.6)) / 40;  // Adjust the divisor for your desired text size scale
            Font style = new Font("Times New Roman", size, FontStyle.Italic);
            // Define diamond vertices

            PointF[] diamondVertices = new PointF[]
            {
            new PointF(bounds.Left + bounds.Width / 2, bounds.Top), // Top
            new PointF(bounds.Right, bounds.Top + bounds.Height / 2), // Right
            new PointF(bounds.Left + bounds.Width / 2, bounds.Bottom), // Bottom
            new PointF(bounds.Left, bounds.Top + bounds.Height / 2), // Left
            };
            float cationDiamondGap = (diamondVertices[0].X - cationTriangleBounds.Right) / 2;
            float anionDiamondGap = (anionTriangleBounds.Left - diamondVertices[0].X) / 2;
            // Draw diamond
            g.DrawPolygon(Pens.Black, diamondVertices);

            GraphicsState gstate = g.Save();

            // Translate to the position of the text
            g.TranslateTransform(diamondVertices[0].X - 0.08f*frmMainForm.mainChartPlotting.Width, diamondVertices[0].Y + 0.19f*frmMainForm.mainChartPlotting.Height);

            // Rotate counterclockwise by 90 degrees
            g.RotateTransform(-62);

            // Draw the rotated text
            g.DrawString("Sulphate (So4) + Chloride (Cl)", new Font("Times New Roman", fontSize, FontStyle.Bold), Brushes.Black, new PointF(0, 0));

            g.Restore(gstate);

            gstate = g.Save();

            g.TranslateTransform(diamondVertices[0].X + 0.04f * frmMainForm.mainChartPlotting.Width, diamondVertices[0].Y + 0.07f * frmMainForm.mainChartPlotting.Height);

            // Rotate counterclockwise by 90 degrees
            g.RotateTransform(62);

            // Draw the rotated text
            g.DrawString("Calcium (Ca) + Magnesium (Mg)", new Font("Times New Roman", fontSize, FontStyle.Bold), Brushes.Black, new PointF(0, 0));

            // Restore the graphics state
            g.Restore(gstate);

            int gridLines = 10; // Number of divisions
            Pen gridPen = new Pen(Color.LightGray, 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash };

            for (int i = 0; i <= gridLines; i += 2)
            {
                float fraction = i / (float)gridLines;

                // Interpolate points along the edges
                PointF topToRight = new PointF(
                    diamondVertices[1].X + fraction * (diamondVertices[0].X - diamondVertices[1].X),
                    diamondVertices[1].Y + fraction * (diamondVertices[0].Y - diamondVertices[1].Y)
                );

                PointF rightToBottom = new PointF(
                    diamondVertices[1].X + fraction * (diamondVertices[2].X - diamondVertices[1].X),
                    diamondVertices[1].Y + fraction * (diamondVertices[2].Y - diamondVertices[1].Y)
                );

                PointF bottomToLeft = new PointF(
                    diamondVertices[2].X + fraction * (diamondVertices[3].X - diamondVertices[2].X),
                    diamondVertices[2].Y + fraction * (diamondVertices[3].Y - diamondVertices[2].Y)
                );

                PointF leftToTop = new PointF(
                    diamondVertices[0].X + fraction * (diamondVertices[3].X - diamondVertices[0].X),
                    diamondVertices[0].Y + fraction * (diamondVertices[3].Y - diamondVertices[0].Y)
                );

                // Draw diagonals
                g.DrawLine(gridPen, topToRight, bottomToLeft); // Diagonal from top-right to bottom-left
                g.DrawLine(gridPen, rightToBottom, leftToTop); // Diagonal from right-bottom to left-top

                // Draw labels on the edges
                if (i > 0 && i < gridLines)
                {
                    g.DrawString((i * 10).ToString(), rangeFont, Brushes.Black, topToRight.X - 5, topToRight.Y - 15); // Top-to-right edge
                    g.DrawString((i * 10).ToString(), rangeFont, Brushes.Black, leftToTop.X - 15, leftToTop.Y - 10);
                }
            }
            Brush[] regionBrushes = new Brush[]
            {
                new SolidBrush(Color.FromArgb(100, Color.Yellow)), // Top region
                new SolidBrush(Color.FromArgb(100, Color.LightBlue)),  // Bottom region
                new SolidBrush(Color.FromArgb(100, Color.MediumPurple)),   // Left region
                new SolidBrush(Color.FromArgb(100, Color.Magenta)),   // Right region
                new SolidBrush(Color.FromArgb(100, Color.Gray)),    // Center region
            };

            // Draw colored regions
            FillRegionForDiamond(g, regionBrushes[0], diamondVertices[0], diamondVertices[1], diamondVertices[2], diamondVertices[3], RegionType.Top);
            FillRegionForDiamond(g, regionBrushes[1], diamondVertices[0], diamondVertices[1], diamondVertices[2], diamondVertices[3], RegionType.Bottom);
            FillRegionForDiamond(g, regionBrushes[2], diamondVertices[0], diamondVertices[1], diamondVertices[2], diamondVertices[3], RegionType.Left);
            FillRegionForDiamond(g, regionBrushes[3], diamondVertices[0], diamondVertices[1], diamondVertices[2], diamondVertices[3], RegionType.Right);
            FillRegionForDiamond(g, regionBrushes[4], diamondVertices[0], diamondVertices[1], diamondVertices[2], diamondVertices[3], RegionType.Center);
            PointF[] calciumChlorideVertices = new PointF[]
                {
                    new PointF(bounds.Left + bounds.Width / 2, bounds.Top), // Top
                    new PointF((diamondVertices[0].X+diamondVertices[2].X)/2,(diamondVertices[0].Y+diamondVertices[2].Y)/2), //bottom left
                    new PointF((diamondVertices[1].X+diamondVertices[2].X)/2,(diamondVertices[1].Y+diamondVertices[2].Y)/2)//bottom right
                };
            // Label the edges with percentages
            g.DrawString("100", rangeFont, Brushes.Black, diamondVertices[0].X - 20, diamondVertices[0].Y - 15); // Top
            g.DrawString("0", rangeFont, Brushes.Black, diamondVertices[1].X + 5, diamondVertices[1].Y - 10);    // Right
            g.DrawString("0", rangeFont, Brushes.Black, diamondVertices[2].X - 10, diamondVertices[2].Y + 5);   // Bottom
            g.DrawString("100", rangeFont, Brushes.Black, diamondVertices[3].X - 30, diamondVertices[3].Y - 10); // Left

            
            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                Color brush = frmImportSamples.WaterData[i].color;
                PointF diamondCenter = new PointF((diamondVertices[1].X + diamondVertices[3].X) / 2, (diamondVertices[0].Y + diamondVertices[2].Y) / 2);
                FindIntersection(g,bounds, frmImportSamples.WaterData[i].Na + frmImportSamples.WaterData[i].K, frmImportSamples.WaterData[i].Ca, frmImportSamples.WaterData[i].Mg, frmImportSamples.WaterData[i].Cl + frmImportSamples.WaterData[i].So4, frmImportSamples.WaterData[i].HCO3, frmImportSamples.WaterData[i].CO3, frmImportSamples.WaterData[i].color, frmImportSamples.WaterData[i].shape);
            }
        }
        public static void FillRegionForDiamond(Graphics g, Brush brush, PointF top, PointF right, PointF bottom, PointF left, RegionType region)
        {
            PointF[] points;
            PointF toptemp = top;
            PointF righttemp = right;
            PointF lefttemp = left;
            PointF bottomtemp = bottom;
            float size = Math.Min((int)(availableWidth * 0.4), (int)(availableHeight * 0.4)) / 40;  // Adjust the divisor for your desired text size scale
            Font style = new Font("Times New Roman", size, FontStyle.Italic);
            switch (region)
            {

                case RegionType.Top:

                    righttemp.X = (top.X + right.X) / 2;
                    righttemp.Y = (top.Y + right.Y) / 2;
                    lefttemp.X = (top.X + left.X) / 2;
                    lefttemp.Y = (top.Y + left.Y) / 2;
                    points = new PointF[] { toptemp, righttemp, lefttemp };
                    g.FillPolygon(brush, points);
                    g.DrawString("Calcium\nChloride\ntype", style, Brushes.Black, ((lefttemp.X+righttemp.X)/2)-(int)(0.2*(righttemp.X-lefttemp.X)), (top.Y+righttemp.Y)/2);
                    break;
                case RegionType.Bottom:
                    righttemp.X = (bottom.X + right.X) / 2;
                    righttemp.Y = (bottom.Y + right.Y) / 2;
                    lefttemp.X = (bottom.X + left.X) / 2;
                    lefttemp.Y = (bottom.Y + left.Y) / 2;
                    points = new PointF[] { righttemp, lefttemp, bottomtemp };
                    g.FillPolygon(brush, points);
                    g.DrawString("Sodium\nBicarbonate\ntype", style, Brushes.Black, ((lefttemp.X + righttemp.X) / 2) - (int)(0.2 * (righttemp.X - lefttemp.X)), (bottomtemp.Y + righttemp.Y) / 2 - (int)(0.3 * (bottomtemp.Y - lefttemp.Y)));
                    break;
                case RegionType.Left:
                    toptemp.X = (top.X + left.X) / 2;
                    toptemp.Y = (top.Y + left.Y) / 2;
                    bottomtemp.X = (bottom.X + left.X) / 2;
                    bottomtemp.Y = (bottom.Y + left.Y) / 2;
                    righttemp.X = (right.X + left.X) / 2;
                    righttemp.Y = (right.Y + left.Y) / 2;
                    points = new PointF[] { toptemp, righttemp, bottomtemp, lefttemp };
                    g.FillPolygon(brush, points);
                    g.DrawString("Magnesium\nBicarbonate\ntype", style, Brushes.Black, (lefttemp.X + righttemp.X) / 2 - (int)(0.2 * (righttemp.X - lefttemp.X)), (bottomtemp.Y + toptemp.Y) / 2);
                    break;
                case RegionType.Right:
                    toptemp.X = (top.X + right.X) / 2;
                    toptemp.Y = (top.Y + right.Y) / 2;
                    bottomtemp.X = (bottom.X + right.X) / 2;
                    bottomtemp.Y = (bottom.Y + right.Y) / 2;
                    lefttemp.X = (right.X + left.X) / 2;
                    lefttemp.Y = (right.Y + left.Y) / 2;
                    points = new PointF[] { toptemp, lefttemp, bottomtemp, righttemp };
                    g.FillPolygon(brush, points);
                    g.DrawString("Calcium\nChloride\ntype", style, Brushes.Black, (lefttemp.X + righttemp.X) / 2 - (int)(0.2 * (righttemp.X - lefttemp.X)), (bottomtemp.Y + toptemp.Y) / 2);
                    break;
                case RegionType.Center:
                    //top region
                    righttemp.X = (top.X + right.X) / 2;
                    righttemp.Y = (top.Y + right.Y) / 2;
                    lefttemp.X = (top.X + left.X) / 2;
                    lefttemp.Y = (top.Y + left.Y) / 2;
                    toptemp.X = (right.X + left.X) / 2;
                    toptemp.Y = (right.Y + left.Y) / 2;
                    points = new PointF[] { toptemp, righttemp, lefttemp };
                    g.FillPolygon(new SolidBrush(Color.FromArgb(100, Color.LightGreen)), points);
                    g.DrawString("Mixed\ntype", style, Brushes.Black, (lefttemp.X + righttemp.X) / 2 - (int)(0.2 * (righttemp.X - lefttemp.X)), (toptemp.Y + righttemp.Y) / 2 - (int)(0.2 * (toptemp.Y - righttemp.Y)));
                    //bottom region
                    righttemp.X = (right.X + bottom.X) / 2;
                    righttemp.Y = (right.Y + bottom.Y) / 2;
                    lefttemp.X = (left.X + bottom.X) / 2;
                    lefttemp.Y = (left.Y + bottom.Y) / 2;
                    toptemp.X = (right.X + left.X) / 2;
                    toptemp.Y = (right.Y + left.Y) / 2;
                    points = new PointF[] { toptemp, righttemp, lefttemp };
                    g.FillPolygon(new SolidBrush(Color.FromArgb(100, Color.DarkOrange)), points);
                    g.DrawString("Mixed\ntype", style, Brushes.Black, (lefttemp.X + righttemp.X) / 2 - (int)(0.2 * (righttemp.X - lefttemp.X)), (toptemp.Y + righttemp.Y) / 2);
                    break;
                default:
                    points = new PointF[] { top, right, bottom, left };
                    break;
            }

        }

        // Enum to Define Region Types
        public enum RegionType
        {
            Top,
            Bottom,
            Left,
            Right,
            Center
        }

        public static void PlotPointInTriangle(Graphics g, Rectangle bounds, double A, double B, double C, Color brush, string label,string shape)
        {
            // Step 1: Normalize the values
            double total = A + B + C;
            //if (total == 0)
            //{
            //    MessageBox.Show("The total of the items is Zero", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}// Prevent division by zero
            double aNormalized = A / total;
            double bNormalized = B / total;
            double cNormalized = C / total;

            // Step 2: Calculate the triangle vertices
            PointF bottomLeft = new PointF(bounds.Left, bounds.Bottom);
            PointF bottomRight = new PointF(bounds.Right, bounds.Bottom);
            PointF top = new PointF(bounds.Left + bounds.Width / 2, bounds.Top);

            // Step 3: Interpolate position within the triangle
            float x = (float)(
                cNormalized * bottomLeft.X +
                bNormalized * bottomRight.X +
                aNormalized * top.X
            );

            float y = (float)(
                cNormalized * bottomLeft.Y +
                bNormalized * bottomRight.Y +
                aNormalized * top.Y
            );
            if (label == "Cations")
            {
                cationVertices.Add(new PointF(x, y));
            }
            else
            {
                anionVertices.Add(new PointF(x, y));
            }
            // Step 4: Draw the point
            Brush squareBrush = new SolidBrush(brush);
            if (shape != null && shape!="Plus")
            {
                for (int i = 0; i < frmSymbolPicker.symbolNames.Count; i++)
                {
                    if (shape == frmSymbolPicker.symbolNames.ElementAt(i))
                    {
                        frmSymbolPicker.DrawSymbol(g, i, (int)x-12, (int)y-12, 25, squareBrush);
                        break;
                    }
                }
                
                
            }
            else 
            {
                g.FillRectangle(squareBrush, x - (0.005f * frmMainForm.mainChartPlotting.Width)/2, y - (0.02f * frmMainForm.mainChartPlotting.Height) / 2, 0.005f * frmMainForm.mainChartPlotting.Width, 0.02f * frmMainForm.mainChartPlotting.Height);
                g.FillRectangle(squareBrush, x - (0.02f * frmMainForm.mainChartPlotting.Height) / 2, y - (0.005f * frmMainForm.mainChartPlotting.Width)/2, 0.02f * frmMainForm.mainChartPlotting.Height, 0.005f * frmMainForm.mainChartPlotting.Width);
            }
            
        }

        public static void FindIntersection(Graphics g,Rectangle bounds, double NaK,double Ca,double Mg,double ClSo4,double HCO3,double CO3,Color brush,string shape)
        {
            PointF[] diamondVertices = new PointF[]
            {
            new PointF(bounds.Left + bounds.Width / 2, bounds.Top), // Top
            new PointF(bounds.Right, bounds.Top + bounds.Height / 2), // Right
            new PointF(bounds.Left + bounds.Width / 2, bounds.Bottom), // Bottom
            new PointF(bounds.Left, bounds.Top + bounds.Height / 2), // Left
            };
            float Xc = (float)(NaK / (Ca + Mg + NaK));
            float Ya = (float)(ClSo4 / (HCO3 + CO3 + ClSo4));
            // Calculate intersection point
            // Bilinear interpolation
            float x = (1 - Xc) * ((1 - Ya) * diamondVertices[3].X + Ya * diamondVertices[0].X) + Xc * ((1 - Ya) * diamondVertices[2].X + Ya * diamondVertices[1].X);
            float y = (1 - Xc) * ((1 - Ya) * diamondVertices[3].Y + Ya * diamondVertices[0].Y) + Xc * ((1 - Ya) * diamondVertices[2].Y + Ya * diamondVertices[1].Y);

            // Step 4: Draw the point
            Brush squareBrush = new SolidBrush(brush);
            if (shape != null && shape != "Plus")
            {
                for (int i = 0; i < frmSymbolPicker.symbolNames.Count; i++)
                {
                    if (shape == frmSymbolPicker.symbolNames.ElementAt(i))
                    {
                        frmSymbolPicker.DrawSymbol(g, i, (int)x - 12, (int)y - 12, 25, squareBrush);
                        break;
                    }
                }
                

            }
            else
            {
                g.FillRectangle(squareBrush, x - (0.005f * frmMainForm.mainChartPlotting.Width) / 2, y - (0.02f * frmMainForm.mainChartPlotting.Height) / 2, 0.005f * frmMainForm.mainChartPlotting.Width, 0.02f * frmMainForm.mainChartPlotting.Height);
                g.FillRectangle(squareBrush, x - (0.02f * frmMainForm.mainChartPlotting.Height) / 2, y - (0.005f * frmMainForm.mainChartPlotting.Width) / 2, 0.02f * frmMainForm.mainChartPlotting.Height, 0.005f * frmMainForm.mainChartPlotting.Width);
            }
        }
        public static void ExportPiperDiagramToPowerpoint(PowerPoint.Slide slide, PowerPoint.Presentation presentation)
        {

            PowerPoint.Application pptApplication = new PowerPoint.Application();

            #region Variables

            // Get chart drawing area (ClientRectangle)
            float slideWidth = presentation.PageSetup.SlideWidth;
            float slideHeight = presentation.PageSetup.SlideHeight;
            int marginPowerPoint = (int)(0.02 * slideWidth);
            int availableWidthPowerPoint = (int)slideWidth - 4 * marginPowerPoint;
            int availableHeightPowerPoint = (int)slideHeight - 4 * marginPowerPoint;
            // Set chartBounds equal to the slide bounds
            Rectangle chartBounds = new Rectangle(0, 0, (int)slideWidth, (int)slideHeight);

            int triangleHeight = availableHeightPowerPoint / 2 - 100;
            int triangleWidth = triangleHeight;

            int diamondHeight = triangleWidth * 2;
            int diamondWidth = triangleWidth;
            int yOrigin = 100; // Starting position of the chart
            #endregion
            #region Text and tittle
            triangleHeight = Math.Min(triangleHeight, availableHeight / 2);
            Font titleFont = new Font("Times New Roman", 25, FontStyle.Italic);
            string title = "PIPER DIAGRAM";
            SizeF titleSize = TextRenderer.MeasureText(title, titleFont);
            var titletextbox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                presentation.PageSetup.SlideWidth / 2 - 300, yOrigin - 100, 600, 50);
            titletextbox.TextFrame.TextRange.Text = title;
            titletextbox.TextFrame.TextRange.Font.Size = 40;
            titletextbox.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            titletextbox.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;


            var text = slide.Shapes.AddTextbox(
                                        Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                        0, 30, 400, 100 // Adjust the position and size as needed
                                    );

            //text.Select();
            text.TextFrame.TextRange.Text = "Classification of water\n" +
                                   "Ca-SO₄ waters - typical of gypsum ground waters and mine drainage.\n" +
                                   "Ca-HCO₃ waters - typical of shallow, fresh ground waters.\n" +
                                   "Na-Cl waters - typical of marine and\n" +
                                   "deep ancient ground waters.\n" +
                                   "Na-HCO₃ waters - typical of deeper ground waters influenced by ion exchange.\n";

            // Increase the general font size
            text.TextFrame.TextRange.Font.Size = 15; // Adjust the size as needed
            FormatKeyword(text.TextFrame.TextRange, "Classification of water", "Black");
            FormatKeyword(text.TextFrame.TextRange, "Ca-SO₄ waters", "Red");
            FormatKeyword(text.TextFrame.TextRange, "Ca-HCO₃ waters", "Blue");
            FormatKeyword(text.TextFrame.TextRange, "Na-Cl waters", "Green");
            FormatKeyword(text.TextFrame.TextRange, "Na-HCO₃ waters", "Black");



            #endregion
            #region Define triangle and diamond bounds
            Rectangle cationTriangleBounds = new Rectangle(
                chartBounds.Left + (int)slideWidth / 2 - (int)(0.2 * (int)slideWidth) - (int)(diamondWidth / 1.7),
                chartBounds.Top + availableHeightPowerPoint - triangleHeight - 10,
                triangleWidth,
                triangleHeight);


            Rectangle diamondBounds = new Rectangle(
                chartBounds.Left + (int)slideWidth / 2 - (int)(0.2 * (int)slideWidth),
                chartBounds.Top + 170,
                diamondWidth,
                diamondHeight);
            Rectangle anionTriangleBounds = new Rectangle(
                chartBounds.Left + (int)slideWidth / 2 - (int)(0.2 * (int)slideWidth) + (int)(diamondWidth / 1.7),
                chartBounds.Top + availableHeightPowerPoint - triangleHeight - 10,
                triangleWidth,
                triangleHeight);
            #endregion

            PointF[] cationvertices = new PointF[]
            {
                new PointF(cationTriangleBounds.Left, cationTriangleBounds.Bottom), // Bottom-left
                new PointF(cationTriangleBounds.Right, cationTriangleBounds.Bottom), // Bottom-right
                new PointF(cationTriangleBounds.Left + cationTriangleBounds.Width / 2, cationTriangleBounds.Top), // Top
            };
            PointF[] anionvertices = new PointF[]
            {
                new PointF(anionTriangleBounds.Left, anionTriangleBounds.Bottom), // Bottom-left
                new PointF(anionTriangleBounds.Right, anionTriangleBounds.Bottom), // Bottom-right
                new PointF(anionTriangleBounds.Left + anionTriangleBounds.Width / 2, anionTriangleBounds.Top), // Top
            };


            string[] cations = { "Mg", "Ca", "Na+k" };
            string[] anions = { "SO4", "HCO3+CO3", "CL" };
            cationVertices.Clear();
            anionVertices.Clear();
            #region Text Positioning for Cations and Anions Triangles
            // Cations Labels (Magnesium, Calcium, Sodium/Potassium) - Place near triangle edges
            // Adjusting positions as in the image:
            var magnesiumLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, cationTriangleBounds.Left, cationTriangleBounds.Top, 100, 30);
            magnesiumLabel.TextFrame.TextRange.Text = "Magnesium";
            magnesiumLabel.TextFrame.TextRange.Font.Size = 15;

            var calciumLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, cationTriangleBounds.Right - 100, cationTriangleBounds.Top, 100, 30);
            calciumLabel.TextFrame.TextRange.Text = "Calcium";
            calciumLabel.TextFrame.TextRange.Font.Size = 15;

            var sodiumPotassiumLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, cationTriangleBounds.Left + 30, cationTriangleBounds.Bottom - 50, 150, 30);
            sodiumPotassiumLabel.TextFrame.TextRange.Text = "Sodium/Potassium";
            sodiumPotassiumLabel.TextFrame.TextRange.Font.Size = 15;

            // Anions Labels (Sulfate, Bicarbonate, Chloride) - Place near triangle edges
            var sulfateLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, anionTriangleBounds.Left, anionTriangleBounds.Top, 100, 30);
            sulfateLabel.TextFrame.TextRange.Text = "Sulfate";
            sulfateLabel.TextFrame.TextRange.Font.Size = 15;

            var bicarbonateLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, anionTriangleBounds.Left + 30, anionTriangleBounds.Bottom - 50, 150, 30);
            bicarbonateLabel.TextFrame.TextRange.Text = "Bicarbonate";
            bicarbonateLabel.TextFrame.TextRange.Font.Size = 15;

            var chlorideLabel = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, anionTriangleBounds.Right - 100, anionTriangleBounds.Top, 100, 30);
            chlorideLabel.TextFrame.TextRange.Text = "Chloride";
            chlorideLabel.TextFrame.TextRange.Font.Size = 15;
            #endregion

            ExportTriangleToPowerpoint(slide, cationTriangleBounds, "Cations", cations);
            ExportTriangleToPowerpoint(slide, anionTriangleBounds, "Anions", anions);
            ExportDiamondToPowerpoint(slide, diamondBounds, cationTriangleBounds, anionTriangleBounds);
            //#region Draw Legend
            //if (frmImportSamples.WaterData.Count > 0)
            //{
            //    int legendX = anionTriangleBounds.Left + anionTriangleBounds.Width + 25;
            //    int legendY = yOrigin + 100;
            //    int s = 0;
            //    for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            //    {
            //        if (frmImportSamples.WaterData[i].Well_Name.Length + frmImportSamples.WaterData[i].ClientID.Length + frmImportSamples.WaterData[i].Depth.Length + 5 > s)
            //        {
            //            s = frmImportSamples.WaterData[i].Well_Name.Length + frmImportSamples.WaterData[i].ClientID.Length + frmImportSamples.WaterData[i].Depth.Length + 5;
            //        }
            //    }


            //    int legendBoxHeight = (int)(frmImportSamples.WaterData.Count * 30 * (chartBounds.Height / 800.0)); // Make legend height relative
            //    int fontSize = Math.Max(8, Math.Min(12, legendBoxHeight / frmImportSamples.WaterData.Count));
            //    int legendBoxWidth = (int)(s * (fontSize - 3) * (chartBounds.Width / 1200.0)); // Make legend width relative

            //    PowerPoint.Shape legendBox = slide.Shapes.AddShape(
            //    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle,
            //    anionTriangleBounds.Left + anionTriangleBounds.Width + 12, legendY - 10, legendBoxWidth, (frmImportSamples.WaterData.Count * 30) + 5
            //    );
            //    legendBox.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // Make it transparent
            //    legendBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue); // Blue border

            //    int ysample = legendY;
            //    for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            //    {
            //        // Step 4: Plot the point in PowerPoint
            //        Office.MsoAutoShapeType bubbleType = Office.MsoAutoShapeType.msoShapeRectangle; // Default shape (rectangle)


            //        switch (frmImportSamples.WaterData[i].shape)
            //        {
            //            case "circle":
            //                bubbleType = Office.MsoAutoShapeType.msoShapeOval; // Perfect circle
            //                break;
            //            case "cube":
            //                bubbleType = Office.MsoAutoShapeType.msoShapeRectangle; // Cube shape as a rectangle
            //                break;
            //            case "hexagon":
            //                bubbleType = Office.MsoAutoShapeType.msoShapeHexagon; // Hexagon shape
            //                break;
            //            case "merkaba":
            //                bubbleType = Office.MsoAutoShapeType.msoShape5pointStar; // Star shape for Merkaba
            //                break;
            //            case "triangle":
            //                bubbleType = Office.MsoAutoShapeType.msoShapeIsoscelesTriangle; // Triangle shape
            //                break;
            //        }
            //        if (frmImportSamples.WaterData[i].shape != "circle" && frmImportSamples.WaterData[i].shape != "cube" && frmImportSamples.WaterData[i].shape != "hexagon" && frmImportSamples.WaterData[i].shape != "merkaba" && frmImportSamples.WaterData[i].shape != "triangle")
            //        {
            //            // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
            //            var horizontalRectangle = slide.Shapes.AddShape(bubbleType, legendX - 10, ysample - 5, 20, 10); // Adjust for your specific use case

            //            horizontalRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color);
            //            horizontalRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            //            // Set the border thickness
            //            horizontalRectangle.Line.Weight = 1; // Adjust thickness as needed
            //            // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
            //            var verticalRectangle = slide.Shapes.AddShape(bubbleType, legendX - 5, ysample - 10, 10, 20); // Adjust for your specific use case

            //            verticalRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color);
            //            verticalRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

            //            // Set the border thickness
            //            verticalRectangle.Line.Weight = 1; // Adjust thickness as needed


            //        }
            //        else
            //        {
            //            // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
            //            var bubble = slide.Shapes.AddShape(bubbleType, legendX - 17, ysample - 17, 35, 35); // Adjust for your specific use case

            //            bubble.Fill.ForeColor.RGB = ColorTranslator.ToOle(frmImportSamples.WaterData[i].color);
            //            bubble.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
            //        }

            //        var Label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, legendX + 30, ysample, 400, 30);
            //        Label.TextFrame.TextRange.Text = (frmImportSamples.WaterData[i].Well_Name) + ", " + (frmImportSamples.WaterData[i].ClientID) + ", " + (frmImportSamples.WaterData[i].Depth);
            //        Label.TextFrame.TextRange.Font.Name = "Times New Roman";
            //        Label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            //        Label.TextFrame.TextRange.Font.Size = fontSize;
            //        ysample += 30;
            //    }
            //}
            //#endregion
        }
        public static void ExportTriangleToPowerpoint(PowerPoint.Slide slide, Rectangle bounds, string label, string[] data)
        {
            #region Draw Triangle Sides
            PointF[] vertices = new PointF[]
            {
                new PointF(bounds.Left, bounds.Bottom), // Bottom-left
                new PointF(bounds.Right, bounds.Bottom), // Bottom-right
                new PointF(bounds.Left + bounds.Width / 2, bounds.Top), // Top
            };
            float[,] points = {
                { vertices[0].X, vertices[0].Y },  // Point 1
                { vertices[1].X, vertices[1].Y },   // Point 2
                { vertices[2].X, vertices[2].Y }  // Point 3
            };
            PowerPoint.Shape polygon = slide.Shapes.AddPolyline(points);
            polygon.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            polygon.Line.Weight = 2;

            PowerPoint.Shape p = slide.Shapes.AddPolyline(new float[,]
                {
                    { vertices[2].X,vertices[2].Y },
                    { vertices[0].X,vertices[0].Y }
                });
            p.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            p.Line.Weight = 2;
            #endregion
            int fontsize = 8;
            // Label vertices
            if (data[1] == "HCO3+CO3")
            {
                var Label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, vertices[2].X - 270, vertices[2].Y + 55, 400, 30);
                Label.TextFrame.TextRange.Text = "Carbonate (CO3) + Bicarbonate(HCO3)";
                Label.TextFrame.TextRange.Font.Name = "Times New Roman";
                Label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                Label.TextFrame.TextRange.Font.Size = fontsize;
                Label.Rotation = -62;

            }
            else
            {
                var Label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, vertices[2].X, vertices[0].Y + 10, 100, 30);
                Label.TextFrame.TextRange.Font.Size = fontsize;
                Label.TextFrame.TextRange.Font.Name = "Times New Roman";
                Label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                Label.TextFrame.TextRange.Text = data[1];
            }
            if (data[2] == "CL")
            {
                var Label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, vertices[2].X, vertices[0].Y + 10, 100, 30);
                Label.TextFrame.TextRange.Font.Size = fontsize;
                Label.TextFrame.TextRange.Font.Name = "Times New Roman";
                Label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                Label.TextFrame.TextRange.Text = data[2];
            }
            else
            {
                var Label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, vertices[2].X - 60, vertices[2].Y + 180, 400, 30);
                Label.TextFrame.TextRange.Text = "Sodium (Na)+Potassium(K)";
                Label.TextFrame.TextRange.Font.Size = fontsize;
                Label.TextFrame.TextRange.Font.Name = "Times New Roman";
                Label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                Label.Rotation = 62;
            }

            if (data[0] == "Mg")
            {
                var Label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, vertices[2].X - 150, vertices[2].Y + 130, 100, 30);
                Label.TextFrame.TextRange.Text = data[0];
                Label.TextFrame.TextRange.Font.Name = "Times New Roman";
                Label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                Label.TextFrame.TextRange.Font.Size = fontsize;
            }
            else if (data[0] == "SO4")
            {
                var Label = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, vertices[2].X + 70, vertices[2].Y + 70, 250, 30);
                Label.TextFrame.TextRange.Text = data[0];
                Label.TextFrame.TextRange.Font.Name = "Times New Roman";
                Label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                Label.TextFrame.TextRange.Font.Size = fontsize;
            }


            #region Draw grid lines and numbered ranges
            int gridLines = 10; // Number of divisions

            float ii = (float)8;
            float c = (float)10;
            for (int i = 0; i < gridLines; i += 2)
            {
                // Fraction for positioning
                float fraction = i / (float)gridLines;


                // Bottom-left to top
                PointF leftToTop = new PointF(
                    vertices[0].X + fraction * (vertices[2].X - vertices[0].X),
                    vertices[0].Y - fraction * (vertices[0].Y - vertices[2].Y)
                );

                // Bottom-right to top
                PointF rightToTop = new PointF(
                    vertices[1].X - fraction * (vertices[1].X - vertices[2].X),
                    vertices[1].Y - fraction * (vertices[1].Y - vertices[2].Y)
                );

                // Bottom-left to bottom-right
                PointF leftToRight = new PointF(
                    vertices[0].X + fraction * (vertices[1].X - vertices[0].X),
                    vertices[0].Y
                );

                // Draw grid lines

                PowerPoint.Shape diagonal1 = slide.Shapes.AddLine((float)leftToTop.X, (float)leftToTop.Y, (float)leftToRight.X, (float)leftToRight.Y);
                diagonal1.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                if (i == 0)
                {
                    PowerPoint.Shape diagonal2 = slide.Shapes.AddLine((float)leftToTop.X, (float)leftToTop.Y, (float)rightToTop.X, (float)rightToTop.Y);
                    diagonal2.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    PowerPoint.Shape diagonal2 = slide.Shapes.AddLine((float)leftToTop.X, (float)leftToTop.Y, (float)rightToTop.X, (float)rightToTop.Y);
                    diagonal2.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                }

                // Labels for sides
                PowerPoint.Shape leftside = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    leftToTop.X - 30, leftToTop.Y - 20, 300, 15);
                leftside.TextFrame.TextRange.Text = (i * 10).ToString("0");
                leftside.TextFrame.TextRange.Font.Size = fontsize;
                PowerPoint.Shape rightside = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    rightToTop.X + 5, rightToTop.Y - 10, 300, 15);
                rightside.TextFrame.TextRange.Text = ((c) * 10).ToString("0");
                rightside.TextFrame.TextRange.Font.Size = fontsize;
                PowerPoint.Shape bottomside = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    leftToRight.X - 5, leftToRight.Y + 5, 300, 15);
                bottomside.TextFrame.TextRange.Text = ((c) * 10).ToString("0");
                bottomside.TextFrame.TextRange.Font.Size = fontsize;

                c -= 2;
                if (i != 0)
                {
                    ii /= 10;
                    rightToTop = new PointF(
                    vertices[1].X - ii * (vertices[1].X - vertices[2].X),
                    vertices[1].Y - ii * (vertices[1].Y - vertices[2].Y)
                    );
                    PowerPoint.Shape diagonal3 = slide.Shapes.AddLine((float)leftToRight.X, (float)leftToRight.Y, (float)rightToTop.X, (float)rightToTop.Y);
                    diagonal3.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                    ii *= 10;
                    ii -= 2;
                }


            }


            // Bottom-left to top
            PointF topToLeft = new PointF(
                vertices[0].X + (vertices[2].X - vertices[0].X),
                vertices[0].Y - (vertices[0].Y - vertices[2].Y)
            );

            // Bottom-right to top
            PointF TopToRight = new PointF(
                vertices[1].X - (vertices[1].X - vertices[2].X),
                vertices[1].Y - (vertices[1].Y - vertices[2].Y)
            );

            // Bottom-left to bottom-right
            PointF RightToLeft = new PointF(
                vertices[0].X + (vertices[1].X - vertices[0].X),
                vertices[0].Y
            );
            PowerPoint.Shape label1 = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    topToLeft.X - 35, topToLeft.Y - 10, 100, 15);
            label1.TextFrame.TextRange.Text = (100).ToString("0");
            label1.TextFrame.TextRange.Font.Size = fontsize;
            PowerPoint.Shape label2 = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    TopToRight.X + 5, TopToRight.Y - 10, 100, 15);
            label2.TextFrame.TextRange.Text = (0).ToString("0");
            label2.TextFrame.TextRange.Font.Size = fontsize;
            PowerPoint.Shape label3 = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    RightToLeft.X - 5, RightToLeft.Y + 5, 100, 15);
            label3.TextFrame.TextRange.Text = (0).ToString("0");
            label3.TextFrame.TextRange.Font.Size = fontsize;

            #endregion

            if (label == "Cations")
            {
                // Define polygon points for Magnesium section (top)
                float[,] magnesiumPoints = {
                    { bounds.Left + bounds.Width / 2, bounds.Top },  // Top
                    { (vertices[0].X + vertices[2].X) / 2, (vertices[0].Y + vertices[2].Y) / 2 }, // Bottom left
                    { (vertices[1].X + vertices[2].X) / 2, (vertices[1].Y + vertices[2].Y) / 2 }  // Bottom right
                };

                // Define polygon points for Calcium section (left bottom)
                float[,] calciumPoints = {
                    { (vertices[0].X + vertices[2].X) / 2, (vertices[0].Y + vertices[2].Y) / 2 }, // Top
                    { bounds.Left, bounds.Bottom }, // Bottom left
                    { (vertices[0].X + vertices[1].X) / 2, (vertices[0].Y + vertices[1].Y) / 2 } // Bottom right
                };

                // Define polygon points for Sodium/Potassium section (right bottom)
                float[,] sodiumPotassiumPoints = {
                    { (vertices[1].X + vertices[2].X) / 2, (vertices[1].Y + vertices[2].Y) / 2 }, // Top
                    { bounds.Right, bounds.Bottom }, // Bottom right
                    { (vertices[0].X + vertices[1].X) / 2, (vertices[0].Y + vertices[1].Y) / 2 } // Bottom left
                };

                // Add polygons to PowerPoint Slide
                PowerPoint.Shape magnesiumShape = slide.Shapes.AddPolyline(magnesiumPoints);
                magnesiumShape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                magnesiumShape.Fill.Transparency = 0.5f; // 50% Transparent
                magnesiumShape.Line.Visible = Office.MsoTriState.msoFalse; // No border

                PowerPoint.Shape calciumShape = slide.Shapes.AddPolyline(calciumPoints);
                calciumShape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                calciumShape.Fill.Transparency = 0.5f;
                calciumShape.Line.Visible = Office.MsoTriState.msoFalse;

                PowerPoint.Shape sodiumPotassiumShape = slide.Shapes.AddPolyline(sodiumPotassiumPoints);
                sodiumPotassiumShape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Cyan);
                sodiumPotassiumShape.Fill.Transparency = 0.5f;
                sodiumPotassiumShape.Line.Visible = Office.MsoTriState.msoFalse;

                //// Add Text Labels for Sections
                //PowerPoint.Shape text1 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                //magnesiumPoints[1, 0] + 20, magnesiumPoints[1, 1] - 60, 150, 30);
                //text1.TextFrame.TextRange.Text = "Magnesium";
                //text1.TextFrame.TextRange.Font.Name = "Times New Roman";
                //text1.TextFrame.TextRange.Font.Size = fontsize;
                //text1.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                //PowerPoint.Shape text2 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                //calciumPoints[1, 0] + 30, calciumPoints[1, 1] - 70, 100, 30);
                //text2.TextFrame.TextRange.Text = "Calcium type";
                //text2.TextFrame.TextRange.Font.Name = "Times New Roman";
                //text2.TextFrame.TextRange.Font.Size = fontsize;
                //text2.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;

                //PowerPoint.Shape text3 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                // sodiumPotassiumPoints[2, 0] + 40, sodiumPotassiumPoints[2, 1] - 90, 100, 30);
                //text3.TextFrame.TextRange.Text = "Sodium and Potassium";
                //text3.TextFrame.TextRange.Font.Name = "Times New Roman";
                //text3.TextFrame.TextRange.Font.Size = fontsize;
                //text3.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;

                //PowerPoint.Shape text4 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                //calciumPoints[0, 0] + 30, calciumPoints[0, 1] + 10, 100, 30);
                //text4.TextFrame.TextRange.Text = "No dominant type";
                //text4.TextFrame.TextRange.Font.Name = "Times New Roman";
                //text4.TextFrame.TextRange.Font.Size = fontsize;
                //text4.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;


            }
            else if (label == "Anions")
            {
                // Define polygon points for Sulphate section (top)
                float[,] sulphatePoints = {
                    { bounds.Left + bounds.Width / 2, bounds.Top },
                    { (vertices[0].X + vertices[2].X) / 2, (vertices[0].Y + vertices[2].Y) / 2 },
                    { (vertices[1].X + vertices[2].X) / 2, (vertices[1].Y + vertices[2].Y) / 2 }
                };

                // Define polygon points for Bicarbonate section (left bottom)
                float[,] bicarbonatePoints = {
                    { (vertices[0].X + vertices[2].X) / 2, (vertices[0].Y + vertices[2].Y) / 2 },
                    { bounds.Left, bounds.Bottom },
                    { (vertices[0].X + vertices[1].X) / 2, (vertices[0].Y + vertices[1].Y) / 2 }
                };

                // Define polygon points for Chloride section (right bottom)
                float[,] chloridePoints = {
                    { (vertices[1].X + vertices[2].X) / 2, (vertices[1].Y + vertices[2].Y) / 2 },
                    { bounds.Right, bounds.Bottom },
                    { (vertices[0].X + vertices[1].X) / 2, (vertices[0].Y + vertices[1].Y) / 2 }
                };

                // Add polygons to PowerPoint Slide
                PowerPoint.Shape sulphateShape = slide.Shapes.AddPolyline(sulphatePoints);
                sulphateShape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Pink);
                sulphateShape.Fill.Transparency = 0.5f;
                sulphateShape.Line.Visible = Office.MsoTriState.msoFalse;

                PowerPoint.Shape bicarbonateShape = slide.Shapes.AddPolyline(bicarbonatePoints);
                bicarbonateShape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Magenta);
                bicarbonateShape.Fill.Transparency = 0.5f;
                bicarbonateShape.Line.Visible = Office.MsoTriState.msoFalse;

                PowerPoint.Shape chlorideShape = slide.Shapes.AddPolyline(chloridePoints);
                chlorideShape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkOrange);
                chlorideShape.Fill.Transparency = 0.5f;
                chlorideShape.Line.Visible = Office.MsoTriState.msoFalse;

                // Add Text Labels with "Times New Roman" font
                //PowerPoint.Shape text1 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                //    sulphatePoints[1, 0] + 30, sulphatePoints[1, 1] - 70, 100, 40);
                //text1.TextFrame.TextRange.Text = "Sulphate type";
                //text1.TextFrame.TextRange.Font.Name = "Times New Roman";
                //text1.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                //text1.TextFrame.TextRange.Font.Size = fontsize;

                //PowerPoint.Shape text2 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                //    bicarbonatePoints[1, 0] + 20, bicarbonatePoints[1, 1] - 60, 120, 40);
                //text2.TextFrame.TextRange.Text = "Bicarbonate type";
                //text2.TextFrame.TextRange.Font.Name = "Times New Roman";
                //text2.TextFrame.TextRange.Font.Size = fontsize;
                //text2.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;

                //PowerPoint.Shape text3 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                //    chloridePoints[2, 0] + 40, chloridePoints[2, 1] - 70, 100, 60);
                //text3.TextFrame.TextRange.Text = "Chloride type";
                //text3.TextFrame.TextRange.Font.Name = "Times New Roman";
                //text3.TextFrame.TextRange.Font.Size = fontsize;
                //text3.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                //PowerPoint.Shape text4 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                //    bicarbonatePoints[0, 0] + 45, bicarbonatePoints[0, 1] + 10, 100, 40);
                //text4.TextFrame.TextRange.Text = "No dominant type";
                //text4.TextFrame.TextRange.Font.Name = "Times New Roman";
                //text4.TextFrame.TextRange.Font.Size = fontsize;
                //text4.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
            }

            PowerPoint.Shape Labeltext = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    vertices[2].X - 20, vertices[0].Y + 30, 100, 30);
            Labeltext.TextFrame.TextRange.Text = label;
            Labeltext.TextFrame.TextRange.Font.Size = fontsize;
            Labeltext.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            #region plot the points

            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            {
                Color brush = frmImportSamples.WaterData[i].color;
                if (label == "Cations")
                {

                    PlotPointInTrianglePowerpoint(slide, bounds, frmImportSamples.WaterData[i].Mg, frmImportSamples.WaterData[i].Na + frmImportSamples.WaterData[i].K, frmImportSamples.WaterData[i].Ca, brush, "Cations",frmImportSamples.WaterData[i].shape);
                }
                else
                {
                    PlotPointInTrianglePowerpoint(slide, bounds, frmImportSamples.WaterData[i].So4, frmImportSamples.WaterData[i].Cl, frmImportSamples.WaterData[i].HCO3 + frmImportSamples.WaterData[i].CO3, brush, "Anions", frmImportSamples.WaterData[i].shape);
                }
            }
            #endregion
        }
        public static void FindIntersectionPowerpoint(PowerPoint.Slide slide, Rectangle bounds, double NaK, double Ca, double Mg, double ClSo4, double HCO3, double CO3, Color brush, string shape)
        {
            PointF[] diamondVertices = new PointF[]
            {
            new PointF(bounds.Left + bounds.Width / 2, bounds.Top), // Top
            new PointF(bounds.Right, bounds.Top + bounds.Height / 2), // Right
            new PointF(bounds.Left + bounds.Width / 2, bounds.Bottom), // Bottom
            new PointF(bounds.Left, bounds.Top + bounds.Height / 2), // Left
            };
            float Xc = (float)(NaK / (Ca + Mg + NaK));
            float Ya = (float)(ClSo4 / (HCO3 + CO3 + ClSo4));
            // Calculate intersection point
            // Bilinear interpolation
            float x = (1 - Xc) * ((1 - Ya) * diamondVertices[3].X + Ya * diamondVertices[0].X) + Xc * ((1 - Ya) * diamondVertices[2].X + Ya * diamondVertices[1].X);
            float y = (1 - Xc) * ((1 - Ya) * diamondVertices[3].Y + Ya * diamondVertices[0].Y) + Xc * ((1 - Ya) * diamondVertices[2].Y + Ya * diamondVertices[1].Y);

            // Step 4: Plot the point in PowerPoint
            Office.MsoAutoShapeType bubbleType = Office.MsoAutoShapeType.msoShapeRectangle; // Default shape (rectangle)

            // Check the shape and create corresponding shape in PowerPoint
            switch (shape)
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

            if (shape != "circle" && shape != "cube" && shape != "hexagon" && shape != "merkaba" && shape != "triangle")
            {
                // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
                var horizontalRectangle = slide.Shapes.AddShape(bubbleType, x - 10, y - 5, 20, 10); // Adjust for your specific use case

                horizontalRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                horizontalRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

                // Set the border thickness
                horizontalRectangle.Line.Weight = 1; // Adjust thickness as needed
                                                     // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
                var verticalRectangle = slide.Shapes.AddShape(bubbleType, x - 5, y - 10, 10, 20); // Adjust for your specific use case

                verticalRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                verticalRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

                // Set the border thickness
                verticalRectangle.Line.Weight = 1; // Adjust thickness as needed
            }
            else
            {
                // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
                var bubble = slide.Shapes.AddShape(bubbleType, x - 17, y - 17, 35, 35); // Adjust for your specific use case

                bubble.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                bubble.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
            }
        }


        public static void PlotPointInTrianglePowerpoint(PowerPoint.Slide slide, Rectangle bounds, double A, double B, double C, Color brush, string label, string shape)
        {
            // Step 1: Normalize the values
            double total = A + B + C;
            double aNormalized = A / total;
            double bNormalized = B / total;
            double cNormalized = C / total;

            // Step 2: Calculate the triangle vertices
            PointF bottomLeft = new PointF(bounds.Left, bounds.Bottom);
            PointF bottomRight = new PointF(bounds.Right, bounds.Bottom);
            PointF top = new PointF(bounds.Left + bounds.Width / 2, bounds.Top);

            // Step 3: Interpolate position within the triangle
            float x = (float)(
                cNormalized * bottomLeft.X +
                bNormalized * bottomRight.X +
                aNormalized * top.X
            );

            float y = (float)(
                cNormalized * bottomLeft.Y +
                bNormalized * bottomRight.Y +
                aNormalized * top.Y
            );

            // Store points in the corresponding list
            if (label == "Cations")
            {
                cationVertices.Add(new PointF(x, y));
            }
            else
            {
                anionVertices.Add(new PointF(x, y));
            }

            // Step 4: Plot the point in PowerPoint
            Office.MsoAutoShapeType bubbleType = Office.MsoAutoShapeType.msoShapeRectangle; // Default shape (rectangle)


            switch (shape)
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
            if (shape != "circle" && shape != "cube" && shape != "hexagon" && shape != "merkaba" && shape != "triangle")
            {
                // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
                var horizontalRectangle = slide.Shapes.AddShape(bubbleType, x - 10, y - 5, 20, 10); // Adjust for your specific use case

                horizontalRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                horizontalRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

                // Set the border thickness
                horizontalRectangle.Line.Weight = 1; // Adjust thickness as needed
                // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
                var verticalRectangle = slide.Shapes.AddShape(bubbleType, x - 5, y - 10, 10, 20); // Adjust for your specific use case

                verticalRectangle.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush);
                verticalRectangle.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();

                // Set the border thickness
                verticalRectangle.Line.Weight = 1; // Adjust thickness as needed


            }
            else 
            {
                // Assuming 'bubble' is a shape object in a slide or document that can be assigned to the shape type
                var bubble = slide.Shapes.AddShape(bubbleType, x - 17, y - 17, 35, 35); // Adjust for your specific use case

                bubble.Fill.ForeColor.RGB = ColorTranslator.ToOle(brush == Color.Blue ? Color.Red : brush == Color.Red ? Color.Blue : brush);
                bubble.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
            }



        }
        public static void FormatKeyword(Microsoft.Office.Interop.PowerPoint.TextRange textRange, string keyword, string colorName)
        {
            int start = textRange.Text.IndexOf(keyword);
            if (start >= 0)
            {
                int length = keyword.Length;
                Microsoft.Office.Interop.PowerPoint.TextRange keywordRange = textRange.Characters(start + 1, length);

                // Convert the colorName to an RGB value
                Color color = Color.FromName(colorName);
                keywordRange.Font.Color.RGB = ColorTranslator.ToOle(color);

                keywordRange.Font.Size = 15;
                keywordRange.Font.Bold = Office.MsoTriState.msoTrue;
                keywordRange.Font.Italic = Office.MsoTriState.msoTrue;

                // Restore the underline state
                if (keyword == "Classification of water")
                {
                    keywordRange.Font.Underline = Office.MsoTriState.msoTrue;
                }
                else
                {
                    keywordRange.Font.Underline = Office.MsoTriState.msoFalse;
                }

                // Access and resize the parent shape
                Microsoft.Office.Interop.PowerPoint.Shape parentShape = textRange.Parent as Microsoft.Office.Interop.PowerPoint.Shape;
                if (parentShape != null && parentShape.Type == Office.MsoShapeType.msoTextBox)
                {
                    // Decrease the width and height of the textbox
                    parentShape.Width -= 20;  // Reduce the width by 20 units
                    parentShape.Height -= 10; // Reduce the height by 10 units
                }
            }
        }
        public static void ExportDiamondToPowerpoint(PowerPoint.Slide slide, Rectangle bounds, Rectangle cationTriangleBounds, Rectangle anionTriangleBounds)
        {
            PointF[] vertices = new PointF[]
            {
                new PointF(bounds.Left + bounds.Width / 2, bounds.Top), // Top
                new PointF(bounds.Right, bounds.Top + bounds.Height / 2), // Right
                new PointF(bounds.Left + bounds.Width / 2, bounds.Bottom), // Bottom
                new PointF(bounds.Left, bounds.Top + bounds.Height / 2), // Left
            };
            float[,] points = {
                { vertices[0].X, vertices[0].Y },  // Point 1
                { vertices[1].X, vertices[1].Y },   // Point 2
                { vertices[2].X, vertices[2].Y },
                { vertices[3].X, vertices[3].Y}// Point 3
            };
            PowerPoint.Shape polygon = slide.Shapes.AddPolyline(points);
            polygon.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            polygon.Line.Weight = 2;

            PowerPoint.Shape p = slide.Shapes.AddPolyline(new float[,]
                {
                    { vertices[3].X,vertices[3].Y },
                    { vertices[0].X,vertices[0].Y }
                });
            p.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

            // Draw diamond
            // Define circle colors and positions
            int fontsize = 8;
            var Label1 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, vertices[0].X - 270, vertices[0].Y + 50, 400, 30);
            Label1.TextFrame.TextRange.Text = "Sulphate (So4) + Chloride (Cl)";
            Label1.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            Label1.TextFrame.TextRange.Font.Size = fontsize;
            Label1.Rotation = -62;
            var Label2 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, vertices[0].X - 50, vertices[0].Y + 200, 400, 30);
            Label2.TextFrame.TextRange.Text = "Calcium (Ca) + Magnesium (Mg)";
            Label2.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            Label2.TextFrame.TextRange.Font.Size = fontsize;
            Label2.Rotation = 62;


            int gridLines = 10; // Number of divisions


            for (int i = 0; i <= gridLines; i += 2)
            {
                float fraction = i / (float)gridLines;

                // Interpolate points along the edges
                PointF topToRight = new PointF(
                    vertices[1].X + fraction * (vertices[0].X - vertices[1].X),
                    vertices[1].Y + fraction * (vertices[0].Y - vertices[1].Y)
                );

                PointF rightToBottom = new PointF(
                    vertices[1].X + fraction * (vertices[2].X - vertices[1].X),
                    vertices[1].Y + fraction * (vertices[2].Y - vertices[1].Y)
                );

                PointF bottomToLeft = new PointF(
                    vertices[2].X + fraction * (vertices[3].X - vertices[2].X),
                    vertices[2].Y + fraction * (vertices[3].Y - vertices[2].Y)
                );

                PointF leftToTop = new PointF(
                    vertices[0].X + fraction * (vertices[3].X - vertices[0].X),
                    vertices[0].Y + fraction * (vertices[3].Y - vertices[0].Y)
                );
                // Draw diagonals
                if (i != 0)
                {
                    PowerPoint.Shape diagonal1 = slide.Shapes.AddLine((float)topToRight.X, (float)topToRight.Y, (float)bottomToLeft.X, (float)bottomToLeft.Y);
                    diagonal1.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                    PowerPoint.Shape diagonal2 = slide.Shapes.AddLine((float)rightToBottom.X, (float)rightToBottom.Y, (float)leftToTop.X, (float)leftToTop.Y);
                    diagonal2.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);
                }
                var L1 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, topToRight.X - 5, topToRight.Y - 15, 250, 30);
                L1.TextFrame.TextRange.Text = (i * 10).ToString();
                L1.TextFrame.TextRange.Font.Size = fontsize;
                var L2 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, leftToTop.X - 25, leftToTop.Y - 10, 250, 30);
                L2.TextFrame.TextRange.Text = (i * 10).ToString();
                L2.TextFrame.TextRange.Font.Size = fontsize;
            }
            // Define colors for each region
            Color[] regionColors = new Color[]
            {
                Color.Yellow,     // Calcium-Chloride
                Color.LightGreen,  // Mixed Type 1
                Color.MediumPurple,  // Magnesium-Bicarbonate
                Color.Magenta,   // Sodium-Chloride
                Color.DarkOrange,   // Mixed Type 2
                Color.LightBlue   // Sodium-Bicarbonate
            };

            // Define all diamond region points
            float[][,] diamondRegions = new float[][,]
            {
                // Calcium-Chloride Region
                new float[,] {
                    { bounds.Left + bounds.Width / 2, bounds.Top },
                    { (vertices[0].X + vertices[1].X) / 2, (vertices[0].Y + vertices[1].Y) / 2 },
                    { (vertices[0].X + vertices[3].X) / 2, (vertices[0].Y + vertices[3].Y) / 2 }
                },

                // Mixed Type 1 Region
                new float[,] {
                    { (vertices[0].X + vertices[1].X) / 2, (vertices[0].Y + vertices[1].Y) / 2 },
                    { (vertices[0].X + vertices[3].X) / 2, (vertices[0].Y + vertices[3].Y) / 2 },
                    { bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2 }
                },

                // Magnesium-Bicarbonate Region
                new float[,] {
                    { (vertices[0].X + vertices[3].X) / 2, (vertices[0].Y + vertices[3].Y) / 2 },
                    { bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2 },
                    { (vertices[2].X + vertices[3].X) / 2, (vertices[2].Y + vertices[3].Y) / 2 },
                    { vertices[3].X, vertices[3].Y }
                },

                // Sodium-Chloride Region
                new float[,] {
                    { (vertices[0].X + vertices[1].X) / 2, (vertices[0].Y + vertices[1].Y) / 2 },
                    { bounds.Right, bounds.Top + (bounds.Height / 2) },
                    { (vertices[1].X + vertices[2].X) / 2, (vertices[1].Y + vertices[2].Y) / 2 },
                    { bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2 }
                },

                // Mixed Type 2 Region
                new float[,] {
                    { bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2 },
                    { (vertices[2].X + vertices[3].X) / 2, (vertices[2].Y + vertices[3].Y) / 2 },
                    { (vertices[2].X + vertices[1].X) / 2, (vertices[2].Y + vertices[1].Y) / 2 }
                },

                // Sodium-Bicarbonate Region
                new float[,] {
                    { (vertices[2].X + vertices[3].X) / 2, (vertices[2].Y + vertices[3].Y) / 2 },
                    { (vertices[2].X + vertices[1].X) / 2, (vertices[2].Y + vertices[1].Y) / 2 },
                    { bounds.Left + bounds.Width / 2, bounds.Bottom }
                }
            };

            // Loop through regions and draw polygons in PowerPoint
            for (int i = 0; i < diamondRegions.Length; i++)
            {
                PowerPoint.Shape regionShape = slide.Shapes.AddPolyline(diamondRegions[i]);
                regionShape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(regionColors[i]);
                regionShape.Fill.Transparency = 0.5f; // Set transparency
                regionShape.Line.Visible = Office.MsoTriState.msoFalse; // Hide border
            }

            // Add Labels inside the Diamond regions
            string[] regionLabels = { "Calcium\nchloride\nType", "Mixed\nType", "Magnesium Bicarbonate\nType", "Sodium\nchloride\nType", "Mixed\nType", "Sodium Bicarbonate\nType" };
            


            for (int i = 0; i < diamondRegions.Length; i++)
            {
                // Get the centroid (average) position of each region
                float avgX = 0, avgY = 0;
                for (int j = 0; j < diamondRegions[i].GetLength(0); j++)
                {
                    avgX += diamondRegions[i][j, 0];
                    avgY += diamondRegions[i][j, 1];
                }

                avgX /= diamondRegions[i].GetLength(0);
                avgY /= diamondRegions[i].GetLength(0);

                // Place label in the centroid of the region
                PowerPoint.Shape label = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    avgX-20, avgY-10, 150, 30
                );
                label.TextFrame.TextRange.Text = regionLabels[i];
                label.TextFrame.TextRange.Font.Name = "Times New Roman";
                label.TextFrame.TextRange.Font.Size = fontsize;
                label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            }
            if (frmImportSamples.WaterData.Count > 0)
            {
                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                {
                    var data = frmImportSamples.WaterData[i];
                    Color brush = frmImportSamples.WaterData[i].color;
                    PointF diamondCenter = new PointF((vertices[1].X + vertices[3].X) / 2, (vertices[0].Y + vertices[2].Y) / 2);
                    FindIntersectionPowerpoint(slide,bounds,data.Na+data.K,data.Ca,data.Mg,data.Cl+data.So4,data.HCO3,data.CO3,brush,data.shape);
                }
            }


        }
        public static PointF[] GetHexagonPoints(float x, float y, float size)
        {
            float width = size * (float)Math.Sqrt(3) / 2;
            return new PointF[]
            {
                new PointF(x + width / 2, y),
                new PointF(x + width, y + size / 4),
                new PointF(x + width, y + 3 * size / 4),
                new PointF(x + width / 2, y + size),
                new PointF(x, y + 3 * size / 4),
                new PointF(x, y + size / 4),
            };
        }

        public static PointF[] GetStarPoints(float x, float y, float size)
        {
            float innerRadius = size / 2.5f;
            float outerRadius = size;
            PointF[] points = new PointF[10];
            double angle = -Math.PI / 2;

            for (int i = 0; i < 10; i++)
            {
                float radius = (i % 2 == 0) ? outerRadius : innerRadius;
                points[i] = new PointF(
                    x + size / 2 + (float)(Math.Cos(angle) * radius),
                    y + size / 2 + (float)(Math.Sin(angle) * radius)
                );
                angle += Math.PI / 5;
            }

            return points;
        }
    }
}
