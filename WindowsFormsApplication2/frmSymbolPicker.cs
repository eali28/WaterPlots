using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public partial class frmSymbolPicker : Form
    {
        public static PictureBox symbolPictureBox;
        public static HashSet<string> symbolNames = new HashSet<string>();
        public static Brush symbolColor;
        public static string selectedShape=null;
        public frmSymbolPicker(Brush brush)
        {
            int cols = 5;
            int rows = 6;
            int cellSize = 40;
            this.ClientSize = new System.Drawing.Size(cols * cellSize, rows * cellSize);
            this.ShowInTaskbar = true;
            this.TopMost = false;
            InitializeComponent();
            InitializeSymbolPictureBox();
            
            if (brush == null || ReferenceEquals(brush, Brushes.Transparent))
            {
                symbolColor = Brushes.Red;
            }
            else 
            {
                symbolColor = brush;
            }
            
            loadSymbols();
        }

        private void InitializeSymbolPictureBox()
        {
            symbolPictureBox = new PictureBox();
            symbolPictureBox.Dock = DockStyle.Fill;
            symbolPictureBox.SizeMode = PictureBoxSizeMode.Normal;
            symbolPictureBox.MouseClick += pictureBox1_MouseClick;
            this.Controls.Add(symbolPictureBox);
        }

        public static void loadSymbols()
        {
            int cols = 5; // number of columns
            int rows = 6; // number of rows
            int cellSize = 40; // size of each cell
            Bitmap symbolGrid = new Bitmap(cols * cellSize, rows * cellSize);

            using (Graphics g = Graphics.FromImage(symbolGrid))
            {
                g.Clear(Color.White);
                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        int x = col * cellSize;
                        int y = row * cellSize;
                        int index = row * cols + col;
                        DrawSymbol(g, index, x, y, cellSize,symbolColor);
                        g.DrawRectangle(Pens.Black, x, y, cellSize, cellSize);
                    }
                }
            }
            symbolPictureBox.Image = symbolGrid;
        }

        public static void DrawSymbol(Graphics g, int index, int x, int y, int size,Brush color)
        {
            // Red fill, blue border
            Brush fill = color;
            Pen border = new Pen(Color.Blue, 1);
            int cx = x + size / 2;
            int cy = y + size / 2;
            int r = (size - 10) / 2;
            switch (index)
            {
                case 0: // Circle
                    symbolNames.Add("Circle");
                    g.FillEllipse(fill, cx - r, cy - r, 2 * r, 2 * r);
                    g.DrawEllipse(border, cx - r, cy - r, 2 * r, 2 * r);
                    break;
                case 1: // Diamond
                    symbolNames.Add("Diamond");
                    Point[] diamond = { new Point(cx, y + 5), new Point(x + size - 5, cy), new Point(cx, y + size - 5), new Point(x + 5, cy) };
                    g.FillPolygon(fill, diamond);
                    g.DrawPolygon(border, diamond);
                    break;
                case 2: // Pentagon
                    symbolNames.Add("Pentagon");
                    g.FillPolygon(fill, GetPolygon(cx, cy, r, 5, -Math.PI / 2));
                    g.DrawPolygon(border, GetPolygon(cx, cy, r, 5, -Math.PI / 2));
                    break;
                case 3: // Hexagon
                    symbolNames.Add("Hexagon");
                    g.FillPolygon(fill, GetPolygon(cx, cy, r, 6, 0));
                    g.DrawPolygon(border, GetPolygon(cx, cy, r, 6, 0));
                    break;
                case 4: // Octagon
                    symbolNames.Add("Octagon");
                    g.FillPolygon(fill, GetPolygon(cx, cy, r, 8, 0));
                    g.DrawPolygon(border, GetPolygon(cx, cy, r, 8, 0));
                    break;
                case 5: // Up triangle
                    symbolNames.Add("Up triangle");
                    g.FillPolygon(fill, new Point[] { new Point(cx, y + 5), new Point(x + size - 5, y + size - 5), new Point(x + 5, y + size - 5) });
                    g.DrawPolygon(border, new Point[] { new Point(cx, y + 5), new Point(x + size - 5, y + size - 5), new Point(x + 5, y + size - 5) });
                    break;
                case 6: // Left triangle
                    symbolNames.Add("Right triangle");
                    g.FillPolygon(fill, new Point[] { new Point(x + size - 5, cy), new Point(x + 5, y + 5), new Point(x + 5, y + size - 5) });
                    g.DrawPolygon(border, new Point[] { new Point(x + size - 5, cy), new Point(x + 5, y + 5), new Point(x + 5, y + size - 5) });
                    break;
                case 7: // Down triangle
                    symbolNames.Add("Down triangle");
                    g.FillPolygon(fill, new Point[] { new Point(cx, y + size - 5), new Point(x + size - 5, y + 5), new Point(x + 5, y + 5) });
                    g.DrawPolygon(border, new Point[] { new Point(cx, y + size - 5), new Point(x + size - 5, y + 5), new Point(x + 5, y + 5) });
                    break;
                case 8: // Right triangle
                    symbolNames.Add("Left triangle");
                    g.FillPolygon(fill, new Point[] { new Point(x + 5, cy), new Point(x + size - 5, y + 5), new Point(x + size - 5, y + size - 5) });
                    g.DrawPolygon(border, new Point[] { new Point(x + 5, cy), new Point(x + size - 5, y + 5), new Point(x + size - 5, y + size - 5) });
                    break;
                case 9: // Star (5-point)
                    symbolNames.Add("Star (5-point)");
                    g.FillPolygon(fill, GetStar(cx, cy, r, r / 2, 5));
                    g.DrawPolygon(border, GetStar(cx, cy, r, r / 2, 5));
                    break;
                case 10: // Star (6-point)
                    symbolNames.Add("Star (6-point)");
                    g.FillPolygon(fill, GetStar(cx, cy, r, r / 2, 6));
                    g.DrawPolygon(border, GetStar(cx, cy, r, r / 2, 6));
                    break;
                case 11: // Star (8-point)
                    symbolNames.Add("Star (8-point)");
                    g.FillPolygon(fill, GetStar(cx, cy, r, r / 2, 8));
                    g.DrawPolygon(border, GetStar(cx, cy, r, r / 2, 8));
                    break;
                case 12: // Trapezoid (up)
                    symbolNames.Add("Trapezoid (up)");
                    g.FillPolygon(fill, new Point[] { new Point(x + 10, y+5), new Point(x + size - 10, y+5), new Point(x + size - 5, y + size - 5), new Point(x + 5, y + size - 5) });
                    g.DrawPolygon(border, new Point[] { new Point(x + 10, y+5), new Point(x + size - 10,y+5), new Point(x + size - 5, y + size - 5), new Point(x + 5, y + size - 5) });
                    break;
                case 13: // Trapezoid (right)
                    symbolNames.Add("Trapezoid (right)");
                    g.FillPolygon(fill, new Point[] { new Point(x + 5, y + 5), new Point(x + size - 5, y + 10), new Point(x + size - 5, y + size - 10), new Point(x + 5, y + size - 5) });
                    g.DrawPolygon(border, new Point[] { new Point(x + 5, y + 5), new Point(x + size - 5, y + 10), new Point(x + size - 5, y + size - 10), new Point(x + 5, y + size - 5) });
                    break;
                case 14: // Trapezoid (down)
                    symbolNames.Add("Trapezoid (down)");
                    g.FillPolygon(fill, new Point[] { new Point(x + 5, y + 5), new Point(x + size - 5, y + 5), new Point(x + size - 10, y + size - 5), new Point(x + 10, y + size - 5) });
                    g.DrawPolygon(border, new Point[] { new Point(x + 5, y + 5), new Point(x + size - 5, y + 5), new Point(x + size - 10, y + size - 5), new Point(x + 10, y + size - 5) });
                    break;
                case 15: // Trapezoid (left)
                    symbolNames.Add("Trapezoid (left)");
                    g.FillPolygon(fill, new Point[] { new Point(x + 5, y + 10), new Point(x + size - 5, y + 5), new Point(x + size - 5, y + size - 5), new Point(x + 5, y + size - 10) });
                    g.DrawPolygon(border, new Point[] { new Point(x + 5, y + 10), new Point(x + size - 5, y + 5), new Point(x + size - 5, y + size - 5), new Point(x + 5, y + size - 10) });
                    break;
                case 16: // Rectangle
                    symbolNames.Add("Rectangle");
                    g.FillRectangle(fill, x + 8, y + 8, size - 16, size - 16);
                    g.DrawRectangle(border, x + 8, y + 8, size - 16, size - 16);
                    break;
                case 17: // Vertical rectangle
                    symbolNames.Add("Vertical rectangle");
                    g.FillRectangle(fill, cx - 6, y + 8, 12, size - 16);
                    g.DrawRectangle(border, cx - 6, y + 8, 12, size - 16);
                    break;
                case 18: // Plus
                    symbolNames.Add("Plus");
                    g.FillRectangle(fill, cx - 6, y + 8, 12, size - 16);
                    g.FillRectangle(fill, x + 8, cy - 6, size - 16, 12);
                    g.DrawRectangle(border, cx - 6, y + 8, 12, size - 16);
                    g.DrawRectangle(border, x + 8, cy - 6, size - 16, 12);
                    break;
                case 19: // X
                    symbolNames.Add("X");
                    int thickness = 8;
                    // First diagonal: top-left to bottom-right
                    Point[] bar1 = new Point[] {
                        new Point(x + 8, y + 8 + thickness),
                        new Point(x + 8 + thickness, y + 8),
                        new Point(x + size - 8, y + size - 8 - thickness),
                        new Point(x + size - 8 - thickness, y + size - 8)
                    };
                    g.FillPolygon(fill, bar1);
                    g.DrawPolygon(border, bar1);
                    // Second diagonal: top-right to bottom-left
                    Point[] bar2 = new Point[] {
                        new Point(x + size - 8, y + 8 + thickness),
                        new Point(x + size - 8 - thickness, y + 8),
                        new Point(x + 8, y + size - 8 - thickness),
                        new Point(x + 8 + thickness, y + size - 8)
                    };
                    g.FillPolygon(fill, bar2);
                    g.DrawPolygon(border, bar2);
                    break;
                case 20: // Horizontal bar
                    symbolNames.Add("Horizontal bar");
                    g.FillRectangle(fill, x + 8, cy - 6, size - 16, 12);
                    g.DrawRectangle(border, x + 8, cy - 6, size - 16, 12);
                    break;
                case 21: // Up arrow
                    symbolNames.Add("Up arrow");
                    DrawArrow(g, x, y, size, "up", fill, border);
                    break;
                case 22: // Right arrow
                    symbolNames.Add("Right arrow");
                    DrawArrow(g, x, y, size, "right", fill, border);
                    break;
                case 23: // Down arrow
                    symbolNames.Add("Down arrow");
                    DrawArrow(g, x, y, size, "down", fill, border);
                    break;
                case 24: // Left arrow
                    symbolNames.Add("Left arrow");
                    DrawArrow(g, x, y, size, "left", fill, border);
                    break;
                case 25: // Arrow with tail (up)
                    symbolNames.Add("Arrow with tail (up)");
                    DrawArrow(g, x, y, size, "up", fill, border, true);
                    break;
                case 26: // Arrow with tail (right)
                    symbolNames.Add("Arrow with tail (right)");
                    DrawArrow(g, x, y, size, "right", fill, border, true);
                    break;
                case 27: // Arrow with tail (down)
                    symbolNames.Add("Arrow with tail (down)");
                    DrawArrow(g, x, y, size, "down", fill, border, true);
                    break;
                case 28: // Arrow with tail (left)
                    symbolNames.Add("Arrow with tail (left)");
                    DrawArrow(g, x, y, size, "left", fill, border, true);
                    break;
                case 29: // Upward fat arrow
                    symbolNames.Add("Upward fat arrow");
                    DrawFatArrow(g, x, y, size, "up", fill, border);
                    break;
                default:
                    // fallback: circle
                    
                    g.FillEllipse(fill, cx - r, cy - r, 2 * r, 2 * r);
                    g.DrawEllipse(border, cx - r, cy - r, 2 * r, 2 * r);
                    break;
            }
        }

        public static Point[] GetPolygon(int cx, int cy, int r, int sides, double startAngle)
        {
            Point[] pts = new Point[sides];
            for (int i = 0; i < sides; i++)
            {
                double angle = startAngle + i * 2 * Math.PI / sides;
                pts[i] = new Point(
                    cx + (int)(r * Math.Cos(angle)),
                    cy + (int)(r * Math.Sin(angle))
                );
            }
            return pts;
        }

        public static Point[] GetStar(int cx, int cy, int rOuter, int rInner, int points)
        {
            Point[] pts = new Point[points * 2];
            for (int i = 0; i < points * 2; i++)
            {
                double angle = -Math.PI / 2 + i * Math.PI / points;
                int r = (i % 2 == 0) ? rOuter : rInner;
                pts[i] = new Point(
                    cx + (int)(r * Math.Cos(angle)),
                    cy + (int)(r * Math.Sin(angle))
                );
            }
            return pts;
        }

        public static void DrawArrow(Graphics g, int x, int y, int size, string direction, Brush fill, Pen border, bool tail = false)
        {
            int cx = x + size / 2;
            int cy = y + size / 2;
            int arrowSize = size - 12;
            Point[] pts;
            switch (direction)
            {
                case "up":
                    pts = new Point[] { new Point(cx, y + 6), new Point(x + size - 8, y + size - 8), new Point(cx, y + size - 18), new Point(x + 8, y + size - 8) };
                    break;
                case "right":
                    pts = new Point[] { new Point(x + size - 6, cy), new Point(x + 8, y + size - 8), new Point(x + size - 18, cy), new Point(x + 8, y + 8) };
                    break;
                case "down":
                    pts = new Point[] { new Point(cx, y + size - 6), new Point(x + 8, y + 8), new Point(cx, y + size - 18), new Point(x + size - 8, y + 8) };
                    break;
                case "left":
                    pts = new Point[] { new Point(x + 6, cy), new Point(x + size - 8, y + 8), new Point(x + 18, cy), new Point(x + size - 8, y + size - 8) };
                    break;
                default:
                    pts = new Point[] { new Point(cx, y + 6), new Point(x + size - 8, y + size - 8), new Point(cx, y + size - 18), new Point(x + 8, y + size - 8) };
                    break;
            }
            g.FillPolygon(fill, pts);
            g.DrawPolygon(border, pts);
            if (tail)
            {
                switch (direction)
                {
                    case "up":
                        g.FillRectangle(fill, cx - 5, y + size / 2, 10, size / 2 - 8);
                        g.DrawRectangle(border, cx - 5, y + size / 2, 10, size / 2 - 8);
                        break;
                    case "right":
                        g.FillRectangle(fill, x + 8, cy - 5, size / 2 - 8, 10);
                        g.DrawRectangle(border, x + 8, cy - 5, size / 2 - 8, 10);
                        break;
                    case "down":
                        g.FillRectangle(fill, cx - 5, y + 8, 10, size / 2 - 8);
                        g.DrawRectangle(border, cx - 5, y + 8, 10, size / 2 - 8);
                        break;
                    case "left":
                        g.FillRectangle(fill, x + size / 2, cy - 5, size / 2 - 8, 10);
                        g.DrawRectangle(border, x + size / 2, cy - 5, size / 2 - 8, 10);
                        break;
                }
            }
        }

        public static void DrawFatArrow(Graphics g, int x, int y, int size, string direction, Brush fill, Pen border)
        {
            int cx = x + size / 2;
            int cy = y + size / 2;
            Point[] pts = new Point[] {
                new Point(cx, y + 5),
                new Point(x + size - 10, cy - 5),
                new Point(cx + 5, cy - 5),
                new Point(cx + 5, y + size - 10),
                new Point(cx - 5, y + size - 10),
                new Point(cx - 5, cy - 5),
                new Point(x + 10, cy - 5)
            };
            g.FillPolygon(fill, pts);
            g.DrawPolygon(border, pts);
        }

        private void pictureBox1_MouseClick(object sender, MouseEventArgs e)
        {
            int cols = 5;
            int rows = 6;
            int cellSize = 40;
            if (e.X < 0 || e.Y < 0 || e.X >= cols * cellSize || e.Y >= rows * cellSize)
                return; // Click outside grid
            int col = e.X / cellSize;
            int row = e.Y / cellSize;
            if (col < 0 || col >= cols || row < 0 || row >= rows)
                return; // Click outside grid
            int symbolIndex = row * cols + col;
            
            if (frmPiperLegend.dgvJobsInDetails.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow selectedRow in frmPiperLegend.dgvJobsInDetails.SelectedRows)
                {
                    if (selectedRow.Cells[1].Value != null)
                    {
                        for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                        {
                            if (frmImportSamples.WaterData[i].sampleID == selectedRow.Cells[1].Value.ToString())
                            {
                                frmImportSamples.WaterData[i].shape = symbolNames.ElementAt(symbolIndex);
                                break;
                            }
                        }
                    }
                }
            }
            this.Close();
        }
    }
}
