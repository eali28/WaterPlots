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
    public partial class frmPiperLegend : Form
    {
        private List<Image> shapeImages = new List<Image>();
        public static Color selectedColor;
        public static bool IsUpdateClicked;
        private static frmSymbolPicker openSymbolPicker = null;  // Track the open symbol picker form
        public frmPiperLegend()
        {
            InitializeComponent();
            Loaddgv();
            selectedColor = Color.Transparent;

            // Add a button column to the DataGridView
            //DataGridViewButtonColumn actionColumn = new DataGridViewButtonColumn();
            //actionColumn.Name = "Symbol";
            //actionColumn.HeaderText = "Symbol";
            //actionColumn.Text = "";
            //actionColumn.UseColumnTextForButtonValue = true;
            //dgvJobsInDetails.Columns.Add(actionColumn);

            //// Register the CellContentClick event handler
            //dgvJobsInDetails.CellContentClick += dgvJobsInDetails_CellContentClick;
        }


        private void shapeCombobox_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            e.ItemHeight = 350; // Increased height

        }
        public static void shape_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            e.DrawBackground();

            // Define colors for each shape
            Color[] shapeColors = {
                Color.Red,       // Circle
                Color.Blue,      // Cube
                Color.Green,     // Hexagon
                Color.Purple,    // Icosahedron
                Color.Brown,     // Merkaba
                Color.Magenta    // Triangle
            };

            // Get the corresponding color
            Color shapeColor = shapeColors[e.Index % shapeColors.Length];
            Brush shapeBrush = new SolidBrush(shapeColor);
            Pen shapePen = new Pen(Color.Black, 1); // Border color

            int shapeSize = 12;
            int centerX = e.Bounds.X + 15;
            int centerY = e.Bounds.Y + (e.Bounds.Height / 2) - (shapeSize / 2);

            Graphics g = e.Graphics;

            // Draw shapes manually instead of using images
            switch (e.Index)
            {
                case 0: // Circle
                    g.FillEllipse(shapeBrush, centerX, centerY, shapeSize, shapeSize);
                    g.DrawEllipse(shapePen, centerX, centerY, shapeSize, shapeSize);
                    break;
                case 1: // Cube (Square)
                    g.FillRectangle(shapeBrush, centerX, centerY, shapeSize, shapeSize);
                    g.DrawRectangle(shapePen, centerX, centerY, shapeSize, shapeSize);
                    break;
                case 2: // Hexagon
                    PointF[] hexagon = GetHexagonPoints(centerX, centerY, shapeSize);
                    g.FillPolygon(shapeBrush, hexagon);
                    g.DrawPolygon(shapePen, hexagon);
                    break;

                case 3: // Merkaba (Star Shape)
                    PointF[] star = GetStarPoints(centerX, centerY, shapeSize);
                    g.FillPolygon(shapeBrush, star);
                    g.DrawPolygon(shapePen, star);
                    break;
                case 4: // Triangle
                    PointF[] simpleTriangle = {
                        new PointF(centerX, centerY),
                        new PointF(centerX + shapeSize, centerY + shapeSize),
                        new PointF(centerX - shapeSize, centerY + shapeSize)
                    };
                    g.FillPolygon(shapeBrush, simpleTriangle);
                    g.DrawPolygon(shapePen, simpleTriangle);
                    break;
            }

            e.DrawFocusRectangle();
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
        private void Loaddgv()
        {
            for (int i = 0; i < frmImportSamples.copyOfJobs.Count; i++)
            {
                int rowIndex = dgvJobsInDetails.Rows.Add();
                dgvJobsInDetails.Rows[rowIndex].Cells[0].Value = frmImportSamples.copyOfJobs[i].jobID;
                dgvJobsInDetails.Rows[rowIndex].Cells[1].Value = frmImportSamples.copyOfJobs[i].sampleID;
                dgvJobsInDetails.Rows[rowIndex].Cells[2].Value = frmImportSamples.copyOfJobs[i].clientID;
                dgvJobsInDetails.Rows[rowIndex].Cells[3].Value = frmImportSamples.copyOfJobs[i].wellName;
                dgvJobsInDetails.Rows[rowIndex].Cells[4].Value = frmImportSamples.copyOfJobs[i].lat;
                dgvJobsInDetails.Rows[rowIndex].Cells[5].Value = frmImportSamples.copyOfJobs[i].Long;
                dgvJobsInDetails.Rows[rowIndex].Cells[6].Value = frmImportSamples.copyOfJobs[i].sampleType;
                dgvJobsInDetails.Rows[rowIndex].Cells[7].Value = frmImportSamples.copyOfJobs[i].formationName;
                dgvJobsInDetails.Rows[rowIndex].Cells[8].Value = frmImportSamples.copyOfJobs[i].depth;
                dgvJobsInDetails.Rows[rowIndex].Cells[9].Value = frmImportSamples.copyOfJobs[i].prep;
            }
        }
        private void colorPanel_Click(object sender, EventArgs e)
        {
            //this.Activate();  // Activate the form
            this.BringToFront();  // Bring the form to the front
            this.Focus();  // Ensure the form has focus

            // Show the color dialog
            //Form1.chart1.Paint -= PiperDrawer.DrawPiperDiagram; // Comment this out

            if (colorDialog1.ShowDialog(this) == DialogResult.OK)
            {
                selectedColor = colorDialog1.Color;
                colorPanel.BackColor = selectedColor;
            }

            //Form1.chart1.Paint += PiperDrawer.DrawPiperDiagram; // Comment this out
            this.Invalidate();  // Redraw the form if necessary
            if (dgvJobsInDetails.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in dgvJobsInDetails.SelectedRows) // Change 'DataGridView' to 'DataGridViewRow'
                {
                    if (row.Cells[1].Value != null) // Ensure the cell is not null
                    {
                        for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                        {
                            if (frmImportSamples.WaterData[i].sampleID == row.Cells[1].Value.ToString())
                            {
                                frmImportSamples.WaterData[i].color = selectedColor != Color.Transparent ? selectedColor : frmImportSamples.WaterData[i].color;
                            }
                        }
                    }
                }
            }
        }

        private void updateButton_Click(object sender, EventArgs e)
        {
            frmMainForm.selectedSamples.Clear();
            IsUpdateClicked = true;
            
            //if (dgvJobsInDetails.SelectedRows.Count > 0)
            //{
            //    foreach (DataGridViewRow row in dgvJobsInDetails.SelectedRows) // Change 'DataGridView' to 'DataGridViewRow'
            //    {
            //        if (row.Cells[1].Value != null) // Ensure the cell is not null
            //        {
            //            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
            //            {
            //                if (frmImportSamples.WaterData[i].sampleID == row.Cells[1].Value.ToString())
            //                {
            //                    frmImportSamples.WaterData[i].shape = frmSymbolPicker.selectedShape;
            //                    frmImportSamples.WaterData[i].color = selectedColor != Color.Transparent ? selectedColor : frmImportSamples.WaterData[i].color;
            //                }
            //            }
            //        }
            //    }
            //}
            frmMainForm.UpdatePiperDiagram();
            this.Close();
        }

        private void PiperDetails_Load(object sender, EventArgs e)
        {

        }

        private void colorPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void symbol_change_Click(object sender, EventArgs e)
        {
            // Check if the click is on the button column
            Brush brush=Brushes.Red;

            //if (e.ColumnIndex == dgvJobsInDetails.Columns["Symbol"].Index && e.RowIndex >= 0)
            //{
                //var row = dgvJobsInDetails.Rows[e.RowIndex];
                //for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                //{
                //    if (row.Cells[1].Value.ToString() == frmImportSamples.WaterData[i].sampleID)
                //    {
                //        brush = new SolidBrush(frmImportSamples.WaterData[i].color);
                //        break;
                //    }
                //}

                // Close existing symbol picker if open
                if (openSymbolPicker != null && !openSymbolPicker.IsDisposed)
                {
                    openSymbolPicker.Close();
                }

                openSymbolPicker = new frmSymbolPicker(new SolidBrush(selectedColor));
                openSymbolPicker.Show();
            //}
        }

        private void dgvJobsInDetails_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
