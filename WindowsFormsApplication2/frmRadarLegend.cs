using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Collections.Generic;

using System.Linq;
using System.Text;

using System.IO;

namespace WindowsFormsApplication2
{
    public partial class frmRadarLegend : Form
    {
        public static Color selectedColor;
        public static DashStyle selectedStyle;
        public static float lineWidth; // Default line width
        public static int index;
        public static bool IsUpdateClicked { get; private set; }
        private List<Image> shapeImages = new List<Image>();
        public static bool typeComboboxClicked=false;
        public static string selectedShape;

        public frmRadarLegend()
        {
            InitializeComponent();
            LoadTypesIntoComboBox();
            Loaddgv();
            selectedColor = Color.Transparent;
            colorPanel.BackColor = selectedColor;
            colorPanel.Refresh();
            colorPanel.Enabled = true;
            colorPanel.Visible = true;
            IsUpdateClicked = false;

            selectedStyle = DashStyle.Solid;
            lineWidth = 2;
        }
        

        private void LoadTypesIntoComboBox()
        {
            // Clear previous items
            typeCombobox.Items.Clear();

            // Add DashStyles
            typeCombobox.Items.Add(DashStyle.Solid);
            typeCombobox.Items.Add(DashStyle.Dash);
            typeCombobox.Items.Add(DashStyle.Dot);
            typeCombobox.Items.Add(DashStyle.DashDot);
            typeCombobox.Items.Add(DashStyle.DashDotDot);

        }
        private void shapeCombobox_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            e.ItemHeight = 350; // Increased height
            
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

            // Load scales into ScalesDataGridView
        }

        

        

        private void typeCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (typeCombobox.SelectedItem != null)
            {
                typeComboboxClicked = true;
                selectedStyle = (DashStyle)typeCombobox.SelectedItem;
                foreach (DataGridViewRow row in dgvJobsInDetails.SelectedRows) // Change 'DataGridView' to 'DataGridViewRow'
                {
                    if (row.Cells[1].Value != null) // Ensure the cell is not null
                    {
                        for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                        {
                            if (frmImportSamples.WaterData[i].sampleID == row.Cells[1].Value.ToString())
                            {
                                frmImportSamples.WaterData[i].selectedStyle = selectedStyle;
                            }
                        }
                    }
                }
                this.Invalidate(); // Redraw the form to update the line style
            }
        }

        private void typeCombobox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            e.DrawBackground();

            using (Pen pen = new Pen(Color.Black, 2))
            {
                pen.DashStyle = (DashStyle)typeCombobox.Items[e.Index]; // Get style from combobox item
                if (selectedColor != Color.Transparent)
                {
                    pen.Color = selectedColor;
                }
                int xStart = e.Bounds.Left + 5;
                int xEnd = e.Bounds.Right - 5;
                int y = e.Bounds.Top + (e.Bounds.Height / 2); // Center the line vertically

                e.Graphics.DrawLine(pen, xStart, y, xEnd, y); // Draw line with style
            }

            e.DrawFocusRectangle();
        }


        

        private void updateButton_Click(object sender, EventArgs e)
        {
            frmMainForm.selectedSamples.Clear();
            IsUpdateClicked = true;
            this.Close();
        }

        private void widthTextBox_TextChanged(object sender, EventArgs e)
        {
            float newWidth;

            if (float.TryParse(widthTextBox.Text, out newWidth)) // Correct TryParse usage
            {
                if (newWidth > 0) // Ensure width is positive
                {
                    lineWidth = newWidth;
                    foreach (DataGridViewRow row in dgvJobsInDetails.SelectedRows) // Change 'DataGridView' to 'DataGridViewRow'
                    {
                        if (row.Cells[1].Value != null) // Ensure the cell is not null
                        {
                            for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                            {
                                if (frmImportSamples.WaterData[i].sampleID == row.Cells[1].Value.ToString())
                                {
                                    frmImportSamples.WaterData[i].lineWidth = lineWidth;
                                }
                            }
                        }
                    }
                    this.Invalidate(); // Redraw the line
                }
            }

        }



        private void Item_Details_Load(object sender, EventArgs e)
        {

        }
        private void colorPanel_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                // Ensure the form is activated and brought to the front
                this.Activate();
                this.BringToFront();

                // Show the color dialog
                
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {

                    // Update the selected color and the panel's background color
                    selectedColor = colorDialog.Color;
                    colorPanel.BackColor = selectedColor;

                    // Redraw the form to update the line
                    this.Invalidate();
                }
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
        }

        private void colorPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void ScalesDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }






    }
}
