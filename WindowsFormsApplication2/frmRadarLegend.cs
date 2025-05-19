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
            dgvJobsInDetails.ColumnHeaderMouseClick += DgvJobsInDetails_ColumnHeaderMouseClick;
        }

        private void DgvJobsInDetails_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex == -1) // Only handle header clicks
            {
                string headerText = dgvJobsInDetails.Columns[e.ColumnIndex].HeaderText;

                // Check if header is already in the list
                if (!clsConstants.clickedHeaders.Contains(headerText))
                {
                    clsConstants.clickedHeaders.Add(headerText);
                    MessageBox.Show($"Header '{headerText}' has been added to the legend.", "Header Clicked");
                }
                else
                {
                    clsConstants.clickedHeaders.Remove(headerText);
                    MessageBox.Show($"Header '{headerText}' has been removed from the legend.", "Header Clicked");
                }
            }
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
            if(!frmMainForm.isExcelFileImported)
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
            else
            {
                dgvJobsInDetails.Rows.Clear();
                dgvJobsInDetails.Columns.Clear();
                List<string> headers = new List<string>();
                // Use headers from Excel (already in order in columnIndices)
                headers.Add("ID");
                headers.Add("Client ID");
                headers.Add("Sample Type");
                headers.Add("Label");
                headers.Add("K");
                headers.Add("Na");
                headers.Add("Mg");
                headers.Add("Ca");
                headers.Add("Al");
                headers.Add("Co");
                headers.Add("Cu");
                headers.Add("Mn");
                headers.Add("Ni");
                headers.Add("Sr");
                headers.Add("Zn");
                headers.Add("Ba");
                headers.Add("Pb");
                headers.Add("Fe");
                headers.Add("Cd");
                headers.Add("Cr");
                headers.Add("Tl");
                headers.Add("Be");
                headers.Add("Se");
                headers.Add("B");
                headers.Add("Li");
                headers.Add("Cl");

                if (headers == null || headers.Count == 0)
                {
                    MessageBox.Show("No headers found from WaterData.");
                    return;
                }

                // Create columns using header names
                foreach (var header in headers)
                {
                    dgvJobsInDetails.Columns.Add(header, header);
                }
                int ID = 1;
                // Populate rows manually from your WaterData
                foreach (var sample in frmImportSamples.WaterData)
                {
                    int rowIndex = dgvJobsInDetails.Rows.Add();
                    dgvJobsInDetails.Rows[rowIndex].Cells[0].Value = ID.ToString();
                    
                    dgvJobsInDetails.Rows[rowIndex].Cells[1].Value = sample.ClientID?.ToString() ?? "";
                    dgvJobsInDetails.Rows[rowIndex].Cells[2].Value = sample.sampleType?.ToString() ?? "";
                    dgvJobsInDetails.Rows[rowIndex].Cells[3].Value = sample.Label?.ToString() ?? "";
                    dgvJobsInDetails.Rows[rowIndex].Cells[4].Value = sample.K.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[5].Value = sample.Na.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[6].Value = sample.Mg.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[7].Value = sample.Ca.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[8].Value = sample.Al.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[9].Value = sample.Co.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[10].Value = sample.Cu.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[11].Value = sample.Mn.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[12].Value = sample.Ni.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[13].Value = sample.Sr.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[14].Value = sample.Zn.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[15].Value = sample.Ba.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[16].Value = sample.Pb.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[17].Value = sample.Fe.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[18].Value = sample.Cd.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[19].Value = sample.Cr.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[20].Value = sample.Tl.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[21].Value = sample.Be.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[22].Value = sample.Se.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[23].Value = sample.B.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[24].Value = sample.Li.ToString();
                    dgvJobsInDetails.Rows[rowIndex].Cells[25].Value = sample.Cl.ToString();
                    ID++;
                }
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
                            if (frmImportSamples.WaterData[i].sampleID == row.Cells[1].Value.ToString() || frmImportSamples.WaterData[i].ID == row.Cells[0].Value.ToString())
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
                                if (frmImportSamples.WaterData[i].sampleID == row.Cells[1].Value.ToString() || frmImportSamples.WaterData[i].ID == row.Cells[0].Value.ToString())
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
                                if (frmImportSamples.WaterData[i].sampleID == row.Cells[1].Value.ToString() || frmImportSamples.WaterData[i].ID == row.Cells[0].Value.ToString())
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
