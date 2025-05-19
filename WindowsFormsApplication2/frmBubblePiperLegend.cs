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
    public partial class frmBubblePiperLegend : Form
    {
        private List<Image> shapeImages = new List<Image>();
        public static Color selectedColor;
        public static bool IsUpdateClicked;
        private static frmSymbolPicker openSymbolPicker = null;  // Track the open symbol picker form
        

        public frmBubblePiperLegend(string diagramName)
        {
            InitializeComponent();
            Loaddgv();
            selectedColor = Color.Transparent;
            this.Text = diagramName;
            IsUpdateClicked = false;
            // Add ColumnHeaderMouseClick event handler
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

        private void Loaddgv()
        {
            if (!frmMainForm.isExcelFileImported)
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
                            if (frmImportSamples.WaterData[i].sampleID == row.Cells[1].Value.ToString() || frmImportSamples.WaterData[i].ID == row.Cells[0].Value.ToString())
                            {
                                frmImportSamples.WaterData[i].color = selectedColor != Color.Transparent ? selectedColor : frmImportSamples.WaterData[i].color;
                                if(this.Text=="Piper Legend")
                                {
                                    frmImportSamples.WaterData[i].piper = true;
                                }
                                else
                                {
                                    frmImportSamples.WaterData[i].bubble = true;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void updateButton_Click(object sender, EventArgs e)
        {
            //frmMainForm.selectedSamples.Clear();
            IsUpdateClicked = true;
            //if (this.Text == "Piper Legend")
            //{
            //    frmMainForm.UpdatePiperDiagram();
            //}
            //else if (this.Text == "Bubble Legend")
            //{
            //    frmMainForm.UpdateBubbleDiagram();
            //}
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
