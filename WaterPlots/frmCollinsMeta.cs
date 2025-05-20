using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WaterPlots
{
    public partial class frmCollinsMeta : Form
    {
        public static bool isUpdateClicked;
        public frmCollinsMeta(string diagramName)
        {
            this.Text = diagramName;
            InitializeComponent();
            Loaddgv();
            isUpdateClicked = false;
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

        private void updateButton_Click(object sender, EventArgs e)
        {
            isUpdateClicked = true;

            this.Close();
        }

        private void updateButton_Click_1(object sender, EventArgs e)
        {
            isUpdateClicked = true;

            this.Close();
        }
    }
}
