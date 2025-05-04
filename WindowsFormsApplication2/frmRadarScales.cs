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

    public partial class frmRadarScales : Form
    {
        public string[] labelsRadar1 =
            {
                "Cl (mol/L)",
                "Na (mol/L)",
                "K (mol/L)",
                "Ca (mol/L)",
                "Mg (mol/L)",
                "Ba (mol/L)",
                "Sr (mol/L)"
            };
        public string[] labelsRadar2 =
            {
            "EV_Na-Cl",
            "EV_Cl-Ca",
            "EV_HCO3-Cl",
            "EV_Cl-Sr",
            "EV_Na-Ca",
            "GT_K-Na",
            "SS_Sr-Mg",
            "SS_Mg-Cl",
            "SS_Sr-Cl",
            "Lith_Sr-K",
            "Lith_Mg-K",
            "Lith_Ca-K",
            "Wt%K",
            "OM_B-Cl",
            "OM_B-Na",
            "OM_B-Mg"
            

        };
        string[] labelsRadar3 =
            {
            "Na",
            "K",
            "Ca",
            "Mg",
            "Al",
            "Co",
            "Cu",
            "Mn",
            "Ni",
            "Sr",
            "Zn",
            "Ba",
            "Pb",
            "Fe",
            "Cd",
            "Cr",
            "Tl",
            "Be",
            "Se",
            "B",
            "Li"
        };
        public static bool isScaleChanged { get; private set; }
        public static bool IsUpdateClicked { get; private set; }
        public frmRadarScales()
        {
            InitializeComponent();
            isScaleChanged = false;
            LoadScalesGrid();
        }
        private void ScalesDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1 && e.RowIndex >= 0) // Check if it's the Scale column
            {
                try
                {
                    string value = ScalesDatagridView.Rows[e.RowIndex].Cells[1].Value.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        double scaleValue = double.Parse(value);

                        // Update the appropriate scale array based on the row index
                        if (frmMainForm.listBoxCharts.SelectedItem == "Radar Diagram 1") // Radar1 scales
                        {
                            clsRadarDrawer.Radar1Scales[e.RowIndex] = scaleValue;
                        }
                        else if (frmMainForm.listBoxCharts.SelectedItem == "Radar Diagram 2")// Radar2 scales
                        {
                            clsRadarDrawer.Radar2Scales[e.RowIndex] = scaleValue;
                        }
                        else
                        {
                            clsRadarDrawer.Radar3Scales[e.RowIndex] = scaleValue;
                        }

                        // Refresh the radar chart
                        frmMainForm.mainChartPlotting.Invalidate();
                        isScaleChanged = true;
                    }
                }
                catch (FormatException)
                {
                    MessageBox.Show("Please enter a valid number for the scale value.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }
        private void LoadScalesGrid()
        {
            ScalesDatagridView.Rows.Clear();
            ScalesDatagridView.Columns.Clear();

            // Add columns
            ScalesDatagridView.Columns.Add("Parameter", "Parameter");
            ScalesDatagridView.Columns.Add("Scale", "Scale");

            // Add rows for Radar1
            if (frmMainForm.listBoxCharts.SelectedItem == "Radar Diagram 1")
            {
                for (int i = 0; i < labelsRadar1.Length; i++)
                {
                    ScalesDatagridView.Rows.Add(labelsRadar1[i], clsRadarDrawer.Radar1Scales[i].ToString("F5"));
                }
            }
            else if (frmMainForm.listBoxCharts.SelectedItem == "Radar Diagram 2")
            {
                for (int i = 0; i < labelsRadar2.Length; i++)
                {
                    ScalesDatagridView.Rows.Add(labelsRadar2[i], clsRadarDrawer.Radar2Scales[i].ToString("F5"));
                }
            }
            else
            {
                for (int i = 0; i < labelsRadar3.Length; i++)
                {
                    ScalesDatagridView.Rows.Add(labelsRadar3[i], clsRadarDrawer.Radar3Scales[i].ToString("F5"));
                }
            }
            int rowHeight = ScalesDatagridView.RowTemplate.Height;
            int headerHeight = ScalesDatagridView.ColumnHeadersHeight;
            int rowCount = ScalesDatagridView.Rows.Count;

            int totalHeight = headerHeight + (rowHeight * rowCount);

            ScalesDatagridView.Height = totalHeight;
            // Attach the CellValueChanged event handler
            ScalesDatagridView.CellValueChanged += ScalesDataGridView_CellValueChanged;
        }
        private void updateButton_Click(object sender, EventArgs e)
        {
            if (isScaleChanged)
            {
                frmMainForm.UpdateScalesinRadar(frmMainForm.listBoxCharts.SelectedItem.ToString());
                IsUpdateClicked = true;
            }

            this.Close();
        }
    }
}
