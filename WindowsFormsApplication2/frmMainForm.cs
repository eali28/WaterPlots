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
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms.DataVisualization.Charting;

namespace WindowsFormsApplication2
{
    public partial class frmMainForm : Form
    {

        public static int counter = 0;

        public RectangleF textBounds;
        public string text;
        
        public static bool flag = false;
        public static string connectionString = "Server=SQL-STRATOCHEM;Database=BRI;Integrated Security=True;";
        public static List<string> selectedSamples = new List<string>();
        public static Bitmap bmbpic;
        public static Panel legendPanel;
        public static Panel metaPanel;
        private bool isLoaded = false;

        public frmMainForm()
        {

            InitializeComponent();


        }
        private void Chart1_SizeChanged(object sender, EventArgs e)
        {
            legendPictureBox.Invalidate();
        }
        private void radioButtonWater_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonWater.Checked)
            {
                listBoxCharts.Items.Clear();
                listBoxCharts.Items.Add("Pie Chart");
                listBoxCharts.Items.Add("Piper Diagram");
                listBoxCharts.Items.Add("Collins Diagram");
                listBoxCharts.Items.Add("Stiff Diagram");
                listBoxCharts.Items.Add("Schoeller Diagram");
                listBoxCharts.Items.Add("Bubble Diagram");
                listBoxCharts.Items.Add("log Na Vs log Cl");
                listBoxCharts.Items.Add("log Mg Vs log Cl");
                listBoxCharts.Items.Add("log Ca Vs log Cl");
                listBoxCharts.Items.Add("Radar Diagram 1");
                listBoxCharts.Items.Add("Radar Diagram 2");
                listBoxCharts.Items.Add("ICP Reproducibility");
                listBoxCharts.Items.Add("Major Element Logs");
            }
        }
        private void listBoxCharts_DoubleClick(object sender, EventArgs e)
        {
            if (listBoxCharts.SelectedItem != null)
            {
                try
                {
                    string selectedChart = listBoxCharts.SelectedItem.ToString();
                    listBoxSelected.Visible = true;

                    // Add item if it doesn't already exist in the second listbox
                    if (!listBoxSelected.Items.Contains(selectedChart))
                    {
                        listBoxSelected.Items.Add(selectedChart);
                    }
                    else 
                    {
                        MessageBox.Show("Already exists!");
                    }
                    listBoxCharts.Refresh();
                    listBoxSelected.Refresh();
                }
                catch (Exception ex)
                {
                    // Handle any unexpected exceptions
                    MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listBoxCharts.Refresh();
            listBoxSelected.Refresh();
            if (isLoaded)
                return;

            // Initialization logic here
            isLoaded = true;
            if (frmImportSamples.WaterData != null && frmImportSamples.WaterData.Count > 0)
            {
                saveIcon.Image = Properties.Resources.saveActivated;
            }
            else
            {
                saveIcon.Image = global::WindowsFormsApplication2.Resources.saveIcon;
            }


        }
        private void buttonopenFile_Click(object sender, EventArgs e)
        {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Title = "Open SCSH File";
            openFileDialog.Filter = "SCSH Files (*.scsh)|*.scsh|All Files (*.*)|*.*";
            openFileDialog.DefaultExt = "SCSH";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                try
                {
                    // Read the file content
                    string[] fileLines = File.ReadAllLines(filePath);
                
                    frmImportSamples.WaterData.Clear(); // Clear existing data before loading

                    string selectedCompany = "";
                    string selectedJob = "";

                    bool readingLegend = false;
                    bool readingDataSection = false;
                    foreach (string line in fileLines)
                    {
                        if (line.StartsWith("Company Name:"))
                        {
                            selectedCompany = line.Substring(13).Trim();
                        }
                        else if (line.StartsWith("Job ID:"))
                        {
                            selectedJob = line.Substring(7).Trim();
                        }
                        else if (line.StartsWith("# Data"))
                        {
                            readingDataSection = true; // Start reading water data
                            readingLegend = false;
                            continue;
                        }
                        else if (line.StartsWith("# Legend"))
                        {
                            readingLegend = true;
                            continue;
                        }
                        else if (readingLegend && !string.IsNullOrWhiteSpace(line) && !line.StartsWith("Data"))
                        {

                            string[] parts = line.Split(',');
                            if (parts.Length == 4)
                            {
                                clsWater waterSample = new clsWater
                                {
                                    sampleID = parts[0].Trim(),
                                    Well_Name = parts[1].Trim(),
                                    ClientID = parts[2].Trim(),
                                    Depth = parts[3].Trim()
                                };
                                frmImportSamples.WaterData.Add(waterSample);
                            }
                        }
                        else if(line.StartsWith("# Pie Chart"))
                        {
                            readingDataSection=false;
                            break;
                        }
                        else if (readingDataSection && !string.IsNullOrWhiteSpace(line))
                        {
                            // Read chemical analysis data (Sample ID, Na, K, Ca, Mg, Cl, HCO3, CO3, SO4)
                            string[] parts = line.Split(new string[] { ",Na:", " K:", " Ca:", " Mg:", " CL:", " HCO3:", " CO3:", " SO4:", " Ba:", " Sr:", " B:","TDS:" }, StringSplitOptions.None);
                            string sampleID = parts[0].Replace("Sample ID:", "").Trim();
                            for (int i = 1; i < parts.Length; i++)
                            {
                                parts[i] = parts[i].Length > 2 ? parts[i].Substring(0, parts[i].Length - 1) : parts[i].Substring(0, parts[i].Length - 1);
                            }
                            if (parts.Length == 13)
                            {

                                clsWater waterSample = frmImportSamples.WaterData.FirstOrDefault(w => w.sampleID == sampleID);
                                if (waterSample != null)
                                {
                                    waterSample.Na = parts[1] != null ? Convert.ToDouble(parts[1]) : 0;
                                    waterSample.K = parts[2] != null ? Convert.ToDouble(parts[2]) : 0;
                                    waterSample.Ca = parts[3] != null ? Convert.ToDouble(parts[3]) : 0;
                                    waterSample.Mg = parts[4] != null ? Convert.ToDouble(parts[4]) : 0;
                                    waterSample.Cl = parts[5] != null ? Convert.ToDouble(parts[5]) : 0;
                                    waterSample.HCO3 = parts[6] != null ? Convert.ToDouble(parts[6]) : 0;
                                    waterSample.CO3 = parts[7] != null ? Convert.ToDouble(parts[7]) : 0;
                                    waterSample.So4 = parts[8] != null ? Convert.ToDouble(parts[8]) : 0;
                                    waterSample.Ba = parts[9] != null ? Convert.ToDouble(parts[9]) : 0;
                                    waterSample.Sr = parts[10] != null ? Convert.ToDouble(parts[10]) : 0;
                                    waterSample.B = parts[11] != null ? Convert.ToDouble(parts[11]) : 0;
                                    waterSample.TDS = parts[12] != null ? Convert.ToDouble(parts[12]) : 0;
                                    waterSample.color = frmImportSamples.GetRandomColor(false);

                                }
                            }
                        }
                    }

                    frmImportSamples.selectedCompany = selectedCompany;
                    frmImportSamples.JOBID = selectedJob;

                    frmImportSamples.isCalculateAndPlotClicked = true;
                    MessageBox.Show("File loaded successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            // Example: adjust picturebox size to match form
            mainChartPlotting.Width = this.ClientSize.Width;
            mainChartPlotting.Height = this.ClientSize.Height;
        }




        private void checkBoxSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            // Handle the "Select All" checkbox
            if (checkBoxSelectAll.Checked)
            {
                listBoxSelected.SelectionMode = SelectionMode.MultiSimple;
                for (int i = 0; i < listBoxSelected.Items.Count; i++)
                {
                    listBoxSelected.SetSelected(i, true);
                }
            }
            else
            {
                listBoxSelected.ClearSelected();
            }
            mainChartPlotting.Invalidate();
        }

        private void listBoxCharts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxCharts.SelectedItem != null)
            {
               
                listBoxCharts.Visible = true;
                string selectedChart = listBoxCharts.SelectedItem.ToString();
                listBoxCharts.Refresh();
                listBoxSelected.Refresh();
                legendPictureBox.MouseDoubleClick -= pictureBoxPiper_Click;
                legendPictureBox.MouseDoubleClick -= pictureBoxPie_Click;
                legendPictureBox.MouseDoubleClick -= pictureBoxSchoeller_Click;
                legendPictureBox.MouseDoubleClick -= pictureBoxCollins_Click;
                mainChartPlotting.MouseDoubleClick -= legendPictureBoxRadar;
                mainChartPlotting.MouseDoubleClick -= PictureBoxRadarScales;
                mainChartPlotting.Controls.Remove(legendPanel);
                mainChartPlotting.Controls.Remove(metaPanel);
                legendPanel = new Panel();
                metaPanel = new Panel();

                legendPictureBox = new PictureBox();
                
                bmbpic = new Bitmap(legendPanel.Width, legendPanel.Height);
                metaPanel.Controls.Clear();

                int screenWidth = Screen.PrimaryScreen.WorkingArea.Width;
                int screenHeight = Screen.PrimaryScreen.WorkingArea.Height;

                mainChartPlotting.Size = new Size((int)(this.Width*0.9), (int)(this.Height*0.7));
                mainChartPlotting.BackColor = Color.White;
                //Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Bitmap chartBitmap = new Bitmap(1728, 756);
                Graphics graphics = Graphics.FromImage(chartBitmap);

                // Clear the bitmap before drawing the new chart
                graphics.Clear(Color.White);

                if (selectedChart == "Pie Chart")
                {
                    clsPieDrawer.DrawPieChart(graphics, mainChartPlotting.Width, mainChartPlotting.Height);
                }
                else if (selectedChart == "Collins Diagram")
                {
                    clsCollinsDrawer.DrawCollinsDiagram(graphics, mainChartPlotting.Width, mainChartPlotting.Height);
                }
                else if (selectedChart == "Schoeller Diagram")
                {

                    clsSchoellerDrawer.DrawSchoellerDiagram(graphics);
                }
                else if (selectedChart == "Stiff Diagram")
                {
                    clsStiffDrawer.DrawStiffDiagram(graphics);
                }
                else if (selectedChart == "Bubble Diagram")
                {

                    clsBubbleDrawer.DrawBubbleDiagram(graphics);

                }
                else if (selectedChart == "Radar Diagram 1")
                {
                    mainChartPlotting.MouseDoubleClick += PictureBoxRadarScales;

                    Rectangle bounds = new Rectangle((int)(0.01f*mainChartPlotting.Width), (int)(0.08f*mainChartPlotting.Height), (int)(0.9 * mainChartPlotting.Width), (int)(0.9 * mainChartPlotting.Height));
                    clsRadarDrawer.DrawRadarChart1(graphics, bounds, flag);

                }
                else if (selectedChart == "Piper Diagram")
                {
                    clsPiperDrawer.DrawPiperDiagram(graphics);
                }
                else if (selectedChart == "log Na Vs log Cl")
                {
                    int diagramWidth = (int)(0.5f * mainChartPlotting.Width);
                    int diagramHeight = (int)(0.7f * mainChartPlotting.Height);
                    int x = (int)(0.03f * mainChartPlotting.Width);
                    int y = (mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * mainChartPlotting.Height);
                    clsLogsDrawer.DrawlogNa_VS_logCl(graphics,diagramWidth, diagramHeight, x, y);
                }
                else if (selectedChart == "log Mg Vs log Cl")
                {
                    int diagramWidth = (int)(0.5f * mainChartPlotting.Width);
                    int diagramHeight = (int)(0.7f * mainChartPlotting.Height);
                    int x = (int)(0.03f * mainChartPlotting.Width);
                    int y = (mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * mainChartPlotting.Height);
                    clsLogsDrawer.DrawlogMg_VS_logCl(graphics, diagramWidth, diagramHeight, x, y);
                }
                else if (selectedChart == "log Ca Vs log Cl")
                {
                    int diagramWidth = (int)(0.5f * mainChartPlotting.Width);
                    int diagramHeight = (int)(0.7f * mainChartPlotting.Height);
                    int x = (int)(0.03f * mainChartPlotting.Width);
                    int y = (mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * mainChartPlotting.Height);
                    clsLogsDrawer.DrawlogCa_VS_logCl(graphics, diagramWidth, diagramHeight, x, y);
                }
                else if (selectedChart == "Radar Diagram 2")
                {
                    mainChartPlotting.MouseDoubleClick += PictureBoxRadarScales;
                    Rectangle bounds = new Rectangle((int)(0.01f * mainChartPlotting.Width), (int)(0.02f * mainChartPlotting.Height), (int)(0.9 * mainChartPlotting.Width), (int)(0.9 * mainChartPlotting.Height));
                    clsRadarDrawer.DrawRadarChart2(graphics, bounds, flag);
                }
                else if (selectedChart == "Major Element Logs")
                {
                    clsLogsDrawer.DrawlogNa_VS_logCl(graphics, (int)(0.2 * mainChartPlotting.Width), (int)(0.3 * mainChartPlotting.Height), (int)(0.04 * mainChartPlotting.Width), (int)(0.05 * mainChartPlotting.Height));
                    clsLogsDrawer.DrawlogMg_VS_logCl(graphics, (int)(0.2 * mainChartPlotting.Width), (int)(0.3 * mainChartPlotting.Height), (int)(0.4 * mainChartPlotting.Width), (int)(0.05 * mainChartPlotting.Height));
                    clsLogsDrawer.DrawlogCa_VS_logCl(graphics, (int)(0.2 * mainChartPlotting.Width), (int)(0.3 * mainChartPlotting.Height), (int)(0.2 * mainChartPlotting.Width), (int)(0.5 * mainChartPlotting.Height));
                }
                else
                {
                    mainChartPlotting.MouseDoubleClick += PictureBoxRadarScales;
                    Rectangle bounds = new Rectangle((int)(0.01f * mainChartPlotting.Width), (int)(0.02f * mainChartPlotting.Height), (int)(0.9 * mainChartPlotting.Width), (int)(0.9 * mainChartPlotting.Height));
                    clsRadarDrawer.DrawRadarChart3(graphics, bounds, flag);
                }
                Bitmap resized = new Bitmap(chartBitmap, mainChartPlotting.Width, mainChartPlotting.Height);
                mainChartPlotting.Image = resized;
                listBoxCharts.Refresh();
                listBoxSelected.Refresh();
            }
        }


        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (listBoxSelected.SelectedItem != null)
            {
                var selectedItems = listBoxSelected.SelectedItems.Cast<object>().ToList();

                // Loop through and remove each item
                foreach (var item in selectedItems)
                {
                    listBoxSelected.Items.Remove(item);
                }
                listBoxSelected.Refresh();

            }
            else
            {
                MessageBox.Show("Please select an item to delete.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            checkBoxSelectAll.Checked = false;

        }



        private void listBoxSelected_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxSelected.SelectedItem != null)
            {
                // Get the selected chart type
                string selectedChart = listBoxSelected.SelectedItem.ToString();

                // Check if the chart already displays the same diagram
                if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == selectedChart)
                {
                    return; // Do nothing if the diagram is already displayed
                }


                mainChartPlotting.Size = new Size(1700, 1000);

                // Update listBoxCharts selection and configure the chart
                switch (selectedChart)
                {
                    case "Pie Chart":
                        listBoxCharts.SelectedIndex = 0;
                        break;
                    case "Collins Diagram":
                        listBoxCharts.SelectedIndex = 2;
                        break;
                    case "Schoeller Diagram":
                        listBoxCharts.SelectedIndex = 4;
                        break;
                    case "Stiff Diagram":
                        listBoxCharts.SelectedIndex = 3;
                        break;
                    case "Bubble Diagram":
                        listBoxCharts.SelectedIndex = 5;
                        break;
                    case "Radar Diagram 1":
                        listBoxCharts.SelectedIndex = 9;
                        break;
                    case "Piper Diagram":
                        listBoxCharts.SelectedIndex = 1;
                        break;
                    case "log Na Vs log Cl":
                        listBoxCharts.SelectedIndex = 6;
                        break;
                    case "log Mg Vs log Cl":
                        listBoxCharts.SelectedIndex = 7;
                        break;
                    case "log Ca Vs log Cl":
                        listBoxCharts.SelectedIndex = 8;
                        break;
                    case "Radar Diagram 2":
                        listBoxCharts.SelectedIndex = 10;
                        break;
                    case "ICP Reproducibility":
                        listBoxCharts.SelectedIndex = 11;
                        break;
                    case "Major Element Logs":
                        listBoxCharts.SelectedIndex = 12;
                        break;
                }

                // Refresh UI to reflect the changes
                listBoxCharts.Refresh();
                listBoxSelected.Refresh();
                mainChartPlotting.Invalidate();
            }
        }

        private void printPowerpoint_Click(object sender, EventArgs e)
        {
            string userName = Environment.UserName;
            string pptPath = string.Format(@"C:\Users\{0}\Documents\Diagrams.pptx", userName);
            PowerPoint.Application pptApplication = new PowerPoint.Application();
            PowerPoint.Presentation presentation;

            // Open existing PowerPoint if available, otherwise create a new one
            if (File.Exists(pptPath))
            {
                presentation = pptApplication.Presentations.Open(pptPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
            }
            else
            {
                presentation = pptApplication.Presentations.Add(Office.MsoTriState.msoTrue);
            }
            presentation.PageSetup.SlideWidth = (float)(10.84) * 72f;
            presentation.PageSetup.SlideHeight = (float)(7.5)*72f;
            for (int i = 0; i < listBoxSelected.Items.Count; i++)
            {
                string selectedChart = listBoxSelected.Items[i].ToString();

                // Add a new slide
                int newSlideIndex = presentation.Slides.Count + 1;
                PowerPoint.Slide slide = presentation.Slides.Add(newSlideIndex, PowerPoint.PpSlideLayout.ppLayoutBlank);
                // Calculate proper dimensions and positions for three charts
                float slideWidth = presentation.PageSetup.SlideWidth;
                float slideHeight = presentation.PageSetup.SlideHeight;
                if (listBoxSelected.Items[i].ToString() == "Collins Diagram")
                {
                    clsCollinsDrawer.ExportCollinsToPowerPoint(slide, presentation);
                }
                else if (listBoxSelected.Items[i].ToString() == "Pie Chart")
                {
                    clsPieDrawer.ExportPieChartToPowerPoint(slide, presentation);
                }
                else if (listBoxSelected.Items[i].ToString() == "Stiff Diagram")
                {
                    clsStiffDrawer.ExportStiffDiagramToPowerPoint(slide, presentation);
                }
                else if (listBoxSelected.Items[i].ToString() == "Bubble Diagram")
                {
                    clsBubbleDrawer.ExportBubbleDiagramToPowerPoint(slide, presentation);
                }
                else if (listBoxSelected.Items[i].ToString() == "log Na Vs log Cl")
                {
                    clsLogsDrawer.ExportLogNaVsLogClChartToPowerPoint(slide, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight,(int)(0.1f*slideWidth),(int)(0.1f*slideHeight));
                }
                else if (listBoxSelected.Items[i].ToString() == "log Mg Vs log Cl")
                {
                    clsLogsDrawer.ExportlogMgVslogCltoPowerpoint(slide, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight, (int)(0.1f * slideWidth), (int)(0.1f * slideHeight));
                }
                else if (listBoxSelected.Items[i].ToString() == "log Ca Vs log Cl")
                {
                    clsLogsDrawer.ExportlogCaVslogCltoPowerPoint(slide, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight, (int)(0.1f * slideWidth), (int)(0.1f * slideHeight));
                }
                else if (listBoxSelected.Items[i].ToString() == "Schoeller Diagram")
                {

                    clsSchoellerDrawer.ExportSchoellerDiagramToPowerPoint(slide, presentation);
                }
                else if (listBoxSelected.Items[i].ToString() == "Radar Diagram 1")
                {
                    Rectangle bounds = new Rectangle((int)(0.01f * mainChartPlotting.Width), (int)(0.08f * mainChartPlotting.Height), (int)(0.9 * mainChartPlotting.Width), (int)(0.9 * mainChartPlotting.Height));
                    clsRadarDrawer.ExportRadar1ToPowerpoint(bounds, slide, presentation,flag);
                }
                else if (listBoxSelected.Items[i].ToString() == "Radar Diagram 2")
                {
                    Rectangle bounds = new Rectangle((int)(0.01f * mainChartPlotting.Width), (int)(0.02f * mainChartPlotting.Height), (int)(0.9 * mainChartPlotting.Width), (int)(0.9 * mainChartPlotting.Height));
                    clsRadarDrawer.ExportRadar2ToPowerpoint(bounds, slide, presentation,flag);
                }
                else if (listBoxSelected.Items[i].ToString() == "Piper Diagram")
                {
                    clsPiperDrawer.ExportPiperDiagramToPowerpoint(slide, presentation);
                }
                else if (listBoxSelected.Items[i].ToString() == "Major Element Logs")
                {

                    int diagramWidth = 600, diagramHeight = 400;
                    // First chart (Log Na vs Log Cl) - Top left
                    clsLogsDrawer.ExportLogNaVsLogClChartToPowerPoint(slide, diagramWidth, diagramHeight, 200, 100);

                    // Second chart (Log Mg vs Log Cl) - Top right
                    clsLogsDrawer.ExportlogMgVslogCltoPowerpoint(slide, diagramWidth, diagramHeight, (int)(slideWidth * 0.6f), 100);

                    // Third chart (Log Ca vs Log Cl) - Bottom right
                    clsLogsDrawer.ExportlogCaVslogCltoPowerPoint(slide, diagramWidth, diagramHeight, (int)(slideWidth * 0.6f), (int)(slideHeight * 0.6f));
                }
                else
                {
                    Rectangle bounds = new Rectangle((int)(0.01f * presentation.PageSetup.SlideWidth), (int)(0.02f * presentation.PageSetup.SlideHeight), (int)(0.9 * presentation.PageSetup.SlideWidth), (int)(0.9 * presentation.PageSetup.SlideHeight));
                    clsRadarDrawer.ExportRadar3ToPowerpoint(bounds, slide, presentation, flag);
                }
            }
            listBoxCharts.Refresh();
            listBoxSelected.Refresh();
        }
        // Event handler to update chart scales based on user input from textboxes
        public static void UpdateScalesinRadar(string name)
        {
            // Check and update each scale based on the user input from corresponding textboxes
            if (name == "Radar Diagram 1")
            {

                for (int i = 0; i < clsRadarDrawer.Radar1Scales.Length; i++)
                {

                    switch (i)
                    {
                        case 0:
                            clsRadarDrawer.maxCl = clsRadarDrawer.Radar1Scales[i];
                                break;
                        case 1:
                                clsRadarDrawer.maxNa = clsRadarDrawer.Radar1Scales[i];
                                break;
                        case 2:
                                clsRadarDrawer.maxK = clsRadarDrawer.Radar1Scales[i];
                                break;
                        case 3:
                                clsRadarDrawer.maxCa = clsRadarDrawer.Radar1Scales[i];
                                break;
                        case 4:
                                clsRadarDrawer.maxMg = clsRadarDrawer.Radar1Scales[i];
                                break;
                        case 5:
                                clsRadarDrawer.maxBa = clsRadarDrawer.Radar1Scales[i];
                                break;
                        case 6:
                                clsRadarDrawer.maxSr = clsRadarDrawer.Radar1Scales[i];
                                break;

                    }
                    
                }



            }

            else if (name == "Radar Diagram 2")
            {
                for (int i = 0; i < clsRadarDrawer.Radar2Scales.Length; i++)
                {

                    switch (i)
                    {
                        case 0:
                            clsRadarDrawer.maxNaCl = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 1:
                            clsRadarDrawer.maxClCa = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 2:
                            clsRadarDrawer.maxHCO3Cl = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 3:
                            clsRadarDrawer.maxClSr = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 4:
                            clsRadarDrawer.maxNaCa = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 5:
                            clsRadarDrawer.maxKNa = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 6:
                            clsRadarDrawer.maxSrMg = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 7:
                            clsRadarDrawer.maxMgCl = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 8:
                            clsRadarDrawer.maxSrCl = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 9:
                            clsRadarDrawer.maxSrK = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 10:
                            clsRadarDrawer.maxMgK = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 11:
                            clsRadarDrawer.maxCaK = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 12:
                            clsRadarDrawer.maxtK = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 13:
                            clsRadarDrawer.maxBCl = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 14:
                            clsRadarDrawer.maxBNa = clsRadarDrawer.Radar2Scales[i];
                            break;
                        case 15:
                            clsRadarDrawer.maxBMg = clsRadarDrawer.Radar2Scales[i];
                            break;

                    }

                }

            }
            else
            {
                for (int i = 0; i < clsRadarDrawer.Radar3Scales.Length; i++)
                {

                    switch (i)
                    {
                        case 0:
                            clsRadarDrawer.maxNa = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 1:
                            clsRadarDrawer.maxK = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 2:
                            clsRadarDrawer.maxCa = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 3:
                            clsRadarDrawer.maxMg = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 4:
                            clsRadarDrawer.maxAl = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 5:
                            clsRadarDrawer.maxCo = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 6:
                            clsRadarDrawer.maxCu = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 7:
                            clsRadarDrawer.maxMn = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 8:
                            clsRadarDrawer.maxNi = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 9:
                            clsRadarDrawer.maxSr = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 10:
                            clsRadarDrawer.maxZn = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 11:
                            clsRadarDrawer.maxBa = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 12:
                            clsRadarDrawer.maxPb = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 13:
                            clsRadarDrawer.maxFe = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 14:
                            clsRadarDrawer.maxCd = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 15:
                            clsRadarDrawer.maxCr = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 16:
                            clsRadarDrawer.maxTl = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 17:
                            clsRadarDrawer.maxBe = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 18:
                            clsRadarDrawer.maxSe = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 19:
                            clsRadarDrawer.maxB = clsRadarDrawer.Radar3Scales[i];
                            break;
                        case 20:
                            clsRadarDrawer.maxLi = clsRadarDrawer.Radar3Scales[i];
                            break;
                    }

                }
            }

            flag = true;
            // After updating the values, refresh the radar chart with the updated scales

            Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
            Graphics graphics = Graphics.FromImage(chartBitmap);
            graphics.Clear(Color.White);
            if (name == "Radar Diagram 1")
            {
                Rectangle bounds = new Rectangle(50, 50, (int)(0.9 * mainChartPlotting.Width), (int)(0.9 * mainChartPlotting.Height));
                clsRadarDrawer.DrawRadarChart1(graphics, bounds, flag);

            }
            else if (name == "Radar Diagram 2")
            {
                Rectangle bounds = new Rectangle(50, 50, (int)(0.9 * mainChartPlotting.Width), (int)(0.9 * mainChartPlotting.Height));
                clsRadarDrawer.DrawRadarChart2(graphics, bounds, flag);

            }
            else if(name=="ICP Reproducibility")
            {
                Rectangle bounds = new Rectangle(50, 50, (int)(0.9 * mainChartPlotting.Width), (int)(0.9 * mainChartPlotting.Height));
                clsRadarDrawer.DrawRadarChart3(graphics, bounds, flag);
            }
            mainChartPlotting.Image = chartBitmap;
        }

        private void ButtonUpdateScales_Click(object sender, EventArgs e)
        {
            string selectedChart = listBoxCharts.SelectedItem.ToString();
            UpdateScalesinRadar(selectedChart);
        }

        public static void legendPictureBoxRadar(object sender, EventArgs e)
        {
            frmRadarLegend itemDetails = new frmRadarLegend();
            itemDetails.ShowDialog();
            itemDetails.BringToFront();
            itemDetails.Activate();
            if (frmRadarLegend.IsUpdateClicked)
            {
                UpdateRadarDiagram();
            }
        }
        public static void PictureBoxRadarScales(object sender, EventArgs e)
        {
            frmRadarScales itemDetails = new frmRadarScales();
            itemDetails.ShowDialog();
            itemDetails.BringToFront();
            itemDetails.Activate();
            if (frmRadarScales.IsUpdateClicked)
            {
                UpdateRadarDiagram();
            }
        }
        public static void pictureBoxSchoeller_Click(object sender, EventArgs e)
        {
            SchoellerDetails itemDetails = new SchoellerDetails();
            itemDetails.ShowDialog();
            itemDetails.BringToFront();
            itemDetails.Activate();
            if (SchoellerDetails.IsUpdateClicked)
            {
                UpdateSchoellerDiagram();
            }
        }

        public static void pictureBoxPie_Click(object sender, EventArgs e)
        {
            frmPieLegend itemDetails = new frmPieLegend();
            itemDetails.ShowDialog();
            itemDetails.BringToFront();
            itemDetails.Activate();
            UpdatePieDiagram();
        }
        
        public static void pictureBoxPiper_Click(object sender, EventArgs e)
        {
            frmPiperLegend itemDetails = new frmPiperLegend();
            itemDetails.ShowDialog();
            itemDetails.BringToFront();
            itemDetails.Activate();
               
            UpdatePiperDiagram();
        }

        public static void pictureBoxCollins_Click(object sender, EventArgs e)
        {
            frmCollinsLegend itemDetails = new frmCollinsLegend();
            itemDetails.ShowDialog();
            itemDetails.BringToFront();
            itemDetails.Activate();

            UpdateCollinsDiagram();
        }

        public static void UpdateRadarDiagram()
        {

            Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
            Graphics graphics = Graphics.FromImage(chartBitmap);
            graphics.Clear(Color.White);
            Rectangle bounds = new Rectangle(50, 50, (int)(0.9 * mainChartPlotting.Width), (int)(0.9 * mainChartPlotting.Height));
            if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == "Radar Diagram 1")
            {

                
                clsRadarDrawer.DrawRadarChart1(graphics, bounds, flag);
            }
            else if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == "Radar Diagram 2")
            {

                clsRadarDrawer.DrawRadarChart2(graphics, bounds, flag);
            }
            else if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == "ICP Reproducibility")
            {

                clsRadarDrawer.DrawRadarChart3(graphics, bounds, flag);
            }
            mainChartPlotting.Image = chartBitmap;

            
        }

        public static void UpdatePiperDiagram()
        {
            if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == "Piper Diagram")
            {

                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                clsPiperDrawer.DrawPiperDiagram(graphics);
                mainChartPlotting.Image = chartBitmap;
            }
        }
        public static void UpdatePieDiagram()
        {

            if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == "Pie Chart")
            {
                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                clsPieDrawer.DrawPieChart(graphics, mainChartPlotting.Width, mainChartPlotting.Height);
                mainChartPlotting.Image = chartBitmap;

            }
            
        }
        public static void UpdateSchoellerDiagram()
        {
            if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == "Schoeller Diagram")
            {
                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                clsSchoellerDrawer.DrawSchoellerDiagram(graphics);
                mainChartPlotting.Image = chartBitmap;
            }
        }
        public static void UpdateCollinsDiagram()
        {
            if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == "Collins Diagram")
            {
                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                clsCollinsDrawer.DrawCollinsDiagram(graphics, mainChartPlotting.Width, mainChartPlotting.Height);
                mainChartPlotting.Image = chartBitmap;
            }
        }
        public static void UpdateBubbleDiagram()
        {
            if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == "Bubble Diagram")
            {
                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                clsBubbleDrawer.DrawBubbleDiagram(graphics);
                mainChartPlotting.Image = chartBitmap;
            }
        }
        public static void UpdateLogsDiagram()
        {
            if (listBoxCharts.SelectedItem == "log Na Vs log Cl")
            {
                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                int diagramWidth = (int)(0.5f * mainChartPlotting.Width);
                int diagramHeight = (int)(0.7f * mainChartPlotting.Height);
                int x = (int)(0.03f * mainChartPlotting.Width);
                int y = (mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * mainChartPlotting.Height);
                clsLogsDrawer.DrawlogNa_VS_logCl(graphics, diagramWidth, diagramHeight, x, y);
                mainChartPlotting.Image = chartBitmap;
            }
            else if (listBoxCharts.SelectedItem == "log Mg Vs log Cl")
            {
                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                int diagramWidth = (int)(0.5f * mainChartPlotting.Width);
                int diagramHeight = (int)(0.7f * mainChartPlotting.Height);
                int x = (int)(0.03f * mainChartPlotting.Width);
                int y = (mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * mainChartPlotting.Height);
                clsLogsDrawer.DrawlogMg_VS_logCl(graphics, diagramWidth, diagramHeight, x, y);
                mainChartPlotting.Image = chartBitmap;
            }
            else if (listBoxCharts.SelectedItem == "log Ca Vs log Cl")
            {
                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                int diagramWidth = (int)(0.5f * mainChartPlotting.Width);
                int diagramHeight = (int)(0.7f * mainChartPlotting.Height);
                int x = (int)(0.03f * mainChartPlotting.Width);
                int y = (mainChartPlotting.Height - diagramHeight) / 2 - (int)(0.02 * mainChartPlotting.Height);
                clsLogsDrawer.DrawlogCa_VS_logCl(graphics, diagramWidth, diagramHeight, x, y);
                mainChartPlotting.Image = chartBitmap;
            }
            else if (listBoxCharts.SelectedItem == "Major Element Logs")
            {
                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                clsLogsDrawer.DrawlogNa_VS_logCl(graphics, (int)(0.2 * mainChartPlotting.Width), (int)(0.3 * mainChartPlotting.Height), (int)(0.02 * mainChartPlotting.Width), (int)(0.05 * mainChartPlotting.Height));
                clsLogsDrawer.DrawlogMg_VS_logCl(graphics, (int)(0.2 * mainChartPlotting.Width), (int)(0.3 * mainChartPlotting.Height), (int)(0.5 * mainChartPlotting.Width), (int)(0.05 * mainChartPlotting.Height));
                clsLogsDrawer.DrawlogCa_VS_logCl(graphics, (int)(0.2 * mainChartPlotting.Width), (int)(0.3 * mainChartPlotting.Height), (int)(0.5 * mainChartPlotting.Width), (int)(0.5 * mainChartPlotting.Height));
                mainChartPlotting.Image = chartBitmap;
            }
        }
        public static void UpdateStiffDiagram()
        {
            if (listBoxCharts.SelectedItem != null && listBoxCharts.SelectedItem.ToString() == "Stiff Diagram")
            {
                Bitmap chartBitmap = new Bitmap(mainChartPlotting.Width, mainChartPlotting.Height);
                Graphics graphics = Graphics.FromImage(chartBitmap);
                graphics.Clear(Color.White);
                clsStiffDrawer.DrawStiffDiagram(graphics);
                mainChartPlotting.Image = chartBitmap;
            }
        }


    }
}
