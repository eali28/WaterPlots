using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;
namespace WindowsFormsApplication2
{
    partial class frmMainForm
    {
        private System.Windows.Forms.CheckBox checkBoxSelectAll;
        private System.Windows.Forms.RadioButton radioButtonWater;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton buttonOpenProject;
        private System.Windows.Forms.ToolStripButton buttonImport;
        private System.Windows.Forms.Button buttonDelete;



        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.buttonOpenProject = new System.Windows.Forms.ToolStripButton();
            saveIcon = new System.Windows.Forms.ToolStripButton();
            this.buttonImport = new System.Windows.Forms.ToolStripButton();
            this.radioButtonWater = new System.Windows.Forms.RadioButton();
            listBoxCharts = new System.Windows.Forms.ListBox();
            this.listBoxSelected = new System.Windows.Forms.ListBox();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.exportMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.printPowerpoint = new System.Windows.Forms.ToolStripMenuItem();
            this.fileMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.openMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.importFromDbMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importFromExcelMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.checkBoxSelectAll = new System.Windows.Forms.CheckBox();
            this.buttonDelete = new System.Windows.Forms.Button();
            mainChartPlotting = new System.Windows.Forms.PictureBox();
            legendPictureBox = new System.Windows.Forms.PictureBox();
            this.toolStrip1.SuspendLayout();
            this.menuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(mainChartPlotting)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(legendPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(70, 70);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.buttonOpenProject,
            saveIcon,
            this.buttonImport});
            this.toolStrip1.Location = new System.Drawing.Point(0, 24);
            this.toolStrip1.MaximumSize = new System.Drawing.Size(2000, 70);
            this.toolStrip1.MinimumSize = new System.Drawing.Size(2000, 70);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(2000, 70);
            this.toolStrip1.TabIndex = 2;
            this.toolStrip1.Text = "Main Toolbar";
            this.toolStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.toolStrip1_ItemClicked);
            // 
            // buttonOpenProject
            // 
            this.buttonOpenProject.Image = global::WindowsFormsApplication2.Resources.openProject;
            this.buttonOpenProject.Name = "buttonOpenProject";
            this.buttonOpenProject.Size = new System.Drawing.Size(74, 67);
            this.buttonOpenProject.ToolTipText = "Open Project";
            this.buttonOpenProject.Click += new System.EventHandler(this.buttonopenFile_Click);
            // 
            // saveIcon
            // 
            saveIcon.Image = global::WindowsFormsApplication2.Resources.saveIcon;
            saveIcon.Name = "saveIcon";
            saveIcon.Size = new System.Drawing.Size(74, 67);
            saveIcon.ToolTipText = "Save";
            saveIcon.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // buttonImport
            // 
            this.buttonImport.Image = global::WindowsFormsApplication2.Resources.importfromdatabase;
            this.buttonImport.Name = "buttonImport";
            this.buttonImport.Size = new System.Drawing.Size(74, 67);
            this.buttonImport.ToolTipText = "Import From Database";
            this.buttonImport.Click += new System.EventHandler(this.buttonImport_Click);
            // 
            // radioButtonWater
            // 
            this.radioButtonWater.AutoSize = true;
            this.radioButtonWater.Location = new System.Drawing.Point(20, 114);
            this.radioButtonWater.Name = "radioButtonWater";
            this.radioButtonWater.Size = new System.Drawing.Size(54, 17);
            this.radioButtonWater.TabIndex = 3;
            this.radioButtonWater.Text = "Water";
            this.radioButtonWater.CheckedChanged += new System.EventHandler(this.radioButtonWater_CheckedChanged);
            // 
            // listBoxCharts
            // 
            listBoxCharts.FormattingEnabled = true;
            listBoxCharts.Location = new System.Drawing.Point(12, 137);
            listBoxCharts.Name = "listBoxCharts";
            listBoxCharts.Size = new System.Drawing.Size(165, 199);
            listBoxCharts.TabIndex = 4;
            listBoxCharts.SelectedIndexChanged += new System.EventHandler(this.listBoxCharts_SelectedIndexChanged);
            listBoxCharts.DoubleClick += new System.EventHandler(this.listBoxCharts_DoubleClick);
            // 
            // listBoxSelected
            // 
            this.listBoxSelected.FormattingEnabled = true;
            this.listBoxSelected.Location = new System.Drawing.Point(12, 397);
            this.listBoxSelected.Name = "listBoxSelected";
            this.listBoxSelected.Size = new System.Drawing.Size(165, 199);
            this.listBoxSelected.TabIndex = 5;
            this.listBoxSelected.SelectedIndexChanged += new System.EventHandler(this.listBoxSelected_SelectedIndexChanged);
            // 
            // menuStrip
            // 
            this.menuStrip.Dock = System.Windows.Forms.DockStyle.None;
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exportMenu,
            this.fileMenu});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.MaximumSize = new System.Drawing.Size(2000, 0);
            this.menuStrip.MinimumSize = new System.Drawing.Size(2000, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(2000, 24);
            this.menuStrip.TabIndex = 0;
            // 
            // exportMenu
            // 
            this.exportMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.printPowerpoint});
            this.exportMenu.Name = "exportMenu";
            this.exportMenu.Size = new System.Drawing.Size(53, 20);
            this.exportMenu.Text = "Export";
            // 
            // printPowerpoint
            // 
            this.printPowerpoint.Name = "printPowerpoint";
            this.printPowerpoint.Size = new System.Drawing.Size(163, 22);
            this.printPowerpoint.Text = "Print Powerpoint";
            this.printPowerpoint.Click += new System.EventHandler(this.printPowerpoint_Click);
            // 
            // fileMenu
            // 
            this.fileMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openMenuItem,
            this.saveMenuItem,
            this.saveAsMenuItem,
            this.importMenu});
            this.fileMenu.Name = "fileMenu";
            this.fileMenu.Size = new System.Drawing.Size(37, 20);
            this.fileMenu.Text = "File";
            // 
            // openMenuItem
            // 
            this.openMenuItem.Name = "openMenuItem";
            this.openMenuItem.Size = new System.Drawing.Size(180, 22);
            this.openMenuItem.Text = "Open";
            // 
            // saveMenuItem
            // 
            this.saveMenuItem.Name = "saveMenuItem";
            this.saveMenuItem.Size = new System.Drawing.Size(180, 22);
            this.saveMenuItem.Text = "Save";
            // 
            // saveAsMenuItem
            // 
            this.saveAsMenuItem.Name = "saveAsMenuItem";
            this.saveAsMenuItem.Size = new System.Drawing.Size(180, 22);
            this.saveAsMenuItem.Text = "Save As";
            // 
            // importMenu
            // 
            this.importMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importFromDbMenuItem,
            this.importFromExcelMenuItem});
            this.importMenu.Name = "importMenu";
            this.importMenu.Size = new System.Drawing.Size(180, 22);
            this.importMenu.Text = "Import";
            // 
            // importFromDbMenuItem
            // 
            this.importFromDbMenuItem.Name = "importFromDbMenuItem";
            this.importFromDbMenuItem.Size = new System.Drawing.Size(192, 22);
            this.importFromDbMenuItem.Text = "Import From Database";
            // 
            // importFromExcelMenuItem
            // 
            this.importFromExcelMenuItem.Name = "importFromExcelMenuItem";
            this.importFromExcelMenuItem.Size = new System.Drawing.Size(192, 22);
            this.importFromExcelMenuItem.Text = "Import From Excel";
            this.importFromExcelMenuItem.Click += new System.EventHandler(this.importFromExcelMenuItem_Click);
            // 
            // checkBoxSelectAll
            // 
            this.checkBoxSelectAll.Location = new System.Drawing.Point(0, 352);
            this.checkBoxSelectAll.Name = "checkBoxSelectAll";
            this.checkBoxSelectAll.Size = new System.Drawing.Size(70, 20);
            this.checkBoxSelectAll.TabIndex = 0;
            this.checkBoxSelectAll.Text = "Select All";
            this.checkBoxSelectAll.CheckedChanged += new System.EventHandler(this.checkBoxSelectAll_CheckedChanged);
            // 
            // buttonDelete
            // 
            this.buttonDelete.Location = new System.Drawing.Point(100, 352);
            this.buttonDelete.Name = "buttonDelete";
            this.buttonDelete.Size = new System.Drawing.Size(70, 30);
            this.buttonDelete.TabIndex = 0;
            this.buttonDelete.Text = "Delete";
            this.buttonDelete.UseVisualStyleBackColor = true;
            this.buttonDelete.Click += new System.EventHandler(this.buttonDelete_Click);
            // 
            // mainChartPlotting
            // 
            mainChartPlotting.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            mainChartPlotting.Location = new System.Drawing.Point(177, 137);
            mainChartPlotting.Name = "mainChartPlotting";
            mainChartPlotting.Size = new System.Drawing.Size(100, 50);
            mainChartPlotting.TabIndex = 0;
            mainChartPlotting.TabStop = false;
            mainChartPlotting.Text = "chart1";
            // 
            // legendPictureBox
            // 
            legendPictureBox.Location = new System.Drawing.Point(0, 0);
            legendPictureBox.Name = "legendPictureBox";
            legendPictureBox.Size = new System.Drawing.Size(100, 50);
            legendPictureBox.TabIndex = 0;
            legendPictureBox.TabStop = false;
            // 
            // frmMainForm
            // 
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(235)))), ((int)(((byte)(250)))));
            this.ClientSize = new System.Drawing.Size(1453, 710);
            this.Controls.Add(mainChartPlotting);
            this.Controls.Add(this.buttonDelete);
            this.Controls.Add(this.checkBoxSelectAll);
            this.Controls.Add(this.menuStrip);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.radioButtonWater);
            this.Controls.Add(listBoxCharts);
            this.Controls.Add(this.listBoxSelected);
            this.MainMenuStrip = this.menuStrip;
            this.Name = "frmMainForm";
            this.Text = "Water Plots";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Activated += new System.EventHandler(this.Form1_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(mainChartPlotting)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(legendPictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void buttonImport_Click(object sender, EventArgs e)
        {
            frmImportSamples getDBForm = new frmImportSamples();
            getDBForm.ShowDialog(); // Use ShowDialog() if you want it as a modal form
            if (frmImportSamples.isCalculateAndPlotClicked)
            {
                // Trigger radar diagram update immediately
                frmMainForm.flag = false;
                //clsRadarDrawer.maxAl = 0; clsRadarDrawer.maxCo = 0; clsRadarDrawer.maxCu = 0; clsRadarDrawer.maxMn = 0; clsRadarDrawer.maxNi = 0; clsRadarDrawer.maxZn = 0; clsRadarDrawer.maxPb = 0; clsRadarDrawer.maxFe = 0; clsRadarDrawer.maxCd = 0; clsRadarDrawer.maxCr = 0; clsRadarDrawer.maxTl = 0; clsRadarDrawer.maxBe = 0; clsRadarDrawer.maxSe = 0; clsRadarDrawer.maxLi = 0; clsRadarDrawer.maxB = 0;
                //clsRadarDrawer.maxNaCl = 0; clsRadarDrawer.maxClCa = 0; clsRadarDrawer.maxHCO3Cl = 0; clsRadarDrawer.maxClSr = 0; clsRadarDrawer.maxNaCa = 0; clsRadarDrawer.maxKNa = 0; clsRadarDrawer.maxSrMg = 0; clsRadarDrawer.maxMgCl = 0; clsRadarDrawer.maxSrCl = 0; clsRadarDrawer.maxSrK = 0; clsRadarDrawer.maxMgK = 0; clsRadarDrawer.maxCaK = 0; clsRadarDrawer.maxtK = 0; clsRadarDrawer.maxBCl = 0; clsRadarDrawer.maxBNa = 0; clsRadarDrawer.maxBMg = 0;
                //clsRadarDrawer.maxCl = 0; clsRadarDrawer.maxNa1 = 0; clsRadarDrawer.maxK1 = 0; clsRadarDrawer.maxCa1 = 0; clsRadarDrawer.maxMg1 = 0; clsRadarDrawer.maxBa1 = 0; clsRadarDrawer.maxSr1 = 0;
                //clsRadarDrawer.maxNa3 = 0; clsRadarDrawer.maxK3 = 0; clsRadarDrawer.maxCa3 = 0; clsRadarDrawer.maxMg3 = 0; clsRadarDrawer.maxBa3 = 0; clsRadarDrawer.maxSr3 = 0;
                UpdateRadarDiagram();
                UpdateCollinsDiagram();
                UpdatePieDiagram();
                UpdatePiperDiagram();
                UpdateSchoellerDiagram();
                UpdateLogsDiagram();
                UpdateStiffDiagram();
                UpdateBubbleDiagram();
                if (listBoxCharts.SelectedItem != null)
                {
                    UpdateScalesinRadar(listBoxCharts.SelectedItem.ToString());
                }
                
            }

        }
        private void buttonSave_Click(object sender, EventArgs e)
        {
            if (frmImportSamples.WaterData.Count > 0)
            {

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Title = "Save SCSH File";
                    saveFileDialog.Filter = "SCSH Files (*.scsh)|*.scsh|All Files (*.*)|*.*";
                    saveFileDialog.DefaultExt = "SCSH";
                    saveFileDialog.AddExtension = true;
                    saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                    // Show the dialog and get result
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;

                        try
                        {
                            // Write content to the selected file
                            using (StreamWriter writer = new StreamWriter(filePath))
                            {
                                writer.WriteLine(string.Format(@"Company Name: {0}", frmImportSamples.selectedCompany));
                                writer.WriteLine(string.Format(@"Job ID: {0}", frmImportSamples.selectedJob));
                                writer.WriteLine("# Legend");
                                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                                {
                                    writer.WriteLine(string.Format(@"{0},{1},{2},{3}", frmImportSamples.WaterData[i].sampleID, frmImportSamples.WaterData[i].Well_Name, frmImportSamples.WaterData[i].ClientID, frmImportSamples.WaterData[i].Depth));
                                }
                                writer.WriteLine("# Data");
                                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                                {
                                    writer.WriteLine(string.Format(@"Sample ID: {0},Na: {1}, K: {2}, Ca: {3}, Mg: {4}, CL: {5}, HCO3: {6}, CO3: {7}, SO4: {8}, Ba: {9}, Sr: {10}, B: {11}, TDS: {12}", frmImportSamples.WaterData[i].sampleID, frmImportSamples.WaterData[i].Na, frmImportSamples.WaterData[i].K, frmImportSamples.WaterData[i].Ca, frmImportSamples.WaterData[i].Mg, frmImportSamples.WaterData[i].Cl, frmImportSamples.WaterData[i].HCO3, frmImportSamples.WaterData[i].CO3, frmImportSamples.WaterData[i].So4,frmImportSamples.WaterData[i].Ba,frmImportSamples.WaterData[i].Sr,frmImportSamples.WaterData[i].B,frmImportSamples.WaterData[i].TDS));
                                
                                }
                                writer.WriteLine("");
                                writer.WriteLine("# Pie Chart");
                                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                                {
                                    writer.WriteLine(string.Format(@"Sample ID: {0},Na: {1}, K: {2}, Ca: {3}, Mg: {4}, CL: {5}, HCO3: {6}, CO3: {7}, SO4: {8}", frmImportSamples.WaterData[i].sampleID,frmImportSamples.WaterData[i].Na,frmImportSamples.WaterData[i].K,frmImportSamples.WaterData[i].Ca,frmImportSamples.WaterData[i].Mg,frmImportSamples.WaterData[i].Cl,frmImportSamples.WaterData[i].HCO3,frmImportSamples.WaterData[i].CO3,frmImportSamples.WaterData[i].So4));
                                }
                                writer.WriteLine("");
                                writer.WriteLine("# Piper Diagram");
                                
                                for (int i = 0; i < frmImportSamples.WaterData.Count; i++)
                                {
                                    writer.WriteLine(string.Format(@"Sample ID: {0},Cations: Na+K: {1}, Ca: {2}, Mg: {3},Anions:  CL: {4}, HCO3 + CO3: {5}, SO4: {6}", frmImportSamples.WaterData[i].sampleID, frmImportSamples.WaterData[i].Na + frmImportSamples.WaterData[i].K, frmImportSamples.WaterData[i].Ca, frmImportSamples.WaterData[i].Mg, frmImportSamples.WaterData[i].Cl, frmImportSamples.WaterData[i].HCO3 + frmImportSamples.WaterData[i].CO3, frmImportSamples.WaterData[i].So4));
                                }
                                writer.WriteLine("Points on triangle:");
                                writer.WriteLine("totalCations = Na + K + Ca + Mg");
                                writer.WriteLine("Normalized Na+K= Na+K / totalCations");
                                writer.WriteLine("Normalized Ca= Ca / totalCations");
                                writer.WriteLine("Normalized Mg= Mg / totalCations");
                                writer.WriteLine("X=Normalized Na+K*bootomrightX + Normalized Ca*bottomLeftX + Normalized Mg*topX");
                                writer.WriteLine("Y=Normalized Na+K*bootomrightY + Normalized Ca*bottomLeftY + Normalized Mg*topY");
                                writer.WriteLine("");
                                writer.WriteLine("totalAnions = CL + HCO3 + CO3 + SO4");
                                writer.WriteLine("Normalized CL= CL / totalAnions");
                                writer.WriteLine("Normalized HCO3 + CO3= HCO3 + CO3 / totalAnions");
                                writer.WriteLine("Normalized SO4= SO4 / totalAnions");
                                writer.WriteLine("X=Normalized CL*bootomrightX + Normalized HCO3 + CO3*bottomLeftX + Normalized SO4*topX");
                                writer.WriteLine("X=Normalized CL*bootomrightY + Normalized HCO3 + CO3*bottomLeftY + Normalized SO4*topY");
                                writer.WriteLine("");
                                
                                writer.WriteLine("# Collins Diagram");
                                writer.WriteLine("Calculations:");
                                writer.WriteLine("Nafac = 22.99, Kfac = 39.0983, Cafac = 20.039, Mgfac = 12.1525, Clfac = 35.453, HCO3fac = 61.01684, CO3fac = 30.004, SO4fac = 48.0313");
                                writer.WriteLine("Normalized Na=Na/Nafac");
                                writer.WriteLine("Normalized K=K/Kfac");
                                writer.WriteLine("Normalized Ca= Ca / Cafac");
                                writer.WriteLine("Normalized Mg= Mg / Mgfac");
                                writer.WriteLine("Normalized CL= CL / Clfac");
                                writer.WriteLine("Normalized HCO3=HCO3/HCO3fac");
                                writer.WriteLine("Normalized CO3=CO3/CO3fac");
                                writer.WriteLine("Normalized SO4= SO4 / SO4fac");
                                writer.WriteLine("");
                                writer.WriteLine("# Stiff Diagram");
                                writer.WriteLine("Calculations:");
                                writer.WriteLine("Nafac = 22.99, Kfac = 39.0983, Cafac = 20.039, Mgfac = 12.1525, Clfac = 35.453, HCO3fac = 61.01684, CO3fac = 30.004, SO4fac = 48.0313");
                                writer.WriteLine("Normalized Na=Na/Nafac");
                                writer.WriteLine("Normalized K=K/Kfac");
                                writer.WriteLine("Normalized Ca= Ca / Cafac");
                                writer.WriteLine("Normalized Mg= Mg / Mgfac");
                                writer.WriteLine("Normalized CL= CL / Clfac");
                                writer.WriteLine("Normalized HCO3=HCO3/HCO3fac");
                                writer.WriteLine("Normalized CO3=CO3/CO3fac");
                                writer.WriteLine("Normalized SO4= SO4 / SO4fac");
                                writer.WriteLine("total = Normalized Na + Normalized K + Normalized Mg + Normalized Ca + Normalized CL + Normalized SO4 + Normalized HCO3 + Normalized CO3");
                                writer.WriteLine("Na+K Percentage=(Normalized Na+Normalized K)/total");
                                writer.WriteLine("Ca Percentage=Normalized Ca/total*100");
                                writer.WriteLine("Mg Percentage=Normalized Mg/total*100");
                                writer.WriteLine("CL Percentage=Normalized CL/total*100");
                                writer.WriteLine("HCO3 Percentage=Normalized HCO3/total*100");
                                writer.WriteLine("CO3 Percentage=Normalized CO3/total*100");
                                writer.WriteLine("SO4 Percentage=Normalized SO4/total*100");
                                writer.WriteLine("");
                                writer.WriteLine("# Schoeller Driagram");
                                writer.WriteLine("Calculations:");
                                writer.WriteLine("Nafac = 22.99, Kfac = 39.0983, Cafac = 20.039, Mgfac = 12.1525, Clfac = 35.453, HCO3fac = 61.01684, CO3fac = 30.004, SO4fac = 48.0313");
                                writer.WriteLine("Normalized Na=Na/Nafac");
                                writer.WriteLine("Normalized K=K/Kfac");
                                writer.WriteLine("Normalized Ca= Ca / Cafac");
                                writer.WriteLine("Normalized Mg= Mg / Mgfac");
                                writer.WriteLine("Normalized CL= CL / Clfac");
                                writer.WriteLine("Normalized HCO3=HCO3/HCO3fac");
                                writer.WriteLine("Normalized CO3=CO3/CO3fac");
                                writer.WriteLine("Normalized SO4= SO4 / SO4fac");
                                writer.WriteLine("");
                                writer.WriteLine("# Bubble Driagram");
                                writer.WriteLine("X(metamorphic) = (Cl - Na) / Mg");
                                writer.WriteLine("Y(desulphurization) = (So4 * 100) / Cl");
                                writer.WriteLine("");
                                writer.WriteLine("# Elements Molar Concentration");
                                writer.WriteLine("Calculations: ");
                                writer.WriteLine("Bm = 35453, Bn = 22989.7, Bo = 39098.3, Bp = 40078, Bq = 24305, Br = 137327, Bs = 87620");
                                writer.WriteLine("Normalized Na=Na/Bn");
                                writer.WriteLine("Normalized K=K/Bo");
                                writer.WriteLine("Normalized Ca= Ca / Bp");
                                writer.WriteLine("Normalized Mg= Mg / Bq");
                                writer.WriteLine("Normalized CL= CL / Bm");
                                writer.WriteLine("Normalized Sr=Sr/Bs");
                                writer.WriteLine("Normalized Ba=Ba/Br");
                                writer.WriteLine("");
                                writer.WriteLine("#Genetic Origin and Alteration Plot");
                                writer.WriteLine("Calculations: ");
                                writer.WriteLine("EV_Na-Ca=Na / Ca");
                                writer.WriteLine("GT_K-Na=K / Na");
                                writer.WriteLine("SS_Sr-Mg=Sr / Mg");
                                writer.WriteLine("SS_Mg-Cl=Mg / Cl");
                                writer.WriteLine("SS_Sr-Cl=Sr / Cl");
                                writer.WriteLine("Lith_Sr-K=Sr / K");
                                writer.WriteLine("Lith_Mg-K=Mg / K");
                                writer.WriteLine("Lith_Ca-K=Ca / K");
                                writer.WriteLine("Wt%K=K / 10000");
                                writer.WriteLine("OM_B-Cl=B / Cl");
                                writer.WriteLine("OM_B-Na=B / Na");
                                writer.WriteLine("OM_B-Mg=B / Mg");
                                writer.WriteLine("EV_Na-Cl=Na / Cl");
                                writer.WriteLine("EV_Cl-Ca=Cl / Ca");
                                writer.WriteLine("EV_HCO3-Cl=HCO3 / Cl");
                                writer.WriteLine("EV_Cl-Sr=Cl / Sr");

                            }
                            string message = string.Format(@"File saved successfully at: {0}", filePath);
                            MessageBox.Show(message, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error saving file: "+ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            else 
            {
                MessageBox.Show("There is no data to save", "Error");
            }
        }


        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        #endregion

        private MenuStrip menuStrip;
        private ToolStripMenuItem fileMenu;
        private ToolStripMenuItem openMenuItem;
        private ToolStripMenuItem saveMenuItem;
        private ToolStripMenuItem saveAsMenuItem;
        private ToolStripMenuItem importMenu;
        private ToolStripMenuItem importFromDbMenuItem;
        private ToolStripMenuItem importFromExcelMenuItem;
        private ToolStripMenuItem exportMenu;
        private ToolStripMenuItem printPowerpoint;
        public ListBox listBoxSelected;
        public static ToolStripButton saveIcon;
        public static ListBox listBoxCharts;
        public static PictureBox mainChartPlotting;
        public static PictureBox legendPictureBox;
    }
}

