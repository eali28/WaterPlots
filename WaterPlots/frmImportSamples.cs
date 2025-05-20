using System;
using System.Collections.Generic;
using System.ComponentModel;
//using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace WaterPlots
{
    public partial class frmImportSamples : Form
    {
        private static Random random = new Random();
        public static string title = "";
        public static string JOBID = "";
        public static string connectionString = "Server=SQL-STRATOCHEM;Database=BRI;Integrated Security=True;";
        public static List<clsSampleData> samplesData = new List<clsSampleData>();
        public static List<clsSampleData> selectedSamples = new List<clsSampleData>();
        public static List<clsWater> WaterData = new List<clsWater>();
        private static frmImportSamples instance;
        private BackgroundWorker backgroundWorker;
        public static bool isCalculateAndPlotClicked=false;
        public static string selectedCompany;
        public static string selectedJob;
        public static List<clsJobs> copyOfJobs = new List<clsJobs>();
        
        public static frmImportSamples Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new frmImportSamples();
                }
                return instance;
            }
        }
        public frmImportSamples()
        {
            InitializeComponent();
            copyOfJobs = new List<clsJobs>();
            InitializeDataGridView();
            LoadData();
            cbCompanyName.SelectedIndexChanged += CbCompanyName_SelectedIndexChanged;
            cbJobNumber.SelectedIndexChanged += cbJobNumber_SelectedIndexChanged;
            // Initialize BackgroundWorker
            backgroundWorker = new BackgroundWorker();
            //backgroundWorker.DoWork += backgroundWorker_DoWork;
            isCalculateAndPlotClicked = false;
            for (int i = 0; i < clsConstants.oldData.Count; i++)
            {
                var data = clsConstants.oldData[i];
                DataGridViewRow newRow = new DataGridViewRow();
                newRow.CreateCells(dgvJobs);
                newRow.Cells[0].Value = data.jobID;
                newRow.Cells[1].Value = data.sampleID;
                newRow.Cells[2].Value = data.clientID;
                newRow.Cells[3].Value = data.wellName;
                newRow.Cells[4].Value = data.lat;
                newRow.Cells[5].Value = data.Long;
                newRow.Cells[6].Value = data.sampleType;
                newRow.Cells[7].Value = data.formationName;
                newRow.Cells[8].Value = data.depth;
                newRow.Cells[9].Value = data.prep;
                dgvJobs.Rows.Add(newRow);
            }

        }
        public static void InitializeDataGridView()
        {
            // Define Sample_ID column
            DataGridViewTextBoxColumn sampleIdColumn = new DataGridViewTextBoxColumn
            {
                Name = "Sample_ID",
                HeaderText = "Sample ID",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn ClientIDColumn = new DataGridViewTextBoxColumn
            {
                Name = "CLIENT_ID",
                HeaderText = "Client ID",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn WellName = new DataGridViewTextBoxColumn
            {
                Name = "Well_Name",
                HeaderText = "Well_Name",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn Lat = new DataGridViewTextBoxColumn
            {
                Name = "Lat",
                HeaderText = "Lat",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn Long = new DataGridViewTextBoxColumn
            {
                Name = "Long",
                HeaderText = "Long",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn sampleType = new DataGridViewTextBoxColumn
            {
                Name = "Sample_type",
                HeaderText = "Sample Type",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn formationName = new DataGridViewTextBoxColumn
            {
                Name = "Form_name",
                HeaderText = "Form Name",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn Depth = new DataGridViewTextBoxColumn
            {
                Name = "Depth",
                HeaderText = "Depth",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn Prep = new DataGridViewTextBoxColumn
            {
                Name = "Prep",
                HeaderText = "Prep",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn Age = new DataGridViewTextBoxColumn
            {
                Name = "Age",
                HeaderText = "Age",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn Abb = new DataGridViewTextBoxColumn
            {
                Name = "Abb",
                HeaderText = "Abb",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn API = new DataGridViewTextBoxColumn
            {
                Name = "API",
                HeaderText = "API",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn G02A = new DataGridViewTextBoxColumn
            {
                Name = "G02A",
                HeaderText = "G02A",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn SMPL = new DataGridViewTextBoxColumn
            {
                Name = "SMPL",
                HeaderText = "SMPL",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            DataGridViewTextBoxColumn WATER = new DataGridViewTextBoxColumn
            {
                Name = "WATER",
                HeaderText = "WATER",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            // Add column to DataGridView
            dgvSamples.Columns.Add(sampleIdColumn);
            dgvSamples.Columns.Add(ClientIDColumn);
            dgvSamples.Columns.Add(WellName);
            dgvSamples.Columns.Add(Lat);
            dgvSamples.Columns.Add(Long);
            dgvSamples.Columns.Add(sampleType);
            dgvSamples.Columns.Add(formationName);
            dgvSamples.Columns.Add(Depth);
            dgvSamples.Columns.Add(Prep);
            dgvSamples.Columns.Add(Age);
            dgvSamples.Columns.Add(Abb);
            dgvSamples.Columns.Add(API);
            dgvSamples.Columns.Add(G02A);
            dgvSamples.Columns.Add(SMPL);
            dgvSamples.Columns.Add(WATER);
        }
        private void GetDB_Load(object sender, EventArgs e)
        {

        }
        private void gbGasTest_Enter(object sender, EventArgs e)
        {

        }

        private void cbGas_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void LoadData()
        {

            // SQL query for company names
            string queryCompanyNames = "SELECT DISTINCT Company FROM REQUEST WHERE Company IS NOT NULL ORDER BY Company";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Populate Company Names in ComboBox
                    using (SqlCommand cmdCompanyNames = new SqlCommand(queryCompanyNames, connection))
                    {
                        SqlDataReader reader = cmdCompanyNames.ExecuteReader();
                        cbCompanyName.Items.Clear(); // Clear existing items

                        while (reader.Read())
                        {
                            cbCompanyName.Items.Add(reader["Company"].ToString());
                        }
                        reader.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while fetching data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void PopulateJobNumbers(string selectedCompany)
        {
            if (string.IsNullOrEmpty(selectedCompany))
                return;

            // SQL query to fetch JOBIDs for the selected company
            
            string queryJobNumbers = "SELECT DISTINCT JOBID FROM REQUEST WHERE Company = @Company ORDER BY JOBID";


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    using (SqlCommand cmdJobNumbers = new SqlCommand(queryJobNumbers, connection))
                    {
                        cmdJobNumbers.Parameters.AddWithValue("@Company", selectedCompany);

                        SqlDataReader reader = cmdJobNumbers.ExecuteReader();
                        cbJobNumber.Items.Clear(); // Clear existing items

                        while (reader.Read())
                        {
                            cbJobNumber.Items.Add(reader["JOBID"].ToString());
                        }
                        reader.Close();
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while fetching job numbers: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        public static void PopulateFirstDataGridView(string JobID, string Company ,bool flag)
        {
            if (string.IsNullOrEmpty(JobID) || string.IsNullOrEmpty(Company))
                return;

            JOBID = JobID;
            if (!flag)
            {
                string queryJobTittle = "SELECT DISTINCT Title FROM JOB WHERE JobID=@JobID";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    try
                    {
                        connection.Open();

                        using (SqlCommand cmd = new SqlCommand(queryJobTittle, connection))
                        {
                            cmd.Parameters.AddWithValue("@JobID", JobID);

                            SqlDataReader reader = cmd.ExecuteReader();
                            title = string.Empty;
                            while (reader.Read())
                            {
                                title = reader["Title"].ToString();
                            }

                            reader.Close();
                        }

                        JobTitletext.Text = title;
                        JobTitletext.AutoSize = true;
                        JobTitletext.MaximumSize = new Size(500, 0);



                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("An error occurred while fetching job title: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }


            // SQL queries
            string querySamples = @"
        SELECT 
            R.Sample_ID, 
            ISAMP.ClientID, 
            ISAMP.Sample_type,
            ISAMP.Form_name,
            ISAMP.Base_depth,
            ISAMP.Depth_units,
            R.TEST,
            R.TESTDATA,
            R.PREP,
            W.Well_Name,
            W.Lat, 
            W.Long,
            W.Age
            
        FROM 
            REQUEST R
        LEFT JOIN 
            [INFO SAMPLE] ISAMP 
        ON 
            R.Sample_ID = ISAMP.Sample_ID
        LEFT JOIN
            [INFO WELL] W
        ON
            W.Well_ID=ISAMP.Well_ID
        WHERE 
            R.Company = @Company AND R.JOBID = @JOBID";
            samplesData=new List<clsSampleData>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    // Clear the DataGridView before populating new data
                    if (!flag)
                    {
                        dgvSamples.Rows.Clear();
                    }
                    else 
                    {
                        dgvSamples = new System.Windows.Forms.DataGridView();
                        InitializeDataGridView();
                    }


                    // Fetch Sample_ID, ClientID, and Sample_type
                    using (SqlCommand cmdSamples = new SqlCommand(querySamples, connection))
                    {
                        cmdSamples.Parameters.AddWithValue("@Company", Company);
                        cmdSamples.Parameters.AddWithValue("@JOBID", JobID);

                        SqlDataReader reader = cmdSamples.ExecuteReader();
                        while (reader.Read())
                        {
                            int idx = -1;
                            for (int i = 0; i < samplesData.Count; i++)
                            {
                                if ((reader["Sample_ID"] as string) == samplesData[i].SampleID && (reader["PREP"] as string) == samplesData[i].Prep)
                                {
                                    idx = i;
                                    break;
                                }
                            }
                            if (idx != -1)
                            {

                                if ((reader["TEST"] as string) == "API")
                                {
                                    // Get the existing tuple
                                    samplesData[idx].API = reader["TEST"] != DBNull.Value ? (string)reader["TEST"] : "";
                                }
                                else if ((reader["TEST"] as string) == "G02A")
                                {
                                    // Get the existing tuple
                                    samplesData[idx].G02A = reader["TEST"] != DBNull.Value ? (string)reader["TEST"] : "";
                                }
                                else if ((reader["TEST"] as string)== "SMPL")
                                {
                                    samplesData[idx].SMPL = reader["TEST"] != DBNull.Value ? (string)reader["TEST"] : "";
                                }
                                else
                                {
                                    samplesData[idx].WATER = reader["TEST"] != DBNull.Value ? (string)reader["TEST"] : "";
                                }

                            }
                            else
                            {
                                if ((reader["TEST"] as string) == "G02A")
                                {
                                    samplesData.Add(new clsSampleData
                                    {
                                        SampleID = reader["Sample_ID"] != DBNull.Value ? (string)reader["Sample_ID"] : "",
                                        ClientID = reader["ClientID"] != DBNull.Value ? (string)reader["ClientID"] : "",
                                        SampleType = reader["Sample_type"] != DBNull.Value ? (string)reader["Sample_type"] : "",
                                        FormationName = reader["Form_name"] != DBNull.Value ? (string)reader["Form_name"] : "",
                                        Depth = ((string)(reader["Base_depth"] != DBNull.Value ? (string)reader["Base_depth"] : "") + " FT") ?? "",
                                        Prep = reader["PREP"] != DBNull.Value ? (string)reader["PREP"] : "",
                                        API = "",
                                        G02A = ("C" + (reader["TESTDATA"]!=DBNull.Value?(string)reader["TESTDATA"]:"")) ?? "",
                                        SMPL = "",
                                        WATER = "",
                                        wellName = reader["Well_Name"] != DBNull.Value ? (string)reader["Well_Name"] : "",
                                        latitude = reader["Lat"] is float ? ((float)reader["Lat"]).ToString() : "",
                                        longtude = reader["Long"] is float ? ((float)reader["Long"]).ToString() : "",
                                        age = reader["Age"] != DBNull.Value ? reader["Age"].ToString() : ""
                                    });

                                }
                                else if ((reader["TEST"] as string) == "API")
                                {
                                    samplesData.Add(new clsSampleData
                                    {
                                        SampleID = reader["Sample_ID"] != DBNull.Value ? (string)reader["Sample_ID"] : "",
                                        ClientID = reader["ClientID"] != DBNull.Value ? (string)reader["ClientID"] : "",
                                        SampleType = reader["Sample_type"] != DBNull.Value ? (string)reader["Sample_type"] : "",
                                        FormationName = reader["Form_name"] != DBNull.Value ? (string)reader["Form_name"] : "",
                                        Depth = ((string)(reader["Base_depth"] != DBNull.Value ? (string)reader["Base_depth"] : "") + " FT") ?? "",
                                        Prep = reader["PREP"] != DBNull.Value ? (string)reader["PREP"] : "",
                                        API = reader["TEST"] != DBNull.Value ? (string)reader["TEST"] : "",
                                        G02A = "",
                                        SMPL = "",
                                        WATER = "",
                                        wellName = reader["Well_Name"] != DBNull.Value ? (string)reader["Well_Name"] : "",
                                        latitude = reader["Lat"] is float ? ((float)reader["Lat"]).ToString() : "",
                                        longtude = reader["Long"] is float ? ((float)reader["Long"]).ToString() : "",
                                        age = reader["Age"] != DBNull.Value ? reader["Age"].ToString() : ""
                                    });
                                }
                                else if ((reader["TEST"] as string) == "SMPL")
                                {
                                    samplesData.Add(new clsSampleData
                                    {
                                        SampleID = reader["Sample_ID"] != DBNull.Value ? (string)reader["Sample_ID"] : "",
                                        ClientID = reader["ClientID"] != DBNull.Value ? (string)reader["ClientID"] : "",
                                        SampleType = reader["Sample_type"] != DBNull.Value ? (string)reader["Sample_type"] : "",
                                        FormationName = reader["Form_name"] != DBNull.Value ? (string)reader["Form_name"] : "",
                                        Depth = ((string)(reader["Base_depth"] != DBNull.Value ? (string)reader["Base_depth"] : "") + " FT") ?? "",
                                        Prep = reader["PREP"] != DBNull.Value ? (string)reader["PREP"] : "",
                                        API = "",
                                        G02A = "",
                                        SMPL = reader["TEST"] != DBNull.Value ? (string)reader["TEST"] : "",
                                        WATER = "",
                                        wellName = reader["Well_Name"] != DBNull.Value ? (string)reader["Well_Name"] : "",
                                        latitude = reader["Lat"] is float ? ((float)reader["Lat"]).ToString() : "",
                                        longtude = reader["Long"] is float ? ((float)reader["Long"]).ToString() : "",
                                        age = reader["Age"] != DBNull.Value ? reader["Age"].ToString() : ""
                                    });
                                }
                                else
                                {
                                    samplesData.Add(new clsSampleData
                                    {
                                        SampleID = reader["Sample_ID"] != DBNull.Value ? (string)reader["Sample_ID"] : "",
                                        ClientID = reader["ClientID"] != DBNull.Value ? (string)reader["ClientID"] : "",
                                        SampleType = reader["Sample_type"] != DBNull.Value ? (string)reader["Sample_type"] : "",
                                        FormationName = reader["Form_name"] !=DBNull.Value?(string)reader["Form_name"]: "",
                                        Depth = ((string)(reader["Base_depth"] != DBNull.Value ? (string)reader["Base_depth"] : "") + " FT") ?? "",
                                        Prep = reader["PREP"] != DBNull.Value ? (string)reader["PREP"] : "",
                                        API = "",
                                        G02A = "",
                                        SMPL = "",
                                        WATER = reader["TEST"] != DBNull.Value ? (string)reader["TEST"] : "",
                                        wellName = reader["Well_Name"] != DBNull.Value ? (string)reader["Well_Name"] : "",
                                        latitude = reader["Lat"] is float ? ((float)reader["Lat"]).ToString() : "",
                                        longtude = reader["Long"] is float ? ((float)reader["Long"]).ToString() : "",
                                        age = reader["Age"] != DBNull.Value ? reader["Age"].ToString() : ""
                                    });

                                }
                            }



                        }
                        reader.Close();
                    }

                    

                    // Populate the DataGridView by matching rows from both queries
                    for (int i = 0; i < samplesData.Count; i++)
                    {
                        string sampleID = i < samplesData.Count ? samplesData[i].SampleID : string.Empty;
                        string clientID = i < samplesData.Count ? samplesData[i].ClientID : string.Empty;
                        string sampleType = i < samplesData.Count ? samplesData[i].SampleType : string.Empty;
                        string FormName = i < samplesData.Count ? samplesData[i].FormationName : string.Empty;
                        string wellName = i < samplesData.Count ? samplesData[i].wellName : string.Empty;
                        string lat = i < samplesData.Count ? samplesData[i].latitude : string.Empty;
                        string longitude = i < samplesData.Count ? samplesData[i].longtude : string.Empty;
                        string age = i < samplesData.Count ? samplesData[i].age : string.Empty;
                        string depth = i < samplesData.Count ? samplesData[i].Depth : string.Empty;
                        string Prep = i < samplesData.Count ? (string)samplesData[i].Prep : string.Empty;
                        string api = i < samplesData.Count ? samplesData[i].API : string.Empty;
                        string G02A = i < samplesData.Count ? samplesData[i].G02A : string.Empty;
                        string smpl = i < samplesData.Count ? samplesData[i].SMPL : string.Empty;
                        string water = i < samplesData.Count ? samplesData[i].WATER : string.Empty;
                        if (sampleType == "WATER" || flag && sampleID!=string.Empty)
                        {
                            dgvSamples.Rows.Add(sampleID, clientID, wellName, lat, longitude, sampleType, FormName, depth, Prep, age, "", api != string.Empty ? "C" : "", G02A != string.Empty ? G02A : "", smpl != string.Empty ? "C" : "", water != string.Empty ? "C" : "");
                        }
                        //WaterData[i].latitude = lat;
                        //WaterData[i].longtude = longitude;

                    }
                   

                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while fetching data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void DeleteButton_Click(object sender, EventArgs e)
        {
            if (dgvJobs.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow selectedRow in dgvJobs.SelectedRows)
                {
                    // Check if the row is not a new row before deleting
                    if (!selectedRow.IsNewRow)
                    {
                        dgvJobs.Rows.Remove(selectedRow);
                    }
                }
            }
            else
            {
                MessageBox.Show("No rows selected to delete", "Delete Row", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void CbCompanyName_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Fetch JOBIDs based on selected company
            selectedCompany = (string)cbCompanyName.SelectedItem;
            PopulateJobNumbers(selectedCompany);
        }
        private void cbJobNumber_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedJob = (string)cbJobNumber.SelectedItem;
            PopulateFirstDataGridView(selectedJob, selectedCompany,false);
        }

        private void cbCompanyName_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }
        public static void AddButton_Click(object sender, EventArgs e)
        {
            if (dgvSamples.SelectedRows.Count > 0)
            {
                //copyOfJobs.Clear();
                //dgvJobs.Rows.Clear();
                
                
                for (int j = dgvSamples.SelectedRows.Count - 1; j >= 0; j--)
                {

                    DataGridViewRow selectedRow = dgvSamples.SelectedRows[j];
                    
                    DataGridViewRow newRow = new DataGridViewRow();

                    // Clone cell structure
                    newRow.CreateCells(dgvJobs);

                        // Assign Job ID manually
                    // Copy cell values
                    int c = 1;
                    for (int i = 0; i < 9; i++)
                    {
                        newRow.Cells[c].Value = selectedRow.Cells[i].Value;
                        c++;
                    }

                    // Add the new row to both DataGridViews
                    if (!selectedRow.IsNewRow && !IsRowEmpty(selectedRow))
                    {
                        newRow.Cells[0].Value = JOBID;
                        dgvJobs.Rows.Add(newRow);
                        clsJobs newJob = new clsJobs
                        {
                            jobID = JOBID,
                            sampleID = (string)newRow.Cells[1].Value,
                            clientID = (string)newRow.Cells[2].Value,
                            wellName = (string)newRow.Cells[3].Value,
                            lat = (string)newRow.Cells[4].Value,
                            Long = (string)newRow.Cells[5].Value,
                            sampleType = (string)newRow.Cells[6].Value,
                            formationName = (string)newRow.Cells[7].Value,
                            depth = (string)newRow.Cells[8].Value,
                            prep = (string)newRow.Cells[9].Value
                        };
                        clsConstants.oldData.Add(newJob);
                        copyOfJobs.Add(newJob);
                    }
                }
            }
            else
            {
                MessageBox.Show("No rows selected in the first DataGridView.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private static bool IsRowEmpty(DataGridViewRow row)
        {
            foreach (DataGridViewCell cell in row.Cells)
            {
                if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                {
                    return false; // If at least one cell has data, the row is NOT empty
                }
            }
            return true; // Row is empty
        }

        public void btnCalculateAndPlot_Click(object sender, EventArgs e)
        {

            if (dgvJobs.Rows.Count > 0)
            {

                try
                {

                    
                    backgroundWorker.RunWorkerAsync();
                    // Reset and show the existing progress bar
                    progressBar.Value = 0;
                    progressBar.Visible = true;


                    getWaterData();
                    progressBar.Value = 100;
                    frmMainForm.isScalesNeedNoUpdate = false;
                    frmMainForm.UpdateRadarDiagram();
                    frmMainForm.UpdateCollinsDiagram();
                    frmMainForm.UpdatePieDiagram();
                    frmMainForm.UpdatePiperDiagram();
                    frmMainForm.UpdateSchoellerDiagram();
                    frmMainForm.UpdateLogsDiagram();
                    frmMainForm.UpdateStiffDiagram();
                    frmMainForm.UpdateBubbleDiagram();
                    if (frmMainForm.listBoxCharts.SelectedItem != null)
                    {
                        frmMainForm.UpdateScalesinRadar(frmMainForm.listBoxCharts.SelectedItem.ToString());
                    }
                    if (WaterData.Count > 0)
                    {
                        isCalculateAndPlotClicked = true;
                        frmMainForm.saveIcon.Image = Properties.Resources.saveActivated;
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("No Data to calculate.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    frmMainForm.mainChartPlotting.Refresh();
                }
                catch (Exception ex)
                {

                    MessageBox.Show("An error occurred during calculation: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("No Data to calculate.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public static void getWaterData()
        {
            WaterData = new List<clsWater>();
            string waterDataQuery = @"SELECT SAMPLE_ID,Na, K, Ca, Mg, SO4, Bicarbonate, Carbonate, Cl, Sr,B, Ba,TDS,Al,Co,Cu,Fe,Mn,Pb,Ni,Zn,Cd,Cr,Tl,Be,Se,Li FROM [ANL_WATER_ANALYSIS] WHERE REQNUM = @JobID AND SAMPLE_ID=@SAMPLE_ID";
            string querysamplewell = @"
            SELECT 
                ISAMP.SAMPLE_ID,
                W.Well_Name,
                ISAMP.Base_depth,
                ISAMP.ClientID
            FROM 
                [INFO SAMPLE] ISAMP 
            LEFT JOIN 
                [INFO WELL] W 
            ON 
                ISAMP.Well_ID=W.Well_ID 
            LEFT JOIN 
                REQUEST R 
            ON 
                ISAMP.SAMPLE_ID=R.Sample_ID 
            WHERE 
                R.JOBID=@JOBID";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    for (int i = 0; i < dgvJobs.Rows.Count; i++)
                    {
                        using (SqlCommand cmdWater = new SqlCommand(waterDataQuery, connection))
                        {
                            cmdWater.Parameters.AddWithValue("@JobID", dgvJobs.Rows[i].Cells[0].Value??DBNull.Value);
                            cmdWater.Parameters.AddWithValue("@SAMPLE_ID", dgvJobs.Rows[i].Cells[1].Value ?? DBNull.Value);
                            SqlDataReader reader = cmdWater.ExecuteReader();
                            int idx = 0;
                            if (!reader.HasRows)
                            {
                                Console.WriteLine("No rows returned for SAMPLE_ID:", dgvJobs.Rows[i].Cells[1].Value);
                            }
                            while (reader.Read())
                            {
                                string sampleid = reader["SAMPLE_ID"].ToString();
                                string Nastring = reader["Na"].ToString();
                                Nastring = Nastring.Replace("<", "").Replace(">", "");

                                double NaValue;
                                if (!double.TryParse(Nastring, out NaValue))
                                {
                                    NaValue = 0; // or handle error as needed
                                }
                                string Kstring = reader["K"].ToString();
                                Kstring = Kstring.Replace("<", "").Replace(">", "");
                                double KValue = reader["K"] != DBNull.Value ? Convert.ToDouble(reader["K"]) : 0;
                                if (!double.TryParse(Kstring, out KValue))
                                {
                                    KValue = 0; // or handle error as needed
                                }
                                string Castring = reader["Ca"].ToString();
                                Castring = Castring.Replace("<", "").Replace(">", "");
                                double CaValue;
                                if (!double.TryParse(Castring, out CaValue))
                                {
                                    CaValue = 0; // or handle error as needed
                                }
                                string Mgstring = reader["Mg"].ToString();
                                Mgstring = Mgstring.Replace("<", "").Replace(">", "");
                                double MgValue;
                                if (!double.TryParse(Mgstring, out MgValue))
                                {
                                    MgValue = 0; // or handle error as needed
                                }
                                string SO4string = reader["SO4"].ToString();
                                SO4string = SO4string.Replace("<", "").Replace(">", "");
                                double SO4Value;
                                if (!double.TryParse(SO4string, out SO4Value))
                                {
                                    SO4Value = 0; // or handle error as needed
                                }
                                string HCO3string = reader["Bicarbonate"].ToString();
                                HCO3string=HCO3string.Replace("<","").Replace(">","");
                                double HCO3Value;
                                if (!double.TryParse(HCO3string, out HCO3Value))
                                {
                                    HCO3Value = 0; // or handle error as needed
                                }
                                //double CO3 = reader["Carbonate"] != DBNull.Value ? Convert.ToDouble(reader["Carbonate"]) : 0;
                                string Alstring = reader["Al"].ToString();
                                Alstring = Alstring.Replace("<", "").Replace(">", "");
                                double ALValue;
                                if (!double.TryParse(Alstring, out ALValue))
                                {
                                    ALValue = 0; // or handle error as needed
                                }
                                // Co
                                string Costring = reader["Co"].ToString().Replace("<", "").Replace(">", "");
                                double CoValue;
                                if (!double.TryParse(Costring, out CoValue))
                                {
                                    CoValue = 0;
                                }

                                // Cu
                                string Custring = reader["Cu"].ToString().Replace("<", "").Replace(">", "");
                                double CuValue;
                                if (!double.TryParse(Custring, out CuValue))
                                {
                                    CuValue = 0;
                                }

                                // Mn
                                string Mnstring = reader["Mn"].ToString().Replace("<", "").Replace(">", "");
                                double MnValue;
                                if (!double.TryParse(Mnstring, out MnValue))
                                {
                                    MnValue = 0;
                                }

                                // Ni
                                string Nistring = reader["Ni"].ToString().Replace("<", "").Replace(">", "");
                                double NiValue;
                                if (!double.TryParse(Nistring, out NiValue))
                                {
                                    NiValue = 0;
                                }

                                // Zn
                                string Znstring = reader["Zn"].ToString().Replace("<", "").Replace(">", "");
                                double ZnValue;
                                if (!double.TryParse(Znstring, out ZnValue))
                                {
                                    ZnValue = 0;
                                }

                                // Pb
                                string Pbstring = reader["Pb"].ToString().Replace("<", "").Replace(">", "");
                                double PbValue;
                                if (!double.TryParse(Pbstring, out PbValue))
                                {
                                    PbValue = 0;
                                }

                                // Fe
                                string Festring = reader["Fe"].ToString().Replace("<", "").Replace(">", "");
                                double FeValue;
                                if (!double.TryParse(Festring, out FeValue))
                                {
                                    FeValue = 0;
                                }

                                // Cd
                                string Cdstring = reader["Cd"].ToString().Replace("<", "").Replace(">", "");
                                double CdValue;
                                if (!double.TryParse(Cdstring, out CdValue))
                                {
                                    CdValue = 0;
                                }

                                // Cr
                                string Crstring = reader["Cr"].ToString().Replace("<", "").Replace(">", "");
                                double CrValue;
                                if (!double.TryParse(Crstring, out CrValue))
                                {
                                    CrValue = 0;
                                }

                                // Tl
                                string Tlstring = reader["Tl"].ToString().Replace("<", "").Replace(">", "");
                                double TlValue;
                                if (!double.TryParse(Tlstring, out TlValue))
                                {
                                    TlValue = 0;
                                }

                                // Be
                                string Bestring = reader["Be"].ToString().Replace("<", "").Replace(">", "");
                                double BeValue;
                                if (!double.TryParse(Bestring, out BeValue))
                                {
                                    BeValue = 0;
                                }

                                // Se
                                string Sestring = reader["Se"].ToString().Replace("<", "").Replace(">", "");
                                double SeValue;
                                if (!double.TryParse(Sestring, out SeValue))
                                {
                                    SeValue = 0;
                                }

                                // Li
                                string Listring = reader["Li"].ToString().Replace("<", "").Replace(">", "");
                                double LiValue;
                                if (!double.TryParse(Listring, out LiValue))
                                {
                                    LiValue = 0;
                                }

                                string carbonateString = reader["Carbonate"].ToString().Replace("<", "").Replace(">", ""); ;
                                
                                double CO3Value;
                                if (!double.TryParse(carbonateString, out CO3Value))
                                {
                                    // Successfully converted to double
                                    CO3Value = 0;
                                }

                                // Cl
                                string Clstring = reader["Cl"].ToString().Replace("<", "").Replace(">", "");
                                double ClValue;
                                if (!double.TryParse(Clstring, out ClValue))
                                {
                                    ClValue = 0;
                                }

                                // Sr
                                string Srstring = reader["Sr"].ToString().Replace("<", "").Replace(">", "");
                                double SrValue;
                                if (!double.TryParse(Srstring, out SrValue))
                                {
                                    SrValue = 0;
                                }

                                // Ba
                                string Bastring = reader["Ba"].ToString().Replace("<", "").Replace(">", "");
                                double BaValue;
                                if (!double.TryParse(Bastring, out BaValue))
                                {
                                    BaValue = 0;
                                }

                                // TDS
                                string TDSstring = reader["TDS"].ToString().Replace("<", "").Replace(">", "");
                                double tdsValue;
                                if (!double.TryParse(TDSstring, out tdsValue))
                                {
                                    tdsValue = 0;
                                }

                                // B
                                string Bstring = reader["B"].ToString().Replace("<", "").Replace(">", "");
                                double BValue;
                                if (!double.TryParse(Bstring, out BValue))
                                {
                                    BValue = 0;
                                }



                                // Add the parsed values to the WaterData list
                                clsWater existingWaterData = new clsWater
                                {
                                    sampleID = sampleid,
                                    Na = NaValue,
                                    K = KValue,
                                    Ca = CaValue,
                                    Mg = MgValue,
                                    So4 = SO4Value,
                                    HCO3 = HCO3Value,
                                    CO3 = CO3Value,
                                    Cl = ClValue,
                                    Sr = SrValue,
                                    Ba = BaValue,
                                    B = BValue,
                                    TDS = tdsValue,
                                    Well_Name = "",
                                    Depth = "",
                                    ClientID = "",
                                    color = Color.Blue,
                                    Al = ALValue,
                                    Co = CoValue,
                                    Cu = CuValue,
                                    Mn = MnValue,
                                    Ni = NiValue,
                                    Zn = ZnValue,
                                    Pb = PbValue,
                                    Fe = FeValue,
                                    Cd = CdValue,
                                    Cr = CrValue,
                                    Tl = TlValue,
                                    Be = BeValue,
                                    Se = SeValue,
                                    Li = LiValue,
                                    jobID = JOBID,
                                    latitude = dgvJobs.Rows[i].Cells[4].Value.ToString(),
                                    longtude=dgvJobs.Rows[i].Cells[5].Value.ToString(),
                                    formName=dgvJobs.Rows[i].Cells[7].Value.ToString(),
                                    prep=dgvJobs.Rows[i].Cells[9].Value.ToString()
                                };

                                WaterData.Add(existingWaterData);

                                // Update progress bar
                                progressBar.Invoke((Action)(() => progressBar.Value = (idx + 1) * 100 / samplesData.Count));
                                idx++;

                            }

                            reader.Close();

                        }
                    }
                    for(int j=0;j<dgvJobs.RowCount;j++)
                    {
                        using (SqlCommand cmdsamplewell = new SqlCommand(querysamplewell, connection))
                        {
                            cmdsamplewell.Parameters.AddWithValue("@JOBID", dgvJobs.Rows[j].Cells[0].Value ?? DBNull.Value);
                            SqlDataReader reader = cmdsamplewell.ExecuteReader();
                            while (reader.Read())
                            {
                                for (int i = 0; i < WaterData.Count; i++)
                                {
                                    if (reader["SAMPLE_ID"].ToString() == WaterData[i].sampleID)
                                    {
                                        var existingTuple = WaterData[i];
                                        WaterData[i].Well_Name = reader["Well_Name"].ToString();
                                        WaterData[i].Depth = reader["Base_depth"].ToString() + " FT";
                                        WaterData[i].ClientID = reader["ClientID"].ToString();

                                        WaterData[i].color = GetRandomColor(false);

                                    }
                                }
                            }
                            reader.Close();
                        }
                    }
                    

                    progressBar.Invoke((Action)(() => progressBar.Value = 100));

                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while fetching water data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }


        }

        public static Color GetRandomColor(bool found)
        {
            Color newColor = Color.FromArgb(random.Next(256), random.Next(256), random.Next(256));
            found = false;
            foreach (var data in WaterData)
            {
                if (data.color == newColor)
                {
                    found = true;
                    break;
                }
                else
                {

                    found = false;
                }
            }
            if (found)
            {
                return GetRandomColor(found);
            }
            else 
            {
                return newColor;
            }
            
        }
        //private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    getWaterData();
        //}
    }
}
