namespace WindowsFormsApplication2
{
    partial class frmBubblePiperLegend
    {
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
            this.colorPanel = new System.Windows.Forms.Panel();
            this.colorLabel = new System.Windows.Forms.Label();
            dgvJobsInDetails = new System.Windows.Forms.DataGridView();
            this.JobIDColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SampleIDColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ClientIDColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.WellNameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.latColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LongColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SampleTypeColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.formNameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DepthColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.prepColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.updateButton = new System.Windows.Forms.Button();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.symbolChange = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(dgvJobsInDetails)).BeginInit();
            this.SuspendLayout();
            // 
            // colorPanel
            // 
            this.colorPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.colorPanel.Location = new System.Drawing.Point(421, 248);
            this.colorPanel.Name = "colorPanel";
            this.colorPanel.Size = new System.Drawing.Size(157, 20);
            this.colorPanel.TabIndex = 13;
            this.colorPanel.Click += new System.EventHandler(this.colorPanel_Click);
            this.colorPanel.Paint += new System.Windows.Forms.PaintEventHandler(this.colorPanel_Paint);
            // 
            // colorLabel
            // 
            this.colorLabel.AutoSize = true;
            this.colorLabel.Location = new System.Drawing.Point(357, 248);
            this.colorLabel.Name = "colorLabel";
            this.colorLabel.Size = new System.Drawing.Size(37, 13);
            this.colorLabel.TabIndex = 14;
            this.colorLabel.Text = "Color: ";
            // 
            // dgvJobsInDetails
            // 
            dgvJobsInDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvJobsInDetails.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.JobIDColumn,
            this.SampleIDColumn,
            this.ClientIDColumn,
            this.WellNameColumn,
            this.latColumn,
            this.LongColumn,
            this.SampleTypeColumn,
            this.formNameColumn,
            this.DepthColumn,
            this.prepColumn});
            dgvJobsInDetails.Location = new System.Drawing.Point(2, 0);
            dgvJobsInDetails.Name = "dgvJobsInDetails";
            dgvJobsInDetails.Size = new System.Drawing.Size(900, 222);
            dgvJobsInDetails.TabIndex = 15;
            dgvJobsInDetails.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvJobsInDetails_CellContentClick_1);
            // 
            // dataGridViewTextBoxColumn16
            // 
            this.JobIDColumn.HeaderText = "Job ID";
            this.JobIDColumn.Name = "dataGridViewTextBoxColumn16";
            // 
            // dataGridViewTextBoxColumn17
            // 
            this.SampleIDColumn.HeaderText = "Sample ID";
            this.SampleIDColumn.Name = "dataGridViewTextBoxColumn17";
            // 
            // dataGridViewTextBoxColumn18
            // 
            this.ClientIDColumn.HeaderText = "Client ID";
            this.ClientIDColumn.Name = "dataGridViewTextBoxColumn18";
            // 
            // dataGridViewTextBoxColumn19
            // 
            this.WellNameColumn.HeaderText = "Well Name";
            this.WellNameColumn.Name = "dataGridViewTextBoxColumn19";
            // 
            // dataGridViewTextBoxColumn20
            // 
            this.latColumn.HeaderText = "Lat";
            this.latColumn.Name = "dataGridViewTextBoxColumn20";
            // 
            // dataGridViewTextBoxColumn21
            // 
            this.LongColumn.HeaderText = "Long";
            this.LongColumn.Name = "dataGridViewTextBoxColumn21";
            // 
            // dataGridViewTextBoxColumn22
            // 
            this.SampleTypeColumn.HeaderText = "Sample Type";
            this.SampleTypeColumn.Name = "dataGridViewTextBoxColumn22";
            // 
            // dataGridViewTextBoxColumn23
            // 
            this.formNameColumn.HeaderText = "Formation Name";
            this.formNameColumn.Name = "dataGridViewTextBoxColumn23";
            // 
            // dataGridViewTextBoxColumn24
            // 
            this.DepthColumn.HeaderText = "Depth";
            this.DepthColumn.Name = "dataGridViewTextBoxColumn24";
            // 
            // dataGridViewTextBoxColumn25
            // 
            this.prepColumn.HeaderText = "Prep";
            this.prepColumn.Name = "dataGridViewTextBoxColumn25";
            // 
            // updateButton
            // 
            this.updateButton.Location = new System.Drawing.Point(380, 312);
            this.updateButton.Name = "updateButton";
            this.updateButton.Size = new System.Drawing.Size(75, 23);
            this.updateButton.TabIndex = 16;
            this.updateButton.Text = "Apply";
            this.updateButton.UseVisualStyleBackColor = true;
            this.updateButton.Click += new System.EventHandler(this.updateButton_Click);
            // 
            // button1
            // 
            this.symbolChange.Location = new System.Drawing.Point(742, 298);
            this.symbolChange.Name = "symbolChange";
            this.symbolChange.Size = new System.Drawing.Size(114, 23);
            this.symbolChange.TabIndex = 17;
            this.symbolChange.Text = "Change Symbol";
            this.symbolChange.UseVisualStyleBackColor = true;
            this.symbolChange.Click += new System.EventHandler(this.symbol_change_Click);
            // 
            // frmPiperLegend
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(235)))), ((int)(((byte)(250)))));
            this.ClientSize = new System.Drawing.Size(968, 364);
            this.Controls.Add(this.symbolChange);
            this.Controls.Add(this.updateButton);
            this.Controls.Add(dgvJobsInDetails);
            this.Controls.Add(this.colorLabel);
            this.Controls.Add(this.colorPanel);
            this.Name = "frmPiperLegend";
            
            this.Load += new System.EventHandler(this.PiperDetails_Load);
            ((System.ComponentModel.ISupportInitialize)(dgvJobsInDetails)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel colorPanel;
        private System.Windows.Forms.Label colorLabel;
        public static System.Windows.Forms.DataGridView dgvJobsInDetails;
        public System.Windows.Forms.DataGridViewTextBoxColumn JobIDColumn;
        public System.Windows.Forms.DataGridViewTextBoxColumn SampleIDColumn;
        public System.Windows.Forms.DataGridViewTextBoxColumn ClientIDColumn;
        public System.Windows.Forms.DataGridViewTextBoxColumn WellNameColumn;
        public System.Windows.Forms.DataGridViewTextBoxColumn latColumn;
        public System.Windows.Forms.DataGridViewTextBoxColumn LongColumn;
        public System.Windows.Forms.DataGridViewTextBoxColumn SampleTypeColumn;
        public System.Windows.Forms.DataGridViewTextBoxColumn formNameColumn;
        public System.Windows.Forms.DataGridViewTextBoxColumn DepthColumn;
        public System.Windows.Forms.DataGridViewTextBoxColumn prepColumn;
        public System.Windows.Forms.Button updateButton;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Button symbolChange;
    }
}