namespace WaterPlots
{
    partial class SchoellerDetails
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
            this.dgvJobsInDetails = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn16 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn17 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn18 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn19 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn20 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn21 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn22 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn23 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn24 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn25 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colorPanel = new System.Windows.Forms.Panel();
            this.colorLabel = new System.Windows.Forms.Label();
            this.updateButton = new System.Windows.Forms.Button();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.widthLabel = new System.Windows.Forms.Label();
            this.typeCombobox = new System.Windows.Forms.ComboBox();
            this.typeLabel = new System.Windows.Forms.Label();
            this.widthTextBox = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvJobsInDetails)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvJobsInDetails
            // 
            this.dgvJobsInDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvJobsInDetails.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn16,
            this.dataGridViewTextBoxColumn17,
            this.dataGridViewTextBoxColumn18,
            this.dataGridViewTextBoxColumn19,
            this.dataGridViewTextBoxColumn20,
            this.dataGridViewTextBoxColumn21,
            this.dataGridViewTextBoxColumn22,
            this.dataGridViewTextBoxColumn23,
            this.dataGridViewTextBoxColumn24,
            this.dataGridViewTextBoxColumn25});
            this.dgvJobsInDetails.Location = new System.Drawing.Point(12, 12);
            this.dgvJobsInDetails.Name = "dgvJobsInDetails";
            this.dgvJobsInDetails.Size = new System.Drawing.Size(846, 222);
            this.dgvJobsInDetails.TabIndex = 16;
            // 
            // dataGridViewTextBoxColumn16
            // 
            this.dataGridViewTextBoxColumn16.HeaderText = "Job ID";
            this.dataGridViewTextBoxColumn16.Name = "dataGridViewTextBoxColumn16";
            // 
            // dataGridViewTextBoxColumn17
            // 
            this.dataGridViewTextBoxColumn17.HeaderText = "Sample ID";
            this.dataGridViewTextBoxColumn17.Name = "dataGridViewTextBoxColumn17";
            // 
            // dataGridViewTextBoxColumn18
            // 
            this.dataGridViewTextBoxColumn18.HeaderText = "Client ID";
            this.dataGridViewTextBoxColumn18.Name = "dataGridViewTextBoxColumn18";
            // 
            // dataGridViewTextBoxColumn19
            // 
            this.dataGridViewTextBoxColumn19.HeaderText = "Well Name";
            this.dataGridViewTextBoxColumn19.Name = "dataGridViewTextBoxColumn19";
            // 
            // dataGridViewTextBoxColumn20
            // 
            this.dataGridViewTextBoxColumn20.HeaderText = "Lat";
            this.dataGridViewTextBoxColumn20.Name = "dataGridViewTextBoxColumn20";
            // 
            // dataGridViewTextBoxColumn21
            // 
            this.dataGridViewTextBoxColumn21.HeaderText = "Long";
            this.dataGridViewTextBoxColumn21.Name = "dataGridViewTextBoxColumn21";
            // 
            // dataGridViewTextBoxColumn22
            // 
            this.dataGridViewTextBoxColumn22.HeaderText = "Sample Type";
            this.dataGridViewTextBoxColumn22.Name = "dataGridViewTextBoxColumn22";
            // 
            // dataGridViewTextBoxColumn23
            // 
            this.dataGridViewTextBoxColumn23.HeaderText = "Formation Name";
            this.dataGridViewTextBoxColumn23.Name = "dataGridViewTextBoxColumn23";
            // 
            // dataGridViewTextBoxColumn24
            // 
            this.dataGridViewTextBoxColumn24.HeaderText = "Depth";
            this.dataGridViewTextBoxColumn24.Name = "dataGridViewTextBoxColumn24";
            // 
            // dataGridViewTextBoxColumn25
            // 
            this.dataGridViewTextBoxColumn25.HeaderText = "Prep";
            this.dataGridViewTextBoxColumn25.Name = "dataGridViewTextBoxColumn25";
            // 
            // colorPanel
            // 
            this.colorPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.colorPanel.Location = new System.Drawing.Point(365, 278);
            this.colorPanel.Name = "colorPanel";
            this.colorPanel.Size = new System.Drawing.Size(157, 20);
            this.colorPanel.TabIndex = 17;
            this.colorPanel.Click += new System.EventHandler(this.colorPanel_Click);
            // 
            // colorLabel
            // 
            this.colorLabel.AutoSize = true;
            this.colorLabel.Location = new System.Drawing.Point(291, 285);
            this.colorLabel.Name = "colorLabel";
            this.colorLabel.Size = new System.Drawing.Size(37, 13);
            this.colorLabel.TabIndex = 15;
            this.colorLabel.Text = "Color: ";
            // 
            // updateButton
            // 
            this.updateButton.Location = new System.Drawing.Point(365, 405);
            this.updateButton.Name = "updateButton";
            this.updateButton.Size = new System.Drawing.Size(75, 23);
            this.updateButton.TabIndex = 18;
            this.updateButton.Text = "Apply";
            this.updateButton.UseVisualStyleBackColor = true;
            this.updateButton.Click += new System.EventHandler(this.updateButton_Click);
            // 
            // widthLabel
            // 
            this.widthLabel.AutoSize = true;
            this.widthLabel.Location = new System.Drawing.Point(287, 361);
            this.widthLabel.Name = "widthLabel";
            this.widthLabel.Size = new System.Drawing.Size(72, 13);
            this.widthLabel.TabIndex = 22;
            this.widthLabel.Text = "Width of line: ";
            // 
            // typeCombobox
            // 
            this.typeCombobox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.typeCombobox.FormattingEnabled = true;
            this.typeCombobox.ItemHeight = 20;
            this.typeCombobox.Location = new System.Drawing.Point(365, 319);
            this.typeCombobox.Name = "typeCombobox";
            this.typeCombobox.Size = new System.Drawing.Size(157, 26);
            this.typeCombobox.TabIndex = 20;
            this.typeCombobox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.typeCombobox_DrawItem);
            this.typeCombobox.SelectedIndexChanged += new System.EventHandler(this.typeCombobox_SelectedIndexChanged);
            // 
            // typeLabel
            // 
            this.typeLabel.AutoSize = true;
            this.typeLabel.Location = new System.Drawing.Point(291, 322);
            this.typeLabel.Name = "typeLabel";
            this.typeLabel.Size = new System.Drawing.Size(68, 13);
            this.typeLabel.TabIndex = 19;
            this.typeLabel.Text = "Type of line: ";
            // 
            // widthTextBox
            // 
            this.widthTextBox.Location = new System.Drawing.Point(365, 361);
            this.widthTextBox.Name = "widthTextBox";
            this.widthTextBox.Size = new System.Drawing.Size(157, 20);
            this.widthTextBox.TabIndex = 21;
            this.widthTextBox.TextChanged += new System.EventHandler(this.widthTextBox_TextChanged);
            // 
            // SchoellerDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(235)))), ((int)(((byte)(250)))));
            this.ClientSize = new System.Drawing.Size(985, 488);
            this.Controls.Add(this.widthLabel);
            this.Controls.Add(this.typeCombobox);
            this.Controls.Add(this.typeLabel);
            this.Controls.Add(this.widthTextBox);
            this.Controls.Add(this.updateButton);
            this.Controls.Add(this.colorLabel);
            this.Controls.Add(this.colorPanel);
            this.Controls.Add(this.dgvJobsInDetails);
            this.Name = "SchoellerDetails";
            this.Text = "Schoeller Legend";
            this.Load += new System.EventHandler(this.SchoellerDetails_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvJobsInDetails)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.DataGridView dgvJobsInDetails;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn16;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn17;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn18;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn19;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn20;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn21;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn22;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn23;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn24;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn25;
        private System.Windows.Forms.Panel colorPanel;
        private System.Windows.Forms.Label colorLabel;
        public System.Windows.Forms.Button updateButton;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Label widthLabel;
        private System.Windows.Forms.ComboBox typeCombobox;
        private System.Windows.Forms.Label typeLabel;
        private System.Windows.Forms.TextBox widthTextBox;
    }
}