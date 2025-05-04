namespace WindowsFormsApplication2
{
    partial class frmRadarScales
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
            this.ScalesDatagridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.updateButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.ScalesDatagridView)).BeginInit();
            this.SuspendLayout();
            // 
            // ScalesDatagridView
            // 
            this.ScalesDatagridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ScalesDatagridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn11,
            this.dataGridViewTextBoxColumn12});
            this.ScalesDatagridView.Location = new System.Drawing.Point(34, 12);
            this.ScalesDatagridView.Name = "ScalesDatagridView";
            this.ScalesDatagridView.Size = new System.Drawing.Size(244, 358);
            this.ScalesDatagridView.TabIndex = 13;
            // 
            // dataGridViewTextBoxColumn11
            // 
            this.dataGridViewTextBoxColumn11.HeaderText = "Item";
            this.dataGridViewTextBoxColumn11.Name = "dataGridViewTextBoxColumn11";
            // 
            // dataGridViewTextBoxColumn12
            // 
            this.dataGridViewTextBoxColumn12.HeaderText = "Scale";
            this.dataGridViewTextBoxColumn12.Name = "dataGridViewTextBoxColumn12";
            // 
            // updateButton
            // 
            this.updateButton.Location = new System.Drawing.Point(106, 544);
            this.updateButton.Name = "updateButton";
            this.updateButton.Size = new System.Drawing.Size(75, 23);
            this.updateButton.TabIndex = 14;
            this.updateButton.Text = "Apply";
            this.updateButton.UseVisualStyleBackColor = true;
            this.updateButton.Click += new System.EventHandler(this.updateButton_Click);
            // 
            // frmRadarScales
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(235)))), ((int)(((byte)(250)))));
            this.ClientSize = new System.Drawing.Size(330, 589);
            this.Controls.Add(this.updateButton);
            this.Controls.Add(this.ScalesDatagridView);
            this.Name = "frmRadarScales";
            this.Text = "frmRadarScales";
            ((System.ComponentModel.ISupportInitialize)(this.ScalesDatagridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.DataGridView ScalesDatagridView;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn11;
        public System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn12;
        public System.Windows.Forms.Button updateButton;
    }
}