namespace WindowsFormsApplication2
{
    partial class frmPieLegend
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
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.colorLabel = new System.Windows.Forms.Label();
            this.colorPanel = new System.Windows.Forms.Panel();
            this.typeCombobox = new System.Windows.Forms.ComboBox();
            this.updateButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // colorLabel
            // 
            this.colorLabel.AutoSize = true;
            this.colorLabel.Location = new System.Drawing.Point(188, 41);
            this.colorLabel.Name = "colorLabel";
            this.colorLabel.Size = new System.Drawing.Size(37, 13);
            this.colorLabel.TabIndex = 16;
            this.colorLabel.Text = "Color: ";
            // 
            // colorPanel
            // 
            this.colorPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.colorPanel.Location = new System.Drawing.Point(252, 41);
            this.colorPanel.Name = "colorPanel";
            this.colorPanel.Size = new System.Drawing.Size(77, 20);
            this.colorPanel.TabIndex = 15;
            this.colorPanel.Click += new System.EventHandler(this.colorPanel_Click);
            // 
            // typeCombobox
            // 
            this.typeCombobox.FormattingEnabled = true;
            this.typeCombobox.ItemHeight = 13;
            this.typeCombobox.Location = new System.Drawing.Point(65, 40);
            this.typeCombobox.Name = "typeCombobox";
            this.typeCombobox.Size = new System.Drawing.Size(91, 21);
            this.typeCombobox.TabIndex = 18;
            this.typeCombobox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.typeCombobox_DrawItem);
            this.typeCombobox.SelectedIndexChanged += new System.EventHandler(this.typeCombobox_SelectedIndexChanged);
            // 
            // updateButton
            // 
            this.updateButton.Location = new System.Drawing.Point(154, 88);
            this.updateButton.Name = "updateButton";
            this.updateButton.Size = new System.Drawing.Size(75, 23);
            this.updateButton.TabIndex = 19;
            this.updateButton.Text = "Apply";
            this.updateButton.UseVisualStyleBackColor = true;
            this.updateButton.Click += new System.EventHandler(this.updateButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 20;
            this.label1.Text = "Element:";
            // 
            // frmPieLegend
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(235)))), ((int)(((byte)(250)))));
            this.ClientSize = new System.Drawing.Size(372, 136);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.updateButton);
            this.Controls.Add(this.typeCombobox);
            this.Controls.Add(this.colorLabel);
            this.Controls.Add(this.colorPanel);
            this.Name = "frmPieLegend";
            this.Text = "Pie Legend";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Label colorLabel;
        private System.Windows.Forms.Panel colorPanel;
        private System.Windows.Forms.ComboBox typeCombobox;
        public System.Windows.Forms.Button updateButton;
        private System.Windows.Forms.Label label1;
    }
}