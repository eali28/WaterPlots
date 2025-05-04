namespace WindowsFormsApplication2
{
    partial class frmCollinsLegend
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
            this.updateButton = new System.Windows.Forms.Button();
            this.ItemComboBox = new System.Windows.Forms.ComboBox();
            this.colorLabel = new System.Windows.Forms.Label();
            this.colorPanel = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // updateButton
            // 
            this.updateButton.Location = new System.Drawing.Point(147, 104);
            this.updateButton.Name = "updateButton";
            this.updateButton.Size = new System.Drawing.Size(75, 23);
            this.updateButton.TabIndex = 23;
            this.updateButton.Text = "Apply";
            this.updateButton.UseVisualStyleBackColor = true;
            this.updateButton.Click += new System.EventHandler(this.updateButton_Click);
            // 
            // ItemComboBox
            // 
            this.ItemComboBox.FormattingEnabled = true;
            this.ItemComboBox.ItemHeight = 13;
            this.ItemComboBox.Location = new System.Drawing.Point(99, 55);
            this.ItemComboBox.Name = "ItemComboBox";
            this.ItemComboBox.Size = new System.Drawing.Size(78, 21);
            this.ItemComboBox.TabIndex = 22;
            this.ItemComboBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.typeCombobox_DrawItem);
            this.ItemComboBox.SelectedIndexChanged += new System.EventHandler(this.typeCombobox_SelectedIndexChanged);
            // 
            // colorLabel
            // 
            this.colorLabel.AutoSize = true;
            this.colorLabel.Location = new System.Drawing.Point(209, 55);
            this.colorLabel.Name = "colorLabel";
            this.colorLabel.Size = new System.Drawing.Size(37, 13);
            this.colorLabel.TabIndex = 21;
            this.colorLabel.Text = "Color: ";
            // 
            // colorPanel
            // 
            this.colorPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.colorPanel.Location = new System.Drawing.Point(273, 55);
            this.colorPanel.Name = "colorPanel";
            this.colorPanel.Size = new System.Drawing.Size(64, 20);
            this.colorPanel.TabIndex = 20;
            this.colorPanel.Click += new System.EventHandler(this.colorPanel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 58);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 13);
            this.label1.TabIndex = 24;
            this.label1.Text = "Element:";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // frmCollinsLegend
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(235)))), ((int)(((byte)(250)))));
            this.ClientSize = new System.Drawing.Size(378, 157);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.updateButton);
            this.Controls.Add(this.ItemComboBox);
            this.Controls.Add(this.colorLabel);
            this.Controls.Add(this.colorPanel);
            this.Name = "frmCollinsLegend";
            this.Text = "Collins Legend";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ColorDialog colorDialog1;
        public System.Windows.Forms.Button updateButton;
        private System.Windows.Forms.ComboBox ItemComboBox;
        private System.Windows.Forms.Label colorLabel;
        private System.Windows.Forms.Panel colorPanel;
        private System.Windows.Forms.Label label1;
    }
}