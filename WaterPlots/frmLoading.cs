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
    public partial class frmLoading : Form
    {
        private Label messageLabel;
        private PictureBox spinner;
        public frmLoading()
        {
            InitializeComponent();
            BuildUI();
        }
        private void BuildUI()
        {
            // Form setup
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = Color.White;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(300, 200);
            this.TopMost = true;

            // Spinner setup
            spinner = new PictureBox();
            spinner.Size = new Size(64, 64);
            spinner.Location = new Point((this.ClientSize.Width - spinner.Width) / 2, 30);
            spinner.SizeMode = PictureBoxSizeMode.StretchImage;

            // You can use a loading GIF in your resources
            spinner.Image = Properties.Resources.spinner; // Add a GIF to your Resources.resx

            // Label
            messageLabel = new Label();
            messageLabel.Text = "Please wait...";
            messageLabel.Font = new Font("Segoe UI", 10, FontStyle.Regular);
            messageLabel.ForeColor = Color.DimGray;
            messageLabel.TextAlign = ContentAlignment.MiddleCenter;
            messageLabel.Dock = DockStyle.Bottom;
            messageLabel.Padding = new Padding(0, 0, 0, 20);
            messageLabel.Height = 40;

            this.Controls.Add(spinner);
            this.Controls.Add(messageLabel);
        }

        public void SetMessage(string message)
        {
            if (messageLabel != null)
                messageLabel.Text = message;
        }
    }
}
