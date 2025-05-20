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
    public partial class frmPieLegend : Form
    {
        string selectedItem;
        public static Color selectedColor;
        public static bool IsUpdateClicked { get; private set; }
        public static List<Color> PieColor = new List<Color>();
        public static string[] types = { "Na+K", "Ca", "Mg", "Cl", "SO4", "HCO3 + CO3" };
        public frmPieLegend()
        {
            
            InitializeComponent();
            LoadTypesIntoComboBox();
            foreach (Color color in clsPieDrawer.pieColors)
            {
                PieColor.Add(color);
            }
        }
        private void LoadTypesIntoComboBox()
        {
            // Clear previous items
            typeCombobox.Items.Clear();

            
            foreach (var item in types)
            {
                typeCombobox.Items.Add(item);
                
            }
            
        }
        
        private void typeCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (typeCombobox.SelectedItem != null)
            {
                selectedItem = (string)typeCombobox.SelectedItem;
                for (int i = 0; i < types.Length; i++)
                {
                    if (selectedItem == types[i])
                    {
                        
                        this.colorPanel.BackColor = PieColor[i];
                    }
                }
            }
        }
        private void typeCombobox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            e.DrawBackground();
            e.Graphics.DrawString(typeCombobox.Items[e.Index].ToString(),
                                  e.Font,
                                  Brushes.Black,
                                  e.Bounds);
            e.DrawFocusRectangle();
        }


        private void colorPanel_Click(object sender, EventArgs e)
        {
            this.Activate();  // Activate the form
            this.BringToFront();  // Bring the form to the front
            this.Focus();  // Ensure the form has focus

            if (colorDialog1.ShowDialog(this) == DialogResult.OK)
            {
                selectedColor = colorDialog1.Color;
                colorPanel.BackColor = selectedColor;
            }
            this.Invalidate();  // Redraw the form if necessary
        }


        private void updateButton_Click(object sender, EventArgs e)
        {
            frmMainForm.selectedSamples.Clear();
            IsUpdateClicked = true;
            for (int i = 0; i < clsPieDrawer.pieColors.Length; i++)
            {
                if (selectedItem == clsPieDrawer.labels[i])
                {
                    PieColor[i] = selectedColor;
                }
            }
            this.Close();
        }
    }
}