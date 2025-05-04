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

namespace WindowsFormsApplication2
{
    public class ButtonControls : Button
    {
        public static Button MyButton { get; set; }
        public ButtonControls(int X, int Y, Size ButtonSize, int i)
        {
            MyButton = new Button
            {
                Location = new Point(X * ButtonSize.Width, Y * ButtonSize.Height),
                Name = X + "." + Y,
                BackColor = Color.White,
                Size = ButtonSize,
                Font = new Font("Times New Roman", 6),
                Text = "Change line_" + i
            };
        }

        public static int coordsX { get; set; }
        public static int coordsY { get; set; }
        //public string Name { get; set; }
        //public Color BackColor { get; set; }
        public Size size { get; set; }
        public Font font { get; set; }
        //public string Text { get; set; }
        //public PointF Location { get; set; }
    }
}
