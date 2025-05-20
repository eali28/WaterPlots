using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WindowsFormsApplication2
{
    public class clsConstants
    {
        public const int legendTextSize = 10;
        public static float collinscol = 0.008f * frmMainForm.mainChartPlotting.Width;
        public static int chartYPowerpoint = 0;
        public static int metaY = (int)(0.13f * frmMainForm.mainChartPlotting.Height);
        public static List<string> clickedHeaders = new List<string>(); // List to store clicked headers
        public static List<clsJobs> oldData = new List<clsJobs>();
    }
}
