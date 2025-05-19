using System.Drawing;
using System.Drawing.Drawing2D;
namespace WindowsFormsApplication2
{
    public class clsWater
    {
        public string sampleID;
        public double Na; 
        public double K;
        public double Ca;
        public double Mg;
        public double So4;
        public double HCO3;
        public double CO3; 
        public double Cl;
        public double Sr;
        public double Ba;
        public double B;
        public double TDS;
        public double Al;
        public double Co;
        public double Cu;
        public double Mn;
        public double Ni;
        public double Zn;
        public double Pb;
        public double Fe;
        public double Cd;
        public double Cr;
        public double Tl;
        public double Be;
        public double Se;
        public double Li;
        public string Well_Name;
        public string Depth; 
        public string ClientID;
        public string sampleType;
        public string Label;
        public string ID;
        public string jobID;
        public string latitude;
        public string longtude;
        public string formName;
        public string prep;
        public bool bubble = false;
        public bool piper=false;
        public Color color;
        public DashStyle selectedStyle = DashStyle.Solid;
        public float lineWidth = 2; // Default line width
        public string shape=null;
    }
}
