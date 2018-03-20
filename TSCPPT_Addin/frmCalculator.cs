using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace TSCPPT_Addin
{
    public partial class frmCalculator : Form
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        public frmCalculator()
        {
            InitializeComponent();
        }

        private void frmCalculator_Load(object sender, EventArgs e)
        {
            cmb_method.Items.Add("Compound");
            cmb_method.Items.Add("Average");
            cmb_method.Items.Add("Simple");
            
            txt_stValue.Text = Convert.ToString(1000);
            txt_endValue.Text = Convert.ToString(5000);
            txt_Period.Text= Convert.ToString(5);

            txt_Period.BackColor = System.Drawing.Color.FromArgb(255, 255, 204);
            txt_endValue.BackColor = System.Drawing.Color.FromArgb(255, 255, 204);
            txt_stValue.BackColor = System.Drawing.Color.FromArgb(255, 255, 204);
            cmb_method.SelectedIndex = 0;

        }

        private void txt_Period_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string tText = txt_Period.Text.Trim();
                if (!String.IsNullOrEmpty(tText)) { Calculate_CAGRV(); }
            }
            catch(Exception ex) { }
        }

        public void Calculate_CAGRV()
        {
            string cCalc = cmb_method.Text;
            int sValue = Convert.ToInt32(txt_stValue.Text);
            int eValue = Convert.ToInt32(txt_endValue.Text);
            int nPeriod = Convert.ToInt32(txt_Period.Text);
            int gCalc = 0;
            if (cCalc=="Compound")
            {
                lbl_text.Text = "CAGR =";
                if(nPeriod>0)
                {
                    if(eValue>0 && sValue > 0) { gCalc = ((eValue / sValue) ^ (1 / (nPeriod))) - 1; }
                    if (gCalc >= 0) { lbl_text.Text = "CARG ="; }
                    if (gCalc <= 0) { lbl_text.Text = "CARC ="; }
                }
                else { gCalc = 0; }
                
            }
            else if(cCalc=="Average")
            {
                if (nPeriod > 0)
                {
                    if (sValue > 0) { gCalc = (eValue / sValue - 1) / nPeriod; }
                    else { gCalc = -(eValue / sValue - 1) / nPeriod; }

                    if (gCalc >= 0) { lbl_text.Text = "AAGR ="; }
                    if (gCalc < 0) { lbl_text.Text = "AARC ="; }
                }
                else { gCalc = 0; }
            }
            else if (cCalc == "Simple")
            {
                if (nPeriod > 0)
                {
                    if (sValue > 0) { gCalc = (eValue / sValue - 1); }
                    else { gCalc = -(eValue / sValue - 1); }

                    if (gCalc >= 0) { lbl_text.Text = "Growth ="; }
                    if (gCalc < 0) { lbl_text.Text = "Decline ="; }
                }
                else { gCalc = 0; }
            }
            //lbl_Value.Text = String.Format(Convert.ToString(gCalc), "P");
            lbl_Value.Text = gCalc.ToString("P");
        }

        private void cmb_method_SelectedIndexChanged(object sender, EventArgs e)
        {
            Calculate_CAGRV();
            
        }

        private void txt_stValue_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string tText = txt_stValue.Text.Trim();
                if (!String.IsNullOrEmpty(tText)) { Calculate_CAGRV(); }
            }
            catch(Exception ex) { }
        }

        private void txt_endValue_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string tText = txt_endValue.Text.Trim();
                if (!String.IsNullOrEmpty(tText)) { Calculate_CAGRV(); }
            }
            catch(Exception ex) { }
        }

        private void btnCalculater_Click(object sender, EventArgs e)
        {
            string cCalc = cmb_method.Text;
            string gCalc = lbl_Value.Text;
            string dText=null;
            if (cCalc == "Compound")
            {
                if (gCalc.PadLeft(1) != "-") { dText = "CAGR : " + String.Format(Convert.ToString(gCalc), "P"); }
                if (gCalc.PadLeft(1) == "-") { dText = "CARC : " + String.Format(Convert.ToString(gCalc), "P"); }
            }
            else if (cCalc == "Average") { dText = "AAGR : " + String.Format(Convert.ToString(gCalc), "P"); }
            else if(cCalc == "Simple") { dText = "Growth : " + String.Format(Convert.ToString(gCalc), "P"); }
            float lf = ppApp.ActivePresentation.PageSetup.SlideWidth;
            PowerPoint.Shape aShp = ppApp.ActiveWindow.View.Slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, (lf - 115), 0, 115, 18.13);
            aShp.TextFrame.TextRange.Text = dText;
            aShp.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            aShp.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
            aShp.TextFrame.TextRange.Font.Size = 12;
            aShp.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.FromArgb(0,0,0).ToArgb();
            aShp.TextFrame.TextRange.Font.Name = "Calibri";
            aShp.Fill.Visible = MsoTriState.msoFalse;
            aShp.Line.Weight = (float) 0.75;
            aShp.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
            aShp.Line.Style = MsoLineStyle.msoLineSingle;
            aShp.Line.Transparency = 1;
            aShp.Line.Visible = MsoTriState.msoTrue;
            aShp.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(23, 94, 84).ToArgb();
            aShp.Line.BackColor.RGB = System.Drawing.Color.FromArgb(23, 94, 84).ToArgb();
            //aShp.Name = shp_nam + " CAGR Box";
            aShp.Name = " CAGR Box1";
            this.Close();

        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                this.Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
