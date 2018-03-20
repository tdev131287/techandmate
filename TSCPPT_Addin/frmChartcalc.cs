using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace TSCPPT_Addin
{
    public partial class frmChartcalc : Form
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        Shapecheck PPTshpchk = new Shapecheck();
        string[,] chtData;
        int num_points, num_series;
        public frmChartcalc()
        {
            InitializeComponent();
        }

        private void frmChartcalc_Load(object sender, EventArgs e)
        {
            List<string> SelectedCharts = new List<string>();
            List<String> xxVals = new List<string>();
            List<String> yyVals = new List<string>();
            List<string> Seriesname = new List<string>();
            try
            {
                int sld_num = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
                SelectedCharts = PPTshpchk.FindSelectedCharts();
                string shp_nam = SelectedCharts[0];
                PowerPoint.Chart myChart = ActivePPT.Slides[sld_num].Shapes[shp_nam].Chart;
                PowerPoint.SeriesCollection series = ActivePPT.Slides[sld_num].Shapes[shp_nam].Chart.SeriesCollection();
                num_points = ActivePPT.Slides[sld_num].Shapes[shp_nam].Chart.SeriesCollection(1).Points.Count;
                num_series = series.Count;
                chtData = new string[num_points, num_series + 1];

                for (int x = 0; x < num_series; x++)
                {
                    string sName = myChart.SeriesCollection(x + 1).Name;
                    Seriesname.Add(sName);
                    PowerPoint.Series tseries = (PowerPoint.Series)myChart.SeriesCollection(x + 1);
                    xxVals.Clear();
                    foreach (object item in tseries.Values as Array) { xxVals.Add(Convert.ToString(item)); }
                    string[] yVals = xxVals.ToArray();
                    yyVals.Clear();
                    foreach (object item in tseries.XValues as Array) { yyVals.Add(Convert.ToString(item)); }
                    string[] xVals = yyVals.ToArray();
                    //string[] yVals = myChart.SeriesCollection(x).Values.toArray();
                    //string[] xVals = myChart.SeriesCollection(x).XValues.toArray();
                    //chtData[0, x ] = sName;
                    for (int y = 0; y < yVals.Length; y++)
                    {
                        chtData[y, 0] = xVals[y];
                        chtData[y, x + 1] = yVals[y];
                    }
                }
                //for(int x=0;x< chtData.Length; x++) { cmb_stDate.Items.Add(Convert.ToString(chtData[x, 0])); }
                //for (int x = 0; x < chtData.Length; x++) { cmb_stDate.Items.Add(Convert.ToString(chtData[x, 0])); }
                cmb_stDate.Items.Clear();
                cmb_endDate.Items.Clear();
                cmb_Series.Items.Clear();
                foreach (string item in yyVals)
                {
                    cmb_stDate.Items.Add(item);
                    cmb_stDate.SelectedIndex = 0;
                    //cmb_endDate.Items.Add(item);
                    //cmb_endDate.SelectedIndex = 0;
                }
                txt_period.Enabled = false;

                foreach (string item in Seriesname) { cmb_Series.Items.Add(item); cmb_Series.SelectedIndex = 0; }
                //Calculate_CAGR();
                string cmbText = cmb_calcType.Text;
                cmb_calcType.Items.Add("CAGR");
                cmb_calcType.Items.Add("AAGR");
                cmb_calcType.SelectedIndex = 0;
                if (cmbText == "AAGR") { Calculate_AAGR(); }
                else { Calculate_CAGR(); }
            }
            catch (Exception err)
            {
                this.Close();
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "frmChartcalc_Load");
                MessageBox.Show("Check chart type and chart value", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        public void Calculate_CAGR(int diffVal = 0)
        {
            string cSeries = cmb_Series.Text;
            string pStart = cmb_stDate.Text;
            int stVal = 0, endVal = 0;
            float CAGR = 0;
            string pEnd = cmb_endDate.Text;
            try
            {
                int type = cmb_Series.SelectedIndex + 1;
                for (int x = 0; x < num_points; x++) { if (pStart == chtData[x, 0]) { stVal = x; } }
                for (int x = 0; x < num_points; x++)
                {
                    if (pEnd == chtData[x, 0])
                    {
                        endVal = x;
                    }
                }
                float sValue = (float)Convert.ToDouble(chtData[stVal, type]);
                float eValue = (float)Convert.ToDouble(chtData[endVal, type]);
                if (diffVal == 0) { diffVal = endVal - stVal; }


                txt_period.Text = Convert.ToString(diffVal);
                if (diffVal > 0)
                {
                    if (eValue >= 0 && sValue >= 0)
                    {
                        float xx = (eValue / sValue);
                        float yy = (float)1 / diffVal;
                        CAGR = (float)Math.Pow(xx, yy);
                        CAGR = CAGR - 1;
                        //CAGR = (xx ^ yy)-1;
                        if (CAGR < 0) { lbl_type.Text = "CARC ="; }
                        if (CAGR >= 0) { lbl_type.Text = "CAGR ="; }
                        lbl_Value.Text = CAGR.ToString("P");
                    }
                    else
                    {
                        lbl_Value.Text = "NA";
                    }

                }
                else
                {
                    CAGR = 0;
                    lbl_Value.Text = CAGR.ToString("P");
                }
            }
            catch (Exception err)
            {
                this.Close();
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Calculate_CAGR");
                MessageBox.Show("Check chart type and chart value", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void Calculate_AAGR()
        {
            List<float> diffVal = new List<float>();
            lbl_Value.Text = "";
            float sum = 0;
            int stYear, lYear;
            try
            {
                int type = cmb_Series.SelectedIndex + 1;
                stYear = Convert.ToInt32(cmb_stDate.Text);
                lYear = Convert.ToInt32(cmb_endDate.Text);
                for (int x = 1; x < num_points; x++)
                {
                    int year = (int)Convert.ToDouble(chtData[x, 0]);
                    if (year >= stYear && year <= lYear)
                    {
                        float diff = (float)Convert.ToDouble(chtData[x, type]) / (float)Convert.ToDouble(chtData[x - 1, type]);
                        diff = diff - 1;
                        diffVal.Add(diff);
                    }
                }
                foreach (float x in diffVal)
                {
                    sum = sum + x;
                }
                float AAGR = sum / diffVal.Count;
                lbl_type.Text = "AAGR";
                lbl_Value.Text = AAGR.ToString("P");
            }
            catch (Exception err)
            {
                this.Close();
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Calculate_AAGR");
                MessageBox.Show("Check chart type and chart value", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void cmb_stDate_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int l = cmb_stDate.SelectedIndex;
                cmb_endDate.Items.Clear();
                for (int i = l; i < num_points; i++) { cmb_endDate.Items.Add(chtData[i, 0]); }
                cmb_endDate.SelectedIndex = cmb_endDate.Items.Count - 1;
                //Calculate_CAGR();

                string cmbText = cmb_calcType.Text;
                if (cmbText == "AAGR") { Calculate_AAGR(); }
                else { Calculate_CAGR(); }
            }
            catch (Exception err)
            {
                this.Close();
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "cmb_stDate_SelectedIndexChanged");
                MessageBox.Show("Check chart type and chart value", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmb_endDate_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cmbText = cmb_calcType.Text;
            if (cmbText == "AAGR") { Calculate_AAGR(); }
            else { Calculate_CAGR(); }
        }

        private void btn_Caption_Click(object sender, EventArgs e)
        {
            
            int sld_num = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
            float top_adjust = 0;
            int cnt = 0;
            string CAGR, dText = null;
            List<string> SelectedCharts = new List<string>();
            try
            {
                SelectedCharts = PPTshpchk.FindSelectedCharts();
                string shp_nam = SelectedCharts[0];
                PowerPoint.Shape myShape = ActivePPT.Slides[sld_num].Shapes[shp_nam];
                int num_shapes = ActivePPT.Slides[sld_num].Shapes.Count;
                for (int i = 1; i <= num_shapes; i++)
                {
                    if (ActivePPT.Slides[sld_num].Shapes[i].Name == shp_nam + " CAGR Box") { cnt++; }
                }
                if (cnt > 0) { top_adjust = (float)18.13 * cnt + 5; }
                CAGR = lbl_Value.Text;
                string cmbText = cmb_calcType.Text;
                if (cmbText == "AAGR")
                {
                    dText = "AAGR (" + cmb_stDate.Text + "–" + cmb_endDate.Text + "): " + CAGR;
                }
                else
                {
                    if (CAGR.Substring(0, 1) != "-")
                    {
                        dText = "CAGR (" + cmb_stDate.Text + "–" + cmb_endDate.Text + "): " + CAGR;
                    }
                    else if (CAGR.Substring(0, 1) == "-")
                    {
                        dText = "CARC (" + cmb_stDate.Text + "–" + cmb_endDate.Text + "): " + CAGR;
                    }
                }

                float tp = ActivePPT.Slides[sld_num].Shapes[shp_nam].Top;
                PowerPoint.Shape aShp = ActivePPT.Slides[sld_num].Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, myShape.Left + myShape.Width, myShape.Top + top_adjust, 120, (float)23.5);
                aShp.TextFrame.TextRange.Text = dText;
                aShp.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
                aShp.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                aShp.TextFrame.TextRange.Font.Size = 12;
                aShp.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                aShp.TextFrame.TextRange.Font.Name = "Calibri";
                aShp.Fill.Visible = MsoTriState.msoFalse;
                aShp.Line.Weight = (float)0.75;
                aShp.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
                aShp.Line.Style = MsoLineStyle.msoLineSingle;
                aShp.Line.Transparency = 0;
                aShp.Line.Visible = MsoTriState.msoTrue;
                aShp.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(23, 94, 84).ToArgb();
                aShp.Line.BackColor.RGB = System.Drawing.Color.FromArgb(23, 94, 84).ToArgb();
                aShp.Name = shp_nam + " CAGR Box";
                this.Close();
            }
            catch (Exception err)
            {
                this.Close();
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "btn_Caption_Click");
                MessageBox.Show("Check chart type and chart value", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmb_Series_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cmbText = cmb_calcType.Text;
            if (cmbText == "AAGR") { Calculate_AAGR(); }
            else { Calculate_CAGR(); }
        }

        private void chk_Period_CheckedChanged(object sender, EventArgs e)
        {
            if(chk_Period.Checked==true)
            {
                cmb_endDate.Enabled = false;
                cmb_endDate.BackColor= System.Drawing.Color.FromArgb(242, 242, 242);
                cmb_endDate.ForeColor = System.Drawing.Color.FromArgb(0,0,0);
                txt_period.Enabled = true;
                txt_period.BackColor = System.Drawing.Color.FromArgb(255, 255, 204);
            }
            else
            {
                cmb_stDate.Enabled = true;
                cmb_endDate.Enabled = true;
                cmb_stDate.BackColor= System.Drawing.Color.FromArgb(255, 255, 204);
                cmb_endDate.BackColor = System.Drawing.Color.FromArgb(255, 255, 204);
                txt_period.Enabled = false;
                txt_period.BackColor = System.Drawing.Color.FromArgb(242, 242, 242);

            }
            Calculate_CAGR();
        }

        private void txt_period_TextChanged(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(txt_period.Text)==false)
            {
                int val = Convert.ToInt32(txt_period.Text);
                //Calculate_CAGR(val);
                string cmbText = cmb_calcType.Text;
                if (cmbText == "AAGR") { Calculate_AAGR(); }
                else { Calculate_CAGR(val); }
            }

        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cmb_calcType_SelectedIndexChanged(object sender, EventArgs e)
        {
            string cmbText = cmb_calcType.Text;
            if (cmbText == "AAGR") { Calculate_AAGR(); }
            else { Calculate_CAGR(); }
        }
    }
}
