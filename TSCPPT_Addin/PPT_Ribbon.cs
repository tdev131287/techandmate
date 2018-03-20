using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using CustomTaskPanes = Microsoft.Office.Tools.CustomTaskPane;
using System.Drawing;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new PPT_Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace TSCPPT_Addin
{
    [ComVisible(true)]
    public class PPT_Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        string tscColors, cDataLabels, cYAxis;
        public PPT_Ribbon()
        {
        }
               
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("TSCPPT_Addin.PPT_Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            
        }
        public Bitmap setImage(Office.IRibbonControl rib)
        {
           switch(rib.Id)
            {
                case "customButton13": return  new Bitmap(PPTAttribute.PiCon + "Title  Slide.png");
                case "customButton14": return new Bitmap(PPTAttribute.PiCon + "Content Slide.png");
                case "customButton15": return new Bitmap(PPTAttribute.PiCon + "Section Heading.png");
                case "customButton16": return new Bitmap(PPTAttribute.PiCon + "Framework Slide.png");
                case "customButton17": return new Bitmap(PPTAttribute.PiCon + "Blank Slide.png");
                case "customButton18": return new Bitmap(PPTAttribute.PiCon + "Title  Slide.png");
                case "customButton21": return new Bitmap(PPTAttribute.PiCon + "Slide Title.png");
                case "customButton22": return new Bitmap(PPTAttribute.PiCon + "Slide Title.png");
                case "customButton23": return new Bitmap(PPTAttribute.PiCon + "Text Box.png");
                case "customButton24": return new Bitmap(PPTAttribute.PiCon + "Note Box.png");
                case "customButton25": return new Bitmap(PPTAttribute.PiCon + "Source Box.png");
                case "customButton26": return new Bitmap(PPTAttribute.PiCon + "Chart Title.png");
                case "customButton27": return new Bitmap(PPTAttribute.PiCon + "Quote Box.png");
                case "customButton31": return new Bitmap(PPTAttribute.PiCon + "Slide Title.png");
                case "customButton32": return new Bitmap(PPTAttribute.PiCon + "Slide Title.png");
                case "customButton33": return new Bitmap(PPTAttribute.PiCon + "Text Box.png");
                case "customButton34": return new Bitmap(PPTAttribute.PiCon + "Note Box.png");
                case "customButton35": return new Bitmap(PPTAttribute.PiCon + "Source Box.png");
                case "customButton36": return new Bitmap(PPTAttribute.PiCon + "Chart Title.png");
                case "customButton37": return new Bitmap(PPTAttribute.PiCon + "Quote Box.png");
                case "customButton41": return new Bitmap(PPTAttribute.PiCon + "Column chart.png");
                case "customButton42": return new Bitmap(PPTAttribute.PiCon + "Stacked Chart.png");
                case "customButton44": return new Bitmap(PPTAttribute.PiCon + "Pie chart.png");
                case "customButton11": return new Bitmap(PPTAttribute.PiCon + "New Theme1.png");
            }
       
            return null;
        }
        #region TSC PPT Callbacks Define 
        //Load TSC 2018 Theme in active presentation.
        public void TSCP_Callback11(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            funObj.ApplyPPT_Theme(rib);
            PPTAttribute.UserTracker(rib);
        }

        //Start a new presentation with TSC 2015 Theme.
        public void TSCP_Callback12(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            funObj.addNewPPT_In_tsc_format(rib);
            PPTAttribute.UserTracker(rib);
        }
        //Insert TSC formatted slides in active presentation.
        public void TSCP_Callback13(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            if (funObj.TSCThemeLoaded()) { funObj.Insert_Selected_Slide(rib); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle,MessageBoxButtons.OK,MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        //Insert slide components on active slide.
        public void TSCP_Callback_Insert(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            if (funObj.TSCThemeLoaded()) { funObj.insert_PPT_Object(rib); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);

        }
        public void TSCP_Callback_Format(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            if (funObj.TSCThemeLoaded()) { funObj.Format_PPT_Object(rib); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        public void TSCP_Callback_Chart(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            if (funObj.TSCThemeLoaded()) { funObj.InsertCharts(rib); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        //Growth Rates
        public void TSCP_Callback51(Office.IRibbonControl rib)
        {
            PPT.Application ppApp = Globals.ThisAddIn.Application;
            pptfunctions funObj = new pptfunctions();
            Shapecheck shpObj = new Shapecheck();
            frmChartcalc chartObj = new frmChartcalc();
            //frmCalculator frmObj = new frmCalculator();
            Shapecheck PPTshpchk = new Shapecheck();
            List<string> SelectedCharts = new List<string>();
            SelectedCharts = PPTshpchk.FindSelectedCharts();
            PPT.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
            int numSelCht = SelectedCharts.Count;
            if (numSelCht == 0) {
                MessageBox.Show("Please select a  chart for CAGR/AAGR calculation.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int sld_num = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
            PPT.Chart myChart = ActivePPT.Slides[sld_num].Shapes[SelectedCharts[0]].Chart;
            string chType = shpObj.chartType(myChart);
            if (funObj.TSCThemeLoaded())
            {
                if (numSelCht == 0)
                {
                    //frmObj.Show();
                    MessageBox.Show("Please select a  chart CAGR/AAGR calculation.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (numSelCht > 1)
                {

                    MessageBox.Show("Please select a single chart CAGR/AAGR calculation.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if(numSelCht == 1)
                {
                    //PPT.Chart myChart = 
                    //string chType = PPTshpchk.chartType();
                    if (chType != "Pie")
                    {
                        chartObj.ShowDialog();
                    }
                    else
                    {
                        MessageBox.Show("Please select a column chart for CAGR/AAGR Calculation.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        //Sum of  Pie Chart
        public void TSCP_Callback52(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            
            if (funObj.TSCThemeLoaded()) { funObj.SumPieChart(); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }

        public void TSCP_Callback71(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            frmChangeLanguage lngObj = new frmChangeLanguage();
            if (funObj.TSCThemeLoaded()) { lngObj.Show(); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }

        public void TSCP_Callback61(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            frmReviewformat lngObj = new frmReviewformat();
            if (funObj.TSCThemeLoaded()) { lngObj.Show(); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        public void TSCP_Callback46(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
           
            if (funObj.TSCThemeLoaded()) { funObj.formatChart(tscColors, cDataLabels, cYAxis,rib); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }


        public void TSCP_Callback82(Office.IRibbonControl rib)
        {
            tscformat choObj = new tscformat();
            choObj.tsc_loadtheme();
        }

        public string GetSelectedItem4D1(Office.IRibbonControl rib, int index)
        {
            tscColors = "item1";
            return (tscColors);
        }
        public void TSCP_Callback4D1(Office.IRibbonControl rib,string selectedID, int selectedIndex)
        {
            tscColors = selectedID;
        }
        // ----
        public string GetSelectedItem4D2(Office.IRibbonControl rib, int index)
        {
            cDataLabels = "item6";
            return (cDataLabels);
        }
        public void TSCP_Callback4D2(Office.IRibbonControl rib, string selectedID, int selectedIndex)
        {
            cDataLabels = selectedID;
        }
        //----
        public string GetSelectedItem4D3(Office.IRibbonControl rib, int index)
        {
            cYAxis = "item9";
            return (cYAxis);
        }
        public void TSCP_Callback4D3(Office.IRibbonControl rib, string selectedID, int selectedIndex)
        {
            cYAxis = selectedID;
        }

        public void TSCP_Callback81(Office.IRibbonControl rib)
        {
            PPTAttribute.SQLConnection();
        }
        #endregion

        #endregion

        #region Helpers
        public void btnloadtheme_Clicked(Office.IRibbonControl rib)
        {
            //MessageBox.Show("Call me");
        }
        public void PPTdictionary_click(Office.IRibbonControl rib)
        {
           
            //Set objPane = Globals.ThisAddIn.Application.ActivePresentation.
            navigationCtrl myUserControl1 = new navigationCtrl();
            //MessageBox.Show("Try to add Navigation Pane");
            
            //Microsoft.Office.Tools.CustomTaskPane tskObj = new Microsoft.Office.Tools.CustomTaskPane;

        }
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public void TSCP_Callback62(Office.IRibbonControl rib)
        {
            frmEditorialReview eReview = new frmEditorialReview();
            eReview.Show();
        }
        #endregion
    }
}
