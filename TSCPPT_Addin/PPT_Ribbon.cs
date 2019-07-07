using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using CustomTaskPanes = Microsoft.Office.Tools.CustomTaskPane;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.Diagnostics;


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

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            
        }
        public Bitmap OnLoadImage(string imageName)
        {
            return new Bitmap(PPTAttribute.themeColor + imageName);
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
                case "customButton28": return new Bitmap(PPTAttribute.PiCon + "Slide Title.png");
                case "btnbullet1": return new Bitmap(PPTAttribute.PiCon + "bullet1.png");
                case "btnbullet2": return new Bitmap(PPTAttribute.PiCon + "bullet2.png");
                case "btnbullet3": return new Bitmap(PPTAttribute.PiCon + "bullet3.png");
                case "btnformatPPT": return new Bitmap(PPTAttribute.PiCon + "New Theme1.png");
                case "galTest": return new Bitmap(PPTAttribute.PiCon + "color scheme.png");
                case "btnAlignCenter": return new Bitmap(PPTAttribute.PiCon + "VerticalCenter.png");
                case "menu41": return new Bitmap(PPTAttribute.PiCon + "InsertChart.png");
                case "menu11": return new Bitmap(PPTAttribute.PiCon + "Add slide1.png");
                case "menuRemove": return new Bitmap(PPTAttribute.PiCon + "IconClean.png");
                case "btnRemoveDoubleSpace": return new Bitmap(PPTAttribute.PiCon + "IconRDspace.png");

            }
            ribbon.Invalidate(); 
            return null;
        }
        public void buttonInvalidate()
        {
            //this.ribbon.Invalidate();
        }

        
        #region TSC PPT Callbacks Define 
        //Load TSC 2018 Theme in active presentation.
        public void TSCP_Callback11(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            funObj.ApplyPPT_Theme(rib);
            PPTAttribute.UserTracker(rib);
            //this.ribbon.InvalidateControl("customButton13");
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
            PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
            pptfunctions funObj = new pptfunctions();
            Shapecheck shpObj = new Shapecheck();
            frmChartcalc chartObj = new frmChartcalc();
            //frmCalculator frmObj = new frmCalculator();
            Shapecheck PPTshpchk = new Shapecheck();
            List<string> SelectedCharts = new List<string>();
            SelectedCharts = PPTshpchk.FindSelectedCharts();
            PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
            int numSelCht = SelectedCharts.Count;
            if (numSelCht == 0) {
                MessageBox.Show("Please select a chart for CAGR/AAGR calculation.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int sld_num = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
            PowerPoint.Chart myChart = ActivePPT.Slides[sld_num].Shapes[SelectedCharts[0]].Chart;
            string chType = shpObj.chartType(myChart);
            if (funObj.TSCThemeLoaded())
            {
                if (numSelCht == 0)
                {
                    //frmObj.Show();
                    MessageBox.Show("Please select a chart CAGR/AAGR calculation.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            if (funObj.TSCThemeLoaded()) { lngObj.ShowDialog(); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }

        public void TSCP_Callback61(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            frmReviewformat lngObj = new frmReviewformat();
            if (funObj.TSCThemeLoaded()) { lngObj.ShowDialog(); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        public void TSCP_Callback46(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
           
            if (funObj.TSCThemeLoaded())
            {
                frmChartformat chartobj = new frmChartformat();
                chartobj.ShowDialog();
                //funObj.formatChart(tscColors, cDataLabels, cYAxis,rib);
            }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }

        public void btnbullet_Click(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();

            if (funObj.TSCThemeLoaded()) { funObj.formatbullettxt(rib); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        public void btnformatPPT_Click(Office.IRibbonControl rib)
        {
            frmPPTFormat frmObj = new frmPPTFormat();
            pptfunctions funObj = new pptfunctions();
            frmObj.ShowDialog();
            PPTAttribute.UserTracker(rib);

        }
        public void ChartGalleryOnAction(Office.IRibbonControl rib, string galleryID, int selectedIndex)
        {
            pptfunctions funObj = new pptfunctions();

            if (funObj.TSCThemeLoaded())
            {
                frmChartformat chartobj = new frmChartformat();
                switch (galleryID)
                {
                    case "chart1": { funObj.formatChart("item1", rib); break; }
                    case "chart2": { funObj.formatChart("item2", rib); break; }
                    case "chart3": { funObj.formatChart("item3", rib); break; }
                    case "chart4": { funObj.formatChart("item4", rib); break; }
                }
            }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }

        public void btnInsertTable_Click(Office.IRibbonControl rib, string galleryID, int selectedIndex)
        {
            pptfunctions funObj = new pptfunctions();
            if (funObj.TSCThemeLoaded())
            {
                frmtable tableObj = new frmtable();
                
                switch (galleryID)
                {
                    case "Atable1": { PPTAttribute.tableType.Add("Atable1"); break; }
                    case "Atable2": { PPTAttribute.tableType.Add("Atable2"); break; }
                    case "Atable3": { PPTAttribute.tableType.Add("Atable3"); break; }
                    case "Atable4": { PPTAttribute.tableType.Add("Atable4"); break; }
                    case "Atable5": { PPTAttribute.tableType.Add("Atable5"); break; }
                    case "Atable6": { PPTAttribute.tableType.Add("Atable6"); break; }
                    case "Atable7": { PPTAttribute.tableType.Add("Atable7"); break; }
                    case "Atable8": { PPTAttribute.tableType.Add("Atable8"); break; }
                }
                tableObj.ShowDialog();
            }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }

        public void btnformatTable_Click(Office.IRibbonControl rib, string galleryID, int selectedIndex)
        {
            pptfunctions funObj = new pptfunctions();
            string tbltype=null;
            if (funObj.TSCThemeLoaded())
            {
                clstblformat tableObj = new clstblformat();

                switch (galleryID)
                {
                    case "ftable1": { tbltype= "Atable1"; break; }
                    case "ftable2": { tbltype = "Atable2"; break; }
                    case "ftable3": { tbltype = "Atable3"; break; }
                    case "ftable4": { tbltype = "Atable4"; break; }
                    case "ftable5": { tbltype = "Atable5"; break; }
                    case "ftable6": { tbltype = "Atable6"; break; }
                    case "ftable7": { tbltype = "Atable7"; break; }
                    case "ftable8": { tbltype = "Atable8"; break; }
                }
                tableObj.tableformat(tbltype);
            }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        public void GalleryOnAction(Office.IRibbonControl rib, string galleryID, int selectedIndex)
        {
            PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
            string bgrCode = null;
            char Csplit = ',';
            switch (galleryID)
            {
                case "itm11": { bgrCode = "102,114,0"; break; }
                case "itm12": { bgrCode =  "133,142,51"; break; }
                case "itm13": { bgrCode =  "163,170,102"; break; }
                case "itm14": { bgrCode = "194,199,153"; break; }
                case "itm15": { bgrCode =  "224,227,204"; break; }

                case "itm21": { bgrCode =  "155,174,0"; break; }
                case "itm22": { bgrCode = "175,190,91"; break; }
                case "itm23": { bgrCode = "195,206,202,"; break; }
                case "itm24": { bgrCode = "215,223,153"; break; }
                case "itm25": { bgrCode = "235,239,204"; break; }

                case "itm31": { bgrCode =  "97,85,75"; break; }
                case "itm32": { bgrCode = "128,119,110"; break; }
                case "itm33": { bgrCode =  "160,153,146"; break; }
                case "itm34": { bgrCode =  "191,186,183"; break; }
                case "itm35": { bgrCode =  "223,221,219"; break; }

                case "itm41": { bgrCode =  "57,42,30"; break; }
                case "itm42": { bgrCode =  "90,76,67"; break; }
                case "itm43": { bgrCode =  "126,115,108"; break; }
                case "itm44": { bgrCode =  "168,159,152"; break; }
                case "itm45": { bgrCode = "210,205,102"; break; }

                case "itm51": { bgrCode =  "193,184,162"; break; }
                case "itm52": { bgrCode =  "205,198,181"; break; }
                case "itm53": { bgrCode =  "218,212,199"; break; }
                case "itm54": { bgrCode =  "230,227,218"; break; }
                case "itm55": { bgrCode = "243,241,236"; break; }

                case "itm61": { bgrCode = "78,204,124"; break; }
                case "itm62": { bgrCode = "113,214,150"; break; }
                case "itm63": { bgrCode =  "149,224,177"; break; }
                case "itm64": { bgrCode =  "184,235,203"; break; }
                case "itm65": { bgrCode =  "220,245,229"; break; }

                case "itm71": { bgrCode = "9,84,187"; break; }
                case "itm72": { bgrCode = "51,187,37"; break; }
                case "itm73": { bgrCode = "251,194,13"; break; }
                case "itm74": { bgrCode = "127,127,127"; break; }
                case "itm75": { bgrCode = "217,217,217"; break; }
            }
            List<string> Ccode = new List<string>();
            Ccode = bgrCode.Split(Csplit).ToList();
            if (ppApp.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select a single shape  to format.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            foreach (PowerPoint.Shape shp in ppApp.ActiveWindow.Selection.ShapeRange)
            {
                try
                {
                    shp.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(Convert.ToInt32(Ccode[0]), Convert.ToInt32(Ccode[1]), Convert.ToInt32(Ccode[2])).ToArgb();
                }
                catch { }
            }
            PPTAttribute.UserTracker(rib);
        }
        public void TSCP_Callback82(Office.IRibbonControl rib)
        {
            tscformat choObj = new tscformat();
            choObj.tsc_loadtheme();
            PPTAttribute.UserTracker(rib);
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
        
        // -I am working on this --
        public void TSCP_Callback62(Office.IRibbonControl rib)
        {
            frmEditorialReview eReview = new frmEditorialReview();
            eReview.Show();
        }
        public void btnfeedback_Click(Office.IRibbonControl rib)
        {
            Outlook.Application mailObj = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            Outlook.MailItem newmail= mailObj.CreateItem(Outlook.OlItemType.olMailItem);
            newmail.Subject = "Feedback - TSC PPT Addin";
            newmail.To = "automation@thesmartcube.com;PresentationDesign@thesmartcube.com";
            newmail.Display();
            PPTAttribute.UserTracker(rib);
        }
        //Developed by vandna
        public void btnselectshape_Click(Office.IRibbonControl rib)
        {
            clsSelectObject selObj = new clsSelectObject();
            pptfunctions funObj = new pptfunctions();
            if (funObj.TSCThemeLoaded()) { selObj.Select_Same_Object(rib); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);

        }
        //Developed by vandna
        public void btnAlignshape_Click(Office.IRibbonControl rib)
        {
            clsAlignObject objAlign = new TSCPPT_Addin.clsAlignObject();
            pptfunctions funObj = new pptfunctions();
           
            if (funObj.TSCThemeLoaded()) { objAlign.funAlignObject(rib); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);

            

        }
        // Work in Progress --
        public void btnGethierarchy_Click(Office.IRibbonControl rib)
        {
            clshierarchy hierobj = new clshierarchy();
            hierobj.setCompneyHierarchy();
        }
        public void btnInserttable_Click(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            frmtable tbobj = new frmtable();
            if (funObj.TSCThemeLoaded()) { tbobj.ShowDialog(); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        public void btnRemove_Click(Office.IRibbonControl rib)  //Maksood
        {
            pptfunctions funObj = new pptfunctions();

            clsClean cClean = new clsClean();
            if (funObj.TSCThemeLoaded())
            {
                if (rib.Id == "btnRemoveDoubleSpace")
                {
                    cClean.pRemoveDoubleSpace_Clean();
                }
                else if (rib.Id == "btnRemoveSlideAnimations")
                {
                    cClean.pRemoveAnimation();
                }
                else if (rib.Id == "btnRemove3D_effects")
                {
                    cClean.pRemove3D_Effect();
                }

                else if (rib.Id == "btnRemoveMargin")
                {

                    cClean.pRemoveMargin();
                }
                else if(rib.Id== "btnParagraphSpacing")
                {
                    cClean.setParagraphSpacing();
                }
            }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            PPTAttribute.UserTracker(rib);
        }
        public void insertInfo_Click(Office.IRibbonControl rib)
        {
            pptfunctions funObj = new pptfunctions();
            Infographics frminfo = new Infographics();
            if (funObj.TSCThemeLoaded()) { frminfo.ShowDialog(); }
            else { MessageBox.Show("This functionality works with TSC Theme. Please Load TSC theme and try again. Thanks", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
        public void btnImage_Click(Office.IRibbonControl rib)
        {
            string url = "http://kms.thesmartcube.com/SitePages/ImageSearchResult.aspx?keyword=jpg";
            Process.Start(url);
        }
        public void btnfeedback_Click1(Office.IRibbonControl rib)
        {
            PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
            Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
            MyUserControl ctrlObj = new MyUserControl();
            //myCustomTaskPane = CustomTaskPanes.Add(ctrlObj, "New Task Pane");
        }
        public void btnAdddraft_Click(Office.IRibbonControl rib)
        {
            pptfunctions obj = new pptfunctions();
            obj.InsertDraft();

        }
        public void btnRemovedraft_Click(Office.IRibbonControl rib)
        {
            pptfunctions obj = new pptfunctions();
            obj.RemoveDraft();
        }
        public void btnSave_Click(Office.IRibbonControl rib)  //Maksood
        {

            clsSave csave = new clsSave();
            if (rib.Id == "btnSaveSelectedSlide")
            {
                csave.pSaveSelectedSlide();
            }
            else if (rib.Id == "btnSaveAllSlide")
            {
                csave.pSeprateEachSlide();
            }
            PPTAttribute.UserTracker(rib);


        }

        #endregion
    }
}
