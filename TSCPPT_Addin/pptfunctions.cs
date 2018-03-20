using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using System.IO;
using Microsoft.Office.Core;
using System.Drawing;
using System.Runtime.InteropServices;

namespace TSCPPT_Addin
{
    class pptfunctions
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation thisPPT = Globals.ThisAddIn.Application.ActivePresentation;
        Shapecheck PPTshpchk = new Shapecheck();
        Formatshapes PPTshpFormat = new Formatshapes();
        string msgTitle = "The Smart Cube";

        public bool TSCThemeLoaded()
        {
            bool tFlag = true;
            try
            {
                int num_slides = ppApp.ActivePresentation.Slides.Count;
                for (int i = 1; i <= num_slides; i++)
                {
                    if (ppApp.ActivePresentation.Slides[i].Design.Name != "The Smart Cube Theme")
                    {
                        tFlag = false;
                        break;
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "TSCThemeLoaded");
            }
            return tFlag;
        }
        public void ApplyPPT_Theme(Office.IRibbonControl rib)
        {
            //String path;
            PowerPoint.Presentation actPPT=null;
            PowerPoint.Presentation thisPPT = Globals.ThisAddIn.Application.ActivePresentation;
            PPTActions PPTAction = new PPTActions();
            PowerPoint.Application oApp = Marshal.GetActiveObject("PowerPoint.Application") as PowerPoint.Application;
            try
            {
                ppApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                thisPPT.PageSetup.SlideWidth = 960;
                thisPPT.PageSetup.SlideHeight = 540;
                oApp.ActivePresentation.ApplyTheme(PPTAttribute.mPPTPath);
                DeleteOldTSCMasters();
                ppApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
            }
            catch (Exception err)
            {
                string errtext = err.Message; 
                PPTAttribute.ErrorLog(errtext, rib.Id);
                
             }

        }

        public void RenameOldTSCMasters()
        {
            string design_name;
            try
            {
                int num_design = thisPPT.Designs.Count;
                for (int i = 1; i <= num_design; i++)
                {
                    design_name = thisPPT.Designs[i].Name;
                    thisPPT.Designs[design_name].Name = design_name + "_xOldX";
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "RenameOldTSCMasters");
            }
        }
        public void DeleteOldTSCMasters()
        {
            string tsc_designName = "TSC Template";
            string design_name=null;
            try
            {
                int num_design = thisPPT.Designs.Count;

                for (int i = 1; i <= num_design; i++)
                {
                    design_name = thisPPT.Designs[i].Name;
                    //IndexOf("_xoldX", 0) > 0)
                    if (design_name != tsc_designName && design_name.Contains("_xOldX") == true)
                    {
                        thisPPT.Designs[i].Delete();

                    }
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "DeleteOldTSCMasters");
            }
        }

        

        //---------------------------------------------------------
        // Add New Slide in New Template -
        public void addNewPPT_In_tsc_format(Office.IRibbonControl rib)
        {
            
            //String path = @"C:\Users\Devendra.Tripathi\Documents\visual studio 2015\Projects\TSCPPT_Addin\TSCPPT_Addin\AppData\Template\Template_Automation.pptx";
            PowerPoint.Presentation newPPT = ppApp.Presentations.Add();
            PowerPoint.Presentation actPPT = ppApp.Presentations.Open(PPTAttribute.standardppt);
            try
            {
                //PowerPoint.CustomLayout customLayout = null;
                newPPT.ApplyTheme(PPTAttribute.standardppt);
                newPPT.PageSetup.SlideWidth = 960;
                newPPT.PageSetup.SlideHeight = 540;
                actPPT.Slides[1].Copy();
                newPPT.Slides.Paste(1);
                newPPT.Slides[1].Name = "Title Slide";

                actPPT.Slides[2].Copy();
                newPPT.Slides.Paste(2);

                actPPT.Slides[3].Copy();
                newPPT.Slides.Paste(3);

                actPPT.Slides[4].Copy();
                newPPT.Slides.Paste(4);

                actPPT.Slides[5].Copy();
                newPPT.Slides.Paste(5);

                actPPT.Slides[6].Copy();
                newPPT.Slides.Paste(6);
                
                actPPT.Close();
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, rib.Id);
                actPPT.Close();
            }


        }

        // Add New Slide in New Template -
        public void Insert_Selected_Slide(Office.IRibbonControl rib)
        {
            string controlID;int SlideIndex=0;
            int selectedSlides;
            controlID = rib.Id;
            String sName, set_msg=null;
            string coverName;
            PPTActions pptObj = new PPTActions();
            
            int cnt = 0; DialogResult iReply=DialogResult.No;
            try { selectedSlides = ppApp.ActiveWindow.Selection.SlideRange.Count; }
            catch(Exception ex) { selectedSlides = 1; }
            int cpySlide;
            try
            {
                SlideIndex = pptObj.get_LastSelectedSlide();

                if (rib.Id != "customButton13" || rib.Id != "customButton18")
                {
                    if (selectedSlides == 0) { MessageBox.Show("Please select a slide and try again.", PPTAttribute.msgTitle,MessageBoxButtons.OK,MessageBoxIcon.Warning); }//|| thisPPT.Slides.Count > 0)
                    else if (thisPPT.Slides.Count == 0) { SlideIndex = 0; }
                }

                if (rib.Id == "customButton13")
                {
                    foreach (PowerPoint.Slide tslide in thisPPT.Slides)
                    {
                        sName = tslide.Name;
                        if (sName.Length >= 11) { coverName = sName.Substring(0, 11); }
                        else { coverName = sName; }
                        //sName.Substring(0,11) == "Title Slide"
                        if (tslide.CustomLayout.Name == "Title Slide" || coverName == "Title Slide") { cnt++; }
                    }
                    if (cnt == 1) { set_msg = "This presentation already contains a title slide. Would you like to add another?"; }
                    if (cnt > 1) { set_msg = "This presentation already contains multiple title slides. Would you like to add more?"; }
                    if (cnt > 0)
                    {
                        iReply = MessageBox.Show(set_msg, PPTAttribute.msgTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    }
                    if (cnt == 0 || iReply == DialogResult.Yes)
                    {
                        DataTable dtSlideIndex = pptObj.get_SlideIndex("Title Slide");
                        cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                        pptObj.insert_Slide_in_ActivePPT(cpySlide, 1, "CSlide", cnt);
                    }
                }
                else if (rib.Id == "customButton14")
                {
                    DataTable dtSlideIndex = pptObj.get_SlideIndex("Content Slide");
                    cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                    pptObj.insert_Slide_in_ActivePPT(cpySlide, (SlideIndex + 1));
                }
                else if (rib.Id == "customButton15")
                {
                    DataTable dtSlideIndex = pptObj.get_SlideIndex("Section Heading");
                    cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                    pptObj.insert_Slide_in_ActivePPT(cpySlide, (SlideIndex + 1));
                }
                else if (rib.Id == "customButton16")
                {
                    DataTable dtSlideIndex = pptObj.get_SlideIndex("FrameWork Slide");
                    cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                    //int cpySlide1 = Convert.ToInt32(dtSlideIndex.Rows[1]["SlideIndex"]);
                    pptObj.insert_Slide_in_ActivePPT(cpySlide, (SlideIndex + 1));
                    
                }
                else if (rib.Id == "customButton17")
                {
                    DataTable dtSlideIndex = pptObj.get_SlideIndex("Blank Slide");
                    cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                    pptObj.insert_Slide_in_ActivePPT(cpySlide, (SlideIndex + 1));
                }
                else if (rib.Id == "customButton18")
                {
                    foreach (PowerPoint.Slide tslide in thisPPT.Slides)
                    {
                        sName = tslide.Name;
                        try { if (tslide.CustomLayout.Name == "End Page" || sName.Substring(0, 9) == "End Page") { cnt++; } }
                        catch (Exception Ex) { }
                    }
                    if (cnt == 1) { set_msg = "This presentation already contains a End slide. Would you like to add another?"; }
                    if (cnt > 1) { set_msg = "This presentation already contains multiple End slides. Would you like to add more?"; }
                    if (cnt > 0)
                    {
                        iReply = MessageBox.Show(set_msg, PPTAttribute.msgTitle, MessageBoxButtons.YesNo,MessageBoxIcon.Warning);
                    }
                    if (cnt == 0 || iReply == DialogResult.Yes)
                    {
                        DataTable dtSlideIndex = pptObj.get_SlideIndex("End Page");
                        cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                        pptObj.insert_Slide_in_ActivePPT(cpySlide, (thisPPT.Slides.Count + 1), "ESlide", cnt);
                    }
                }
            }
            catch(Exception ex)
            {
                string errtext = ex.Message;
                PPTAttribute.ErrorLog(errtext, rib.Id);
                
            }

        }

        public void insert_PPT_Object(Office.IRibbonControl rib)
        {
            List<int> selectSlide = new List<int>();
            bool EshpNum;
            DialogResult iReply = DialogResult.No;
            string set_msg = null,objType=null;
            try
            {
                PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
                PPTActions actionObj = new PPTActions();

                //string value = Convert.ToString(dt.Rows[2][2]);
                //---------------------------------------
                int selectedSlides = ppApp.ActiveWindow.Selection.SlideRange.Count;
                for (int sln = 1; sln <= selectedSlides; sln++)
                {
                    //ActiveWindow.Selection.SlideRange(i).SlideNumber
                    selectSlide.Add(ppApp.ActiveWindow.Selection.SlideRange[sln].SlideNumber);
                }
                for (int sIndex = 0; sIndex < selectedSlides; sIndex++)
                {
                    int sldNum = selectSlide[sIndex];
                    if (rib.Id == "customButton21")
                    {
                        objType = "Road Map";
                        set_msg = "A Road map already exist on Slide " + sldNum + ". Would you like to add additional Road map?";
                    }

                    else if (rib.Id == "customButton22")
                    {
                        objType = "Title Box";
                        set_msg = "A Title box already exist on Slide " + sldNum + ". Would you like to add additional title?";
                    }
                    else if (rib.Id == "customButton23")
                    {
                        objType = "Text Box";
                        set_msg = "A Text Box already exist on Slide " + sldNum + ". Would you like to add additional title?";
                    }
                    else if (rib.Id == "customButton24")
                    {
                        objType = "Note Box";
                        set_msg = "A Note Box already exist on Slide " + sldNum + ". Would you like to add additional title?";
                    }
                    else if (rib.Id == "customButton25")
                    {
                        objType = "Source Box";
                        set_msg = "A Source Box already exist on Slide " + sldNum + ". Would you like to add additional title?";
                    }
                    else if (rib.Id == "customButton26")
                    {
                        objType = "Chart Title";
                        set_msg = "A Chart Title already exist on Slide " + sldNum + ". Would you like to add additional title?";
                    }
                    else if (rib.Id == "customButton27")
                    {
                        objType = "Quote Box";
                        set_msg = "A Quote Box already exist on Slide " + sldNum + ". Would you like to add additional title?";
                    }
                    DataTable dt = actionObj.get_specification(objType);
                    EshpNum = PPTshpchk.CheckIfBoxAlreadyExist(sldNum, dt);
                    if (EshpNum == true)
                    {
                        iReply = MessageBox.Show(set_msg, PPTAttribute.msgTitle, MessageBoxButtons.YesNo,MessageBoxIcon.Warning);
                        if (iReply == DialogResult.Yes)
                        {
                            string shpname = actionObj.InsertPlaceholder(sldNum, dt, objType);
                            PPTshpFormat.FormatShape(sldNum, shpname, dt);
                            if (objType == "Text Box")
                            {
                                PPTshpchk.CreateBullet(sldNum, shpname);
                                PPTshpchk.FormatBulletInShape(sldNum, shpname);
                            }
                            

                        }
                    }

                    else if (EshpNum == false)
                    {
                        string shpname = actionObj.InsertPlaceholder(sldNum, dt, objType);
                        PPTshpFormat.FormatShape(sldNum, shpname, dt);
                        if (objType == "Text Box")
                        {
                            PPTshpchk.CreateBullet(sldNum, shpname);
                            PPTshpchk.FormatBulletInShape(sldNum, shpname);
                        }

                    }

                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, rib.Id);
            }
        }

        public void Format_PPT_Object(Office.IRibbonControl rib)
        {
            PowerPoint.Shape shpType;
            string objType=null, msg_text;
            int sldNum;
            int shpCount=0;
            Algorithms algoObj = new Algorithms();
            PPTActions actionObj = new PPTActions();
            try
            {
                //if (ppApp.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                //{
                //    MessageBox.Show("Please select a single shape  to format.",PPTAttribute.msgTitle,MessageBoxButtons.OK,MessageBoxIcon.Error);
                //    return;
                //}
                try
                {
                    foreach (PowerPoint.Shape shp in ppApp.ActiveWindow.Selection.ShapeRange)
                    {
                        shpCount++;
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Please select a single shape  to format", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (shpCount > 1) { MessageBox.Show("Multipal shapes selection not allow.", PPTAttribute.msgTitle,MessageBoxButtons.OK, MessageBoxIcon.Error); }
                else
                {
                    //shpType = ActiveWindow.Selection.ShapeRange(1).Type
                    shpType = ppApp.ActiveWindow.Selection.ShapeRange[1];
                    if (rib.Id == "customButton31")
                    {
                        msg_text = "A Road map already exist on this slide. Would you like to add additional Road map?";
                        objType = "Road Map";
                    }
                    else if (rib.Id == "customButton32")
                    {
                        msg_text = "A Title Box already exist on this slide. Would you like to add additional Road map?";
                        objType = "Title Box";
                    }
                    else if (rib.Id == "customButton33")
                    {
                        msg_text = "A Text Box already exist on this slide. Would you like to add additional Road map?";
                        objType = "Text Box";
                    }
                    else if (rib.Id == "customButton34")
                    {
                        msg_text = "A Note Box already exist on this slide. Would you like to add additional Road map?";
                        objType = "Note Box";
                    }
                    else if (rib.Id == "customButton35")
                    {
                        msg_text = "A Source Box already exist on this slide. Would you like to add additional Road map?";
                        objType = "Source Box";
                    }
                    else if (rib.Id == "customButton36")
                    {
                        msg_text = "A Chart Title already exist on this slide. Would you like to add additional Road map?";
                        objType = "Chart Title";
                    }
                    else if (rib.Id == "customButton37")
                    {
                        msg_text = "A Quote Box already exist on this slide. Would you like to add additional Road map?";
                        objType = "Quote Box";
                    }

                                  
                    DataTable dt = actionObj.get_specification(objType);
                    if (shpType.Type == Office.MsoShapeType.msoAutoShape || shpType.Type == Office.MsoShapeType.msoPlaceholder || shpType.Type == Office.MsoShapeType.msoTextBox)
                    {
                        sldNum = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
                        algoObj.SetNamesUnique(sldNum);
                        bool EshpNum = PPTshpchk.CheckIfBoxAlreadyExist(sldNum, dt);
                        PowerPoint.Shape selShape = ppApp.ActiveWindow.Selection.ShapeRange[1];
                        string shpname = PPTshpchk.SelectedShapeNumber(sldNum, selShape);
                        if (objType == "Quote Box" || objType == "Chart Title")
                        {
                            PPTshpFormat.FormatShape(sldNum, shpname, dt, false, false);
                            PPTshpchk.setBulletTypeNone(sldNum, shpname);
                        }
                        else if(objType == "Text Box")
                        {
                            PPTshpFormat.FormatShape(sldNum, shpname, dt, false, false);
                        }
                        else
                        {
                            PPTshpchk.setBulletTypeNone(sldNum, shpname);
                            PPTshpFormat.FormatShape(sldNum, shpname, dt, false, true);
                            
                        }
                        
                        if(objType == "Text Box")
                        {
                            PPTshpchk.FormatBulletInShape(sldNum, shpname);
                            PPTshpchk.setBulletImage(sldNum, shpname);
                            
                        }
                        
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, rib.Id);
            }
        }

        public void  InsertCharts(Office.IRibbonControl rib)
        {
            PPTActions actionObj = new PPTActions();
            DataTable dt;
            try
            {
                int sldNumber = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
                int slNumber;
                string chartName;
                PowerPoint.Presentation currentPPT = Globals.ThisAddIn.Application.ActivePresentation;
                //string dbpath = System.IO.Directory.GetCurrentDirectory();
                string dbpath = System.AppDomain.CurrentDomain.BaseDirectory;
                //string path = @"C:\Users\Devendra.Tripathi\Documents\visual studio 2015\Projects\TSCPPT_Addin\TSCPPT_Addin\AppData\Template\Template_Automation.pptx";
                PowerPoint.Presentation MasterPPT = Globals.ThisAddIn.Application.Presentations.Open(PPTAttribute.mPPTPath, Office.MsoTriState.msoFalse);
                if (rib.Id == "customButton41")
                {

                    dt = actionObj.get_ChartSpacification("Clustered Column");
                    slNumber = Convert.ToInt32(dt.Rows[0]["SlideNumber"]);
                    chartName = Convert.ToString(dt.Rows[0]["ChartName"]);
                    MasterPPT.Slides[slNumber].Shapes[chartName].Copy();
                }
                else if (rib.Id == "customButton42")
                {
                    dt = actionObj.get_ChartSpacification("Stacked Column");
                    slNumber = Convert.ToInt32(dt.Rows[0]["SlideNumber"]);
                    chartName = Convert.ToString(dt.Rows[0]["ChartName"]);
                    MasterPPT.Slides[slNumber].Shapes[chartName].Copy();

                }
                else if (rib.Id == "customButton43")
                {
                    dt = actionObj.get_ChartSpacification("Line");
                    slNumber = Convert.ToInt32(dt.Rows[0]["SlideNumber"]);
                    chartName = Convert.ToString(dt.Rows[0]["ChartName"]);
                    MasterPPT.Slides[slNumber].Shapes[chartName].Copy();
                }
                else if (rib.Id == "customButton44")
                {
                    dt = actionObj.get_ChartSpacification("Pie");
                    slNumber = Convert.ToInt32(dt.Rows[0]["SlideNumber"]);
                    chartName = Convert.ToString(dt.Rows[0]["ChartName"]);
                    MasterPPT.Slides[slNumber].Shapes[chartName].Copy();
                }
                else if (rib.Id == "customButton45")
                {
                    //dt = actionObj.get_ChartSpacification("Doughnut");
                    //slNumber = Convert.ToInt32(dt.Rows[0]["SlideNumber"]);
                    //chartName = Convert.ToString(dt.Rows[0]["ChartName"]);
                    //MasterPPT.Slides[slNumber].Shapes[chartName].Copy();
                }

                currentPPT.Slides[sldNumber].Shapes.Paste();
                float chtWb = currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Width;
                float chtHt = currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Height;
                float sldWd = currentPPT.PageSetup.SlideWidth;
                float sldHt = currentPPT.PageSetup.SlideHeight;
                currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Left = chtWb - 30;
                currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Top = chtHt - 50;
                MasterPPT.Close();
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, rib.Id);
            }
        }

        public void SumPieChart()
        {
            List<string> SelectedCharts = new List<string>();
            PowerPoint.Shape myShape;
            string msgText = null;
            PowerPoint.Chart myChart;
            object[] xVals;
            List<object> yVals = new List<object>();
            try
            {

                DialogResult iReply = DialogResult.No;
                int sld_num = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
                PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
                SelectedCharts = PPTshpchk.FindSelectedCharts();
                int numSelCht = SelectedCharts.Count;
                if (numSelCht == 0)
                {
                    MessageBox.Show("Please select a pie or doughnut chart for sum calculation.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //System.Environment.Exit(0);
                }
                else if (numSelCht > 1)
                {
                    MessageBox.Show("Please select a single chart Sum calculation.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //System.Environment.Exit(0);
                }
                else if (numSelCht == 1)
                {
                    string shp_nam = SelectedCharts[0];
                    myShape = ActivePPT.Slides[sld_num].Shapes[shp_nam];
                    myChart = ActivePPT.Slides[sld_num].Shapes[shp_nam].Chart;
                    string chType = PPTshpchk.chartType(myChart);
                    if (chType == "Pie" || chType == "Doughnut")
                    {
                        string sName = myChart.SeriesCollection(1).Name;
                        PowerPoint.Series series = (PowerPoint.Series)myChart.SeriesCollection(1);
                        foreach (object item in series.Values as Array) { yVals.Add(item); }
                        decimal pSum = 0;
                        for (int i = 0; i < yVals.Count; i++) { pSum = pSum + Convert.ToDecimal(yVals[i]); }
                        if (sName != "") { msgText = "Sum of the " + Convert.ToString(sName) + " is " + pSum.ToString("0.##") + ". Do you want to add a caption in the chart?"; }
                        else msgText = "Sum of the pie is " + pSum.ToString("0.##") + ". Do you want to add a caption in the chart?";
                        iReply = MessageBox.Show(msgText, PPTAttribute.msgTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (iReply == DialogResult.Yes)
                        {
                            string outText = "100% = " + Convert.ToString(pSum);
                            float outLeft = myShape.Left + myShape.Width - 102;
                            float outTop = myShape.Top + myShape.Height - (float)21.8;
                            PowerPoint.Shape aShp = ActivePPT.Slides[sld_num].Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, outLeft, outTop, 102, (float)21.8);
                            aShp.TextFrame.TextRange.Text = outText;
                            aShp.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
                            aShp.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                            aShp.TextFrame.TextRange.Font.Size = 12;
                            aShp.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                            aShp.TextFrame.TextRange.Font.Name = "Calibri";


                            aShp.Fill.Visible = MsoTriState.msoFalse;
                            aShp.Line.Weight = (float)0.75;
                            aShp.Line.DashStyle = MsoLineDashStyle.msoLineSolid;
                            aShp.Line.Style = MsoLineStyle.msoLineSingle;
                            aShp.Line.Transparency = 1;
                            aShp.Line.Visible = MsoTriState.msoTrue;
                            aShp.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(23, 94, 84).ToArgb();
                            aShp.Line.BackColor.RGB = System.Drawing.Color.FromArgb(23, 94, 84).ToArgb();
                        }
                        else
                        {
                            //MessageBox.Show("Please select a Pie Chart for Sum calculation.", PPTAttribute.msgTitle,MessageBoxButtons.OK,MessageBoxIcon.Error);
                            return;
                        }

                    }
                    else
                    {
                        MessageBox.Show("Please select a pie or doughnut chart for sum calculation.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "SumPieChart");
            }
        }// Close function 

        public void formatChart(string tscColors,string cDataLabels,string cYAxis,Office.IRibbonControl rib)
        {
            tscformat chartObj = new tscformat();
            Shapecheck shpObj = new Shapecheck();
            PPTActions dbObj = new PPTActions();
            List<string> SelectedCharts = new List<string>();
            bool hsDataLabels=true, hasYAxis=false;
            DataTable DTcolorCode=null;
            char splitChar = ',';
            string shp_nam;
            try
            {
                if (tscColors == "item1") { tscColors = "Scheme1"; }
                if (tscColors == "item2") { tscColors = "Scheme2"; }
                if (tscColors == "item3") { tscColors = "Scheme3"; }
                if (tscColors == "item4") { tscColors = "Scheme4"; }
                if (cDataLabels == "item6") { hsDataLabels = true; }
                if (cDataLabels == "item7") { hsDataLabels = false; }
                if (cYAxis == "item8") { hasYAxis = true; }
                if (cYAxis == "item9") { hasYAxis = false; }

                int sld_num = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
                int num_shp = ppApp.ActiveWindow.Selection.SlideRange.Count;
                SelectedCharts = chartObj.FindSelectedCharts();
                int numSelCht = SelectedCharts.Count;
                if (numSelCht == 0) { MessageBox.Show("Please select a chart to format.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                else if (numSelCht > 1)
                {
                    DialogResult res = MessageBox.Show("You have selected multiple charts to format. Would you like to continue with format?", PPTAttribute.msgTitle, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
                    if (res == DialogResult.No || res == DialogResult.Cancel) { return; }

                }
                for (int selCht = 0; selCht < numSelCht; selCht++)
                {
                    shp_nam = SelectedCharts[selCht];                       // Get the shape name from selected chart item
                    PowerPoint.Shape myShape = thisPPT.Slides[sld_num].Shapes[shp_nam];
                    PowerPoint.Chart myChart = thisPPT.Slides[sld_num].Shapes[shp_nam].Chart;
                    string chType = shpObj.chartType(myChart);
                    bool ch3D = shpObj.chart3D(myChart);
                    //- ---------------------   G E T   C O L O R   C O D E --------------------------------------------------------
                    PowerPoint.SeriesCollection tmpsc = myChart.SeriesCollection();
                    int seriesCount = tmpsc.Count;
                    if (tscColors == "Scheme1") { DTcolorCode = dbObj.get_ChatColorCode("Scheme1", seriesCount); }
                    else if (tscColors == "Scheme2") { DTcolorCode = dbObj.get_ChatColorCode("Scheme2", seriesCount); }
                    else if (tscColors == "Scheme3") { DTcolorCode = dbObj.get_ChatColorCode("Scheme3", seriesCount); }
                    else if (tscColors == "Scheme4") { DTcolorCode = dbObj.get_ChatColorCode("Scheme4", seriesCount); }
                    // ----------------------------------------------------- Call the function to Format a Chart -------------------
                    chartObj.Format_Series(myChart, hasYAxis, chType);
                    chartObj.Format_ChartArea(myChart);
                    chartObj.Format_Title(myChart);
                    chartObj.Format_YAxis1(myChart, hasYAxis, chType);
                    chartObj.Format_YAxis2(myChart, hasYAxis, chType);
                    chartObj.Format_YGrids(myChart, hasYAxis, chType, ch3D);
                    chartObj.Format_XAxis(myChart, true, chType);
                    chartObj.Format_XGrids(myChart, false, chType);
                    chartObj.Format_XSeries(myChart, true);
                    //chartObj.Format_Legend(myChart, true);
                    chartObj.Format_PlotArea(myChart, chType);
                    
                    // -----------------------------------------------------------------------------------------------------------------
                    PowerPoint.SeriesCollection sc = myChart.SeriesCollection();
                    for (int i = 1; i <= sc.Count; i++)
                    {
                        int srType = myChart.SeriesCollection(i).Type;
                        if (chType == "Pie" || chType == "Doughnut")
                        {
                            for (int j = 1; j <= myChart.SeriesCollection(i).Points.count; j++)
                            {

                                string rgbCode = Convert.ToString(DTcolorCode.Rows[j - 1][0]);
                                string[] bgr = rgbCode.Split(splitChar).ToArray();
                                myChart.SeriesCollection(i).Points(j).Format.Fill.ForeColor.RGB = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                                myChart.SeriesCollection(i).Points(j).Border.LineStyle = PowerPoint.XlLineStyle.xlLineStyleNone;
                            }
                        }
                        else if (chType == "Radar")
                        {
                            if (myChart.ChartType == XlChartType.xlRadarFilled)
                            {
                                string rgbCode = Convert.ToString(DTcolorCode.Rows[i - 1][0]);
                                string[] bgr = rgbCode.Split(splitChar).ToArray();
                                myChart.SeriesCollection(i).Format.Fill.ForeColor.RGB = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                                myChart.SeriesCollection(i).Border.LineStyle = PowerPoint.XlLineStyle.xlLineStyleNone;
                            }
                            else
                            {
                                string rgbCode = Convert.ToString(DTcolorCode.Rows[i - 1][0]);
                                string[] bgr = rgbCode.Split(splitChar).ToArray();
                                myChart.SeriesCollection(i).Border.Color = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                                myChart.SeriesCollection(i).Format.Line.ForeColor.RGB = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                                myChart.SeriesCollection(i).MarkerBackgroundColor = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                            }
                        }
                        else if (srType == 4 || srType == -4169)
                        {
                            string rgbCode = Convert.ToString(DTcolorCode.Rows[i - 1][0]);
                            string[] bgr = rgbCode.Split(splitChar).ToArray();
                            myChart.SeriesCollection(i).Border.Color = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                            myChart.SeriesCollection(i).Format.Line.ForeColor.RGB = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                            myChart.SeriesCollection(i).MarkerBackgroundColor = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                            myChart.SeriesCollection(i).MarkerForegroundColor = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                            myChart.SeriesCollection(i).MarkerSize = 6;
                            if (myChart.ChartType == XlChartType.xl3DLine)
                            {
                                myChart.SeriesCollection(i).Format.Fill.ForeColor.RGB = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                            }

                        }
                        else if (chType == "Surface")
                        {
                            PowerPoint.LegendEntries leEntries = myChart.Legend.LegendEntries();
                            for (int j = 1; j <= leEntries.Count; j++)
                            {
                                if (myChart.ChartType == XlChartType.xlSurface || myChart.ChartType == XlChartType.xlSurfaceTopView)
                                {
                                    string rgbCode = Convert.ToString(DTcolorCode.Rows[i - 1][0]);
                                    string[] bgr = rgbCode.Split(splitChar).ToArray();
                                    myChart.Legend.LegendEntries(j).LegendKey.Format.Fill.ForeColor.RGB = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                                    myChart.Legend.LegendEntries(j).LegendKey.Border.LineStyle = PowerPoint.XlLineStyle.xlLineStyleNone;
                                }
                                else
                                {
                                    string rgbCode = Convert.ToString(DTcolorCode.Rows[i - 1][0]);
                                    string[] bgr = rgbCode.Split(splitChar).ToArray();
                                    myChart.Legend.LegendEntries(j).LegendKey.Format.Line.ForeColor.RGB = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();

                                }
                            }
                        } // Close else if of "Surface"

                        else if (chType == "Stock")
                        {
                            if (myChart.ChartType == XlChartType.xlStockOHLC || myChart.ChartType == XlChartType.xlStockVOHLC)
                            {
                                int numAxis = 1;
                                if (myChart.ChartType == XlChartType.xlStockVOHLC) { numAxis = 2; }
                                myChart.ChartGroups(numAxis).UpBars.Format.Fill.ForeColor.RGB = Color.FromArgb(87, 94, 23).ToArgb();
                                myChart.ChartGroups(numAxis).DownBars.Format.Fill.ForeColor.RGB = Color.FromArgb(255, 255, 255).ToArgb();
                            }
                            if (myChart.ChartType == XlChartType.xlStockVHLC || myChart.ChartType == XlChartType.xlStockVOHLC)
                            {
                                myChart.SeriesCollection(i).Format.Fill.ForeColor.RGB = Color.FromArgb(176, 214, 196).ToArgb();
                            }
                        }
                        else
                        {
                            string rgbCode = Convert.ToString(DTcolorCode.Rows[i - 1][0]);
                            string[] bgr = rgbCode.Split(splitChar).ToArray();
                            myChart.SeriesCollection(i).Format.Fill.ForeColor.RGB = Color.FromArgb(Convert.ToInt32(bgr[2]), Convert.ToInt32(bgr[1]), Convert.ToInt32(bgr[0])).ToArgb();
                            myChart.SeriesCollection(i).Border.LineStyle = PowerPoint.XlLineStyle.xlLineStyleNone;
                        } // Close Else

                    } // Close For -For - i

                    // -------------------------------- format Data Labels
                    PowerPoint.SeriesCollection srCol = myChart.SeriesCollection();
                    for (int i = 1; i <= srCol.Count; i++)
                    {
                        if (hsDataLabels == true)
                        {
                            myChart.SeriesCollection(i).HasDataLabels = true;
                            if (chType == "Pie" || chType == "Doughnut")
                            {
                                myChart.SeriesCollection(i).DataLabels.ShowCategoryName = true;
                                myChart.SeriesCollection(i).DataLabels.ShowPercentage = true;
                                myChart.SeriesCollection(i).DataLabels.ShowValue = false;
                                myChart.SeriesCollection(i).DataLabels.ShowSeriesName = false;
                            }
                            else
                            {
                                //myChart.SeriesCollection(i).DataLabels.Position = XlDataLabelPosition.xlLabelPositionBestFit;
                                myChart.SeriesCollection(i).DataLabels.Font.Name = "Calibri";
                                myChart.SeriesCollection(i).DataLabels.Font.Size = 11;
                                myChart.SeriesCollection(i).DataLabels.Font.Bold = false;
                                myChart.SeriesCollection(i).DataLabels.Font.Color = Color.FromArgb(0, 0, 0);

                            }
                        }
                        else
                        {
                            myChart.SeriesCollection(i).HasDataLabels = false;
                        }
                    }
                } // Close main for
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, rib.Id);
            } 
        }// Close the method
    }
}