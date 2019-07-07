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
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
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
            PowerPoint.Presentation thisPPT = Globals.ThisAddIn.Application.ActivePresentation;
            PPTActions PPTAction = new PPTActions();
            PowerPoint.Application oApp = Marshal.GetActiveObject("PowerPoint.Application") as PowerPoint.Application;
            try
            {
                //--
                //int xx =ppApp.ActivePresentation.Designs[1].SlideMaster.CustomLayouts.Count;
                //

                ppApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                thisPPT.PageSetup.SlideWidth = 960;
                thisPPT.PageSetup.SlideHeight = 540;
                oApp.ActivePresentation.ApplyTheme(PPTAttribute.mPPTPath);
                DeleteOldTSCMasters();
                ppApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
                //thisPPT.Slides[1].CustomLayout = ppApp.ActivePresentation.Designs[1].SlideMaster.CustomLayouts["Title Slide"];
                thisPPT.Slides[1].CustomLayout = ppApp.ActivePresentation.Designs["The Smart Cube Theme"].SlideMaster.CustomLayouts[1];
                thisPPT.Slides[1].Shapes["Text Placeholder 4"].Name = "Client Name";
                thisPPT.Slides[1].Shapes["Title 3"].Name = "Project Title";
                thisPPT.Slides[1].Shapes["Text Placeholder 5"].Name = "Project Date";
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
                int num_design = ActivePPT.Designs.Count;
                for (int i = 1; i <= num_design; i++)
                {
                    design_name = ActivePPT.Designs[i].Name;
                    ActivePPT.Designs[design_name].Name = design_name + "_xOldX";
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
                int num_design = ActivePPT.Designs.Count;

                for (int i = 1; i <= num_design; i++)
                {
                    design_name = ActivePPT.Designs[i].Name;
                    //IndexOf("_xoldX", 0) > 0)
                    if (design_name != tsc_designName && design_name.Contains("_xOldX") == true)
                    {
                        ActivePPT.Designs[i].Delete();

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
                

                actPPT.Slides[2].Copy();
                newPPT.Slides.Paste(1);

                actPPT.Slides[3].Copy();
                newPPT.Slides.Paste(2);

                actPPT.Slides[4].Copy();
                newPPT.Slides.Paste(3);

                actPPT.Slides[5].Copy();
                newPPT.Slides.Paste(4);

                actPPT.Slides[6].Copy();
                newPPT.Slides.Paste(5);

                actPPT.Slides[7].Copy();
                newPPT.Slides.Paste(6);

                actPPT.Slides[8].Copy();
                newPPT.Slides.Paste(7);

                actPPT.Slides[9].Copy();
                newPPT.Slides.Paste(8);

                actPPT.Slides[10].Copy();
                newPPT.Slides.Paste(9);


                newPPT.Slides[1].Select();

                actPPT.Slides[1].Copy();
                newPPT.Slides.Paste(1);
                newPPT.Slides[1].Name = "Title Slide";

                actPPT.Close();
                // Select first slide afte insert the ppt
               
                ppApp.ActiveWindow.View.GotoSlide(1);

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
                if (SlideIndex == 0) { MessageBox.Show("Please select a slide before proceed.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);return; }
                if (rib.Id != "customButton13" || rib.Id != "customButton18")
                {
                    if (selectedSlides == 0) { MessageBox.Show("Please select a slide before proceed.", PPTAttribute.msgTitle,MessageBoxButtons.OK,MessageBoxIcon.Warning); }//|| thisPPT.Slides.Count > 0)
                    else if (ActivePPT.Slides.Count == 0) { SlideIndex = 0; }
                }

                if (rib.Id == "customButton13")
                {
                    foreach (PowerPoint.Slide tslide in ActivePPT.Slides)
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
                else if (rib.Id == "customButton17")
                {
                    DataTable dtSlideIndex = pptObj.get_SlideIndex("Blank Slide");
                    cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                    pptObj.insert_Slide_in_ActivePPT(cpySlide, (SlideIndex + 1));
                }
                //- New PPT Layout 
                else if (rib.Id == "ncustomButton1")
                {
                    DataTable dtSlideIndex = pptObj.get_SlideIndex("Four columns");
                    cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                    pptObj.insert_Slide_in_ActivePPT(cpySlide, (SlideIndex + 1));
                }
                else if (rib.Id == "ncustomButton2")
                {
                    DataTable dtSlideIndex = pptObj.get_SlideIndex("One Chart One Column");
                    cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                    pptObj.insert_Slide_in_ActivePPT(cpySlide, (SlideIndex + 1));
                }
                else if (rib.Id == "ncustomButton3")
                {
                    DataTable dtSlideIndex = pptObj.get_SlideIndex("One Chart One Column Horizontal");
                    cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                    pptObj.insert_Slide_in_ActivePPT(cpySlide, (SlideIndex + 1));
                }
                else if (rib.Id == "ncustomButton4")
                {
                    DataTable dtSlideIndex = pptObj.get_SlideIndex("Two Columns Chart");
                    cpySlide = Convert.ToInt32(dtSlideIndex.Rows[0]["SlideIndex"]);
                    pptObj.insert_Slide_in_ActivePPT(cpySlide, (SlideIndex + 1));
                }
                //------
                else if (rib.Id == "customButton18")
                {
                    foreach (PowerPoint.Slide tslide in ActivePPT.Slides)
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
                        pptObj.insert_Slide_in_ActivePPT(cpySlide, (ActivePPT.Slides.Count + 1), "ESlide", cnt);
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
                        set_msg = "A Title box already exist on Slide " + sldNum + ". Would you like to add additional Title box?";
                    }
                    else if (rib.Id == "customButton23")
                    {
                        objType = "Text Box";
                        set_msg = "A Text Box already exist on Slide " + sldNum + ". Would you like to add additional Text Box?";
                    }
                    else if (rib.Id == "customButton24")
                    {
                        objType = "Note Box";
                        set_msg = "A Note Box already exist on Slide " + sldNum + ". Would you like to add additional Note Box?";
                    }
                    else if (rib.Id == "customButton25")
                    {
                        objType = "Source Box";
                        set_msg = "A Source Box already exist on Slide " + sldNum + ". Would you like to add additional Source Box?";
                    }
                    else if (rib.Id == "customButton26")
                    {
                        objType = "Chart Title";
                        set_msg = "A Chart Title already exist on Slide " + sldNum + ". Would you like to add additional Chart Title?";
                    }
                    else if (rib.Id == "customButton27")
                    {
                        objType = "Quote Box";
                        set_msg = "A Quote Box already exist on Slide " + sldNum + ". Would you like to add additional Quote Box?";
                    }
                    else if (rib.Id == "customButton28")
                    {
                        objType = "Sub Heading";
                        set_msg = "Sub Heading already exist on Slide " + sldNum + ". Would you like to add additional Sub Heading?";
                    }
                    DataTable dt = actionObj.get_specification(objType);
                    EshpNum = PPTshpchk.CheckIfBoxAlreadyExist(sldNum, dt);
                    if (EshpNum == true && (objType != "Quote Box" && objType != "Text Box"))
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
                            else if (objType == "Sub Heading")
                            {
                                string lineobj = "Line";
                                DataTable dtline = actionObj.get_specification(lineobj);
                                string lineshpname = actionObj.InsertPlaceholder(sldNum, dt, lineobj);
                                PPTshpFormat.FormatShape(sldNum, lineshpname, dtline);
                            }

                        }
                    }

                    else 
                    {
                        string shpname = actionObj.InsertPlaceholder(sldNum, dt, objType);
                        PPTshpFormat.FormatShape(sldNum, shpname, dt);
                        if (objType == "Text Box")
                        {
                            PPTshpchk.CreateBullet(sldNum, shpname);
                            PPTshpchk.FormatBulletInShape(sldNum, shpname);
                        }
                        else if (objType == "Sub Heading")
                        {
                            string lineobj = "Line";
                            DataTable dtline = actionObj.get_specification(lineobj);
                            string lineshpname = actionObj.InsertPlaceholder(sldNum, dt, lineobj);
                            PPTshpFormat.FormatShape(sldNum, lineshpname, dtline);
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
                        msg_text = "A Title Box already exist on this slide. Would you like to add additional Title Box?";
                        objType = "Title Box";
                    }
                    else if (rib.Id == "customButton33")
                    {
                        msg_text = "A Text Box already exist on this slide. Would you like to add additional Text Box?";
                        objType = "Text Box";
                    }
                    else if (rib.Id == "customButton34")
                    {
                        msg_text = "A Note Box already exist on this slide. Would you like to add additional Note Box?";
                        objType = "Note Box";
                    }
                    else if (rib.Id == "customButton35")
                    {
                        msg_text = "A Source Box already exist on this slide. Would you like to add additional Source Box?";
                        objType = "Source Box";
                    }
                    else if (rib.Id == "customButton36")
                    {
                        msg_text = "A Chart Title already exist on this slide. Would you like to add additional Chart Title?";
                        objType = "Chart Title";
                    }
                    else if (rib.Id == "customButton37")
                    {
                        msg_text = "A Quote Box already exist on this slide. Would you like to add additional Quote Box?";
                        objType = "Quote Box";
                    }
                    else if (rib.Id == "customButton38")
                    {
                        msg_text = "A Cagr Box already exist on this slide. Would you like to add additional Cagr map?";
                        objType = "Cagr Box";
                    }

                    DataTable dt = actionObj.get_specification(objType);
                    if (shpType.Type == Office.MsoShapeType.msoAutoShape || shpType.Type == Office.MsoShapeType.msoPlaceholder || shpType.Type == Office.MsoShapeType.msoTextBox)
                    {
                        sldNum = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
                        algoObj.SetNamesUnique(sldNum);
                        bool EshpNum = PPTshpchk.CheckIfBoxAlreadyExist(sldNum, dt);
                        PowerPoint.Shape selShape = ppApp.ActiveWindow.Selection.ShapeRange[1];
                        string shpname = PPTshpchk.SelectedShapeNumber(sldNum, selShape);
                        if (objType == "Quote Box" || objType == "Chart Title" || objType == "Cagr Box")
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
            DataTable dt=null;
            try
            {
                int sldNumber = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
                int slNumber;
                string chartName;
                PowerPoint.Presentation currentPPT = Globals.ThisAddIn.Application.ActivePresentation;
                //string dbpath = System.IO.Directory.GetCurrentDirectory();
                string dbpath = System.AppDomain.CurrentDomain.BaseDirectory;
                //string path = @"C:\Users\Devendra.Tripathi\Documents\visual studio 2015\Projects\TSCPPT_Addin\TSCPPT_Addin\AppData\Template\Template_Automation.pptx";
                PowerPoint.Presentation MasterPPT = Globals.ThisAddIn.Application.Presentations.Open(PPTAttribute.ChartPPTPath, Office.MsoTriState.msoFalse);
                if (rib.Id == "customButton41")
                {
                    dt = actionObj.get_ChartSpacification("Clustered Column");
                }
                else if (rib.Id == "customButton42")
                {
                    dt = actionObj.get_ChartSpacification("Stacked Column");
                }
                else if (rib.Id == "customButton43")
                {
                    dt = actionObj.get_ChartSpacification("100% Stacked Column Chart");
                }
                else if (rib.Id == "customButton44")
                {
                    dt = actionObj.get_ChartSpacification("Clustered Bar Chart");
                }
                else if (rib.Id == "customButton45")
                {
                    dt = actionObj.get_ChartSpacification("Stacked Bar Chart ");
                }
                else if (rib.Id == "customButton46")
                {
                    dt = actionObj.get_ChartSpacification("100% Stacked Bar Chart");
                }
                else if (rib.Id == "customButton47")
                {
                    dt = actionObj.get_ChartSpacification("Line Chart");
                }
                else if (rib.Id == "customButton48")
                {
                    dt = actionObj.get_ChartSpacification("Combination Chart");
                }
                else if (rib.Id == "customButton49")
                {
                    dt = actionObj.get_ChartSpacification("Pie Chart");
                }
                else if (rib.Id == "customButton410")
                {
                    dt = actionObj.get_ChartSpacification("Doughnut");
                }
                else if (rib.Id == "customButton411")
                {
                    dt = actionObj.get_ChartSpacification("Scatter Chart");
                }
                else if (rib.Id == "customButton412")
                {
                    dt = actionObj.get_ChartSpacification("Dimension Scatter Chart");
                }
                else if (rib.Id == "customButton413")
                {
                    dt = actionObj.get_ChartSpacification("Single Ascending Waterfall Chart");

                }
                else if (rib.Id == "customButton414")
                {
                    dt = actionObj.get_ChartSpacification("Single Descending Waterfall Chart");
                }
                else if (rib.Id == "customButton415")
                {
                    dt = actionObj.get_ChartSpacification("Multiple Ascending Waterfall Chart");
                }
                else if (rib.Id == "customButton416")
                {
                    dt = actionObj.get_ChartSpacification("Multiple Descending Waterfall Chart");
                }
                slNumber = Convert.ToInt32(dt.Rows[0]["SlideNumber"]);
                chartName = Convert.ToString(dt.Rows[0]["ChartName"]);
                MasterPPT.Slides[slNumber].Shapes[chartName].Copy();

                currentPPT.Slides[sldNumber].Shapes.Paste();
                //currentPPT.Slides[sldNumber].Shapes.PasteSpecial(Microsoft.Office.Interop.PowerPoint.PpPasteDataType.ppPasteDefault);
                float chtWb = currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Width;
                float chtHt = currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Height;
                float sldWd = currentPPT.PageSetup.SlideWidth;
                float sldHt = currentPPT.PageSetup.SlideHeight;
                //currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Left = chtWb - 30;
                //currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Top = chtHt - 50;
                //currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Height = 300;
                //currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Width = 500;
                currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Left = (float)193.29;
                currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Top = (float)115.83;
                currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Height = (float)321.93;
                currentPPT.Slides[sldNumber].Shapes[currentPPT.Slides[sldNumber].Shapes.Count].Width = (float)528.33;
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
                        if (sName != "") { msgText = "Sum of the  pie is " + pSum.ToString("0.##") + ". Do you want to add a caption in the chart?"; }
                        else msgText = "Sum of the pie is " + pSum.ToString("0.##") + ". Do you want to add a caption in the chart?";
                        iReply = MessageBox.Show(msgText, PPTAttribute.msgTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (iReply == DialogResult.Yes)
                        {
                            foreach (PowerPoint.Shape shp in ActivePPT.Slides[sld_num].Shapes)
                            {
                                if (shp.Name == myChart.Name + "sumofpie")
                                {
                                    shp.Delete();
                                    break;
                                }
                            }
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
                            aShp.Name = myChart.Name + "sumofpie";

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

        public void formatChart(string tscColors,Office.IRibbonControl rib)
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
                //if (cDataLabels == "item6") { hsDataLabels = true; }
                //if (cDataLabels == "item7") { hsDataLabels = true; }
                //if (cYAxis == "item8") { hasYAxis = false; }
                //if (cYAxis == "item9") { hasYAxis = false; }

                hsDataLabels = false;
                hasYAxis = false;
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
                    PowerPoint.Shape myShape = ActivePPT.Slides[sld_num].Shapes[shp_nam];
                    PowerPoint.Chart myChart = ActivePPT.Slides[sld_num].Shapes[shp_nam].Chart;
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
                    
                    chartObj.Format_ChartArea(myChart);
                    chartObj.Format_Title(myChart);
                    //chartObj.Format_YAxis1(myChart, hasYAxis, chType);
                    //chartObj.Format_YAxis2(myChart, hasYAxis, chType);
                    //chartObj.Format_YGrids(myChart, hasYAxis, chType, ch3D);
                    chartObj.Format_XAxis(myChart, true, chType);
                    chartObj.Format_XGrids(myChart, false, chType);
                    chartObj.Format_XSeries(myChart, true);
                    //chartObj.Format_Legend(myChart, true);
                    chartObj.Format_PlotArea(myChart, chType);
                    //chartObj.Format_Series(myChart, hasYAxis, chType);

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
                        //if (myChart.HasDataTable == true)
                        //{
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
                        //}
                        //else
                        //{
                        //    myChart.SeriesCollection(i).HasDataLabels = false;
                        //}
                    }
                    chartObj.Format_Series(myChart, hasYAxis, chType);  // format column chart line

                } // Close main for
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, rib.Id);
            } 
        }// Close the method

        public void formatbullettxt(Office.IRibbonControl rib)
        {

            //PowerPoint.TextRange mytxtRng = ppApp.ActiveWindow.Selection.ShapeRange[1].TextFrame.TextRange;
            //try
            //{
                PPTActions actionObj = new PPTActions();
                string shpname=null;
           
                int sldNum = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
                PowerPoint.Shape selShape = ppApp.ActiveWindow.Selection.ShapeRange[1];
                if (selShape.Type != MsoShapeType.msoGroup)
                {
                if (selShape.Type != MsoShapeType.msoTable)
                {
                    if (ppApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    {
                        PowerPoint.TextRange txtRng = ppApp.ActiveWindow.Selection.TextRange;
                        int SelectedParagraphs = txtRng.Characters(txtRng.Start, txtRng.Length).Paragraphs().Count;
                        ppApp.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.HangingPunctuation = MsoTriState.msoTrue;
                        if (SelectedParagraphs < 2)
                        {
                            shpname = ppApp.ActiveWindow.Selection.ShapeRange.Name;
                            int CurrentCursorPositionInCharacters = ppApp.ActiveWindow.Selection.TextRange.Start;
                            PowerPoint.TextFrame txtFrame = ppApp.ActiveWindow.Selection.ShapeRange.TextFrame;
                            int pnum = txtFrame.TextRange.Characters(-1, CurrentCursorPositionInCharacters).Paragraphs().Count;
                            PowerPoint.ParagraphFormat prgformat = txtRng.Paragraphs(pnum).ParagraphFormat;
                            prgformat.SpaceBefore = 6;
                            prgformat.SpaceAfter = 0;
                            prgformat.SpaceWithin = (float)0.9;

                            if (rib.Id == "btnbullet1")

                            {
                                //-Ruler.Levels[1].FirstMargin
                                char myCharacter = (char)132;
                                
                                txtRng.Paragraphs(pnum).IndentLevel = 1;
                                
                                txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.Character = myCharacter;
                                txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.Font.Color.RGB = System.Drawing.Color.FromArgb(78, 204, 124).ToArgb();
                                //txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.Font.Size = (float)10.84;
                                txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.RelativeSize = (float).9;
                                txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.Font.Name = "Wingdings 3";
                                txtRng.Paragraphs(pnum).Font.Color.RGB = System.Drawing.Color.FromArgb(57, 42, 30).ToArgb();
                                //txtRng.Paragraphs(pnum).Font.Size = 12;
                                txtRng.Paragraphs(pnum).Font.Name = "Corbel";
                                txtRng.Paragraphs(pnum).ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                                //ppApp.ActiveWindow.Selection.ShapeRange.TextFrame.Ruler.Levels[pnum].LeftMargin = (float)18;
                                //ppApp.ActiveWindow.Selection.ShapeRange.TextFrame.Ruler.Levels[pnum].FirstMargin = (float)18;
                                //ppApp.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Paragraphs[pnum].ParagraphFormat.FirstLineIndent = (float)18;
                                //ppApp.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Paragraphs[pnum].ParagraphFormat.RightIndent = (float)18;
                                ppApp.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Paragraphs[pnum].ParagraphFormat.FirstLineIndent = -18;
                                ppApp.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Paragraphs[pnum].ParagraphFormat.LeftIndent = (float)18;

                            }
                            else if (rib.Id == "btnbullet2")
                            {
                                txtRng.Paragraphs(pnum).IndentLevel = 2;
                                txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.Character = 167;
                                txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.Font.Color.RGB = System.Drawing.Color.FromArgb(78, 204, 124).ToArgb();
                                //txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.Font.Size = (float)11;
                                txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.RelativeSize = (float)1.0;
                                txtRng.Paragraphs(pnum).ParagraphFormat.Bullet.Font.Name = "Wingdings";
                                txtRng.Paragraphs(pnum).Font.Color.RGB = System.Drawing.Color.FromArgb(57, 42, 30).ToArgb();
                                //txtRng.Paragraphs(pnum).Font.Size = 11;
                                txtRng.Paragraphs(pnum).ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                                txtRng.Paragraphs(pnum).Font.Name = "Corbel";
                                ppApp.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Paragraphs[pnum].ParagraphFormat.FirstLineIndent = -18;
                                ppApp.ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Paragraphs[pnum].ParagraphFormat.LeftIndent = (float)36;
                            }

                        }
                    }

                }
                else
                {
                    PowerPoint.Table oTbl = ppApp.ActiveWindow.Selection.ShapeRange[1].Table;
                    for (int tbR = 1; tbR < oTbl.Rows.Count; tbR++)
                    {
                        for (int tbC = 1; tbC < oTbl.Columns.Count; tbC++)
                        {
                            if (oTbl.Cell(tbR, tbC).Selected)
                            {

                                int currentposition = ppApp.ActiveWindow.Selection.TextRange.Start;
                                //int currentposition = oTbl.Cell(tbR, tbC).Shape.TextFrame.TextRange.Start;
                                PowerPoint.TextFrame txtFrame = oTbl.Cell(tbR, tbC).Shape.TextFrame;
                                int tpnum = txtFrame.TextRange.Characters(-1, currentposition).Paragraphs().Count;
                                if (rib.Id == "btnbullet1")
                                {
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Bullet.Character = 132;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(78, 204, 124).ToArgb();
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Bullet.RelativeSize = (float).9;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Bullet.Font.Name = "Wingdings 3";
                                    //oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].Font.
                                    //oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].Font.Size = 12;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].Font.Name = "Corbel";
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.FirstLineIndent = -18;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.LeftIndent = (float)18;
                                }
                                else if (rib.Id == "btnbullet2")
                                {
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Bullet.Character = 167;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Bullet.Font.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(78, 204, 124).ToArgb();
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Bullet.RelativeSize = (float)1.0;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Bullet.Font.Name = "Wingdings";
                                    //oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].Font.
                                    //oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].Font.Size = 12;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].Font.Name = "Corbel";
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.FirstLineIndent = -18;
                                    oTbl.Cell(tbR, tbC).Shape.TextFrame2.TextRange.Paragraphs[tpnum].ParagraphFormat.LeftIndent = (float)36;
                                }
                            }
                        }

                    }
                    //MessageBox.Show("Bullet formating can not allow in table", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                }

                else
                {
                    MessageBox.Show("Please ungroup a selected items", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            //}
            //catch(Exception err)
            //{
            //    string errtext = err.Message;
            //    PPTAttribute.ErrorLog(errtext, rib.Id);
            //}
        }

        public void InsertDraft()
        {
            string waterMarkText;
            waterMarkText = "WORKING DRAFT";
            try
            {
                for (int i = 1; i <= ActivePPT.Slides.Count; i++)
                {
                    var test = ActivePPT.Slides[i].CustomLayout.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 194, 263, 519, 87);
                    test.TextFrame.TextRange.Text = waterMarkText;
                    test.Name = "Layout";
                    test.Rotation = 326;
                    test.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.FromArgb(184, 235, 203).ToArgb();
                    test.TextFrame.TextRange.Font.Size = 66;
                    test.TextFrame.TextRange.Font.Name = "Corbel";

                }
            }
            catch(Exception ex)
            {

            }

        }
        public void RemoveDraft()
        {

            try
            {
                for (int i = 1; i <= ActivePPT.Slides.Count; i++)
                {
                    try
                    {
                        ActivePPT.Slides[i].CustomLayout.Shapes["Layout"].Delete();
                    }
                    catch { }

                }
            }
            catch(Exception ex) { }

        }
    }
    
}