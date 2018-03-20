using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;

namespace TSCPPT_Addin
{
    public class Reviewformat
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        List<string> shpNames = new List<string>();
        List<string> errAU = new List<string>();
        Shapecheck chkObj = new Shapecheck();
        PPTActions actionObj = new PPTActions();
       
        public string errText;
        string shpErr = null;
        public void CheckFormat(int sldNum,string method)
        {
            shpNames.Clear();                               // Clear the list box -
            shpNames = NameAllShapes(sldNum);
            char spltChar = '\n';
            List<string> errA = new List<string>();
            string shpName;
            try
            {
                if (shpNames.Count != 0)
                {

                    for (int s = 0; s < shpNames.Count; s++)
                    {
                        shpErr = null;
                        PowerPoint.Shape shp = ActivePPT.Slides[sldNum].Shapes[shpNames[s]];
                        shpName = shp.Name;
                        shpErr = shpErr + CheckFormat_ShapeInSlide(sldNum, shpName);
                        if (ActivePPT.Slides[sldNum].CustomLayout.Name == "Title Slide")
                        {
                            shpErr = shpErr + CheckFormat_TitleSlide(sldNum, shpName) + '\n';
                        }
                        else if (ActivePPT.Slides[sldNum].CustomLayout.Name == "Contents Slide")
                        {
                            shpErr = shpErr + CheckFormat_ContentSlide(sldNum, shpName) + '\n';
                        }
                        else if (ActivePPT.Slides[sldNum].CustomLayout.Name == "Divider Slide")
                        {
                            shpErr = shpErr + CheckFormat_DividerSlide(sldNum, shpName) + '\n';
                        }
                        else
                        {
                            shpErr = shpErr + CheckFormat_MainSlide(sldNum, shpName) + '\n';
                        }

                        errA = shpErr.Split(spltChar).ToList();
                        errAU = UniqueValues(errA);
                        string fshpErr = null;
                        int counter = 1;
                        foreach (string err in errAU)
                        {
                            fshpErr = fshpErr + (counter + 1) + ")  " + err + '\n';
                        }
                        DeleteFormatComments_M1(sldNum, shpName);
                        if (method == "method1")
                        {

                            float lf = ActivePPT.Slides[sldNum].Shapes[shpName].Left;
                            float wb = ActivePPT.Slides[sldNum].Shapes[shpName].Width;
                            float tp = ActivePPT.Slides[sldNum].Shapes[shpName].Top;
                            if (shpErr != null && string.IsNullOrEmpty(shpErr.Replace("\n", "")) == false)
                            {
                                PowerPoint.Comment cmtNew = ActivePPT.Slides[sldNum].Comments.Add((lf + wb), tp, shpName + " Error", "TFR", shpErr);
                            }
                        }
                        else if (method == "method2")
                        {
                            if (shpErr != null && string.IsNullOrEmpty(shpErr.Replace("\n", "")) == false)
                            {

                                ActivePPT.Slides[sldNum].Shapes[shpName].Select();
                                List<string> lstError = new List<string>();
                                lstError = shpErr.Split('\n').ToList();
                                foreach (string errType in lstError)
                                {
                                    //MessageBox.Show("Static Value: " +Convert.ToString(PPTAttribute.exitFlag));
                                    if (errType != "" && PPTAttribute.exitFlag==false)
                                    {
                                        StreamWriter sw = new StreamWriter(PPTAttribute.supportfile);
                                        //string errtxt = sldNum + "|" + shpErr.Replace("\n", "") + "|" + shpName;
                                        string errtxt = sldNum + "|" + errType + "|" + shpName;
                                        sw.WriteLine(errtxt);
                                        frmErrorlist frmobj = new frmErrorlist();
                                        sw.Close();
                                        frmobj.ShowDialog();
                                    }
                                }

                            }
                        }
                    }// - Close Loop

                    //if (String.IsNullOrWhiteSpace(shpErr)) { MessageBox.Show("There is no error found in format review",PPTAttribute.msgTitle,MessageBoxButtons.OK, MessageBoxIcon.Information); }
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CheckFormat");
            }
        }

        public List<string> NameAllShapes(int sldNum)
        {
            try
            {
                int shpCount = ActivePPT.Slides[sldNum].Shapes.Count;
                for (int i = 1; i <= shpCount; i++)
                {
                    PowerPoint.Shape shp = ActivePPT.Slides[sldNum].Shapes[i];
                    if (shp.Type == MsoShapeType.msoGroup)
                    {
                        for (int x = 1; x <= shp.GroupItems.Count; x++)
                        {
                            if (shp.GroupItems[x].Type == MsoShapeType.msoGroup)
                            {
                                ShapesInGroups(shp.GroupItems[x]);
                            }
                            else { shpNames.Add(shp.GroupItems[x].Name); }
                        }
                    }
                    else { shpNames.Add(shp.Name); }
                } // - Close a loop 
                if (shpNames.Count != 0)
                {
                    //shpNames = sortShapes(sldNum, shpNames);              // Need to check what is requirment -
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "NameAllShapes");
            }
            return (shpNames);
        }
        public void ShapesInGroups(PowerPoint.Shape gShp)
        {
            try
            {
                for (int x = 1; x < gShp.GroupItems.Count; x++)
                {
                    if (gShp.GroupItems[x].Type == MsoShapeType.msoGroup) { ShapesInGroups(gShp.GroupItems[x]); }
                    else { shpNames.Add(gShp.Name); }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "ShapesInGroups");
            }
        }
        
        public string  CheckFormat_ShapeInSlide(int sldNum, string shpName)
        {
            string sErr = null;
            float sngCenterH, sngDistH=0, sngShapeLeft, sngShapeRight, sngCenterV, sngDistV, sngShapeTop, sngShapeBottom;
            // --Margin Error
            float sldLeft = 0, sldTop=0;
            try
            {
                PowerPoint.Shape tShp = ActivePPT.Slides[sldNum].Shapes[shpName];
                float sldRight = ActivePPT.PageSetup.SlideWidth;
                float sldBottom = ActivePPT.PageSetup.SlideHeight;
                // -- 'Check for Shape
                if (tShp.Type == MsoShapeType.msoLine)
                {
                    sngCenterH = tShp.Left + tShp.Width / 2;
                    sngDistH = fFindHorizontalDistance(tShp);                 //This is half of horizontal distance
                    sngShapeLeft = sngCenterH - sngDistH;                    //Actual Left
                    sngShapeRight = sngCenterH + sngDistH;                   //Actual Right

                    //Vertical Distances
                    sngCenterV = tShp.Top + tShp.Height / 2;
                    sngDistV = fFindVerticalDistance(tShp);                 //This is half of vertical distance
                    sngShapeTop = sngCenterV - sngDistV;                    //Actual Top
                    sngShapeBottom = sngCenterV + sngDistV;                 //Actual Bottom

                    //'Horizontal Comparision and Vertical Comparision
                    if (sngShapeLeft < sldLeft) { sErr = sErr + "Left position is outside slide margin" + '\n'; }
                    if (sngShapeRight > sldRight) { sErr = sErr + "Right position is outside slide margin" + '\n'; }
                    if (sngShapeTop < sldTop) { sErr = sErr + "Top position is outside slide margin" + '\n'; }
                    if (sngShapeBottom > sldBottom) { sErr = sErr + "Bottom position is outside slide margin" + '\n'; }

                }
                else
                {
                    if (tShp.Left < sldLeft) { sErr = sErr + "Left position is outside slide margin" + '\n'; }
                    if ((tShp.Left + tShp.Width) > sldRight) { sErr = sErr + "Right position is outside slide margin" + '\n'; }
                    if (tShp.Top < sldTop) { sErr = sErr + "Top position is outside slide margin" + '\n'; }
                    if ((tShp.Top + tShp.Height) > sldBottom) { sErr = sErr + "Bottom position is outside slide margin" + '\n'; }
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CheckFormat_ShapeInSlide");
            }
            return (sErr);
        }
        public string CheckFormat_TitleSlide(int sldNum, string shpName)
        {
            string shperr = null, sErr=null;
            //bool Tchkfind = false, Dchkfind=false,client=false;

            try
            {
                PowerPoint.Shape shp = ActivePPT.Slides[sldNum].Shapes[shpName];
                if (shp.Type == MsoShapeType.msoAutoShape || shp.Type == MsoShapeType.msoPlaceholder || shp.Type == MsoShapeType.msoTextBox)
                {
                    if (shp.HasTextFrame == MsoTriState.msoTrue)
                    {
                        DataTable dt1 = actionObj.get_specification("Project Title");
                        string Tchkfind = chkObj.CheckIfBoxAlreadyExist1(sldNum, dt1);

                        DataTable dt2 = actionObj.get_specification("Project Date");
                        string Dchkfind = chkObj.CheckIfBoxAlreadyExist1(sldNum, dt2);

                        DataTable dt3 = actionObj.get_specification("Client Name");
                        string client = chkObj.CheckIfBoxAlreadyExist1(sldNum, dt3);

                        if (Tchkfind == shpName && dt1.Rows.Count != 0) { sErr = sErr + CheckFormat_SBox(sldNum, shpName, dt1); }
                        else if (Dchkfind == shpName && dt2.Rows.Count != 0) { sErr = sErr + CheckFormat_SBox(sldNum, shpName, dt2); }
                        else if (client == shpName && dt3.Rows.Count != 0) { sErr = sErr + CheckFormat_SBox(sldNum, shpName, dt3); }
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CheckFormat_TitleSlide");
            }
            return (sErr);
        }
        public string CheckFormat_ContentSlide(int sldNum, string shpName)
        {
            string shperr = null;
            try
            {
                PowerPoint.Shape shp = ActivePPT.Slides[sldNum].Shapes[shpName];
                if (shp.Type == MsoShapeType.msoAutoShape || shp.Type == MsoShapeType.msoPlaceholder || shp.Type == MsoShapeType.msoTextBox)
                {
                    if (shp.HasTextFrame == MsoTriState.msoTrue)
                    {
                        DataTable dt = actionObj.get_specification("Content Title");
                        string CTitle = chkObj.CheckIfBoxAlreadyExist1(sldNum, dt);
                        DataTable dt1 = actionObj.get_specification("Content Section");
                        string CSection = chkObj.CheckIfBoxAlreadyExist1(sldNum, dt1);
                        DataTable dt2 = actionObj.get_specification("Content Number");
                        string CNumber = chkObj.CheckIfBoxAlreadyExist1(sldNum, dt2);
                        if (shpName == CTitle && dt.Rows.Count != 0) { shperr = shperr + CheckFormat_SBox(sldNum, shpName, dt); }
                        else if (shpName == CSection && dt1.Rows.Count != 0) { shperr = shperr + CheckFormat_SBox(sldNum, shpName, dt1); }
                        else if (shpName == CNumber && dt2.Rows.Count != 0) { shperr = shperr + CheckFormat_SBox(sldNum, shpName, dt2); }
                    }

                }
                else
                {
                    shpErr = shpErr + "Non Standard object in the slide." + '\n';
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CheckFormat_ContentSlide");
            }
            return (shperr);
        }
        public string CheckFormat_DividerSlide(int sldNum, string shpName)
        {
            string shperr = null;
            try
            {
                PowerPoint.Shape shp = ActivePPT.Slides[sldNum].Shapes[shpName];
                if (shp.Type == MsoShapeType.msoAutoShape || shp.Type == MsoShapeType.msoPlaceholder || shp.Type == MsoShapeType.msoTextBox)
                {
                    if (shp.HasTextFrame == MsoTriState.msoTrue)
                    {
                        DataTable dt1 = actionObj.get_specification("Section Title");
                        string STitle = chkObj.CheckIfBoxAlreadyExist1(sldNum, dt1);

                        DataTable dt2 = actionObj.get_specification("Section Content");
                        string SContent = chkObj.CheckIfBoxAlreadyExist1(sldNum, dt2);

                        if (shpName == STitle && dt1.Rows.Count != 0) { shperr = shperr + CheckFormat_SBox(sldNum, shpName, dt1); }
                        else if (shpName == SContent && dt2.Rows.Count != 0) { shperr = shperr + CheckFormat_SBox(sldNum, shpName, dt2); }

                    }
                }
                else
                {
                    shpErr = shpErr + "Non Standard object in the slide." + '\n';
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CheckFormat_DividerSlide");
            }
            return (shperr);
        }


        public string CheckFormat_MainSlide(int sldNum, string shpName)
        {
            string shperr = null;
            string dbShape;
            try
            {
                DataTable dt1 = new DataTable();                        // Get the specification of Title
                PowerPoint.Shape shp = ActivePPT.Slides[sldNum].Shapes[shpName];
                if (shp.Type == MsoShapeType.msoAutoShape || shp.Type == MsoShapeType.msoPlaceholder || shp.Type == MsoShapeType.msoTextBox)
                {
                    if (shp.HasTextFrame == MsoTriState.msoTrue)
                    {
                        // -----------------------------------------------------------
                        int hasspliter = shpName.IndexOf("_");
                        if (hasspliter != -1) { dbShape = shpName.Substring(0, hasspliter); }
                        else { dbShape = shpName; }
                        if (dbShape == "Text Box") { return (shperr); }                // -Avoid a Text box review 
                        dt1 = actionObj.get_specification(dbShape);
                        if (dt1.Rows.Count != 0) { shperr = shperr + CheckFormat_SBox(sldNum, shpName, dt1); };
                    }
                }
                else if (shp.Type == MsoShapeType.msoChart) { }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CheckFormat_MainSlide");
            }
            return (shperr);
        }
        // Chekc this function why we call it.
        public List<string> UniqueValues(List<String> vArray)
        {
            List<string> uniqueErr = new List<string>();
            return (uniqueErr);
        }
        public void DeleteFormatComments_M1(int sldNum, string shpName)
        {
            int numComments = ActivePPT.Slides[sldNum].Comments.Count;
            try
            {
                if (numComments > 0)
                {
                    for (int xx = numComments; xx>=1; xx--)
                    {
                        PowerPoint.Comment myComment = ActivePPT.Slides[sldNum].Comments[xx];
                        if (myComment.AuthorInitials == "TFR")
                        //if (myComment.Author == shpName + " Error" && myComment.AuthorInitials == "TFR")
                        {
                            myComment.Delete();
                           // break;
                        }
                    }
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "DeleteFormatComments_M1");
            }
        }
        public float fFindHorizontalDistance(PowerPoint.Shape objThisShape)
        {
            float HDistance = 0;
            try
            {
                
                float LargeAngle = 0;
                float sngBase = objThisShape.Width / 2;
                float sngPerpendicular = objThisShape.Height / 2;
                float SmallAngle = (sngPerpendicular / sngBase) * (180 / (22 / 7));       //Inverse Tan - Convert into Degrees
                LargeAngle = objThisShape.Rotation;
                if (LargeAngle >= 0 && LargeAngle <= 90) { LargeAngle = LargeAngle; }
                else if (LargeAngle > 90 && LargeAngle <= 180) { LargeAngle = 180 - LargeAngle; }
                else if (LargeAngle > 180 && LargeAngle <= 270) { LargeAngle = LargeAngle - 180; }
                else if (LargeAngle > 270 && LargeAngle <= 360) { LargeAngle = 360 - LargeAngle; }
                float MainAngle = LargeAngle + SmallAngle;
                float sngHypo = fCalculateHypotenuse(objThisShape.Width, objThisShape.Height) / 2;
                HDistance = sngHypo * (float)Math.Sin(MainAngle * (22 / 7) / 180);
                
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "fFindHorizontalDistance");
            }
            return HDistance;

        }

        public float fCalculateHypotenuse(float sngBase,float sngHeight)
        {
            float fCalculate = 0;
            try
            {
               //fCalculate = Math.Sqrt(sngBase ^ 2 + sngHeight ^ 2);
                fCalculate = (float)Math.Sqrt(Math.Pow(sngBase, 2) + Math.Pow(sngHeight, 2));
               
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "fCalculateHypotenuse");
            }
            return (fCalculate);
        }

        public float fFindVerticalDistance(PowerPoint.Shape objThisShape)
        {
            float VDistance = 0;
            float LargeAngle = 0;
            try
            {
                float sngBase = objThisShape.Width / 2;
                float sngPerpendicular = objThisShape.Height / 2;
                float SmallAngle = (sngPerpendicular / sngBase) * (180 / (22 / 7));      //'Inverse Tan - Convert into Degrees
                LargeAngle = objThisShape.Rotation;
                // ---'Adjust the Angle (Virtually for programming)
                //if(LargeAngle >= 0 && LargeAngle <= 90) { LargeAngle = LargeAngle; }
                if (LargeAngle > 90 && LargeAngle <= 180) { LargeAngle = 180 - LargeAngle; }
                if (LargeAngle > 180 && LargeAngle <= 270) { LargeAngle = LargeAngle - 180; }
                if (LargeAngle > 270 && LargeAngle <= 360) { LargeAngle = 360 - LargeAngle; }
                float MainAngle = LargeAngle + SmallAngle;
                float sngHypo = fCalculateHypotenuse(objThisShape.Width, objThisShape.Height) / 2;
                VDistance = sngHypo * (float)Math.Sin(MainAngle * (22 / 7) / 180);

            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "fFindVerticalDistance");
            }
            return (VDistance);
        }
        public string CheckFormat_SBox(int sldNum, string tshpName, DataTable dt, bool sCheck=false)
        {
            string sErr = null;
            try
            {
                CMsoTriState msoObj = new CMsoTriState();
                string shpName = Convert.ToString(dt.Rows[0]["Name"]);
                PowerPoint.Shape curShape = ActivePPT.Slides[sldNum].Shapes[tshpName];
                if (curShape.Line.Visible != msoObj.getMsoTriState(Convert.ToInt32(dt.Rows[0]["LineVisible"])))
                {
                    sErr = sErr + "Line Border" + '\n';
                }
                if (curShape.Fill.Visible != msoObj.getMsoTriState(Convert.ToInt32(dt.Rows[0]["FillVisible"])))
                    //curShape.Fill.Transparency != Convert.ToInt32(dt.Rows[0]["FillTransparency"]))
                {
                    sErr = sErr + "Shape Fill" + '\n';
                }

                // -'-------------- Shape Size and Position ---------------

                float tempBottom = curShape.Top + curShape.Height;
                float shpLeft = (float)Convert.ToDouble(dt.Rows[0]["ShapeLeft"]);
                float shpTop = (float)Convert.ToDouble(dt.Rows[0]["ShapeTop"]);
                if (Convert.ToInt32(curShape.Left) != Convert.ToInt32(shpLeft) || Convert.ToInt32(curShape.Top) != Convert.ToInt32(shpTop)) { sErr = sErr + "Position Error" + '\n'; }

                //float shpWidth = (float)Convert.ToDouble(dt.Rows[0]["ShapeWidth"]);
                //float shpHeight = (float)Convert.ToDouble(dt.Rows[0]["ShapeHeight"]);
                //-    Get the Margin from mapping -
                float leftMargin = (float)Convert.ToDouble(dt.Rows[0]["MarginLeft"]);
                float RightMargin = (float)Convert.ToDouble(dt.Rows[0]["MarginRight"]);
                float TopMargin = (float)Convert.ToDouble(dt.Rows[0]["MarginTop"]);
                float BottomMargin = (float)Convert.ToDouble(dt.Rows[0]["MarginBottom"]);
                if (curShape.TextFrame.MarginLeft != leftMargin || curShape.TextFrame.MarginRight != RightMargin || curShape.TextFrame.MarginTop != TopMargin
                    || curShape.TextFrame.MarginBottom != BottomMargin)
                {
                    sErr = sErr + "Margin Error" + '\n';
                }
                // '--------------------- Check the Font Error -------------------
                string fontName = dt.Rows[0]["FontName"].ToString();
                int Bold = Convert.ToInt32(dt.Rows[0]["Bold"]);
                int Italics = Convert.ToInt32(dt.Rows[0]["Italics"]);
                int Underline = Convert.ToInt32(dt.Rows[0]["Underline"]);
                int FontSize = Convert.ToInt32(dt.Rows[0]["FontSize"]);
                List<int> rgbVal1 = msoObj.get_RGBValue(Convert.ToString(dt.Rows[0]["FontColor"]));                     // Set the Object 


                PowerPoint.Font txtRangefont = curShape.TextFrame.TextRange.Font;

                int rgbCode = txtRangefont.Color.RGB;
                //int B = rgbCode/65536;
                //int G = (rgbCode - B * 65536) / 256;
                //int R = rgbCode - B * 65536 - G * 256;

                int r = rgbCode & 255;
                int g = rgbCode / 256 & 255;
                int b = rgbCode / 65536;//((rgbCode / 256) ^ 2) & 255;
                if (tshpName == "Text Box" || tshpName == "Quote Box")
                {
                    if (txtRangefont.Name != fontName)
                    {
                        sErr = sErr + "Font Error" + '\n';
                    }
                }
                else
                {
                    if (txtRangefont.Name != fontName || txtRangefont.Bold != msoObj.getMsoTriState(Bold) || txtRangefont.Italic != msoObj.getMsoTriState(Italics) ||
                        txtRangefont.Underline != msoObj.getMsoTriState(Underline) || txtRangefont.Size != FontSize ||
                        (rgbVal1[0] != b || rgbVal1[1] != g || rgbVal1[2] != r))              // Font Color is Fix --
                    {
                        sErr = sErr + "Font Error" + '\n';
                    }
                }
                //----- Check  Paragraph Error
                //--- Get a specification from mapping excel file
                int prgBullet = Convert.ToInt32(dt.Rows[0]["ParagraphBullet"]);
                int prgAlignment = Convert.ToInt32(dt.Rows[0]["ParagraphAlignment"]);
                int prghpun = Convert.ToInt32(dt.Rows[0]["ParagraphHangingPunctuation"]);
                float prgbspace = (float)Convert.ToDouble(dt.Rows[0]["ParagraphSpaceBefore"]);
                float prgaspace = (float)Convert.ToDouble(dt.Rows[0]["ParagraphSpaceAfter"]);
                float prgwspace = (float)Convert.ToDouble(dt.Rows[0]["ParagraphSpaceWithin"]);
                float prgflmargin = (float)Convert.ToDouble(dt.Rows[0]["RulerLevel1FirstMargin"]);                    // It's not part of Paragraph format
                float prgllmargin = (float)Convert.ToDouble(dt.Rows[0]["RulerLevel1LeftMargin"]);                     // It's not part of Paragraph format
                PowerPoint.ParagraphFormat prgformat = curShape.TextFrame.TextRange.ParagraphFormat;
                if (prgformat.Bullet.Type != msoObj.getPpBulletType(prgBullet) || prgformat.Alignment != msoObj.ParagraphFormatAlignment(prgAlignment) ||
                    prgformat.SpaceBefore != prgbspace || prgformat.SpaceAfter != prgaspace || prgformat.SpaceWithin != prgwspace)

                {
                    sErr = sErr + "Paragraph Error" + '\n';
                }
                // ---    Size and rotation type error  
                int Rotaion = Convert.ToInt32(dt.Rows[0]["Rotaion"]);
                int LARatio = Convert.ToInt32(dt.Rows[0]["LockAspectRatio"]);
                int Orientation = Convert.ToInt32(dt.Rows[0]["Orientation"]);
                if (curShape.Rotation != Rotaion || curShape.TextFrame.Orientation != msoObj.getOrientation(Orientation) || curShape.LockAspectRatio != msoObj.getMsoTriState(LARatio))
                {
                    sErr = sErr + "Size and rotation" + '\n';
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CheckFormat_SBox");
            }
            return (sErr);
        }
    }
}
