using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Windows.Forms;

namespace TSCPPT_Addin
{
    class Formatshapes
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        PPTActions actionObj = new PPTActions();
        Shapecheck chkObj = new Shapecheck();
        char splitChar = '|';
        public void FormatShape(int sldNum, string shpNum, DataTable dt, bool textAdjust = true, bool sizeAdjust = true)
        {
            CMsoTriState pptMsoObject = new CMsoTriState();
            try
            {
                PowerPoint.Shape Activeshape = ActivePPT.Slides[sldNum].Shapes[shpNum];
                List<int> rgbVal = new List<int>();
                List<int> rgbVal1 = new List<int>();
                //if (textAdjust == false) { textAdjust = true; }
                //if (sizeAdjust == false) { sizeAdjust = true; }
                Activeshape.Name = Convert.ToString(dt.Rows[0]["Name"]);
                //-- Selected object is line --
                Activeshape.Line.Visible = pptMsoObject.getMsoTriState(Convert.ToInt32(dt.Rows[0]["LineVisible"])); //MsoTriState.msoTrue;                               // Apply a msoValue Condition 
                Activeshape.Fill.Visible = pptMsoObject.getMsoTriState(Convert.ToInt32(dt.Rows[0]["FillVisible"]));//MsoTriState.msoFalse;   
                if (Convert.ToInt32(dt.Rows[0]["LineVisible"]) == -1)
                {
                    rgbVal1 = pptMsoObject.get_RGBValue(Convert.ToString(dt.Rows[0]["LineForeColor"]));
                    Activeshape.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(rgbVal1[0], rgbVal1[1], rgbVal1[2]).ToArgb();               // Apply a msoValue Condition 
                    rgbVal1 = pptMsoObject.get_RGBValue(Convert.ToString(dt.Rows[0]["LineBackColor"]));
                    Activeshape.Line.BackColor.RGB = System.Drawing.Color.FromArgb(rgbVal1[0], rgbVal1[1], rgbVal1[2]).ToArgb();              // Apply a msoValue Condition 
                    Activeshape.Line.Weight = (float)Convert.ToDouble(dt.Rows[0]["LineWeight"]);
                    Activeshape.Line.Style = MsoLineStyle.msoLineSingle;
                                             // Apply a msoValue Condition 
                                                                                                                       //Activeshape.Line.Visible = MsoTriState.msoTrue;
                }// Apply a msoValue Condition 
                 //'Shape Fill
                 //Activeshape.Fill.Visible = MsoTriState.msoCTrue;
                
                if (Convert.ToInt32(dt.Rows[0]["FillVisible"]) == -1)
                {
                    rgbVal1 = pptMsoObject.get_RGBValue(Convert.ToString(dt.Rows[0]["FillColor"]));
                    Activeshape.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(rgbVal1[0], rgbVal1[1], rgbVal1[2]).ToArgb();               // Apply a msoValue Condition
                    Activeshape.Fill.Transparency = Convert.ToInt32(dt.Rows[0]["FillTransparency"]);
                }

                //Shape Shadow
                Activeshape.Shadow.Visible = MsoTriState.msoFalse;                                // Apply a msoValue Condition
                if (Convert.ToInt32(dt.Rows[0]["ShadowVisible"]) == -1)
                {
                    rgbVal1 = pptMsoObject.get_RGBValue(Convert.ToString(dt.Rows[0]["LineForeColor"]));
                    Activeshape.TextFrame.TextRange.Font.Name = Convert.ToString(dt.Rows[0]["FontName"]);
                    //LineForeColor
                    Activeshape.Shadow.ForeColor.RGB = System.Drawing.Color.FromArgb(rgbVal1[0], rgbVal1[1], rgbVal1[2]).ToArgb();            // Apply a msoValue Condition
                    Activeshape.Shadow.Blur = Convert.ToInt32(dt.Rows[0]["ShadowBlur"]);
                    Activeshape.Shadow.Size = Convert.ToInt32(dt.Rows[0]["ShadowSize"]);
                    Activeshape.Shadow.Transparency = Convert.ToInt32(dt.Rows[0]["ShadowTransparency"]);
                    Activeshape.Shadow.OffsetX = Convert.ToInt32(dt.Rows[0]["ShadowOffsetX"]);
                    Activeshape.Shadow.OffsetY = Convert.ToInt32(dt.Rows[0]["ShadowOffsetY"]);
                }
                Activeshape.Rotation = Convert.ToInt32(dt.Rows[0]["Rotaion"]);

                Activeshape.LockAspectRatio = pptMsoObject.getMsoTriState(Convert.ToInt32(dt.Rows[0]["LockAspectRatio"]));

                //Shape Position


                // -- Set a Text Frame -
                Activeshape.TextFrame.Orientation = pptMsoObject.getOrientation(Convert.ToInt32(dt.Rows[0]["Orientation"]));
                Activeshape.TextFrame.VerticalAnchor = pptMsoObject.getVerticalAnchor(Convert.ToInt32(dt.Rows[0]["VerticalAnchor"]));
                Activeshape.TextFrame.AutoSize = pptMsoObject.TxtAutoSize(0);     // Only for test -
                Activeshape.TextFrame.AutoSize = pptMsoObject.TxtAutoSize(Convert.ToInt32(dt.Rows[0]["AutoSize"]));                                          // Apply a msoValue Condition
                Activeshape.TextFrame.MarginLeft = (float)Convert.ToDouble((dt.Rows[0]["MarginLeft"]));
                Activeshape.TextFrame.MarginRight = (float)Convert.ToDouble((dt.Rows[0]["MarginRight"]));
                Activeshape.TextFrame.MarginTop = (float)Convert.ToDouble((dt.Rows[0]["MarginTop"]));
                Activeshape.TextFrame.MarginBottom = (float)Convert.ToDouble((dt.Rows[0]["MarginBottom"]));
                Activeshape.TextFrame.WordWrap = MsoTriState.msoCTrue;

                //----- Set shape text range format --
                if (textAdjust == true)
                {
                    if (Activeshape.Name == "Note Box" || Activeshape.Name == "Text Box")
                    {
                        string defaultText = Convert.ToString(dt.Rows[0]["DefaultText"]).Replace("\n", "");
                        string[] text = defaultText.Split(splitChar).ToArray();
                        if (Activeshape.Name == "Note Box") { Activeshape.TextFrame.TextRange.Text = text[0] + '\n' + text[1]; }
                        else if (Activeshape.Name == "Text Box")
                        {
                            Activeshape.TextFrame2.TextRange.Text = text[0] + "\r\n" + text[1] + "\r\n" + text[2] + "\r\n" + text[3];
                            //Activeshape.TextFrame.TextRange.Text = text[0] + '\n' + text[1] + '\n' + text[2] + '\n' + text[3];
                        }
                    }
                    else
                    {
                        Activeshape.TextFrame.TextRange.Text = Convert.ToString(dt.Rows[0]["DefaultText"]);
                    }
                }
                //List<int> rgbVal = new List<int>();
                rgbVal = pptMsoObject.get_RGBValue(Convert.ToString(dt.Rows[0]["FontColor"]));
                Activeshape.TextFrame.TextRange.Font.Name = Convert.ToString(dt.Rows[0]["FontName"]);
                Activeshape.TextFrame.TextRange.Font.Bold = pptMsoObject.getMsoTriState(Convert.ToInt32(dt.Rows[0]["Bold"]));//MsoTriState.msoCTrue;                        // Apply a msoValue Condition
                Activeshape.TextFrame.TextRange.Font.Italic = pptMsoObject.getMsoTriState(Convert.ToInt32(dt.Rows[0]["Italics"]));//MsoTriState.msoFalse;                       // Apply a msoValue Condition
                Activeshape.TextFrame.TextRange.Font.Underline = pptMsoObject.getMsoTriState(Convert.ToInt32(dt.Rows[0]["Underline"]));//MsoTriState.msoFalse;                    // Apply a msoValue Condition
                Activeshape.TextFrame.TextRange.Font.Size = Convert.ToInt32(dt.Rows[0]["FontSize"]);
                Activeshape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.FromArgb(rgbVal[0], rgbVal[1], rgbVal[2]).ToArgb();          // Apply a msoValue Condition
                Activeshape.TextFrame.TextRange.Font.Shadow = pptMsoObject.getMsoTriState(Convert.ToInt32(dt.Rows[0]["Shadow"])); //MsoTriState.msoFalse;                       // Apply a msoValue Condition

                Activeshape.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Bullet.Type = pptMsoObject.getPpBulletType(Convert.ToInt32(dt.Rows[0]["ParagraphBullet"]));  // Apply a msoValue Condition(Need to Check)

                // ---- Set the ParagraphFormat --
                Activeshape.TextFrame.TextRange.ParagraphFormat.Alignment = pptMsoObject.ParagraphFormatAlignment(Convert.ToInt32(dt.Rows[0]["ParagraphAlignment"]));//PowerPoint.PpParagraphAlignment.ppAlignRight;               // Apply a msoValue Condition
                Activeshape.TextFrame.TextRange.ParagraphFormat.HangingPunctuation = pptMsoObject.getMsoTriState(Convert.ToInt32(dt.Rows[0]["ParagraphHangingPunctuation"]));        // Apply a msoValue Condition
                Activeshape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = (float)Convert.ToDouble((dt.Rows[0]["ParagraphSpaceBefore"]));
                Activeshape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = (float)Convert.ToDouble((dt.Rows[0]["ParagraphSpaceAfter"]));
                Activeshape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = (float)Convert.ToDouble((dt.Rows[0]["ParagraphSpaceWithin"]));
                //Activeshape.TextFrame.Ruler.TabStops.DefaultSpacing = 45;


                Activeshape.TextFrame.Ruler.Levels[1].FirstMargin = (float)Convert.ToDouble((dt.Rows[0]["RulerLevel1FirstMargin"]));
                Activeshape.TextFrame.Ruler.Levels[1].LeftMargin = (float)Convert.ToDouble((dt.Rows[0]["RulerLevel1LeftMargin"]));

                // -'------------------- Specific Adjustments -----------------
                if (sizeAdjust == true)
                {
                    Activeshape.Left = (float)Convert.ToDouble((dt.Rows[0]["ShapeLeft"]));
                    Activeshape.Top = (float)Convert.ToDouble((dt.Rows[0]["ShapeTop"]));
                    Activeshape.Width = (float)Convert.ToDouble((dt.Rows[0]["ShapeWidth"]));
                    Activeshape.Height = (float)Convert.ToDouble((dt.Rows[0]["ShapeHeight"]));

                }
                


                if (Activeshape.TextFrame.TextRange.Text.PadLeft(1) == " ") { Activeshape.TextFrame.TextRange.Text = Activeshape.TextFrame.TextRange.Text.Trim(); }

                var txtRange = Activeshape.TextFrame.TextRange;
                int txtLength = Activeshape.TextFrame.TextRange.Count;
                //                                                      Check it why we do the comment -
                //byte result = Convert.ToByte(txtRange.Characters(1,1));
                //if(result==9 || result==10||result==11 || result==13|| result == 32) { txtRange.Text = txtRange.Text.PadRight(txtRange.Text.Length - 1); }
                //- - IF note box 

                if (Convert.ToString(dt.Rows[0]["Name"]) == "Note Box")
                {
                    //int txtLen = Convert.ToString(dt.Rows[0]["DefaultText"]).Length;
                    //if (txtRange.Text.PadLeft(txtLen) != Convert.ToString(dt.Rows[0]["DefaultText"]))
                    //{
                    //    if (txtRange.Text.PadLeft(6) == "Notes:") { txtRange.Text = txtRange.Text.Substring(7, txtRange.Text.Length); }
                    //    else if (txtRange.Text.PadLeft(5) == "Notes:") { txtRange.Text = txtRange.Text.Substring(6, txtRange.Text.Length); }
                    //    txtRange.Text = Convert.ToString(dt.Rows[0]["DefaultText"]) + txtRange.Text;
                    //}
                }
                //- - IF Source Box 
                if (Convert.ToString(dt.Rows[0]["Name"]) == "Source Box")
                {
                    //int txtLen = Convert.ToString(dt.Rows[0]["DefaultText"]).Length;
                    //if (txtRange.Text.PadLeft(txtLen) != Convert.ToString(dt.Rows[0]["DefaultText"]))
                    //{
                    //    if (txtRange.Text.Substring(0,8) == "Sources:") { txtRange.Text = txtRange.Text.Substring(9, txtRange.Text.Length); }
                    //    else if (txtRange.Text.PadLeft(5) == "Sources:") { txtRange.Text = txtRange.Text.Substring(8, txtRange.Text.Length); }
                    //    txtRange.Text = Convert.ToString(dt.Rows[0]["DefaultText"]) + txtRange.Text;
                    //}
                }
                //- - IF Road Map

                if (Convert.ToString(dt.Rows[0]["Name"]) == "Road Map") { txtRange.ChangeCase(pptMsoObject.txtChangeCase(Convert.ToInt32(dt.Rows[0]["Case"]))); }    // Apply a msoValue Condition  


                // --"Quote Box"
                int sn = 0, en = 0;
                if (Convert.ToString(dt.Rows[0]["Name"]) == "Quote Box")
                {
                    var rngs = txtRange.Find("–", 0, MsoTriState.msoCTrue, MsoTriState.msoFalse);
                    var rngE = txtRange.Find(")", 0, MsoTriState.msoCTrue, MsoTriState.msoFalse);
                    if (rngs != null) { sn = rngs.Start; }
                    if (rngE != null) { en = rngE.Start + 3; }
                    if (rngE != null && rngs != null)
                    {
                        txtRange.Characters(sn, en).Font.Italic = MsoTriState.msoCTrue;
                        txtRange.Characters(sn, en).Font.Italic = MsoTriState.msoFalse;
                        txtRange.Characters(sn, en).Font.Bold = MsoTriState.msoCTrue;
                    }
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "FormatShape");
            }
        }

       
        // -- Main method to correct Each shapes format 
        public void CorrectFormat_Selected(int sldNum, string shpName)
        {
            try
            {
                PowerPoint.Presentation ActivePPT = ppApp.ActivePresentation;
                CorrectFormat_ShapeInSlide(sldNum, shpName);
                if (ActivePPT.Slides[sldNum].CustomLayout.Name == "Title Slide") { CorrectFormat_TitleSlide(sldNum, shpName); }
                else if (ActivePPT.Slides[sldNum].CustomLayout.Name == "Content Slide") { CorrectFormat_ContentSlide(sldNum, shpName); }
                else if (ActivePPT.Slides[sldNum].CustomLayout.Name == "Divider Slide") { CorrectFormat_DividerSlide(sldNum, shpName); }
                else { CorrectFormat_MainSlide(sldNum, shpName); }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CorrectFormat_Selected");
            }

        }

        public void CorrectFormat_ShapeInSlide(int sldNum, string shpName)
        {
            Reviewformat formatobj = new Reviewformat();

            try
            {
                float sldLeft = 0;
                float sldTop = 0;
                float sldRight = ActivePPT.PageSetup.SlideWidth;
                float sldBottom = ActivePPT.PageSetup.SlideHeight;
                
                PowerPoint.Shape tShp = ActivePPT.Slides[sldNum].Shapes[shpName];
                if (tShp.Type != MsoShapeType.msoLine)
                {
                    float sngCenterH = tShp.Left + tShp.Width / 2;
                    //fFindHorizontalDistance
                    float sngDistH = formatobj.fFindHorizontalDistance(tShp);    //This is half of horizontal distance
                    float sngShapeLeft = sngCenterH - sngDistH;        //Actual Left
                    float sngShapeRight = sngCenterH + sngDistH;       //Actual Right
                                                                       //Vertical Distances
                    float sngCenterV = tShp.Top + tShp.Height / 2;
                    float sngDistV = formatobj.fFindVerticalDistance(tShp);      //This is half of vertical distance
                    float sngShapeTop = sngCenterV - sngDistV;         //Actual Top
                    float sngShapeBottom = sngCenterV + sngDistV;      //Actual Bottom
                    if ((tShp.Left + tShp.Width) > sldRight) { tShp.Left = sldRight - tShp.Width; }
                    if ((tShp.Top + tShp.Height) > sldBottom) { tShp.Top = sldBottom - tShp.Height; }
                    if (tShp.Top < sldTop) { tShp.Top = sldTop; }
                    if (tShp.Left < sldLeft) { tShp.Left = sldLeft; }
                    //if (sngShapeLeft < sldLeft) { tShp.Left = tShp.Left + (sldLeft - sngShapeLeft); }
                    //if (sngShapeRight > sldRight) { tShp.Left = sldRight-tShp.Width; }  //tShp.Left - (sngShapeRight - sldRight)
                    //if (sngShapeTop < sldTop) { tShp.Top = tShp.Top + (sldTop - sngShapeTop); }
                    //if (sngShapeBottom > sldBottom) { tShp.Top = tShp.Top - (sngShapeBottom - sldBottom); }
                }
                else
                {
                    if (tShp.Left < sldLeft) { tShp.Left = sldLeft; }
                    if ((tShp.Left + tShp.Width) > sldRight) { tShp.Left = tShp.Left - (tShp.Left + tShp.Width - sldRight); }
                    if (tShp.Top < sldTop) { tShp.Top = sldTop; }
                    if ((tShp.Top + tShp.Height) > sldBottom) { tShp.Top = tShp.Top - (tShp.Top + tShp.Height - sldBottom); }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CorrectFormat_ShapeInSlide");
            }
        }

        public void CorrectFormat_TitleSlide(int sldNum, string shpName)
        {

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
                        if (Tchkfind == shpName && dt1.Rows.Count != 0) { FormatShape(sldNum, shpName, dt1, false); }
                        else if (Dchkfind == shpName && dt2.Rows.Count != 0) { FormatShape(sldNum, shpName, dt2, false); }
                        else if (client == shpName && dt3.Rows.Count != 0) { FormatShape(sldNum, shpName, dt3, false); }
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CorrectFormat_TitleSlide");
            }
        }

        public void CorrectFormat_ContentSlide(int sldNum, string shpName)
        {
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
                        if (shpName == CTitle && dt.Rows.Count != 0) { FormatShape(sldNum, shpName, dt, false); }
                        else if (shpName == CSection && dt1.Rows.Count != 0) { FormatShape(sldNum, shpName, dt1, false); }
                        else if (shpName == CNumber && dt2.Rows.Count != 0) { FormatShape(sldNum, shpName, dt2, false); }
                    }

                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CorrectFormat_ContentSlide");
            }
        }
        public void CorrectFormat_DividerSlide(int sldNum, string shpName)
        {
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

                        if (shpName == STitle && dt1.Rows.Count != 0) { FormatShape(sldNum, shpName, dt1, false); }
                        else if (shpName == SContent && dt2.Rows.Count != 0) { FormatShape(sldNum, shpName, dt2, false); }

                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CorrectFormat_DividerSlide");
            }
        }

        public void CorrectFormat_MainSlide(int sldNum, string shpName)
        {
            DataTable dt1 = new DataTable();                        // Get the specification of Title

            string dbShape;
            try
            {
                PowerPoint.Shape shp = ActivePPT.Slides[sldNum].Shapes[shpName];
                if (shp.Type == MsoShapeType.msoAutoShape || shp.Type == MsoShapeType.msoPlaceholder || shp.Type == MsoShapeType.msoTextBox)
                {
                    if (shp.HasTextFrame == MsoTriState.msoTrue)
                    {
                        // -----------------------------------------------------------
                        int hasspliter = shpName.IndexOf("_");
                        if (hasspliter != -1) { dbShape = shpName.Substring(0, hasspliter); }
                        else { dbShape = shpName; }
                        dt1 = actionObj.get_specification(dbShape);
                        if (dt1.Rows.Count != 0) { FormatShape(sldNum, shpName, dt1, false); };
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CorrectFormat_MainSlide");
            }
        }

        //--- Correct Format -For method 3
        public void CorrectFormat(int sldnum)
        {
            Reviewformat fobj = new Reviewformat();
            List<string> shpNames = new List<string>();
            try
            {
                shpNames = fobj.NameAllShapes(sldnum);
                if (shpNames.Count == 0) { return; }
                foreach (string shp in shpNames)
                {

                    CorrectFormat_ShapeInSlide(sldnum, shp);
                    if (ActivePPT.Slides[sldnum].CustomLayout.Name == "Title Slide") { CorrectFormat_TitleSlide(sldnum, shp); }
                    else if (ActivePPT.Slides[sldnum].CustomLayout.Name == "Content Slide") { CorrectFormat_ContentSlide(sldnum, shp); }
                    else if (ActivePPT.Slides[sldnum].CustomLayout.Name == "Divider Slide") { CorrectFormat_DividerSlide(sldnum, shp); }
                    else { CorrectFormat_MainSlide(sldnum, shp); }
                }
                DeleteFormatComments_M3(sldnum);
                MessageBox.Show("Format review and correction has been done", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CorrectFormat");
            }
        }

        public void DeleteFormatComments_M3(int sldnum)
        {
            int numComments = ActivePPT.Slides[sldnum].Comments.Count;
            PowerPoint.Comment myComment = null;
            try
            {
                if (numComments > 0)
                {
                    for (int i = numComments; i >= 1; i--)
                    {
                        myComment = ActivePPT.Slides[sldnum].Comments[i];
                        if (myComment.AuthorInitials == "TFR") { myComment.Delete(); }
                    } // --- 
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "DeleteFormatComments_M3");
            }
        }
    }// - Close class
} // -Close Name space 
