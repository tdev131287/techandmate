using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace TSCPPT_Addin
{
    public partial class frmPPTFormat : Form
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation OpenedPPT = Globals.ThisAddIn.Application.ActivePresentation;
        //PowerPoint.Presentation OpenedPPT;
        string selectedFolder = null;
        public frmPPTFormat()
        {
            InitializeComponent();
        }

        private void btn_Browse_Click(object sender, EventArgs e)
        {
            
            Microsoft.Office.Core.FileDialog fileDialog = ppApp.get_FileDialog(MsoFileDialogType.msoFileDialogFolderPicker);
            fileDialog.InitialFileName = "c:\\Temp\\";
            int nres = fileDialog.Show();
            if (nres == -1) //ok
            {
                FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
                string[] selectedFolders = selectedItems.Cast<string>().ToArray();
                if (selectedFolders.Length > 0)
                {
                   selectedFolder = selectedFolders[0];
                }
                string[] fileEntries = Directory.GetFiles(selectedFolder);
                foreach (string fileName in fileEntries)
                {
                    if (Path.GetExtension(fileName) == ".pptx")
                    {
                        fileGridView.Rows.Add(Path.GetFileName(fileName));
                        
                    }
                }
                foreach (DataGridViewRow row in fileGridView.Rows)
                {
                    DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[2];
                    chk.Value = !(chk.Value == null ? false : (bool)chk.Value); //because chk.Value is initialy null
                }
            
            }
            this.Width = 395;
            this.Height = 430;

        }

        private void frmPPTFormat_Load(object sender, EventArgs e)
        {
            this.Width = 395;
            this.Height = 147;
            rbselectedslide.Checked = true;
            fileGridView.AllowUserToAddRows = false;
        }

        private void btn_Submit_Click(object sender, EventArgs e)
        {
            pptfunctions funObj = new pptfunctions();
            PPTActions actionObj = new PPTActions();
            Formatshapes PPTshpFormat = new Formatshapes();
            ppApp.DisplayAlerts =PowerPoint.PpAlertLevel.ppAlertsNone;
            try
            {
                if (rbselectedppt.Checked == true)          // Work on Current PPT
                {
                    OpenedPPT.PageSetup.SlideWidth = 960;
                    OpenedPPT.PageSetup.SlideHeight = 540;
                    ppApp.ActivePresentation.ApplyTheme(PPTAttribute.mPPTPath);
                    funObj.DeleteOldTSCMasters();

                    foreach (PowerPoint.Slide sld in OpenedPPT.Slides)
                    {
                        foreach (PowerPoint.Shape shp in sld.Shapes)
                        {
                            if (shp.Name == "Road Map" || shp.Name == "Title Box" || shp.Name == "Note Box" || shp.Name == "Source Box")
                            {
                                DataTable dt = actionObj.get_specification(shp.Name);
                                FormatShape(sld.SlideIndex, shp.Name, dt);
                            }
                        }
                    }
                    this.Close();
                }
                else if (rbselectedslide.Checked == true)   // Work on Current Slide
                {
                    OpenedPPT.PageSetup.SlideWidth = 960;
                    OpenedPPT.PageSetup.SlideHeight = 540;
                    ppApp.ActivePresentation.ApplyTheme(PPTAttribute.mPPTPath);
                    funObj.DeleteOldTSCMasters();

                    int sldnum = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
                    PowerPoint.Slide sld = OpenedPPT.Slides[sldnum];
                    foreach (PowerPoint.Shape shp in sld.Shapes)
                    {
                        if (shp.Name == "Road Map" || shp.Name == "Title Box" || shp.Name == "Note Box" || shp.Name == "Source Box")
                        {
                            DataTable dt = actionObj.get_specification(shp.Name);
                            FormatShape(sld.SlideIndex, shp.Name, dt);
                        }
                    }
                    this.Close();
                }
                else if (rbtn_selectfolder.Checked == true)
                {
                    string pathString = System.IO.Path.Combine(selectedFolder, "TSC_PPT_29032018");
                    System.IO.Directory.CreateDirectory(pathString);
                    string pptPath = null, newfilepath = null;
                    int rowscount = fileGridView.Rows.Count;
                    foreach (DataGridViewRow row in fileGridView.Rows)
                    {
                        //for templated control
                        if (Convert.ToBoolean(row.Cells[2].Value) == true)
                        {
                            string fileName = row.Cells[0].Value.ToString();
                            newfilepath = pathString + "\\" + fileName;
                            pptPath = selectedFolder + "\\" + fileName;
                            OpenedPPT = ppApp.Presentations.Open(pptPath);
                            //-- Apply New TSC Theme 
                            ppApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                            OpenedPPT.PageSetup.SlideWidth = 960;
                            OpenedPPT.PageSetup.SlideHeight = 540;
                            ppApp.ActivePresentation.ApplyTheme(PPTAttribute.mPPTPath);
                            funObj.DeleteOldTSCMasters();
                            // -- Format Slide 
                            foreach (PowerPoint.Slide sld in OpenedPPT.Slides)
                            {
                                foreach (PowerPoint.Shape shp in sld.Shapes)
                                {
                                    if (shp.Name == "Road Map" || shp.Name == "Title Box" || shp.Name == "Note Box" || shp.Name == "Source Box")
                                    {
                                        DataTable dt = actionObj.get_specification(shp.Name);
                                        FormatShape(sld.SlideIndex, shp.Name, dt);

                                    }
                                }
                            }
                            row.Cells[1].Value = "Done";
                            OpenedPPT.SaveAs(newfilepath);
                            OpenedPPT.Close();
                            // If Process has been done the update the status Done
                        }
                    } // Close Outer For loop
                    ppApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format -All Slide");
            }
        }


        public void FormatShape(int sldNum, string shpNum, DataTable dt)
        {
            CMsoTriState pptMsoObject = new CMsoTriState();
            try
            {
                PowerPoint.Shape Activeshape = OpenedPPT.Slides[sldNum].Shapes[shpNum];
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
                   
                    
                }// Apply a msoValue Condition 
                 

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
                if (Activeshape.Name == "Road Map")
                {
                    Activeshape.TextFrame2.TextRange.Font.Spacing = 1;
                }
                //----- Set shape text range format --
                
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
                Activeshape.Left = (float)Convert.ToDouble((dt.Rows[0]["ShapeLeft"]));
                Activeshape.Top = (float)Convert.ToDouble((dt.Rows[0]["ShapeTop"]));
                Activeshape.Width = (float)Convert.ToDouble((dt.Rows[0]["ShapeWidth"]));
                Activeshape.Height = (float)Convert.ToDouble((dt.Rows[0]["ShapeHeight"]));
                

                if (Activeshape.TextFrame.TextRange.Text.PadLeft(1) == " ") { Activeshape.TextFrame.TextRange.Text = Activeshape.TextFrame.TextRange.Text.Trim(); }

                var txtRange = Activeshape.TextFrame.TextRange;
                int txtLength = Activeshape.TextFrame.TextRange.Count;
                
                //- - IF Road Map

                if (Convert.ToString(dt.Rows[0]["Name"]) == "Road Map") { txtRange.ChangeCase(pptMsoObject.txtChangeCase(Convert.ToInt32(dt.Rows[0]["Case"]))); }    // Apply a msoValue Condition  
                if (Convert.ToString(dt.Rows[0]["Name"]) == "Source Box")
                {
                    Activeshape.TextFrame.Ruler.Levels[1].FirstMargin = 0;
                    Activeshape.TextFrame.Ruler.Levels[1].LeftMargin = 0;
                    for (int x = 1; x <= txtRange.Paragraphs().Count; x++)
                    {
                        txtRange.Paragraphs(x).ParagraphFormat.HangingPunctuation = MsoTriState.msoFalse;
                    }
                }


            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "FormatShape");
            }
        }

        private void rbselectedslide_CheckedChanged(object sender, EventArgs e)
        {
            this.Width = 395;
            this.Height = 147;
            this.fileGridView.Rows.Clear();
            btn_Browse.Enabled = false;
        }

        private void rbselectedppt_CheckedChanged(object sender, EventArgs e)
        {
            this.Width = 395;
            this.Height = 147;
            this.fileGridView.Rows.Clear();
            btn_Browse.Enabled = false;
        }

        private void rbtn_selectfolder_CheckedChanged(object sender, EventArgs e)
        {
            btn_Browse.Enabled = true;
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmPPTFormat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }
    }
}
