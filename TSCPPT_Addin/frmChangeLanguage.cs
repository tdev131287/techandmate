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
    public partial class frmChangeLanguage : Form
    {
        public frmChangeLanguage()
        {
            InitializeComponent();
        }

        private void frmChangeLanguage_Load(object sender, EventArgs e)
        {
            Algorithms obj = new Algorithms();
            cmb_lng.Items.Add("US English");
            cmb_lng.Items.Add("UK English");
            cmb_lng.SelectedIndex = 1;
            string docObj = obj.DocumentLanguage();
            lbl_Clng.Text = "Language of active document:  " + docObj;

        }

        private void btn_setlanguage_Click(object sender, EventArgs e)
        {
            PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
            PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;

            string lang = cmb_lng.Text;
            int scount = ppApp.ActivePresentation.Slides.Count;
            MsoLanguageID langID= MsoLanguageID.msoLanguageIDEnglishUS, curr_lang;
            if(lang == "US English") { langID = MsoLanguageID.msoLanguageIDEnglishUS; }
            else if(lang == "UK English") { langID = MsoLanguageID.msoLanguageIDEnglishUK; }
            curr_lang = ActivePPT.DefaultLanguageID;
            ActivePPT.DefaultLanguageID = langID;

            for(int sld=1;sld<=ActivePPT.Slides.Count;sld++)
            {
                foreach(PowerPoint.Shape shp in ActivePPT.Slides[sld].Shapes)
                {
                    // '---------------- Check if it is a table
                    if (shp.Type == MsoShapeType.msoTable)
                    {
                        for (int r = 1; r <= shp.Table.Rows.Count; r++)
                        {
                            for (int c = 1; c <= shp.Table.Columns.Count; c++)
                            {
                                shp.Table.Cell(r, c).Shape.TextFrame.TextRange.LanguageID = langID;
                            }
                        }
                    }
                    //'------------- Check if it is a group of shapes
                    if (shp.Type == MsoShapeType.msoGroup)
                    {
                        if (shp.GroupItems.Count > 0)
                        {
                            for (int i = 1; i <= shp.GroupItems.Count; i++)
                            {
                                if (shp.GroupItems[i].HasTextFrame == MsoTriState.msoTrue)
                                {
                                    shp.GroupItems[i].TextFrame.TextRange.LanguageID = langID;
                                }
                            }
                        }
                    }
                    //'-------------- Check if it is a simple shape
                    if (shp.HasTextFrame == MsoTriState.msoTrue)
                    {
                        shp.TextFrame.TextRange.LanguageID = langID;

                    }
                }
                ActivePPT.Slides[sld].NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.LanguageID = langID;
            }
            MessageBox.Show("Language has been changed to " + lang + " in active presentation.");
            this.Close();
        }
    }
}
