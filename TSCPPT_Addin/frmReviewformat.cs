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
    public partial class frmReviewformat : Form
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        public frmReviewformat()
        {
            InitializeComponent();
        }

        private void frmReviewformat_Load(object sender, EventArgs e)
        {
            cmb_RType.Items.Add("Selected");
            cmb_RType.Items.Add("All");
            cmb_RType.SelectedIndex = 0;
            rb_ReviewCorrect.Checked = true;
        }

        private void btn_Review_Click(object sender, EventArgs e)
        {
            string whichSlides = cmb_RType.Text;
            Algorithms algObj = new Algorithms();
            List<int> selSlides = new List<int>();
            Reviewformat formatObj = new Reviewformat();
            Formatshapes shpobj = new Formatshapes();
            frmErrorlist frmObj = new frmErrorlist();
            //int[] selSlides ;
            int selectedSlides = 0;
            try
            {
                bool uUnload = false, globalErrors = false;
                selectedSlides = ppApp.ActiveWindow.Selection.SlideRange.Count;
                if (whichSlides == "Current" || whichSlides == "Selected")
                {

                    if (selectedSlides == 0) { MessageBox.Show("Please select a slide and try again"); this.Close(); }
                    else
                    {
                        for (int i = 1; i <= selectedSlides; i++)
                        {
                            selSlides.Add(ppApp.ActiveWindow.Selection.SlideRange[i].SlideIndex);
                        }
                        //Array.Sort(selSlides);
                        selSlides.Sort();
                    }
                }
                else if (whichSlides == "All")
                {
                    selectedSlides = ppApp.ActivePresentation.Slides.Count;
                    for (int i = 1; i <= selectedSlides; i++) { selSlides.Add(i); }
                }
                if (selectedSlides == 0) { MessageBox.Show("Please select a slide and try again"); this.Close(); return; }
                else
                {
                    for (int i = 0; i < selectedSlides; i++)
                    {
                        int sldNum = selSlides[i];
                        algObj.SetNamesUnique(sldNum);                            // -Set the unique name of each object for get the property 
                        ppApp.ActiveWindow.View.GotoSlide(sldNum);
                        ActivePPT.Slides[sldNum].Select();
                        if (rb_Review.Checked == true) { formatObj.CheckFormat(sldNum, "method1"); }
                        else if (rb_ReviewCorrect.Checked == true)
                        {
                            formatObj.CheckFormat(sldNum, "method2");
                            if (PPTAttribute.exitFlag == false)
                            {
                                MessageBox.Show("Format review and correction has been done", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else if (rb_CorrectAll.Checked == true) { shpobj.CorrectFormat(sldNum); }
                    }
                }
                // - After review select a first slide 
                this.Close();               // Close the user forms 
                if (whichSlides == "All") { ActivePPT.Slides[1].Select(); }
                
            }
            catch(Exception err)
            {
                PPTAttribute.exitFlag = true;
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Review Format");
            }
            
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmReviewformat_FormClosed(object sender, FormClosedEventArgs e)
        {
            PPTAttribute.exitFlag = false;
        }
    }
}
