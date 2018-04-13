using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace TSCPPT_Addin
{
    public partial class frmEditorialReview : Form
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        EditorialReview EObj = new EditorialReview();
        public frmEditorialReview()
        {
            InitializeComponent();
        }

        private void btn_Review_Click(object sender, EventArgs e)
        {
            string whichSlides = cmb_RType.Text;
            int selectedSlides=0;
            bool globalErrors = false;
            List<int> selSlides = new List<int>();
            Algorithms algobj = new Algorithms();
            if (whichSlides == "Current" || whichSlides == "Selected")
            {
                selectedSlides = ppApp.ActiveWindow.Selection.SlideRange.Count;
                if(selectedSlides==0)
                {
                    MessageBox.Show("Please select a slide and try again.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                for(int i=1;i<= selectedSlides;i++)
                {
                    selSlides.Add(ppApp.ActiveWindow.Selection.SlideRange[i].SlideNumber);
                }
                
                selSlides.Sort();
                //Array.Sort(selSlides);
            }
            else if(whichSlides == "All")
            {
                selectedSlides = ActivePPT.Slides.Count;
                for (int i = 1; i <= selectedSlides; i++) { selSlides.Add(i); }
                //Array.Sort(selSlides);
            }
            if (selectedSlides == 0)
            {
                MessageBox.Show("Please select a slide and try again.", PPTAttribute.msgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                for (int i = 1; i <= selectedSlides; i++)
                {
                    int sldNum = selSlides[i - 1];
                    algobj.SetNamesUnique(sldNum);                      // Set the unique name of Each object in particular slide
                    ppApp.ActiveWindow.View.GotoSlide(sldNum);
                    ActivePPT.Slides[sldNum].Select();
                    if (rb_Review.Checked == true) { EObj.CheckEditorial(sldNum, "Method1"); }
                    else if (rb_ReviewCorrect.Checked == true) { EObj.CheckEditorial(sldNum, "Method2"); }
                    else if (rb_CorrectAll.Checked == true) { EObj.CorrectEditorial(sldNum); }
                }
            }
        }

        private void frmEditorialReview_Load(object sender, EventArgs e)
        {
            PPTAttribute.reviewExitFlag = false;
            cmb_RType.Items.Add("Selected");
            cmb_RType.Items.Add("All");
            cmb_RType.SelectedIndex = 0;
            rb_ReviewCorrect.Checked = true;
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
