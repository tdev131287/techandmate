using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace TSCPPT_Addin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //MessageBox.Show("Call a Startup Method");
            

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        // 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new PPT_Ribbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            Globals.ThisAddIn.Application.AfterNewPresentation += Application_AfterNewPresentation;
            Globals.ThisAddIn.Application.SlideSelectionChanged += Application_SlideSelectionChanged;
            Globals.ThisAddIn.Application.PresentationNewSlide += Application_PresentationNewSlide;
            //Globals.ThisAddIn.Application.PresentationOpen += Application_PresentationOpen;
        }

        private void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            ////MessageBox.Show("Just justify when call this function");
            //PowerPoint.Presentation actPPT = Globals.ThisAddIn.Application.ActivePresentation;
            //int sld = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideNumber;
            //string layout = Convert.ToString(actPPT.Slides[sld].CustomLayout.Name);
            //MessageBox.Show("Call Slide Insert Event : " + layout);
            ////throw new NotImplementedException();
        }

        private void Application_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            //PowerPoint.Presentation actPPT = Globals.ThisAddIn.Application.ActivePresentation;
            //int sld = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideNumber;
            //string layout = Convert.ToString(actPPT.Slides[sld].CustomLayout.Name);
            //MessageBox.Show("Call Slide Insert Event : " + layout);
            ////throw new NotImplementedException();
        }

        private void Application_AfterNewPresentation(PowerPoint.Presentation Pres)
        {
            
            //throw new NotImplementedException();
        }

        private void Application_AfterPresentationOpen(PowerPoint.Presentation Pres)
        {
            //MessageBox.Show("Call a Startup Method");
            //throw new NotImplementedException();
        }



        #endregion
    }
}
