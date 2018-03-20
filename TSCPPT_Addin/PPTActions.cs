using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PPTAttribute = Microsoft.Office.Core;
using Microsoft.Office.Core;

namespace TSCPPT_Addin
{
    class PPTActions
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        public void insert_Slide_in_ActivePPT(int CslideIndex, int PslideIndex, string slideType = null,int cnt=0)
        {
            PowerPoint.Presentation actPPT=null, CurrentPPT;
            try
            {
                CurrentPPT = ppApp.ActivePresentation;
                actPPT = ppApp.Presentations.Open(PPTAttribute.standardppt, MsoTriState.msoFalse);
                actPPT.Slides[CslideIndex].Copy();
                CurrentPPT.Slides.Paste(PslideIndex);
                if (slideType == "CSlide") { CurrentPPT.Slides[PslideIndex].Name = "Title Slide" + cnt + 1; }
                else if (slideType == "ESlide") { CurrentPPT.Slides[PslideIndex].Name = "End Page" + cnt + 1; }
                actPPT.Close();
                CurrentPPT.Slides[PslideIndex].Select();
            }
            catch(Exception ex)
            {
                string errtext = ex.Message;
                PPTAttribute.ErrorLog(errtext, "insert_Slide_in_ActivePPT");
                actPPT.Close();
            }
        }

        public int get_LastSelectedSlide()
        {
            int lSlide=0, temp, selected_Slides_Count;
            try
            {
                try { selected_Slides_Count = ppApp.ActiveWindow.Selection.SlideRange.Count; }
                catch (Exception ex) { selected_Slides_Count = 0; }
                lSlide = 0;
                for (int sld = 1; sld <= selected_Slides_Count; sld++)
                {
                    temp = ppApp.ActiveWindow.Selection.SlideRange[sld].SlideNumber;
                    if (temp > lSlide) { lSlide = temp; }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "get_LastSelectedSlide");
            }
            return lSlide;
        }

        public DataTable get_specification(string cri)
        {
            DataTable dt = new DataTable();
            //string ExSpacificationPath = @"C:\Users\Devendra.Tripathi\Documents\visual studio 2015\Projects\TSCPPT_Addin\TSCPPT_Addin\AppData\Mapping\PPT_Specification.xlsx";
            string cnText = @"Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";Data Source=" + PPTAttribute.dbPath;
            string qText = "select * from [Specification$A1:AU25] where Name='"+cri +"'";
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter(qText, cnText);
                da.Fill(dt);
               
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "get_specification");
            }
            return (dt);
        }
        public DataTable get_SlideIndex(string cri)
        {
            DataTable dt = new DataTable();
            //string ExSpacificationPath = @"C:\Users\Devendra.Tripathi\Documents\visual studio 2015\Projects\TSCPPT_Addin\TSCPPT_Addin\AppData\Mapping\PPT_Specification.xlsx";
            string cnText = @"Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";Data Source=" + PPTAttribute.dbPath;
            string qText = "select * from [Slide$A1:B11] where SlideName='" + cri + "'";
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter(qText, cnText);
                da.Fill(dt);

            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "get_SlideIndex");
            }
            return (dt);
        }
        public DataTable get_ChartSpacification(string cri)
        {
            DataTable dt = new DataTable();
            //string ExSpacificationPath = @"C:\Users\Devendra.Tripathi\Documents\visual studio 2015\Projects\TSCPPT_Addin\TSCPPT_Addin\AppData\Mapping\PPT_Specification.xlsx";
            string cnText = @"Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";Data Source=" + PPTAttribute.dbPath;
            string qText = "select * from [Chart$A1:C11] where ChartType='" + cri + "'";
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter(qText, cnText);
                da.Fill(dt);

            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "get_ChartSpacification");
            }
            return (dt);
        }

        public DataTable get_ChatColorCode(string colname,int srCount)
        {
            DataTable dt = new DataTable();
            string qText;
            string cnText = @"Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";Data Source=" + PPTAttribute.dbPath;
            if (srCount <= 5) { qText = "select " + colname + " from [ChartColor$A1:E11]"; }
            else { qText = "select " + colname + " from [ChartColor$J1:N11]"; }
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter(qText, cnText);
                da.Fill(dt);

            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "get_ChatColorCode");
            }
            return (dt);
        }

        public string InsertPlaceholder(int sldNum,DataTable dt,string shpType)
        {
            PowerPoint.Shape oPlaceholder;
            string shpName = null;
            List<string> txtStr = new List<string>();
            int shpCount = 0, shpNumber=0;
            try
            {
                //char splitChar = '|';
                float shpLeft = (float)Convert.ToDouble((dt.Rows[0]["ShapeLeft"]));
                float shpTop = (float)Convert.ToDouble((dt.Rows[0]["ShapeTop"]));
                float shpWidth = (float)Convert.ToDouble((dt.Rows[0]["ShapeWidth"]));
                float shpHeight = (float)Convert.ToDouble((dt.Rows[0]["ShapeHeight"]));
                string dText = Convert.ToString(dt.Rows[0]["DefaultText"]);
                //ActivePresentation.Slides(sldNum).Shapes.AddTextbox(msoTextOrientationHorizontal, shpLeft, shpTop, shpWidth, shpHeight)

                oPlaceholder = ActivePPT.Slides[sldNum].Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, shpLeft, shpTop, shpWidth, shpHeight);

                oPlaceholder.TextFrame.TextRange.Text = dText;

                shpName = oPlaceholder.Name;
                shpCount = ActivePPT.Slides[sldNum].Shapes.Count;
                if (shpType != "Text Box")
                {
                    for (int shpIndex = 1; shpIndex <= shpCount; shpIndex++)
                    {
                        if (ActivePPT.Slides[sldNum].Shapes[shpIndex].Name == shpName) { shpNumber = shpIndex; break; }
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "InsertPlaceholder");
            }
            return (shpName);
        }
    }
}
