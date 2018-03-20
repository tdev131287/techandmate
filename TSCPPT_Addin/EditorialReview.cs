using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Drawing;
using System.Text.RegularExpressions;

namespace TSCPPT_Addin
{
    class EditorialReview
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        List<string> globalErrorList = new List<string>();
        DataTable wList1Dic = new DataTable();
        DataTable wList2Dic = new DataTable();
        DataTable arrStrListToCheck = new DataTable();
        public DataTable getWrokList(string lstType)
        {

            string qText = null;
            wList1Dic.Clear();
            string cnText = @"Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";Data Source=" + PPTAttribute.WordList;
            if (lstType == "List1") { qText = "select * from [WordList_Style$A1:B10000];"; }
            else if (lstType == "List2") { qText = "select * from [WorkList_USUK$A1:B10000];"; }
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter(qText, cnText);
                if (lstType == "List1") { da.Fill(wList1Dic); };
                if (lstType == "List2") { da.Fill(wList2Dic); };
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "getWrokList");
            }
            return (wList1Dic);
        }

        public void setTableColumn()
        {
            if (arrStrListToCheck.Columns.Count == 0)
            {
                DataColumn dc1 = new DataColumn("SldNum", typeof(int));                         // -Slide Number
                arrStrListToCheck.Columns.Add(dc1);
                DataColumn dc2 = new DataColumn("ShapeName", typeof(String));                   // -Shape Name
                arrStrListToCheck.Columns.Add(dc2);
                DataColumn dc3 = new DataColumn("StrFound", typeof(String));                    // -Str Found Status
                arrStrListToCheck.Columns.Add(dc3);
                DataColumn dc4 = new DataColumn("StrPosition", typeof(int));                    // -Position          (Apply -Sorting on this columns)
                arrStrListToCheck.Columns.Add(dc4);
                DataColumn dc5 = new DataColumn("StrLenght", typeof(String));                   // -Str Length
                arrStrListToCheck.Columns.Add(dc5);

                DataColumn dc6 = new DataColumn("PatternType", typeof(String));                   // -Pattern Type
                arrStrListToCheck.Columns.Add(dc6);
                DataColumn dc7 = new DataColumn("tbRow", typeof(String));                       // -Table -Row Position
                arrStrListToCheck.Columns.Add(dc7);
                DataColumn dc8 = new DataColumn("tbCol", typeof(String));                       // -Table Columns Position
                arrStrListToCheck.Columns.Add(dc8);
            }
            
        }
        public void CorrectEditorial(int sldNum)
        {

        }
        public void CheckEditorial(int sldNum, string method)
        {

            string wlError = null;
            string PatternError = null;
            
            getWrokList("List1");                              //- Call this function to set the word from database-
            #region Only Click on Only Editorial Review --
            if (method == "Method1")
            {
                #region - Find a Error from word list mapping & Double space type error
                foreach (PowerPoint.Shape shp in ActivePPT.Slides[sldNum].Shapes)
                {
                    if (shp.Type == MsoShapeType.msoTable)
                    {
                        for (int r = 1; r <= shp.Table.Rows.Count; r++)
                        {
                            for (int c = 1; c <= shp.Table.Columns.Count; c++)
                            {
                                PowerPoint.TextRange txtRng = shp.Table.Cell(r, c).Shape.TextFrame.TextRange;
                                wlError = CheckEditorial_TextRange_M1_WL1(sldNum, shp, txtRng, wlError);
                                wlError = CheckEditorial_TextRange_M1_WL2(sldNum, shp, txtRng, wlError);
                                wlError = CheckEditorial_TextRange_M1_DS(sldNum, shp, txtRng, wlError);
                            }
                        }
                    }
                    else
                    {
                        PowerPoint.TextRange txtRng = shp.TextFrame.TextRange;
                        wlError = CheckEditorial_TextRange_M1_WL1(sldNum, shp, txtRng, wlError);
                        wlError = CheckEditorial_TextRange_M1_WL2(sldNum, shp, txtRng, wlError);
                        wlError = CheckEditorial_TextRange_M1_DS(sldNum, shp, txtRng, wlError);
                    }
                } // Close Shape Iteration 
                #endregion

                #region   -- Apply a logic for Patter Match
                foreach (PowerPoint.Shape shp in ActivePPT.Slides[sldNum].Shapes)
                {
                    if (shp.Type == MsoShapeType.msoTable)
                    {
                        for (int r = 1; r <= shp.Table.Rows.Count; r++)
                        {
                            for (int c = 1; c <= shp.Table.Columns.Count; c++)
                            {
                                PowerPoint.TextRange txtRng = shp.Table.Cell(r, c).Shape.TextFrame.TextRange;
                                wlError = wlError + '\n' + pPatternH12016ToH1_2016(sldNum, shp, txtRng, wlError, "method1");
                            }
                        }
                    }
                    else
                    {
                        PowerPoint.TextRange txtRng = shp.TextFrame.TextRange;
                        PatternError = PatternError + '\n' + pPatternH12016ToH1_2016(sldNum, shp, txtRng, PatternError, "method1");
                    }
                } // Close Shape Iteration 
                #endregion

                if (String.IsNullOrEmpty(wlError) == false)
                {
                    ActivePPT.Slides[sldNum].Comments.Add(0, 0, "Editorial Review:", "TER", wlError);
                }
            }
            #endregion

            #region Click on Reivew & Correct both of Editorial Error
            else if (method == "Method2" || method == "Method3")
            {
                foreach (PowerPoint.Shape shp in ActivePPT.Slides[sldNum].Shapes)
                {
                    if (shp.Type == MsoShapeType.msoTable)
                    {
                        for (int r = 1; r <= shp.Table.Rows.Count; r++)
                        {
                            for (int c = 1; c <= shp.Table.Columns.Count; c++)
                            {
                                PowerPoint.TextRange txtRng = shp.Table.Cell(r, c).Shape.TextFrame.TextRange;
                                pCreateArray_List(sldNum, shp.Name, txtRng, wlError, "Method2",r,c);
                                
                            }
                        }
                    }
                    else
                    {
                        PowerPoint.TextRange txtRng = shp.TextFrame.TextRange;
                        pCreateArray_List(sldNum, shp.Name, txtRng, wlError, "Method2");

                    }
                } // Close Shape Iteration 
            }
            #endregion
        }

        public string CheckEditorial_TextRange_M1_WL1(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError)
        {

            foreach (DataRow row in wList1Dic.Rows)
            {
                string inCorrect = row["Incorrect_Word"].ToString();
                string Correct = row["Correct_Word"].ToString();
                PowerPoint.TextRange rngFound = txtRng.Find(inCorrect);
                while (rngFound != null)
                {
                    int x = rngFound.Start;                     // Return start position where find the text 
                    string cWord = rngFound.Text;
                    wlError = wlError + "Change " + cWord + " to " + Correct + '\n';
                    rngFound = txtRng.Find(inCorrect, x, MsoTriState.msoFalse, MsoTriState.msoFalse);
                }
            }
            return wlError;
        }

        public string CheckEditorial_TextRange_M1_WL2(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError)
        {
            getWrokList("List2");

            for (int x = 1; x <= txtRng.Paragraphs().Count; x++)
            {
                int wCount = txtRng.Paragraphs(x).Words().Count;
                for (int y = 1; y <= wCount; y++)
                {
                    string cWord = txtRng.Paragraphs(x).Words(y).Text;
                    string expression = "US_Word = '" + cWord + "'";
                    DataRow[] foundRows = wList2Dic.Select(expression);
                    if (foundRows.Count() != 0)
                    {
                        string rWord = foundRows[0][1].ToString();
                        wlError = wlError + "Change " + cWord + " to " + rWord + '\n';
                    }
                }

            }
            return wlError;
        }
        public string CheckEditorial_TextRange_M1_DS(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError)
        {
            int x = 0;
            PowerPoint.TextRange rngFound = txtRng.Find("  ", 0, MsoTriState.msoFalse, MsoTriState.msoFalse);
            while (rngFound != null)
            {
                x = rngFound.Start;
                wlError = wlError + "Change Double Space to Single Space" + '\n';
                rngFound = txtRng.Find("  ", x, MsoTriState.msoFalse, MsoTriState.msoFalse);
            }
            return wlError;
        }

        #region  Pattern Related All Logics 
        public string pPatternH12016ToH1_2016(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type ,
            string strFound=null,int lngStartPos=0,int lngRow=0,int lngCol=0)
        {
            //strFound.Substring(0, 2) + " " + strFound.Substring(2, 4);
            
            string wlErrorGlobal = null;
            if (type == "Method1")
            {
                string pattern = @"[HQhq][1-4][1-2][0-9]{3}";
                wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "PatternH1");
            }
            else if(type=="Method2" || type == "Method3")
            {

                PowerPoint.TextRange rngFound = txtRng.Find(strFound, 0, MsoTriState.msoTrue, MsoTriState.msoFalse);

            }
            return wlErrorGlobal;
        }

        public string pPatternFY_2016ToFY2016(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type)
        {
           
            string wlErrorGlobal = null;
            string pattern = @"FY[ ][1-2][0-9]{3}";
            wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "PatternFY");
            return wlErrorGlobal;

        }
        
        public string pPatternAround00(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            string wlErrorGlobal = null;
            string pattern = @"Around[ ][0-9]{1,20}";
            string strReplace = "About 10 or 'Approximately 10' or '~10'";
            wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, strReplace);
            return wlErrorGlobal;
            
        }
        public string pPatternBewtweenYear(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            string wlErrorGlobal = null;
            string pattern = @"Between[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            string strReplace = "'Between 2005 and 2010' or '~From 2005 to 2010' or 'During 2005–2010' or 'Over 2005–2010'";
            wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, strReplace);
            return wlErrorGlobal;
        }
        public string pPatternEndYear(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            string wlErrorGlobal = null;
            string pattern = @"Between[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            string strReplace = "'Between 2005 and 2010' or '~From 2005 to 2010' or 'During 2005–2010' or 'Over 2005–2010'";
            wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, strReplace);
            return wlErrorGlobal;
        }
        public string pPatternFrom_YYYYToFYYYYY(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            string wlErrorGlobal = null;
            string pattern = @"From[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "PatternFrom_YYYY");
            return wlErrorGlobal;
        }
        public string pPatternTime_AM_PM(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            string wlErrorGlobal = null;
            string pattern = @"[0-9]{1,2}[:][0-9]{2}[ ][AP][M]";
            wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "PatternTime_AM_PM");
            return wlErrorGlobal;
        }
        public string pPattern_Measure_Units_1(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            string wlErrorGlobal = null;
            bool uUnload = false;
            string[] varstrNumber = { "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" };
            int[] varlngNumber = { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            string pattern = "(one|two|three|four|five|six|seven|eight|nine)[ ]?(mile|tonne|yard|gram|litre|gallon|pint|inch|inches|hectare|hour|minute|day|month|year|week|barrel|knot|feet|foot|metric ton|ton|km|kg|lbs|Btu|ha)[e$]?[s$]?";
            Match result = Regex.Match(txtRng.Text, pattern);
            MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
            if (result.Success)
            {
                foreach (Match match in matches)
                {
                    string strFound = match.Value;
                }
            }
                    return wlErrorGlobal;
        }
        public string pPattern_Measure_Units(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            return wlError;
        }
        public string pPattern_Currency_No_Space(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            return wlError;
        }
        public string pPattern_Currency_With_Space(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            return wlError;
        }
        public string pPatternSqFt_CuM_1(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            return wlError;
        }
        public string pPatternSqFt_CuM(int sldNum, string shpName, PowerPoint.TextRange txtRng, string wlError, string type)
        {
            return wlError;
        }

        public string OnlyReviewPatterError(PowerPoint.TextRange txtRng, string pattern,string strReplace=null)
        {
            string wlErrorGlobal = null;
            //string pattern = @"Between[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            Match result = Regex.Match(txtRng.Text, pattern);
            MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
            if (result.Success)
            {
                foreach (Match match in matches)
                {
                    string strFound = match.Value;
                    if (strReplace == "PatternFY") { strReplace = strFound.Replace(" ", ""); }
                    else if (strReplace == "PatternH1") { strReplace = strFound.Substring(0, 2) + " " + strFound.Substring(2, 4); }
                    else if(strReplace== "PatternFrom_YYYY") { strReplace = strFound.Replace("-", " to "); }
                    else if (strReplace == "PatternTime_AM_PM")
                    {
                        strReplace = strFound.Replace(":00", "");
                        strReplace = strFound.Replace(" ", "");
                    }
                    //else if (strReplace == "PatternFrom_YYYY") { strReplace = strFound.Replace("-", " to "); }
                    else
                        wlErrorGlobal = wlErrorGlobal + "Change " + strFound + " to " + strReplace + '\n';
                }

            }
            return wlErrorGlobal;
        }
        public void getErrorList(int sldnum,string shpname,PowerPoint.TextRange txtRng, string pattern, string patternType, int lngRow = 0,int lngCol = 0)
        {
            string wlErrorGlobal = null;
            //string pattern = @"Between[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            Match result = Regex.Match(txtRng.Text, pattern);
            MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
            if (result.Success)
            {
                foreach (Match match in matches)
                {
                    arrStrListToCheck.Rows.Add(sldnum, shpname, match.Value, match.Index, match.Length, patternType, lngRow, lngCol);
                }
            }
        }


        #endregion

                    #region Create a Array List as per error type -
        public void pCreateArray_List(int sldNum,string shpName, PowerPoint.TextRange txtRng,string wlError,string strMethod,int r=0 ,int c=0)
        {
            DataTable wList1Dic = getWrokList("List1");
            setTableColumn();
            string pattern;
            pattern = @"[HQhq][1-4][1-2][0-9]{3}";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternH12016ToH1_2016", r,c);
            pattern = @"Between[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternBewtweenYear", r, c);
            pattern = @"From[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternFrom_YYYYToFYYYYY", r, c);
            pattern = @"end[ ][o]?[f]?[ ]?[1-2][0-9]{3}";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternEndYear", r, c);
            pattern = @"[0-9]{1,2}[:][0-9]{2}[ ][AP][M]";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternTime_AM_PM", r, c);
            pattern = @"[0-9,.]{1,20}[ ]?(mile|tonne|yard|gram|litre|gallon|pint|inch|inches|hectare|hour|minute|day|month|year|week|barrel|knot|feet|foot|metric ton|ton|km|kg|lbs|Btu|ha)[e$]?[s$]?";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPattern_Measure_Units", r, c);
            pattern = @"(one|two|three|four|five|six|seven|eight|nine)[ ]?(mile|tonne|yard|gram|litre|gallon|pint|inch|inches|hectare|hour|minute|day|month|year|week|barrel|knot|feet|foot|metric ton|ton|km|kg|lbs|Btu|ha)[e$]?[s$]?";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPattern_Measure_Units_1", r, c);
            //pPatternH12016ToH1_2016_array(sldNum,shpName,txtRng,r,c);

            arrStrListToCheck.DefaultView.Sort = "StrPosition";
            for(int rowCount=0; rowCount< arrStrListToCheck.Rows.Count; rowCount++)
            {
                if(arrStrListToCheck.Rows[rowCount]["PatternType"].ToString()== "pPatternH12016ToH1_2016")
                {
                    //wlError = pPatternH12016ToH1_2016(arrStrListToCheck(0, lngLoop), arrStrListToCheck(1, lngLoop), txtRng, wList1Dic, wlError, "method2", arrStrListToCheck(2, lngLoop), arrStrListToCheck(3, lngLoop), arrStrListToCheck(6, lngLoop), arrStrListToCheck(7, lngLoop))
                    int sldNo = Convert.ToInt32(arrStrListToCheck.Rows[rowCount]["SldNum"]);
                    string shpname = Convert.ToString(arrStrListToCheck.Rows[rowCount]["ShapeName"]);
                    PowerPoint.Shape shp= ActivePPT.Slides[sldNo].Shapes[shpname];
                    PowerPoint.TextRange ErrtxtRng = shp.TextFrame.TextRange;
                    wlError = pPatternH12016ToH1_2016(sldNo, shp, ErrtxtRng, wlError, strMethod);

                }
            }

        }
       
        #endregion

        public void SelectWord(int sldNum, string shpName, PowerPoint.TextRange txtRng,string cWord,int stPos,int lngRow=0,int lngCol=0)
        {
            PowerPoint.Shape shp = ActivePPT.Slides[sldNum].Shapes[shpName];
            shp.Select();
            if (shp.Type==MsoShapeType.msoTable)
            {
                PowerPoint.TextRange tRng= ppApp.ActiveWindow.Selection.ShapeRange.Table.Cell(lngRow, lngCol).Shape.TextFrame.TextRange.Characters(stPos, 1);
                var colorCode = tRng.Font.Color;
                float fSize =  tRng.Font.Size;
                var fBold = tRng.Font.Bold;
                tRng.Font.Shadow = MsoTriState.msoTrue;
                tRng.Font.Color.RGB= Color.FromArgb(0, 255, 0).ToArgb();
                tRng.Font.Bold = MsoTriState.msoTrue;
                tRng.Font.Size = 16;
            }
            else
            {
                PowerPoint.TextRange tRng = ppApp.ActiveWindow.Selection.TextRange.Characters(stPos, 1);
                tRng.Font.Shadow = MsoTriState.msoTrue;
                tRng.Font.Color.RGB = Color.FromArgb(0, 255, 0).ToArgb();
                tRng.Font.Bold = MsoTriState.msoTrue;
                tRng.Font.Size = 16;
            }
        }
        public void DeSelectWord(int sldnum,string shpname, PowerPoint.TextRange txtRng, string cWord, int stPos)
        {

        }


    }
}
