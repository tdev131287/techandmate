﻿using System;
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
using System.IO;

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
        PowerPoint.ColorFormat colorCode ;
        float fSize ;
        MsoTriState fBold ;

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
            getWrokList("List2");
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
                    if (shp.HasTextFrame == MsoTriState.msoTrue)
                    {
                        
                            for (int r = 1; r <= shp.Table.Rows.Count; r++)
                            {
                                for (int c = 1; c <= shp.Table.Columns.Count; c++)
                                {
                                    PowerPoint.TextRange txtRng = shp.Table.Cell(r, c).Shape.TextFrame.TextRange;
                                    wlError = wlError + '\n' + pPatternH12016ToH1_2016(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPatternFY_2016ToFY2016(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPatternAround00(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPatternBewtweenYear(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPatternEndYear(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPatternFrom_YYYYToFYYYYY(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPatternTime_AM_PM(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPattern_Measure_Units_1(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPattern_Measure_Units(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPattern_Currency_No_Space(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPattern_Currency_With_Space(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPatternSqFt_CuM_1(sldNum, shp, txtRng, wlError, "Method1");
                                    wlError = wlError + '\n' + pPatternSqFt_CuM(sldNum, shp, txtRng, wlError, "Method1");
                                }
                            }
                        }
                        else
                        {
                        if (shp.Type == MsoShapeType.msoTable)
                        {
                            PowerPoint.TextRange txtRng = shp.TextFrame.TextRange;
                            wlError = wlError + '\n' + pPatternH12016ToH1_2016(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPatternFY_2016ToFY2016(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPatternAround00(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPatternBewtweenYear(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPatternEndYear(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPatternFrom_YYYYToFYYYYY(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPatternTime_AM_PM(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPattern_Measure_Units_1(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPattern_Measure_Units(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPattern_Currency_No_Space(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPattern_Currency_With_Space(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPatternSqFt_CuM_1(sldNum, shp, txtRng, wlError, "Method1");
                            wlError = wlError + '\n' + pPatternSqFt_CuM(sldNum, shp, txtRng, wlError, "Method1");
                        }
                    } // Close if check it - shape has 
                } // Close Shape Iteration 
                #endregion
                char split = '\n';
                List<string> AllError = new List<string>();
                List<string> uniqueError = new List<string>();
                string finalError = null;
                AllError = wlError.Split(split).ToList();
                foreach(string err in AllError)
                {
                   
                    if (uniqueError.Contains(err)!=true)
                    {
                        if (String.IsNullOrWhiteSpace(err) != true)
                        {
                            finalError = finalError + '\n' + err;
                            uniqueError.Add(err);
                        }
                    }
                }
                if (String.IsNullOrEmpty(finalError) == false)
                {
                    ActivePPT.Slides[sldNum].Comments.Add(0, 0, "Editorial Review:", "TER", finalError);
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
                                    pCreateArray_List(sldNum, shp.Name, txtRng, wlError, "Method2", r, c);

                                }
                            }
                        }
                        else
                        {
                        if (shp.HasTextFrame == MsoTriState.msoTrue)
                        {
                            PowerPoint.TextRange txtRng = shp.TextFrame.TextRange;
                            pCreateArray_List(sldNum, shp.Name, txtRng, wlError, "Method2");

                        }
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
                PowerPoint.TextRange rngFound = txtRng.Find(inCorrect,0,MsoTriState.msoFalse,MsoTriState.msoTrue);
                while (rngFound != null )
                {
                    int x = rngFound.Start;                     // Return start position where find the text 
                    string cWord = rngFound.Text;
                    if (cWord != " ")
                    {
                        wlError = wlError + "Change " + cWord + " to " + Correct + '\n';
                    }
                    rngFound = txtRng.Find(inCorrect, x, MsoTriState.msoFalse, MsoTriState.msoTrue);
                }
                
            }
            return wlError;
        }

        public string CheckEditorial_TextRange_M1_WL2(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError)
        {
            getWrokList("List2");
            string expression = null;
            for (int x = 1; x <= txtRng.Paragraphs().Count; x++)
            {
                int wCount = txtRng.Paragraphs(x).Words().Count;
                for (int y = 1; y <= wCount; y++)
                {
                    string cWord = txtRng.Paragraphs(x).Words(y).Text;
                    MsoLanguageID langID = MsoLanguageID.msoLanguageIDEnglishUS, curr_lang;
                    curr_lang = ActivePPT.DefaultLanguageID;
                    if (curr_lang == MsoLanguageID.msoLanguageIDEnglishUS) { expression = "US_Word = '" + cWord + "'"; }
                    else if(curr_lang == MsoLanguageID.msoLanguageIDEnglishUK){ expression = "UK_Word = '" + cWord + "'"; }


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
            PowerPoint.TextRange rngFound = txtRng.Find("  ", 0, MsoTriState.msoFalse, MsoTriState.msoTrue);
            while (rngFound != null)
            {
                x = rngFound.Start;
                wlError = wlError + "Change Double Space to Single Space" + '\n';
                rngFound = txtRng.Find("  ", x, MsoTriState.msoFalse, MsoTriState.msoTrue);
            }
            return wlError;
        }

        #region  Pattern Related All Logics 
        //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
        public string pPatternH12016ToH1_2016(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type ,
            string strFound=null,int lngRow=0,int lngCol=0)
        {
            //strFound.Substring(0, 2) + " " + strFound.Substring(2, 4);
            frmReplacewith frmObj = new frmReplacewith();
            List<string> varList = new List<string>();
            string wlErrorGlobal = null;
            if (type == "Method1")
            {
                string pattern = @"[HQhq][1-4][1-2][0-9]{3}";
                wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case1");
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }
            return wlErrorGlobal;
        }

        //Array of years till 2099 and preceded by FY with space  (FY 2016 with FY2016)
        //int lngStartPos = 0,
        public string pPatternFY_2016ToFY2016(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
           
            string wlErrorGlobal = null;
            if (type == "Method1")
            {
                string pattern = @"FY[ ][1-2][0-9]{3}";

                wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case2");
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }
            return wlErrorGlobal;

        }
        //Around followed by any whole number with or without space  (Around 10 with About 10)
        public string pPatternAround00(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            if (type == "Method1")
            {
                string pattern = @"Around[ ][0-9]{1,20}";
                string strReplace = "About 10 or 'Approximately 10' or '~10'";
                wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case3");
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError,lngRow, lngCol);
            }
            return wlErrorGlobal;
            
        }

        public string pPatternBewtweenYear(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            if (type == "Method1")
            {
                string pattern = @"Between[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
                string strReplace = "'Between 2005 and 2010' or '~From 2005 to 2010' or 'During 2005–2010' or 'Over 2005–2010'";
                wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case4");
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }
            return wlErrorGlobal;
        }

        public string pPatternEndYear(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            string strReplace;
            if (type == "Method1")
            {
                string pattern = @"end[ ][o]?[f]?[ ]?[1-2][0-9]{3}";
                wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case5");
               

            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }
            return wlErrorGlobal;
        }

        public string pPatternFrom_YYYYToFYYYYY(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            if (type == "Method1")
            {
                string pattern = @"From[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
                wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case6");
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }
            return wlErrorGlobal;
        }

        public string pPatternTime_AM_PM(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null, int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            if (type == "Method1")
            {
                string pattern = @"[0-9]{1,2}[:][0-9]{2}[ ][AP][M]";
                wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case7");
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }
            return wlErrorGlobal;
        }

        public string pPattern_Measure_Units_1(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            bool uUnload = false;
            string strReplace = null;
            string[] varstrNumber = { "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" };
            int[] varlngNumber = { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            string pattern = "(one|two|three|four|five|six|seven|eight|nine)[ ]?(mile|tonne|yard|gram|litre|gallon|pint|inch|inches|hectare|hour|minute|day|month|year|week|barrel|knot|feet|foot|metric ton|ton|km|kg|lbs|Btu|ha)[e$]?[s$]?";
            if (type == "Method1")
            {
                Match result = Regex.Match(txtRng.Text, pattern);
                MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
                if (result.Success)
                {
                    wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case8");
                    
                }
                //wlErrorGlobal = wlErrorGlobal + "Change " + strFound + " to " + strReplace + '\n';
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }

            return wlErrorGlobal;
        }

        public string pPattern_Measure_Units(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {

            string wlErrorGlobal = null;
            bool uUnload = false;
            string strReplace = null;
            string pattern = @"[0-9,.]{1,20}[ ]?(mile|tonne|yard|gram|litre|gallon|pint|inch|inches|hectare|hour|minute|day|month|year|week|barrel|knot|feet|foot|metric ton|ton|km|kg|lbs|Btu|ha)[e$]?[s$]?";
            if (type == "Method1")
            {
                Match result = Regex.Match(txtRng.Text, pattern);
                MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
                if (result.Success)
                {
                    wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case9");
                    
                }
                //wlErrorGlobal = wlErrorGlobal + "Change " + strFound + " to " + strReplace + '\n';
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }

            return wlErrorGlobal;
        }

        public string pPattern_Currency_No_Space(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            bool uUnload = false;
            string strReplace = null;
            string pattern = @"[$€£][ ][0-9]{1,10}";
            if (type == "Method1")
            {
                Match result = Regex.Match(txtRng.Text, pattern);
                MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
                if (result.Success)
                {
                    wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case10");
                    //strReplace = strFound.Replace(" ", "");
                }
                //wlErrorGlobal = wlErrorGlobal + "Change " + strFound + " to " + strReplace + '\n';
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }

            return wlErrorGlobal;
        }

        public string pPattern_Currency_With_Space(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            bool uUnload = false;
            string pattern = @"[$€£][ ][0-9]{1,10}";
            if (type == "Method1")
            {
                Match result = Regex.Match(txtRng.Text, pattern);
                MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
                if (result.Success)
                {
                    wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case11");
                }
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }

            return wlErrorGlobal;
        }

        public string pPatternSqFt_CuM_1(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
           string wlErrorGlobal = null;
            bool uUnload = false;
            string pattern = @"(one|two|three|four|five|six|seven|eight|nine)[ ]?(sq|cu)[.]?[ ]?(ft|m)[.]?";
            if (type == "Method1")
            {
                Match result = Regex.Match(txtRng.Text, pattern);
                MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
                if (result.Success)
                {
                    wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case12");
                }
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }

            return wlErrorGlobal;
        }
        
        public string pPatternSqFt_CuM(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            bool uUnload = false;
            string pattern = @"[0-9.,]{1,20}[ ]?(sq|cu)[.]?[ ]?(ft|m)[.]?";
            if (type == "Method1")
            {
                Match result = Regex.Match(txtRng.Text, pattern);
                MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
                if (result.Success)
                {
                    wlErrorGlobal = OnlyReviewPatterError(txtRng, pattern, "Case13");
                }
            }
            else if (type == "Method2" || type == "Method3")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }

            return wlErrorGlobal;
        }

        public string CheckEditorialwList1Dic(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            
            string wlErrorGlobal = null;
            showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            return wlErrorGlobal;
        }
        public string CheckEditorialwList2Dic(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            Algorithms obj = new Algorithms();
            string docObj = obj.DocumentLanguage();
            if (docObj != "Mixed (US/UK)")
            {
                showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            }
            return wlErrorGlobal;
        }
        public string CheckEditorialwDic(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string wlError, string type,
            string strFound = null,  int lngRow = 0, int lngCol = 0)
        {
            string wlErrorGlobal = null;
            //showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);  // It's  not wroking
            return wlErrorGlobal;
        }


        public string OnlyReviewPatterError(PowerPoint.TextRange txtRng, string pattern,string CType)
        {
            string wlErrorGlobal = null;
            string strReplace = null;
            string strFound = null;
            //string pattern = @"Between[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            Match result = Regex.Match(txtRng.Text, pattern);
            MatchCollection matches = Regex.Matches(txtRng.Text, pattern);
            if (result.Success)
            {
                foreach (Match match in matches)
                {
                    strFound = match.Value;
                    if (CType == "Case1") { strReplace = strFound.Substring(0, 2) + " " + strFound.Substring(2, 4); }
                    else if (CType == "Case2") { strReplace = strFound.Replace(" ", ""); }
                    else if (CType == "Case3") { strReplace = "About 10 or 'Approximately 10' or '~10'"; }
                    else if (CType == "Case4") { strReplace = "Between 2005 and 2010' or '~From 2005 to 2010' or 'During 2005–2010' or 'Over 2005–2010"; }

                    else if (CType == "Case5")
                    {
                        if (strFound.IndexOf("of") > 0)
                        {
                            strReplace = strFound.Replace(" of ", "-");
                        }
                        else
                        {
                            strReplace = strFound.Replace(" ", "-");
                        }
                    }
                    else if (CType == "Case6") { strReplace = strFound.Replace("–", " to "); }
                    else if (CType == "Case7")
                    {
                        strReplace = strFound.Replace(":00", "");
                        strReplace = strReplace.Replace(" ", "");
                    }
                    else if (CType == "Case8")
                    {
                        string[] array = { "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" };
                        foreach (string val in array)
                        {
                            //int xx = strFound.IndexOf(val);
                            if (strFound.Contains(val))
                            {
                                string getUnit = (strFound.Replace(val, "")).Replace(".", "");
                                string unitVal = strFound.Replace(getUnit, "");
                                strReplace = unitVal + " " + getUnit;
                            }
                        }
                    }
                    else if (CType == "Case9")
                    {
                        int intlen = 0;
                        int txtLenght = strFound.Length;
                        for (int x = 0; x < txtLenght; x++)
                        {
                            string value = strFound.Substring(x, 1);
                            byte[] asciiBytes = Encoding.ASCII.GetBytes(value);
                            int xx = asciiBytes[0];
                            if (xx >= 49 && xx <= 57)
                            {
                                intlen = intlen + 1;
                            }
                            else { break; }
                        }
                        strReplace = strFound.Substring(0, intlen) + " " + strFound.Substring(intlen, strFound.Length - intlen);
                    }
                    else if (CType == "Case10")
                    {
                        strReplace = strFound.Replace(" ", "");
                    }
                    else if (CType == "Case11")
                    {
                        strReplace = strFound.Substring(0, 3) + " " + strFound.Substring(3, strFound.Length - 3);
                    }
                    else if (CType == "Case12")
                    {
                        string[] array = { "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" };
                        foreach (string val in array)
                        {
                            //int xx =strFound.Contains(val);
                            if (strFound.Contains(val))
                            {
                                string getUnit = (strFound.Replace(val, ""));
                                string unitVal = strFound.Replace(getUnit, "");
                                strReplace = unitVal + " " + getUnit;
                                strReplace = strReplace.Replace(".", "");
                            }
                        }
                    }
                    else if (CType == "Case13")
                    {
                        int intlen = 0;
                        int txtLenght = strFound.Length;
                        for (int x = 0; x < txtLenght; x++)
                        {
                            string value = strFound.Substring(x, 1);
                            byte[] asciiBytes = Encoding.ASCII.GetBytes(value);
                            int xx = asciiBytes[0];
                            if (xx >= 49 && xx <= 57)
                            {
                                intlen = intlen + 1;
                            }
                            else { break; }
                        }
                        strReplace = strFound.Substring(0, intlen) + " " + strFound.Substring(intlen, strFound.Length - intlen);
                    }
                    //else if (strReplace == "PatternFrom_YYYY") { strReplace = strFound.Replace("-", " to "); }
                    wlErrorGlobal = wlErrorGlobal + "Change " + strFound + " to " + strReplace + '\n';

                }

            }
            return wlErrorGlobal;
        }
        public void getErrorList(int sldnum,string shpname,PowerPoint.TextRange txtRng, string pattern, string patternType, int lngRow = 0,int lngCol = 0)
        {
            //string wlErrorGlobal = null;
            //string pattern = @"Between[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            string inCorrect=null, Correct=null;
            if (pattern == null)
            {
                DataTable finalDic = new DataTable();
                Algorithms obj = new Algorithms();
                string docObj = obj.DocumentLanguage();
                if (patternType == "M2_WL1") { finalDic = wList1Dic; }
                else if(patternType == "M2_WL2")
                {
                    if(docObj== "Mixed (US/UK)") { return; }
                    finalDic = wList2Dic;
                }
                foreach (DataRow row in finalDic.Rows)
                {
                    if (patternType == "M2_WL1")
                    {
                        inCorrect = row["Incorrect_Word"].ToString();
                        Correct = row["Correct_Word"].ToString();
                    }
                    else
                    {
                        if (docObj == "US English")
                        {
                            inCorrect = row["UK_Word"].ToString();
                            Correct = row["US_Word"].ToString();
                        }
                        else if (docObj == "UK English")
                        {
                            inCorrect = row["US_Word"].ToString();
                            Correct = row["UK_Word"].ToString();
                        }
                    }
                    PowerPoint.TextRange rngFound = txtRng.Find(inCorrect,0,MsoTriState.msoFalse,MsoTriState.msoTrue);
                    while (rngFound != null)
                    {
                        int x = rngFound.Start;                     // Return start position where find the text 
                        string cWord = rngFound.Text;
                        if (cWord != " ")
                        {
                            arrStrListToCheck.Rows.Add(sldnum, shpname, rngFound.Text, x, rngFound.Length, patternType, lngRow, lngCol);
                            return;
                        }
                        
                    }

                }
            }
            else
            {
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
        }

        public void showErrorForms(int sldNum, PowerPoint.Shape shp, PowerPoint.TextRange txtRng, string strFound, string logic, int lngRow = 0, int lngCol = 0)
        {
            int lngStartPos;
            string strReplace = null;
            frmReplacewith frmObj = new frmReplacewith();
            List<string> varList = new List<string>();
            PowerPoint.TextRange rngFound = txtRng.Find(strFound, 0, MsoTriState.msoTrue, MsoTriState.msoFalse);
            try
            {
                lngStartPos = rngFound.Start;
            }
            catch  { lngStartPos = 8; }
            
                SelectWord(sldNum, shp.Name, txtRng, strFound, lngStartPos, lngRow, lngCol);
            if (logic == "L1")                                                          //pPatternH12016ToH1_2016_array
            {
                strReplace = strFound.Substring(0, 2) + " " + strFound.Substring(strFound.Length - 4, 4);
            }
            else if (logic == "L2")                                                     //pPatternBewtweenYear_array
            {
                strReplace = strFound.Replace(" ", "");
            }
            else if (logic == "L3")                                                     //pPatternFY_2016ToFY2016_Array
            {
                strReplace = "About Number,Approximately Number,~Number";
            }
            else if (logic == "L4")                                                     //pPatternAround00_Array
            {
                strReplace = "Between YYYY and YYYY,From YYYY to YYYY,During YYYY–YYYY,Over YYYY–YYYY";
            }
            else if (logic == "L5")                                                     //pPatternFrom_YYYYToFYYYYY_array
            {
                //strReplace = strFound.Replace("-", " to ").Replace("–", " to ").Replace("—", " to ");
                if (strFound.IndexOf("of") > 0)
                {
                    strReplace = strFound.Replace(" of ", "-");
                }
                else
                {
                    strReplace = strFound.Replace(" ", "-");
                }

            }
            else if (logic == "L6")                                                     //pPatternEndYear_array
            {
                //strReplace = strFound.Replace(" of ", "-");
                //strReplace = strReplace.Replace(" ", "-");
                strReplace = strFound.Replace("-", " to ").Replace("–", " to ").Replace("—", " to ");
            }
            else if (logic == "L7")                                                     //pPatternTime_AM_PM_array
            {
                strReplace = strFound.Replace(":00", "");
                strReplace = strReplace.Replace(" ", "");
            }
            //-Work in progress -
            //showErrorForms(sldNum, shp, txtRng, strFound, wlError, lngRow, lngCol);
            else if (logic == "L8")                                                     //pPattern_Measure_Units_1_array
            {
                string[] array = { "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" };
                foreach (string val in array)
                {
                    //int xx = strFound.IndexOf(val);
                    if (strFound.Contains(val))
                    {
                        string getUnit = (strFound.Replace(val, "")).Replace(".", "");
                        string unitVal = strFound.Replace(getUnit, "");
                        strReplace = unitVal + " " + getUnit;

                    }
                }

            }
            else if (logic == "L9")                                                     //pPattern_Measure_Units_1_array
            {
                int intlen = 0;
                int txtLenght = strFound.Length;
                for (int x = 0; x < txtLenght; x++)
                {
                    string value = strFound.Substring(x, 1);
                    byte[] asciiBytes = Encoding.ASCII.GetBytes(value);
                    int xx = asciiBytes[0];
                    if (xx >= 48 && xx <= 57)
                    {
                        intlen = intlen + 1;
                    }
                    else { break; }
                }
                string digitVal = strFound.Substring(0, intlen);
                string measureUnit = strFound.Substring(intlen, strFound.Length - intlen);
                string strStringForS = " km kg lbs Btu ha";
                //strReplace = strFound.Substring(0, intlen) + " " + strFound.Substring(intlen, strFound.Length - intlen);
                if (intlen <= 0) { return; }
                if(Convert.ToString(intlen).IndexOf('.')>0)
                {
                    digitVal = String.Format(digitVal, "#,##0.##############");
                }
                else
                {
                    digitVal = String.Format(digitVal, "#,##0");
                }
                //If Not InStr(strStringForS, strMeasure) > 0 Then 
                if (strStringForS.IndexOf(measureUnit) == -1)
                {
                    if (Convert.ToInt32(digitVal) > 1)
                    {
                        if (measureUnit == "foot") { strReplace = digitVal + " " + "feet"; }
                        if (measureUnit == "feet") { strReplace = digitVal + " " + "feet"; }
                        if (measureUnit.Substring(measureUnit.Length - 2, 2) == "ch" || measureUnit.Substring(measureUnit.Length - 2, 2) == "sh")
                        {
                            strReplace = digitVal + " " + measureUnit + "es";
                        }
                        else if (measureUnit.Substring(measureUnit.Length - 1, 1) != "s")
                        {
                            strReplace = digitVal + " " + measureUnit + "s";
                        }
                        else
                        {
                            strReplace = digitVal + " " + measureUnit;
                        }
                    }
                    else
                    {
                        if (measureUnit == "feet") { strReplace = digitVal + " " + "foot"; }
                        if (measureUnit == "foot") { strReplace = digitVal + " " + "foot"; }
                        else
                        {
                            if (measureUnit.Substring(measureUnit.Length - 1, 1) == "s")
                            {
                                measureUnit = measureUnit.Substring(0, measureUnit.Length - 1);
                            }
                            strReplace = digitVal + " " + measureUnit;
                        }

                    }
                }
                else
                {
                    strReplace = digitVal + " " + measureUnit;
                }
            }
            else if (logic == "L10")                                                    //pPattern_Currency_No_Space_array
            {
                strReplace = strFound.Replace(" ", "");
            }
            else if (logic == "L11")                                                    //pPattern_Currency_With_Space_array
            {
                //strReplace = Left(strFound, 3) & " " & Mid(strFound, 4, Len(strFound))
                strReplace = strFound.Substring(0, 3) + " " + strFound.Substring(3, strFound.Length - 3);
            }
            else if (logic == "L12")                                                    //pPatternSqFt_CuM_1_array
            {
                // "onesq. ft." Solve this case-
                string[] array = { "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" };
                foreach (string val in array)
                {
                    //int xx =strFound.Contains(val);
                    if (strFound.Contains(val))
                    {
                        string getUnit = (strFound.Replace(val, ""));
                        string unitVal = strFound.Replace(getUnit, "");
                        strReplace = unitVal + " " + getUnit;
                        strReplace = strReplace.Replace(".", "");
                    }
                }
                //strReplace = strFound.Substring(0, 3) + " " + strFound.Substring(3, strFound.Length - 3);
            }
            else if (logic == "L13")                                                    //pPatternSqFt_CuM_array
            {
                int intlen = 0;
                int txtLenght = strFound.Length;
                for (int x = 0; x < txtLenght; x++)
                {
                    string value = strFound.Substring(x, 1);
                    byte[] asciiBytes = Encoding.ASCII.GetBytes(value);
                    int xx = asciiBytes[0];
                    if (xx >= 49 && xx <= 57)
                    {
                        intlen = intlen + 1;
                    }
                    else { break; }
                }
                strReplace = strFound.Substring(0, intlen) + " " + strFound.Substring(intlen, strFound.Length - intlen);
                strReplace = strReplace.Replace(".", "");
            }
            else if (logic == "L14")                                                    //pPattern_Preposition_array
            {

            }
            else if (logic == "L15" || logic == "L16")                                                    //CheckEditorial_TextRange_M2_WL1_Array
            {
                int x = 0;
                DataRow[] result=null;
                DataTable dtlist = new DataTable();
                string Correct = null;
                if (logic == "L15") { dtlist = wList1Dic; }
                else if(logic == "L16") { dtlist = wList2Dic; }
                List<string> replaceList = new List<string>();
                Algorithms obj = new Algorithms();
                string docObj = obj.DocumentLanguage();
                if (logic == "L15")
                {
                    result = dtlist.Select("Incorrect_Word='" + strFound + "'");
                }
                else
                {
                    if (docObj == "US English")
                    {
                        result = dtlist.Select("UK_Word='" + strFound + "'");
                    }
                    else if(docObj == "UK English")
                    {
                        result = dtlist.Select("US_Word='" + strFound + "'");
                    }
                }
                foreach (DataRow row in result)
                {
                    //string inCorrect = row["Incorrect_Word"].ToString();
                    if (logic == "L15")
                    {
                        Correct = row["Correct_Word"].ToString();
                    }
                    else
                    {
                        if (docObj == "US English")
                        {
                            Correct = row["US_Word"].ToString();
                            //result = dtlist.Select("UK_Word='" + strFound + "'");
                        }
                        else if (docObj == "UK English")
                        {
                            Correct = row["UK_Word"].ToString();
                            //result = dtlist.Select("US_Word='" + strFound + "'");
                        }
                    }
                    if (replaceList.Contains(Correct) == false) { replaceList.Add(Correct); }
                    
                } 
                foreach(string item in replaceList)
                {
                    if (x == 0) { strReplace = item; }
                    else { strReplace = strReplace + "," + item; }
                    x++;
                }
            }
            else if (logic == "L16")                                                    //CheckEditorial_TextRange_M2_WL2_Array
            {

            }
            else if (logic == "L17")                                                    //CheckEditorial_TextRange_M2_DS_Array
            {

            }
            //varList.Add(strReplace);
                if (strReplace.Replace("  ", " ")== strFound) { DeSelectWord(sldNum, shp.Name, txtRng, strFound, lngStartPos, lngRow, lngCol); return; }
                StreamWriter sw = new StreamWriter(PPTAttribute.supportfile);
                string errtxt = sldNum + "|" + strFound + "|" + strReplace.Replace("  "," ");
                sw.WriteLine(errtxt);
                sw.Close();

                frmObj.ShowDialog();
            if(PPTAttribute.reviewExitFlag == true) { DeSelectWord(sldNum, shp.Name, txtRng, strFound, lngStartPos, lngRow, lngCol); return; }
            if (PPTAttribute.discardFlag == false)
            {
                //SelectWord(sldNum, shp.Name, txtRng, strFound, lngStartPos, lngRow, lngCol);
                DeSelectWord(sldNum, shp.Name, txtRng, strFound, lngStartPos, lngRow, lngCol);
                if (logic == "L3")
                {
                    switch (PPTAttribute.ErrIndex)
                    {
                        case 0:
                            strReplace = strFound.Replace("Around", "About");
                            txtRng.Replace(strFound, strReplace);
                            break;
                        case 1:
                            strReplace = strFound.Replace("Around", "Approximately");
                            txtRng.Replace(strFound, strReplace);
                            break;
                        case 2:
                            strReplace = strFound.Replace("Around ", "~");
                            txtRng.Replace(strFound, strReplace);
                            break;
                    }

                    PPTAttribute.ErrIndex = 0;
                }
                else if (logic == "L4")
                {
                    switch (PPTAttribute.ErrIndex)
                    {
                        case 0:
                            strReplace = strFound.Replace("-", "and").Replace("–", "and").Replace("—", "and");
                            txtRng.Replace(strFound, strReplace);
                            break;
                        case 1:
                            strReplace = strFound.Replace("-", "to").Replace("–", "to").Replace("—", "to");
                            strReplace = strReplace.Replace("Between", "From");
                            txtRng.Replace(strFound, strReplace);
                            break;
                        case 2:
                            strReplace = strFound.Replace(" - ", "–").Replace(" – ", "–").Replace(" — ", "—");
                            strReplace = strReplace.Replace("Between", "During");
                            txtRng.Replace(strFound, strReplace);
                            break;
                        case 3:
                            strReplace = strFound.Replace(" - ", "–").Replace(" – ", "–").Replace(" — ", "—");
                            strReplace = strReplace.Replace("Between", "Over");
                            txtRng.Replace(strFound, strReplace);
                            break;
                    }
                    PPTAttribute.ErrIndex = 0;
                }
                else if(logic=="L15")
                {
                    List<string> replacetext = new List<string>();
                    replacetext = strReplace.Split(',').ToList();
                    strReplace = replacetext[PPTAttribute.ErrIndex];
                    txtRng.Replace(strFound, strReplace);
                    
                }
                else
                {
                    txtRng.Replace(strFound, strReplace.Replace("  "," "));
                }
            } //Close if User click on Discard Button --
            else
            {
                DeSelectWord(sldNum, shp.Name, txtRng, strFound, lngStartPos, lngRow, lngCol);
            }

            PPTAttribute.discardFlag = false;
        }
                

        #endregion

       #region Create a Array List as per error type -
        public void pCreateArray_List(int sldNum,string shpName, PowerPoint.TextRange txtRng,string wlError,string strMethod,int r=0 ,int c=0)
        {
            DataTable wList1Dic = getWrokList("List1");
            setTableColumn();
            string pattern;
            pattern = @"[HQhq][1-4][1-2][0-9]{3}";                                          //pPatternH12016ToH1_2016_array
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternH12016ToH1_2016", r,c);

            pattern = @"FY[ ][1-2][0-9]{3}";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternFY_2016ToFY2016", r, c); //pPatternFY_2016ToFY2016_Array

            pattern =  @"Around[ ][0-9]{1,20}";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternAround00", r, c);               //pPatternAround00_Array

            pattern = @"Between[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternBewtweenYear", r, c);           //pPatternBewtweenYear_array

            pattern = @"From[ ][1-2][0-9]{3}[ ]?[-–—][ ]?[1-2][0-9][0-9][0-9]";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternFrom_YYYYToFYYYYY", r, c);                //pPatternFrom_YYYYToFYYYYY_array

            pattern = @"end[ ][o]?[f]?[ ]?[1-2][0-9]{3}";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternEndYear", r, c);             //pPatternEndYear_array

            pattern = @"[0-9]{1,2}[:][0-9]{2}[ ][AP][M]";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternTime_AM_PM", r, c);             // pPatternTime_AM_PM_array

            pattern = @"[0-9,.]{1,20}[ ]?(mile|tonne|yard|gram|litre|gallon|pint|inch|inches|hectare|hour|minute|day|month|year|week|barrel|knot|feet|foot|metric ton|ton|km|kg|lbs|Btu|ha)[e$]?[s$]?";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPattern_Measure_Units", r, c);       //pPattern_Measure_Units_array

            pattern = @"(one|two|three|four|five|six|seven|eight|nine)[ ]?(mile|tonne|yard|gram|litre|gallon|pint|inch|inches|hectare|hour|minute|day|month|year|week|barrel|knot|feet|foot|metric ton|ton|km|kg|lbs|Btu|ha)[e$]?[s$]?";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPattern_Measure_Units_1", r, c);       //pPattern_Measure_Units_1_array

            pattern = @"[$€£][ ][0-9]{1,10}";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPattern_Currency_No_Space", r, c);           //pPattern_Currency_No_Space_array

            pattern = @"(INR|JPY|CNY|RMB|AUS|BRL|CAD|AED|SGD|ZAR)[0-9]{1,30}";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPattern_Currency_With_Space", r, c);           //pPattern_Currency_With_Space_array

            //pattern = @"[0-9,.]{1,20}[ ]?(sq|cu)[.]?[ ]?(ft|m)[.]?";
            //getErrorList(sldNum, shpName, txtRng, pattern, "pPatternSqFt_CuM", r, c);           //pPatternSqFt_CuM_array

            pattern = @"[0-9.,]{1,20}[ ]?(sq|cu)[.]?[ ]?(ft|m)[.]?";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternSqFt_CuM", r, c);           //pPatternSqFt_CuM_array

            pattern = @"(one|two|three|four|five|six|seven|eight|nine)[ ]?(sq|cu)[.]?[ ]?(ft|m)[.]?";
            getErrorList(sldNum, shpName, txtRng, pattern, "pPatternSqFt_CuM_1", r, c);           //pPatternSqFt_CuM_1_array



            //pattern = @"[ ](" + strWord + ")[ ](\w*)[ ]?";
            //getErrorList(sldNum, shpName, txtRng, pattern, "pPattern_Measure_Units_1", r, c);           //pPattern_Preposition_array

            
            getErrorList(sldNum, shpName, txtRng, null, "M2_WL1", r, c);           //CheckEditorial_TextRange_M2_WL1

            
            getErrorList(sldNum, shpName, txtRng, null, "M2_WL2", r, c);           //CheckEditorial_TextRange_M2_WL2

            
            getErrorList(sldNum, shpName, txtRng, null, "M2_DS", r, c);           //CheckEditorial_TextRange_M2_DS

           

            DataView dv = arrStrListToCheck.DefaultView;
            dv.Sort = "StrPosition";
            PowerPoint.TextRange ErrtxtRng=null;
            arrStrListToCheck = dv.ToTable();
            for (int rowCount=0; rowCount< arrStrListToCheck.Rows.Count; rowCount++)
            {
                if (PPTAttribute.reviewExitFlag == true && PPTAttribute.discardFlag==false) { return; }
                int sldNo = Convert.ToInt32(arrStrListToCheck.Rows[rowCount]["SldNum"]);
                string shpname = Convert.ToString(arrStrListToCheck.Rows[rowCount]["ShapeName"]);
                string strFound = Convert.ToString(arrStrListToCheck.Rows[rowCount]["StrFound"]);
                int tblRow = Convert.ToInt32(arrStrListToCheck.Rows[rowCount]["tbRow"]);
                int tblCol = Convert.ToInt32(arrStrListToCheck.Rows[rowCount]["tbCol"]);
                PowerPoint.Shape shp = ActivePPT.Slides[sldNo].Shapes[shpname];
                if (shp.Type == MsoShapeType.msoTable)
                {
                    ErrtxtRng=ActivePPT.Slides[sldNo].Shapes[shpname].Table.Rows[r].Cells[c].Shape.TextFrame.TextRange;
                    //PowerPoint.DataTable tbl = ActivePPT.Slides[sldNo].Shapes.t
                    //ErrtxtRng = tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange;
                }
                else
                {
                    ErrtxtRng = shp.TextFrame.TextRange;
                }
                //tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPatternH12016ToH1_2016")
                {
                    wlError = pPatternH12016ToH1_2016(sldNo, shp, ErrtxtRng, "L1", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPatternFY_2016ToFY2016")
                {
                    wlError = pPatternFY_2016ToFY2016(sldNo, shp, ErrtxtRng, "L2", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPatternAround00")
                {
                    wlError = pPatternAround00(sldNo, shp, ErrtxtRng, "L3", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPatternBewtweenYear")
                {
                    wlError = pPatternBewtweenYear(sldNo, shp, ErrtxtRng, "L4", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPatternEndYear")
                {
                    wlError = pPatternEndYear(sldNo, shp, ErrtxtRng, "L5", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPatternFrom_YYYYToFYYYYY")
                {
                    wlError = pPatternFrom_YYYYToFYYYYY(sldNo, shp, ErrtxtRng, "L6", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPatternTime_AM_PM")
                {
                    wlError = pPatternTime_AM_PM(sldNo, shp, ErrtxtRng, "L7", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPattern_Measure_Units_1")
                {
                    wlError = pPattern_Measure_Units_1(sldNo, shp, ErrtxtRng, "L8", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPattern_Measure_Units")
                {
                    wlError = pPattern_Measure_Units(sldNo, shp, ErrtxtRng, "L9", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPattern_Currency_No_Space")
                {
                    wlError = pPattern_Currency_No_Space(sldNo, shp, ErrtxtRng, "L10", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPattern_Currency_With_Space")
                {
                    wlError = pPattern_Currency_With_Space(sldNo, shp, ErrtxtRng, "L11", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPatternSqFt_CuM_1")
                {
                    wlError = pPatternSqFt_CuM_1(sldNo, shp, ErrtxtRng, "L12", strMethod, strFound, tblRow, tblCol);
                }
                //Array of years till 2099 and preceded by H1, H2, Q1, Q2, Q3, Q4 without space  (H12012 with H1 2012)
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "pPatternSqFt_CuM")
                {
                    wlError = pPatternSqFt_CuM(sldNo, shp, ErrtxtRng, "L13", strMethod, strFound, tblRow, tblCol);
                }
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "M2_WL1")
                {
                    wlError = CheckEditorialwList1Dic(sldNo, shp, ErrtxtRng, "L15", strMethod, strFound, tblRow, tblCol);

                }
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "M2_WL2")
                {
                    wlError = CheckEditorialwList2Dic(sldNo, shp, ErrtxtRng, "L16", strMethod, strFound, tblRow, tblCol);
                }
                else if (arrStrListToCheck.Rows[rowCount]["PatternType"].ToString() == "M2_DS")
                {
                    wlError = CheckEditorialwList2Dic(sldNo, shp, ErrtxtRng, "L17", strMethod, strFound, tblRow, tblCol);
                }
                PPTAttribute.discardFlag =false;                // Set the value after discard the correction 
                                        
            }
            arrStrListToCheck.Clear();  // Referesh Datatable
        }
       
        #endregion

        public void SelectWord(int sldNum, string shpName, PowerPoint.TextRange txtRng,string cWord,int stPos,int lngRow=0,int lngCol=0)
        {
            PowerPoint.Shape shp = ActivePPT.Slides[sldNum].Shapes[shpName];
            int l = cWord.Length;
            shp.Select();
            if (shp.Type==MsoShapeType.msoTable)
            {
                PowerPoint.TextRange tRng= ppApp.ActiveWindow.Selection.ShapeRange.Table.Cell(lngRow, lngCol).Shape.TextFrame.TextRange.Characters(stPos, l);
                colorCode = tRng.Font.Color;
                fSize =  tRng.Font.Size;
                fBold = tRng.Font.Bold;
                // - take font spacification globly
                tRng.Font.Shadow = MsoTriState.msoTrue;
                tRng.Font.Color.RGB= Color.FromArgb(0, 255, 0).ToArgb();
                tRng.Font.Bold = MsoTriState.msoTrue;
                tRng.Font.Size = 16;
            }
            else
            {
                PowerPoint.TextRange tRng = ppApp.ActiveWindow.Selection.TextRange.Characters(stPos, l);
                colorCode = tRng.Font.Color;
                fSize = tRng.Font.Size;
                fBold = tRng.Font.Bold;

                //tRng.Font.Shadow = MsoTriState.msoTrue;
                tRng.Font.Color.RGB = Color.FromArgb(0, 255, 0).ToArgb();
                tRng.Font.Bold = MsoTriState.msoTrue;
                tRng.Font.Size = 16;
            }
        }
        public void DeSelectWord(int sldnum,string shpname, PowerPoint.TextRange txtRng, string cWord, int stPos,int lngRow=0,int lngCol=0)
        {
            PowerPoint.Shape shp = ActivePPT.Slides[sldnum].Shapes[shpname];
            shp.Select();
            int n = stPos;
            int l = cWord.Length;
            if (shp.Type==MsoShapeType.msoTable)
            {
                PowerPoint.TextRange tRng = ppApp.ActiveWindow.Selection.ShapeRange.Table.Cell(lngRow, lngCol).Shape.TextFrame.TextRange.Characters(n, l);
                tRng.Font.Shadow = MsoTriState.msoFalse;
                txtRng.Font.Color.RGB = Color.FromArgb(0, 0, 0).ToArgb();
                //txtRng.Font.Color.RGB= colorCode;
                txtRng.Font.Bold = fBold;
                txtRng.Font.Size = fSize;
            }
            else
            {
                PowerPoint.TextRange tRng = ppApp.ActiveWindow.Selection.TextRange.Characters(n, l);
                tRng.Font.Shadow = MsoTriState.msoFalse;
                txtRng.Font.Color.RGB = Color.FromArgb(0, 0, 0).ToArgb();
                txtRng.Font.Bold = fBold;
                txtRng.Font.Size = fSize;
            }
        }


    }
}
