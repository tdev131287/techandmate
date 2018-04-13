using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Graph = Microsoft.Office.Interop.Graph;
using Microsoft.Office.Core;
namespace TSCPPT_Addin
{
     class Shapecheck
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        public bool CheckIfBoxAlreadyExist(int sldIndex,DataTable dt)
        {
            int boxNum = 0;
            bool returnValue=false;
            try
            {
                string shpName = Convert.ToString(dt.Rows[0]["Name"]);
                float shpLeft = (float)Convert.ToDecimal((dt.Rows[0]["ShapeLeft"]));
                float shpTop = (float)Convert.ToDecimal((dt.Rows[0]["ShapeTop"]));
                float shpWidth = (float)Convert.ToDecimal((dt.Rows[0]["ShapeWidth"]));
                float shpHeight = (float)Convert.ToDecimal((dt.Rows[0]["ShapeHeight"]));

                int shpCount = ActivePPT.Slides[sldIndex].Shapes.Count;
                int cnt = 0;
                foreach (PowerPoint.Shape shp in ActivePPT.Slides[sldIndex].Shapes)
                {
                    if (shp.Name == shpName) { cnt++; }
                }
                if (cnt == 0) { boxNum = FindWhenNoBox(sldIndex, shpName, shpWidth, shpHeight, shpLeft, shpTop); }
                else if (cnt == 1) { boxNum = FindWhenOneBox(sldIndex, shpName, shpWidth, shpHeight, shpLeft, shpTop); }
                else if (cnt > 1) { boxNum = FindWhenMultipleBox(sldIndex, shpName, shpWidth, shpHeight, shpLeft, shpTop); }
                if (boxNum > 0) { returnValue = true; }
                else { returnValue = false; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CheckIfBoxAlreadyExist");
            }
            return (returnValue);
        }
        // Function chekc if shape allready exist the return shape name
        public string  CheckIfBoxAlreadyExist1(int sldIndex, DataTable dt)
        {
            int boxNum = 0;
            String returnValue =null;
            try
            {
                string shpName = Convert.ToString(dt.Rows[0]["Name"]);
                float shpLeft = (float)Convert.ToDecimal((dt.Rows[0]["ShapeLeft"]));
                float shpTop = (float)Convert.ToDecimal((dt.Rows[0]["ShapeTop"]));
                float shpWidth = (float)Convert.ToDecimal((dt.Rows[0]["ShapeWidth"]));
                float shpHeight = (float)Convert.ToDecimal((dt.Rows[0]["ShapeHeight"]));

                int shpCount = ActivePPT.Slides[sldIndex].Shapes.Count;
                int cnt = 0;
                foreach (PowerPoint.Shape shp in ActivePPT.Slides[sldIndex].Shapes)
                {
                    if (shp.Name == shpName) { cnt++; }
                }
                if (cnt == 0) { boxNum = FindWhenNoBox(sldIndex, shpName, shpWidth, shpHeight, shpLeft, shpTop); }
                else if (cnt == 1) { boxNum = FindWhenOneBox(sldIndex, shpName, shpWidth, shpHeight, shpLeft, shpTop); }
                else if (cnt > 1) { boxNum = FindWhenMultipleBox(sldIndex, shpName, shpWidth, shpHeight, shpLeft, shpTop); }
                if (boxNum > 0) { returnValue = ActivePPT.Slides[sldIndex].Shapes[boxNum].Name; }
                else { returnValue = null; }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CheckIfBoxAlreadyExist1");
            }
            return (returnValue);
        }

        //sld_num, shpName, shpWidth, shpHeight, shpLeft, shpTop
        public int FindWhenNoBox(int sld_Num, string shpName, float shpWidth, float shpHeight, float shpLeft, float shpTop)
        {
            
            int lth = 10, hth = 10;
            int tgtShpNum = 0;
            try
            {
                string shpText = null;
                int shpCount = ActivePPT.Slides[sld_Num].Shapes.Count;
                for (int sh = 1; sh <= shpCount; sh++)
                {
                    PowerPoint.Shape shObj = ActivePPT.Slides[sld_Num].Shapes[sh];
                    if (shObj.Type == MsoShapeType.msoAutoShape || shObj.Type == MsoShapeType.msoPlaceholder || shObj.Type == MsoShapeType.msoTextBox)
                    {
                        shpText = ActivePPT.Slides[sld_Num].Shapes[sh].TextFrame.TextRange.Text;
                    }
                    PowerPoint.Shape currentshp = ActivePPT.Slides[sld_Num].Shapes[sh];
                    string t_shpName = currentshp.Name;
                    float t_shpWidth = currentshp.Width;
                    float t_shpheight = currentshp.Height;
                    float t_shpleft = currentshp.Left;
                    float t_shptop = currentshp.Top;
                    float t_midV = (t_shptop + t_shpheight) / 2;
                    float t_midh = (t_shpleft + t_shpWidth) / 2;
                    float TshpBottom = shpTop + shpHeight;
                    float TshpRight = shpLeft + shpWidth;
                    if (t_shpName.PadLeft(5) == shpName.PadLeft(5))
                    {
                        if (t_shpWidth == shpWidth && t_shpleft == shpLeft && t_shptop == shpTop && t_shpheight == shpHeight) { tgtShpNum = sh; }
                        else if (t_midV >= shpTop && t_midV <= TshpBottom && t_midh >= shpLeft && t_midh <= TshpRight && t_shpleft >= shpLeft - lth && t_shpleft <= shpLeft + lth && shpTop >= t_shptop - hth && t_shptop <= shpTop + hth) { tgtShpNum = sh; }
                    }
                    else if (shpWidth == t_shpWidth && t_shpleft == shpLeft && t_shptop == shpTop && t_shpheight == shpHeight) { tgtShpNum = sh; }
                    else if ((t_midV >= shpTop && t_midV <= TshpBottom) && (t_midh >= shpLeft && t_midh <= TshpRight) && (t_shpleft >= shpLeft - lth && t_shpleft <= shpLeft + lth) && (t_shptop >= shpTop - hth && t_shptop <= shpTop - hth) && (t_shptop <= shpTop + hth)) { tgtShpNum = sh; }
                    else if ((t_shpName == "Note Box") && (shpText.PadLeft(4) == shpName.PadLeft(4)))
                    {
                        if ((t_midV >= t_shptop - 2 * hth && t_midV <= TshpBottom + 2 * hth) && (t_midh >= shpLeft - 2 * lth && t_midh <= TshpRight + 2 * lth) && (t_shpleft >= shpLeft - 3 * lth && t_shpleft <= shpLeft + 3 * lth) && (t_shptop >= t_shptop - 3 * hth && t_shptop <= t_shptop + 3 * hth))
                        {
                            tgtShpNum = sh;
                        }
                        else if (shpName == "Source Box" && shpText.PadLeft(6) == shpName.PadLeft(6))
                        {
                            if ((t_midV >= t_shptop - 2 * hth && t_midV <= TshpBottom + 2 * hth) && (t_midh >= shpLeft - 2 * lth && t_midh <= TshpRight + 2 * lth) && (t_shpleft >= shpLeft - 3 * lth && t_shpleft <= shpLeft + 3 * lth) && (t_shptop >= t_shptop - 3 * hth && t_shptop <= t_shptop + 3 * hth))
                            {
                                tgtShpNum = sh;
                            }
                        }
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "FindWhenNoBox");
            }
            return (tgtShpNum);
        }
        public int FindWhenOneBox(int sld_Num, string shpName, double shpWidth, double shpHeight, double shpLeft, double shpTop)
        {
            int tgtShpNum = 0;
            try
            {
                int shpCount = ActivePPT.Slides[sld_Num].Shapes.Count;
                for (int sh = 1; sh <= shpCount; sh++)
                {
                    string tshpName = ActivePPT.Slides[sld_Num].Shapes[sh].Name;
                    if (tshpName == shpName) { tgtShpNum = sh; }
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "FindWhenOneBox");
            }
            return (tgtShpNum);
           
        }
        public int FindWhenMultipleBox(int sld_Num, string shpName, float shpWidth, float shpHeight, float shpLeft, float shpTop)
        {
            
            int lth = 10, hth = 10;
            int tgtShpNum = 0;
            string shpText = null;
            try
            {
                int shpCount = ActivePPT.Slides[sld_Num].Shapes.Count;
                for (int sh = 1; sh <= shpCount; sh++)
                {
                    PowerPoint.Shape shObj = ActivePPT.Slides[sld_Num].Shapes[sh];
                    if (shObj.Type == MsoShapeType.msoAutoShape || shObj.Type == MsoShapeType.msoPlaceholder || shObj.Type == MsoShapeType.msoTextBox)
                    {
                        shpText = ActivePPT.Slides[sld_Num].Shapes[sh].TextFrame.TextRange.Text;
                    }
                    PowerPoint.Shape currentshp = ActivePPT.Slides[sld_Num].Shapes[sh];
                    string t_shpName = currentshp.Name;
                    float t_shpWidth = currentshp.Width;
                    float t_shpheight = currentshp.Height;
                    float t_shpleft = currentshp.Left;
                    float t_shptop = currentshp.Top;
                    float t_midV = (t_shptop + t_shpheight) / 2;
                    float t_midh = (t_shpleft + t_shpWidth) / 2;
                    float TshpBottom = shpTop + shpHeight;
                    float TshpRight = shpLeft + shpWidth;
                    if (shpName == t_shpName)
                    {
                        if (t_shpWidth == shpWidth && t_shpleft == shpLeft && t_shptop == shpTop && t_shpheight == shpHeight) { tgtShpNum = sh; }
                        else if (t_midV >= shpTop && t_midV <= TshpBottom && t_midh >= shpLeft && t_midh <= TshpRight && t_shpleft >= shpLeft - lth && t_shpleft <= shpLeft + lth && shpTop >= t_shptop - hth && t_shptop <= shpTop + hth) { tgtShpNum = sh; }
                    }
                    else if (shpWidth == t_shpWidth && t_shpleft == shpLeft && t_shptop == shpTop && t_shpheight == shpHeight) { tgtShpNum = sh; }
                    else if ((t_midV >= shpTop && t_midV <= TshpBottom) && (t_midh >= shpLeft && t_midh <= TshpRight) && (t_shpleft >= shpLeft - lth && t_shpleft <= shpLeft + lth) && (t_shptop >= shpTop - hth && t_shptop <= shpTop - hth) && (t_shptop <= shpTop + hth)) { tgtShpNum = sh; }
                    else if ((t_shpName == "Note Box") && (shpText.PadLeft(4) == shpName.PadLeft(4)))
                    {
                        if ((t_midV >= t_shptop - 2 * hth && t_midV <= TshpBottom + 2 * hth) && (t_midh >= shpLeft - 2 * lth && t_midh <= TshpRight + 2 * lth) && (t_shpleft >= shpLeft - 3 * lth && t_shpleft <= shpLeft + 3 * lth) && (t_shptop >= t_shptop - 3 * hth && t_shptop <= t_shptop + 3 * hth))
                        {
                            tgtShpNum = sh;
                        }
                        else if (shpName == "Source Box" && shpText.PadLeft(6) == shpName.PadLeft(6))
                        {
                            if ((t_midV >= t_shptop - 2 * hth && t_midV <= TshpBottom + 2 * hth) && (t_midh >= shpLeft - 2 * lth && t_midh <= TshpRight + 2 * lth) && (t_shpleft >= shpLeft - 3 * lth && t_shpleft <= shpLeft + 3 * lth) && (t_shptop >= t_shptop - 3 * hth && t_shptop <= t_shptop + 3 * hth))
                            {
                                tgtShpNum = sh;
                            }
                        }
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "FindWhenMultipleBox");
            }
            return (tgtShpNum);
        }

        public string SelectedShapeNumber(int sldNum,PowerPoint.Shape selShape)
        {
            string selshpname = null;
            try
            {
                string selShpName = selShape.Name;
                float selShpWd = selShape.Width;
                float selShpHt = selShape.Height;
                float selShpLf = selShape.Left;
                float selShpTp = selShape.Top;

                for (int sn = 1; sn <= ActivePPT.Slides[sldNum].Shapes.Count; sn++)
                {
                    string tempShpName = ActivePPT.Slides[sldNum].Shapes[sn].Name;
                    float tempShpWd = ActivePPT.Slides[sldNum].Shapes[sn].Width;
                    float tempShpHt = ActivePPT.Slides[sldNum].Shapes[sn].Height;
                    float tempShpLf = ActivePPT.Slides[sldNum].Shapes[sn].Left;
                    float tempShpTp = ActivePPT.Slides[sldNum].Shapes[sn].Top;
                    if (tempShpName == selShpName && tempShpWd == selShpWd && tempShpHt == selShpHt && tempShpLf == selShpLf && tempShpTp == selShpTp)
                    {
                        selshpname = ActivePPT.Slides[sldNum].Shapes[sn].Name;
                        break;
                    }
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "SelectedShapeNumber");
            }
            return (selshpname);
        }

        public void CreateBullet(int sldNum, string shpname)
        {
            try
            {
                PowerPoint.TextRange txtRng = ActivePPT.Slides[sldNum].Shapes[shpname].TextFrame.TextRange;

                float indentVal = (float)0.64;
                var paragraphs = txtRng.Paragraphs(-1, -1);
               
                for (int index = 1; index <= txtRng.Paragraphs().Count; index++)
                {
                    //string paraText = txtRng.Paragraphs(index).Text;
                    
                    txtRng.Paragraphs(index).IndentLevel = index;
                    txtRng.Paragraphs(index).ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    txtRng.Paragraphs(index).ParagraphFormat.HangingPunctuation = MsoTriState.msoTrue;
                    txtRng.Paragraphs(index).ParagraphFormat.SpaceBefore = 6;
                    txtRng.Paragraphs(index).ParagraphFormat.SpaceAfter = 0;
                    txtRng.Paragraphs(index).ParagraphFormat.SpaceWithin = (float)0.85;
                    txtRng.Paragraphs(index).Parent.Ruler.Levels[index].FirstMargin = indentVal;//18 * (index - 1);
                    txtRng.Paragraphs(index).Parent.Ruler.Levels[index].LeftMargin = 18 * (index);
                    indentVal = indentVal + (float)0.47;
                }

            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "CreateBullet");
            }
                
        }

        public void setBulletTypeNone(int sldNum, string shpname)
        {
            PowerPoint.TextRange txtRng = ppApp.ActiveWindow.Selection.ShapeRange[1].TextFrame.TextRange;
            //PowerPoint.TextRange txtRng = ActivePPT.Slides[sldNum].Shapes[shpname].TextFrame.TextRange;
           
            //var paragraphs = txtRng.Paragraphs(-1, -1);

            for (int index = 1; index <= txtRng.Paragraphs().Count; index++)
            {
                txtRng.Paragraphs(index).ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
                txtRng.Paragraphs(index).IndentLevel = 1;
            }


        }

        public void setBulletImage(int sldNum, string shpname)
        {
            PowerPoint.TextRange txtRng = ppApp.ActiveWindow.Selection.ShapeRange[1].TextFrame.TextRange;
            //PowerPoint.TextRange txtRng = ActivePPT.Slides[sldNum].Shapes[shpname].TextFrame.TextRange;

            //var paragraphs = txtRng.Paragraphs(-1, -1);
            char myCharacter = (char)132;
            for (int index = 1; index <= txtRng.Paragraphs().Count; index++)
            {
                if (txtRng.Paragraphs(index).IndentLevel == 1)
                {
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Character = myCharacter;
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Font.Color.RGB = System.Drawing.Color.FromArgb(78, 204, 124).ToArgb();
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Font.Size = (float)10.84;
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Font.Name = "Wingdings 3";
                    txtRng.Paragraphs(index).Font.Color.RGB= System.Drawing.Color.FromArgb(57,42,30).ToArgb();
                    txtRng.Paragraphs(index).Font.Size = 12;
                    txtRng.Paragraphs(index).Font.Name = "Corbel";
                }
                else if (txtRng.Paragraphs(index).IndentLevel == 2)
                {
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Character = 167;
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Font.Color.RGB = System.Drawing.Color.FromArgb(78, 204, 124).ToArgb();
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Font.Size = (float)11;
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Font.Name = "Wingdings";
                    txtRng.Paragraphs(index).Font.Color.RGB = System.Drawing.Color.FromArgb(57, 42, 30).ToArgb();
                    txtRng.Paragraphs(index).Font.Size = 11;
                    txtRng.Paragraphs(index).Font.Name = "Corbel";
                }
                else if (txtRng.Paragraphs(index).IndentLevel == 3)
                {
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Character = 2013;
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Font.Color.RGB = System.Drawing.Color.FromArgb(78, 204, 124).ToArgb();
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Font.Size = (float)11;
                    txtRng.Paragraphs(index).ParagraphFormat.Bullet.Font.Name = "Corbel";
                    txtRng.Paragraphs(index).Font.Color.RGB = System.Drawing.Color.FromArgb(57, 42, 30).ToArgb();
                    txtRng.Paragraphs(index).Font.Size = 11;
                    txtRng.Paragraphs(index).Font.Name = "Corbel";
                }


            }


        }
        public void FormatBulletInShape(int sldNum,string shpname)
        {
            PowerPoint.Presentation actPPT = Globals.ThisAddIn.Application.Presentations.Open(PPTAttribute.mPPTPath);
            actPPT.Slides[4].Shapes["Text Placeholder 21"].PickUp();
            try
            {
                ActivePPT.Slides[sldNum].Shapes[shpname].Apply();
                //ActivePPT.Slides[sldNum].Shapes["Text Box"].Apply();
            }
            catch { }
            actPPT.Close();
        }
        public void AdjustPosition(int sldNum, int shpNum, DataTable dt)
        {
            PowerPoint.Shape shp = ppApp.ActivePresentation.Slides[sldNum].Shapes[shpNum];
            shp.Left = (float) 390.0895;
            shp.Top = (float)269.7313;
            shp.Width =(float) Convert.ToDouble(dt.Rows[0]["ShapeWidth"]);
            shp.Height = (float)Convert.ToDouble(dt.Rows[0]["ShapeHeight"]);
           
        }
        public List<string> FindSelectedCharts()
        {
            List<string> selCharts = new List<string>();
            List<string> gselCharts = new List<string>();
            PowerPoint.Shape osh;
            try
            {
                int sld_num = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;

                int num_shp = ppApp.ActiveWindow.Selection.ShapeRange.Count;
                int cntSelCht = 0;
                //int subSelShp = ppApp.ActiveWindow.Selection.ChildShapeRange.Count;
                int subSelShp = ppApp.ActiveWindow.Selection.ShapeRange.Count;
                if (subSelShp != 0)
                {
                   
                    for (int cshp = 1; cshp <= subSelShp; cshp++)
                   
                    {
                        //osh = ppApp.ActiveWindow.Selection.ChildShapeRange[cshp];
                        osh = ppApp.ActiveWindow.Selection.ShapeRange[cshp];
                        if(osh.Type==MsoShapeType.msoGroup)
                        {
                            for(int x=1;x<=osh.GroupItems.Count;x++)
                            {
                                if(osh.GroupItems[x].HasChart == MsoTriState.msoTrue){ selCharts.Add(osh.GroupItems[x].Name); }
                            }
                        }
                        else if (osh.HasChart == MsoTriState.msoTrue) { selCharts.Add(osh.Name); }

                        //if (osh.HasChart == MsoTriState.msoTrue) { selCharts.Add(osh.Name); }

                    }
                }
                else
                {
                    for (int cshp = 1; cshp <= num_shp; cshp++)
                    {
                        osh = ppApp.ActiveWindow.Selection.ShapeRange[cshp];
                        if (osh.HasChart == MsoTriState.msoTrue) { selCharts.Add(osh.Name); }
                        if (osh.Type == MsoShapeType.msoGroup)
                        {
                            for (int gItm = 1; gItm <= osh.GroupItems.Count; gItm++)
                            {
                                if (osh.GroupItems[gItm].Type == MsoShapeType.msoGroup) { gselCharts = FindChartsInGroup(osh); }
                                else
                                {
                                    if (osh.GroupItems[gItm].HasChart == MsoTriState.msoTrue) { selCharts.Add(osh.Name); }
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "FindSelectedCharts");
                return selCharts;
            }
            return (selCharts);
        }
        // Chekc theme is loaded or not

        public List<string> FindChartsInGroup(PowerPoint.Shape osh)
        {
            List<string> gselCharts = new List<string>();
            for(int gItm=1;gItm<= osh.GroupItems.Count;gItm++)
            {
                if (osh.GroupItems[gItm].Type == MsoShapeType.msoGroup)
                {
                    FindChartsInGroup(osh.GroupItems[gItm]);
                }
                else
                {
                    if (osh.GroupItems[gItm].HasChart == MsoTriState.msoTrue) { gselCharts.Add(osh.GroupItems[gItm].Name); }
                    
                }
            }
            return (gselCharts);
        }

        public string chartType(PowerPoint.Chart myChart)
        {
            string chtType = null;
            try
            {
                if (myChart.ChartType == XlChartType.xlBarClustered || myChart.ChartType == XlChartType.xlBarStacked || myChart.ChartType == XlChartType.xlBarStacked100 ||
                myChart.ChartType == XlChartType.xl3DBarClustered || myChart.ChartType == XlChartType.xl3DBarStacked || myChart.ChartType == XlChartType.xl3DBarStacked100 ||
                myChart.ChartType == XlChartType.xlCylinderBarClustered || myChart.ChartType == XlChartType.xlCylinderBarStacked || myChart.ChartType == XlChartType.xlCylinderBarStacked100 ||
                myChart.ChartType == XlChartType.xlConeBarClustered || myChart.ChartType == XlChartType.xlConeBarStacked || myChart.ChartType == XlChartType.xlConeBarStacked100 ||
                myChart.ChartType == XlChartType.xlPyramidBarClustered || myChart.ChartType == XlChartType.xlPyramidBarStacked || myChart.ChartType == XlChartType.xlPyramidBarStacked100)
                { chtType = "Column"; }

                else if (myChart.ChartType == XlChartType.xlColumnClustered || myChart.ChartType == XlChartType.xlColumnStacked || myChart.ChartType == XlChartType.xlColumnStacked100 ||
                    myChart.ChartType == XlChartType.xl3DColumnClustered || myChart.ChartType == XlChartType.xl3DColumnStacked || myChart.ChartType == XlChartType.xl3DColumnStacked100 ||
                    myChart.ChartType == XlChartType.xlCylinderColClustered || myChart.ChartType == XlChartType.xlCylinderColStacked || myChart.ChartType == XlChartType.xlCylinderColStacked100 ||
                    myChart.ChartType == XlChartType.xlConeColClustered || myChart.ChartType == XlChartType.xlConeColStacked || myChart.ChartType == XlChartType.xlConeColStacked100 ||
                    myChart.ChartType == XlChartType.xlPyramidColClustered || myChart.ChartType == XlChartType.xlPyramidColStacked || myChart.ChartType == XlChartType.xlPyramidColStacked100 ||
                    myChart.ChartType == XlChartType.xl3DColumn || myChart.ChartType == XlChartType.xlConeCol || myChart.ChartType == XlChartType.xlCylinderCol || myChart.ChartType == XlChartType.xlPyramidCol)
                { chtType = "Column"; }

                else if (myChart.ChartType == XlChartType.xlLine || myChart.ChartType == XlChartType.xlLineMarkers || myChart.ChartType == XlChartType.xlLineStacked || myChart.ChartType == XlChartType.xl3DLine ||
                    myChart.ChartType == XlChartType.xlLineStacked100 || myChart.ChartType == XlChartType.xlLineMarkersStacked || myChart.ChartType == XlChartType.xlLineMarkersStacked100)
                { chtType = "Line"; }
                else if (myChart.ChartType == XlChartType.xlXYScatter || myChart.ChartType == XlChartType.xlXYScatterLines || myChart.ChartType == XlChartType.xlXYScatterLinesNoMarkers ||
                    myChart.ChartType == XlChartType.xlXYScatterSmooth || myChart.ChartType == XlChartType.xlXYScatterSmoothNoMarkers)
                { chtType = "Line"; }
                else if (myChart.ChartType == XlChartType.xlPie || myChart.ChartType == XlChartType.xl3DPie || myChart.ChartType == XlChartType.xlPieExploded || myChart.ChartType == XlChartType.xl3DPieExploded ||
                myChart.ChartType == XlChartType.xlPieOfPie || myChart.ChartType == XlChartType.xlBarOfPie)
                { chtType = "Pie"; }
                else if (myChart.ChartType == XlChartType.xlDoughnut || myChart.ChartType == XlChartType.xlDoughnutExploded)
                { chtType = "Doughnut"; }
                else if (myChart.ChartType == XlChartType.xlArea || myChart.ChartType == XlChartType.xlAreaStacked || myChart.ChartType == XlChartType.xlAreaStacked100 ||
                myChart.ChartType == XlChartType.xl3DArea || myChart.ChartType == XlChartType.xl3DAreaStacked || myChart.ChartType == XlChartType.xl3DAreaStacked100)
                { chtType = "Area"; }
                else if (myChart.ChartType == XlChartType.xlBubble || myChart.ChartType == XlChartType.xlBubble3DEffect)
                { chtType = "Bubble"; }
                else if (myChart.ChartType == XlChartType.xlStockHLC || myChart.ChartType == XlChartType.xlStockOHLC || myChart.ChartType == XlChartType.xlStockVHLC || myChart.ChartType == XlChartType.xlStockVOHLC)
                { chtType = "Stock"; }

                else if (myChart.ChartType == XlChartType.xlSurface || myChart.ChartType == XlChartType.xlSurfaceWireframe || myChart.ChartType == XlChartType.xlSurfaceTopView || myChart.ChartType == XlChartType.xlSurfaceTopViewWireframe)
                { chtType = "Surface"; }
                else if (myChart.ChartType == XlChartType.xlRadar || myChart.ChartType == XlChartType.xlRadarFilled || myChart.ChartType == XlChartType.xlRadarMarkers)
                { chtType = "Radar"; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "chartType");
            }
            return (chtType);
        }
        public bool chart3D(PowerPoint.Chart myChart)
        {
            bool chart3dVal = false;
            try
            {
                if (myChart.ChartType == XlChartType.xl3DBarClustered || myChart.ChartType == XlChartType.xl3DBarStacked || myChart.ChartType == XlChartType.xl3DBarStacked100 ||
                myChart.ChartType == XlChartType.xlCylinderBarClustered || myChart.ChartType == XlChartType.xlCylinderBarStacked || myChart.ChartType == XlChartType.xlCylinderBarStacked100 ||
                myChart.ChartType == XlChartType.xlConeBarClustered || myChart.ChartType == XlChartType.xlConeBarStacked || myChart.ChartType == XlChartType.xlConeBarStacked100 ||
                myChart.ChartType == XlChartType.xlPyramidBarClustered || myChart.ChartType == XlChartType.xlPyramidBarStacked || myChart.ChartType == XlChartType.xlPyramidBarStacked100)
                {
                    chart3dVal = true;
                }
                else if (myChart.ChartType == XlChartType.xl3DColumnClustered || myChart.ChartType == XlChartType.xl3DColumnStacked || myChart.ChartType == XlChartType.xl3DColumnStacked100 ||
                myChart.ChartType == XlChartType.xlCylinderColClustered || myChart.ChartType == XlChartType.xlCylinderColStacked || myChart.ChartType == XlChartType.xlCylinderColStacked100 ||
                myChart.ChartType == XlChartType.xlConeColClustered || myChart.ChartType == XlChartType.xlConeColStacked || myChart.ChartType == XlChartType.xlConeColStacked100 ||
                myChart.ChartType == XlChartType.xlPyramidColClustered || myChart.ChartType == XlChartType.xlPyramidColStacked || myChart.ChartType == XlChartType.xlPyramidColStacked100 ||
                myChart.ChartType == XlChartType.xl3DColumn || myChart.ChartType == XlChartType.xlConeCol || myChart.ChartType == XlChartType.xlCylinderCol || myChart.ChartType == XlChartType.xlPyramidCol)
                {
                    chart3dVal = true;
                }
                else if (myChart.ChartType == XlChartType.xl3DLine)
                {
                    chart3dVal = true;
                }
                else if (myChart.ChartType == XlChartType.xl3DPie || myChart.ChartType == XlChartType.xl3DPieExploded)
                {
                    chart3dVal = true;
                }
                else if (myChart.ChartType == XlChartType.xl3DArea || myChart.ChartType == XlChartType.xl3DAreaStacked || myChart.ChartType == XlChartType.xl3DAreaStacked100)
                {
                    chart3dVal = true;
                }
                else if (myChart.ChartType == XlChartType.xlBubble3DEffect)
                {
                    chart3dVal = true;
                }
                else if (myChart.ChartType == XlChartType.xlSurface || myChart.ChartType == XlChartType.xlSurfaceWireframe)
                {
                    chart3dVal = true;
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "chart3D");
            }

            return (chart3dVal);
        }

        public string CheckFormat_TBox(int sldnum,string shpname)
        {
            string shpError = null;
            return shpError;
        }

        public string CheckFormat_Bullet(int sldnum, string shpname)
        {
            string shpError = null;
            return shpError;
        }

    }
}
