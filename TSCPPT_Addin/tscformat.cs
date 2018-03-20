using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TSCPPT_Addin
{
    class tscformat
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        public void tsc_loadtheme()
        {

            System.Text.StringBuilder specification = new System.Text.StringBuilder();
            foreach (PowerPoint.Shape shp in ppApp.ActiveWindow.Selection.ShapeRange)
            {

                MsoTriState vstatus = shp.Line.Visible;                 //LineVisible
                MsoTriState fillStatus = shp.Fill.Visible;              //FillVisible
                //str fillColor = shp.Fill.ForeColor.RGB;               //FillColor-----------------------Not
                float fillTrasn = shp.Fill.Transparency;                //FillTransparency
                MsoTriState sVisible = shp.Shadow.Visible;              //ShadowVisible
                float sleft = shp.Left;                                 //ShapeLeft
                float stop = shp.Top;                                   //ShapeTop
                float swidth = shp.Width;                               //ShapeWidth
                float shpHeight = shp.Height;                           //ShapeHeight
                float rot = shp.Rotation;                                 //Rotaion
                MsoTriState lrvalue = shp.LockAspectRatio;              //LockAspectRatio
                MsoTextOrientation txtOrientation = shp.TextFrame.Orientation;      //Orientation
                MsoVerticalAnchor txtAnchor = shp.TextFrame.VerticalAnchor;         //VerticalAnchor
                PowerPoint.PpAutoSize autosize = shp.TextFrame.AutoSize;            //AutoSize
                float mgleft = shp.TextFrame.MarginLeft;                            //MarginLeft
                float mgRight = shp.TextFrame.MarginLeft;                   //MarginRight
                float mgtop = shp.TextFrame.MarginTop;                      //MarginTop
                float mgbottom = shp.TextFrame.MarginBottom;                      //MarginBottom
                MsoTriState wWrap = shp.TextFrame.WordWrap;                     //WordWrap
                string dtext = shp.TextFrame.TextRange.Text;                    //DefaultText
                string fname = shp.TextFrame.TextRange.Font.Name;               //FontName
                MsoTriState txbold = shp.TextFrame.TextRange.Font.Bold;         //Bold
                MsoTriState txItalics = shp.TextFrame.TextRange.Font.Italic;        //Italics
                MsoTriState txUnderline = shp.TextFrame.TextRange.Font.Underline;        //Underline
                float txsize = shp.TextFrame.TextRange.Font.Size;                       //FontSize
                                                                                        //FontColor                                                                      
                PowerPoint.ShadowFormat shadow = shp.Shadow;                     //Shadow
                //PowerPoint.PpChangeCase txcase = shp.TextFrame.TextRange.ChangeCase;                //Case
                PowerPoint.BulletFormat txbuttlet = shp.TextFrame.TextRange.ParagraphFormat.Bullet;           //ParagraphBullet
                PowerPoint.PpParagraphAlignment txtAlig = shp.TextFrame.TextRange.ParagraphFormat.Alignment;            //ParagraphAlignment
                MsoTriState hPun = shp.TextFrame.TextRange.ParagraphFormat.HangingPunctuation;                       //ParagraphHangingPunctuation
                float psb = shp.TextFrame.TextRange.ParagraphFormat.SpaceBefore;                                    //ParagraphSpaceBefore
                float psa = shp.TextFrame.TextRange.ParagraphFormat.SpaceAfter;                                    //ParagraphSpaceAfter
                float psw = shp.TextFrame.TextRange.ParagraphFormat.SpaceWithin;                                    //ParagraphSpaceWithin
                float rlfm = shp.TextFrame.Ruler.Levels[1].FirstMargin;                                    //RulerLevel1FirstMargin
                float rllm = shp.TextFrame.Ruler.Levels[1].LeftMargin;                                    ////RulerLevel1LeftMargin
                for (int i = 1; i <= shp.TextFrame.TextRange.Paragraphs().Count; i++)
                {
                    string paraStr = shp.TextFrame.TextRange.Paragraphs(i).Text;
                    int indentVal = shp.TextFrame.TextRange.Paragraphs(i).IndentLevel;
                }
                specification.Append("LineVisible :" + vstatus + Environment.NewLine + "FillVisible :" + fillStatus + Environment.NewLine + "FillTransparency :" + fillTrasn + Environment.NewLine);
                specification.Append("ShapeLeft :" + sleft + Environment.NewLine + "ShapeTop :" + stop + Environment.NewLine + "ShapeWeight :" + swidth + Environment.NewLine + "ShapeHeight :" + shpHeight + Environment.NewLine);


                specification.Append("Rotaion :" + rot + Environment.NewLine + "LockAspectRatio :" + lrvalue + Environment.NewLine + "Orientation :" + txtOrientation + Environment.NewLine + "VerticalAnchor :" + txtAnchor + Environment.NewLine);
                specification.Append("AutoSize :" + autosize + Environment.NewLine + "MarginLeft :" + mgleft + Environment.NewLine + "MarginRight :" + mgRight + Environment.NewLine + "MarginTop :" + mgtop + Environment.NewLine);
                specification.Append("MarginBottom :" + mgbottom + Environment.NewLine + "WordWrap :" + wWrap + Environment.NewLine + "DefaultText :" + dtext + Environment.NewLine + "FontName :" + fname + Environment.NewLine);
                specification.Append("Bold :" + txbold + Environment.NewLine + "Italics :" + txItalics + Environment.NewLine + "Underline :" + txUnderline + Environment.NewLine + "FontSize :" + txsize + Environment.NewLine);
                specification.Append("Shadow :" + shadow + Environment.NewLine + "ParagraphBullet :" + txbuttlet + Environment.NewLine + "ParagraphAlignment :" + txtAlig + Environment.NewLine + "ParagraphHangingPunctuation :" + hPun + Environment.NewLine);
                specification.Append("ParagraphSpaceBefore :" + psb + Environment.NewLine + "ParagraphSpaceAfter :" + psa + Environment.NewLine + "ParagraphSpaceWithin :" + psw + Environment.NewLine + "RulerLevel1FirstMargin :" + rlfm + Environment.NewLine);
                specification.Append("RulerLevel1LeftMargin :" + rllm);
                PPTAttribute.saveSpacification(specification, shp.Name);
            }

        }

        #region Format a Chart as per selected type
        public List<string> FindSelectedCharts()
        {
            List<string> selCharts = new List<string>();
            int subSelShp;
            int sld_num, num_shp;

            PowerPoint.Shape oSh;
            sld_num = ppApp.ActiveWindow.Selection.SlideRange.SlideNumber;
            try { num_shp = ppApp.ActiveWindow.Selection.ShapeRange.Count; }
            catch (Exception ex) { num_shp = 0; }
            try { subSelShp = ppApp.ActiveWindow.Selection.ChildShapeRange.Count; }
            catch (Exception ex) { subSelShp = 0; }
            if (subSelShp != 0)
            {
                for (int i = 1; i <= subSelShp; i++)
                {
                    oSh = ppApp.ActiveWindow.Selection.ChildShapeRange[i];
                    if (oSh.HasChart == MsoTriState.msoTrue) { selCharts.Add(oSh.Name); }
                }
            }
            else
            {
                for (int i = 1; i <= num_shp; i++)
                {
                    oSh = ppApp.ActiveWindow.Selection.ShapeRange[i];
                    if (oSh.HasChart == MsoTriState.msoTrue) { selCharts.Add(oSh.Name); }
                    if (oSh.Type == MsoShapeType.msoGroup)
                    {
                        for (int x = 1; x <= oSh.GroupItems.Count; x++)
                        {
                            if (oSh.GroupItems[x].Type == MsoShapeType.msoGroup) { }
                            else
                            {
                                if (oSh.GroupItems[x].HasChart == MsoTriState.msoTrue) { selCharts.Add(oSh.GroupItems[x].Name); }
                            } // -Close Else
                        }   // Close For loop
                    } // Close if - Check IF statement 
                } // Close for loop to chekc shape count
            } // Close else statement
            return (selCharts);
        }


        public void DefineChartColor(string tscColors)
        {

        }
        public void Format_ChartArea(PowerPoint.Chart myChart)
        {
            myChart.ChartArea.Border.LineStyle = PowerPoint.XlLineStyle.xlLineStyleNone;
            myChart.ChartArea.Border.ColorIndex = 0;

        }
        public void Format_Title(PowerPoint.Chart myChart)
        {
            try
            {
                myChart.HasTitle = true;
                myChart.ChartTitle.Font.Name = "Calibri";
                myChart.ChartTitle.Font.Bold = MsoTriState.msoTrue;
                myChart.ChartTitle.Font.Size = 11;
                myChart.ChartTitle.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                myChart.ChartTitle.Font.Strikethrough = MsoTriState.msoFalse;
                myChart.ChartTitle.Font.Superscript = MsoTriState.msoFalse;
                myChart.ChartTitle.Font.Subscript = MsoTriState.msoFalse;
                myChart.ChartTitle.Font.Shadow = MsoTriState.msoFalse;
                myChart.ChartTitle.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_Title");
            }
        }
        public void Format_YAxis1(PowerPoint.Chart myChart, bool hasYAxis, string chType)
        {
            //If chType = "Radar" Or chType = "Stock" Then hasYAxis = True  'Or ch3D = True
            try
            {
                Microsoft.Office.Interop.Graph.Axis axis;
                if (chType == "Pie" || chType == "Doughnut" || chType == "Surface") { return; }
                if (chType == "Radar" || chType == "Stock") { hasYAxis = true; }

                if (hasYAxis == true)
                {

                    //axis = (Microsoft.Office.Interop.Graph.Axis)myChart.Axes(XlAxisType.xlValue);
                    myChart.HasAxis[XlAxisType.xlValue] = true;
                    //myChart.HasAxis = true;
                    if (myChart.Axes(XlAxisType.xlValue).HasTitle == true)
                    {
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Name = "Calibri";
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Bold = MsoTriState.msoTrue;
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Size = 11;
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Strikethrough = MsoTriState.msoFalse;
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Superscript = MsoTriState.msoFalse;
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Subscript = MsoTriState.msoFalse;
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.OutlineFont = MsoTriState.msoFalse;
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Shadow = MsoTriState.msoFalse;
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
                        myChart.Axes(XlAxisType.xlValue).AxisTitle.Font.Background = PowerPoint.XlBackground.xlBackgroundAutomatic;
                    }
                    myChart.Axes(XlAxisType.xlValue).Border.LineStyle = PowerPoint.XlLineStyle.xlContinuous;              //15/02/2018
                    myChart.Axes(XlAxisType.xlValue).Border.Color = System.Drawing.Color.FromArgb(127, 127, 127).ToArgb();
                    myChart.Axes(XlAxisType.xlValue).Border.Weight = PowerPoint.XlBorderWeight.xlHairline;
                    myChart.Axes(XlAxisType.xlValue).MajorTickMark = PowerPoint.XlTickMark.xlTickMarkOutside;
                    myChart.Axes(XlAxisType.xlValue).TickLabelPosition = PowerPoint.XlTickLabelPosition.xlTickLabelPositionNextToAxis;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.AutoScaleFont = false;
                    //------
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Name = "Calibri";
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Bold = MsoTriState.msoTrue;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Size = 11;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Strikethrough = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Superscript = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Subscript = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.OutlineFont = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Shadow = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Bold = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
                    myChart.Axes(XlAxisType.xlValue).TickLabels.Font.Background = PowerPoint.XlBackground.xlBackgroundAutomatic;

                } // - Check HasAxis is True of Not 
                else if (hasYAxis == false)
                {
                    myChart.HasAxis[XlAxisType.xlValue] = false;
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_YAxis1");
            }

        }
        public void Format_YAxis2(PowerPoint.Chart myChart, bool hasYAxis, string chType)
        {
            try
            {
                bool secAxis = false;
                PowerPoint.SeriesCollection sc = myChart.SeriesCollection();
                if (chType == "Pie" || chType == "Doughnut" || chType == "Surface") { return; }
                if (chType == "Radar" || chType == "Stock" || hasYAxis == true)
                {
                    for (int i = 1; i <= sc.Count; i++)
                    {
                        if (myChart.SeriesCollection(i).AxisGroup == 2) { secAxis = true; break; }       //XlAxisGroup.xlSecondary
                    }
                    if (secAxis == true)
                    {
                        if (myChart.HasAxis == true)
                        {
                            myChart.HasAxis[XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary] = true;
                            if (myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).HasTitle == true)
                            {
                                PowerPoint.Font fontObj = myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font;
                                fontObj.Name = "Calibri";
                                fontObj.Bold = MsoTriState.msoTrue;
                                fontObj.Size = 11;
                                myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                                myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.Strikethrough = MsoTriState.msoFalse;
                                fontObj.Superscript = MsoTriState.msoFalse;
                                fontObj.Subscript = MsoTriState.msoFalse;
                                myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.OutlineFont = MsoTriState.msoFalse;
                                fontObj.Shadow = MsoTriState.msoFalse;
                                myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
                                myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.Background = PowerPoint.XlBackground.xlBackgroundAutomatic;

                            }
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).Border.LineStyle = PowerPoint.XlLineStyle.xlContinuous;
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).Border.Color = System.Drawing.Color.FromArgb(127, 127, 127).ToArgb();
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).Border.Weight = PowerPoint.XlBorderWeight.xlHairline;
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).MajorTickMark = PowerPoint.XlTickMark.xlTickMarkOutside;
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).TickLabelPosition = PowerPoint.XlTickLabelPosition.xlTickLabelPositionNextToAxis;
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).TickLabels.AutoScaleFont = false;
                            // ------

                            PowerPoint.Font fontObjTB = myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).TickLabels.Font;
                            fontObjTB.Name = "Calibri";
                            fontObjTB.Bold = MsoTriState.msoTrue;
                            fontObjTB.Size = 11;
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.Strikethrough = MsoTriState.msoFalse;
                            fontObjTB.Superscript = MsoTriState.msoFalse;
                            fontObjTB.Subscript = MsoTriState.msoFalse;
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.OutlineFont = MsoTriState.msoFalse;
                            fontObjTB.Shadow = MsoTriState.msoFalse;
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
                            myChart.Axes(XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary).AxisTitle.Font.Background = PowerPoint.XlBackground.xlBackgroundAutomatic;
                        } // Check hasAxis is true
                        else if (hasYAxis == false)
                        {
                            myChart.HasAxis[XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary] = false;
                        }

                    }

                }

            } // Close Main if Check the chart type
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_YAxis2");
            }
        }
        //'This function formats the y-grids line style, tick marks, title and ticklables font name, style, size, color, sub-super script, underline
        public void Format_YGrids(PowerPoint.Chart myChart, bool hasYAxis, string chType, bool ch3D)
        {
            try
            {
                bool hasYGrids = false;
                if (chType == "Pie" || chType == "Doughnut") { return; }
                if (chType == "Radar" || (hasYAxis && ch3D == true)) { hasYGrids = true; }
                if (hasYGrids == true)
                {
                    myChart.Axes(XlAxisType.xlValue).HasMajorGridlines = true;
                    myChart.Axes(XlAxisType.xlValue).MajorGridlines.Border.Color = System.Drawing.Color.FromArgb(127, 127, 127).ToArgb();
                    myChart.Axes(XlAxisType.xlValue).MajorGridlines.Border.LineStyle = PowerPoint.XlLineStyle.xlContinuous;

                }
                else if (hasYGrids == false)
                {
                    myChart.Axes(XlAxisType.xlValue).HasMajorGridlines = false;
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_YGrids");
            }
        }
        public void Format_XAxis(PowerPoint.Chart myChart, bool hasXAxis, string chType)
        {
            try
            {
                if (chType == "Pie" || chType == "Doughnut") { return; }
                if (chType == "Radar")
                {
                    PowerPoint.Font RALFont = myChart.ChartGroups(1).RadarAxisLabels.Font;
                    RALFont.Name = "Calibri";
                    RALFont.Bold = MsoTriState.msoFalse;
                    RALFont.Size = 11;
                    myChart.ChartGroups(1).RadarAxisLabels.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                    myChart.ChartGroups(1).RadarAxisLabels.Font.Strikethrough = MsoTriState.msoFalse;
                    RALFont.Superscript = MsoTriState.msoFalse;
                    RALFont.Subscript = MsoTriState.msoFalse;
                    myChart.ChartGroups(1).RadarAxisLabels.Font.OutlineFont = MsoTriState.msoFalse;
                    RALFont.Shadow = MsoTriState.msoFalse;
                    myChart.ChartGroups(1).RadarAxisLabels.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
                    myChart.ChartGroups(1).RadarAxisLabels.Font.Background = PowerPoint.XlBackground.xlBackgroundAutomatic;
                    //System.Environment.Exit(0);
                }
                if (hasXAxis == true)
                {
                    if (myChart.Axes(XlAxisType.xlCategory).HasTitle == true)
                    {
                        PowerPoint.Font AxTitleFont = myChart.Axes(XlAxisType.xlCategory).AxisTitle.Font;
                        AxTitleFont.Name = "Calibri";
                        AxTitleFont.Bold = MsoTriState.msoFalse;
                        AxTitleFont.Size = 11;
                        myChart.Axes(XlAxisType.xlCategory).AxisTitle.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                        myChart.Axes(XlAxisType.xlCategory).AxisTitle.Font.Strikethrough = MsoTriState.msoFalse;
                        AxTitleFont.Superscript = MsoTriState.msoFalse;
                        AxTitleFont.Subscript = MsoTriState.msoFalse;
                        myChart.ChartGroups(1).RadarAxisLabels.Font.OutlineFont = MsoTriState.msoFalse;
                        AxTitleFont.Shadow = MsoTriState.msoFalse;
                        myChart.Axes(XlAxisType.xlCategory).AxisTitle.Font.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
                        myChart.Axes(XlAxisType.xlCategory).AxisTitle.Font.Font.Background = PowerPoint.XlBackground.xlBackgroundAutomatic;
                    }
                    myChart.Axes(XlAxisType.xlCategory).Border.LineStyle = PowerPoint.XlLineStyle.xlContinuous;
                    myChart.Axes(XlAxisType.xlCategory).Border.Color = System.Drawing.Color.FromArgb(127, 127, 127).ToArgb();
                    myChart.Axes(XlAxisType.xlCategory).Border.Weight = PowerPoint.XlBorderWeight.xlHairline;
                    myChart.Axes(XlAxisType.xlCategory).MajorTickMark = PowerPoint.XlTickMark.xlTickMarkOutside;
                    myChart.Axes(XlAxisType.xlCategory).TickLabelPosition = PowerPoint.XlTickLabelPosition.xlTickLabelPositionNextToAxis;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.AutoScaleFont = false;
                    //--- 
                    //PowerPoint.Font TickLabelsFont = myChart.Axes(XlAxisType.xlCategory).TickLabels.Font;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Name = "Calibri";
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Bold = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Size = 11;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Strikethrough = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Superscript = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Subscript = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.OutlineFont = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Shadow = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
                    myChart.Axes(XlAxisType.xlCategory).TickLabels.Font.Background = PowerPoint.XlBackground.xlBackgroundAutomatic;


                }
                else if (hasXAxis == false)
                {
                    myChart.HasAxis[XlAxisType.xlCategory] = false;
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_XAxis");
            }
        }
        public void Format_XGrids(PowerPoint.Chart myChart, bool hasXGrids, string chType)
        {
            try
            {
                if (chType == "Pie" || chType == "Doughnut")
                {
                    if (hasXGrids == true)
                    {
                        if (myChart.Axes(XlAxisType.xlCategory).HasMajorGridlines == true)
                        {
                            myChart.Axes(XlAxisType.xlCategory).MajorGridlines.Border.Color = System.Drawing.Color.FromArgb(127, 127, 127).ToArgb();
                            myChart.Axes(XlAxisType.xlCategory).MajorGridlines.Border.LineStyle = PowerPoint.XlLineStyle.xlDot;

                        }
                    }
                    else if (hasXGrids == false)
                    {
                        myChart.HasAxis[XlAxisType.xlCategory] = false;
                    }
                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_XGrids");
            }

        }
        public void Format_XSeries(PowerPoint.Chart myChart, bool hasYAxis)
        {
            //if (myChart.Axes(XlAxisType.xlValue).HasTitle == true)
            try
            {
                if (myChart.HasAxis[XlAxisType.xlSeriesAxis] == true)
                {
                    myChart.HasAxis(XlAxisType.xlSeriesAxis).Border.LineStyle = PowerPoint.XlLineStyle.xlContinuous;
                    myChart.HasAxis(XlAxisType.xlSeriesAxis).Border.Color = System.Drawing.Color.FromArgb(127, 127, 127).ToArgb();
                    myChart.HasAxis(XlAxisType.xlSeriesAxis).Border.Weight = PowerPoint.XlBorderWeight.xlHairline;

                    myChart.HasAxis(XlAxisType.xlSeriesAxis).Border.MajorTickMark = PowerPoint.XlTickMark.xlTickMarkOutside;
                    myChart.HasAxis(XlAxisType.xlSeriesAxis).Border.MinorTickMark = PowerPoint.XlTickMark.xlTickMarkNone;
                    myChart.HasAxis(XlAxisType.xlSeriesAxis).Border.TickLabelPosition = PowerPoint.XlTickLabelPosition.xlTickLabelPositionNextToAxis; ;
                    myChart.HasAxis(XlAxisType.xlSeriesAxis).Border.AutoScaleFont = false;

                    PowerPoint.Font TickLabelsFont = myChart.Axes(XlAxisType.xlSeriesAxis).TickLabels.Font;
                    TickLabelsFont.Name = "Calibri";
                    TickLabelsFont.Bold = MsoTriState.msoFalse;
                    TickLabelsFont.Size = 11;
                    myChart.Axes(XlAxisType.xlSeriesAxis).TickLabels.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                    myChart.Axes(XlAxisType.xlSeriesAxis).TickLabels.Font.Strikethrough = MsoTriState.msoFalse;
                    TickLabelsFont.Superscript = MsoTriState.msoFalse;
                    TickLabelsFont.Subscript = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlSeriesAxis).TickLabels.Font.OutlineFont = MsoTriState.msoFalse;
                    TickLabelsFont.Shadow = MsoTriState.msoFalse;
                    myChart.Axes(XlAxisType.xlSeriesAxis).TickLabels.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
                    myChart.Axes(XlAxisType.xlSeriesAxis).TickLabels.Font.Background = PowerPoint.XlBackground.xlBackgroundAutomatic;

                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_XSeries");
            }
        }
        public void Format_Legend(PowerPoint.Chart myChart, bool hasLegend)
        {
            int Num;
            try
            {
                if (hasLegend == true)
                {
                    if (myChart.HasLegend == true)
                    {
                        PowerPoint.LegendEntries ln = myChart.Legend.LegendEntries();
                        Num = ln.Count;
                    }
                    else
                    {
                        PowerPoint.SeriesCollection sc = myChart.SeriesCollection();
                        Num = sc.Count;
                    }
                    if (Num > 1)
                    {
                        if (myChart.HasLegend == false) { myChart.HasLegend = true; }
                        myChart.Legend.Font.Name = "Calibri";
                        myChart.Legend.Font.Size = 11;
                        myChart.Legend.Font.Bold = MsoTriState.msoFalse;
                        myChart.Legend.Position = PowerPoint.XlLegendPosition.xlLegendPositionBottom;

                    }
                    else if (hasLegend == false)
                    {
                        if (myChart.HasLegend == true) { myChart.Legend.Delete(); }
                    }

                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_Legend");
            }
        }

        public void Format_PlotArea(PowerPoint.Chart myChart, string chType)
        {
            try
            {
                myChart.PlotArea.Border.LineStyle = PowerPoint.XlLineStyle.xlLineStyleNone;
                myChart.PlotArea.Interior.ColorIndex = PowerPoint.XlColorIndex.xlColorIndexAutomatic;
                if (chType == "Pie" || chType == "Doughnut") { }
                else
                {
                    myChart.PlotArea.Left = 8;
                    if (myChart.HasTitle == true) { myChart.PlotArea.Top = myChart.ChartArea.Top + myChart.ChartTitle.Top; }
                    else { myChart.PlotArea.Top = 8; }
                    myChart.PlotArea.Width = myChart.ChartArea.Width - 24;
                    if (myChart.HasLegend == true) { myChart.PlotArea.Height = myChart.ChartArea.Height - myChart.Legend.Height * 2; }
                    else { myChart.PlotArea.Height = myChart.ChartArea.Height - 16; }
                }

            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_PlotArea");
            }
        }

        public void Format_Series(PowerPoint.Chart myChart, bool hasYAxis, string chType)
        {
            try
            {
                //bool secAxis = false;
                PowerPoint.SeriesCollection sc = myChart.SeriesCollection();
                if (chType == "Pie" || chType == "Doughnut" || chType == "Surface") { return; }
                if (chType == "Radar" || chType == "Stock" || hasYAxis == true)
                {
                    for (int i = 1; i <= sc.Count; i++)
                    {
                        PowerPoint.Series chartSeries = myChart.SeriesCollection(i);
                        chartSeries.Border.Color= System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                        myChart.SeriesCollection(i).Format.Line.Visible = MsoTriState.msoTrue;
                        myChart.SeriesCollection(i).Format.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                        myChart.SeriesCollection(i).MarkerBackgroundColor = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                    }

                }
            }
            catch (Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "Format_PlotArea");
            }
            
        }
        #endregion
    }
    // Close a Class statement -

}


