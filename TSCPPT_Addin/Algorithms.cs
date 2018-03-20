using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TSCPPT_Addin
{
    class Algorithms
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        public string DocumentLanguage()
        {
            string language = "US English";

            try
            {
                MsoLanguageID shp_lang, not_lang;
                Dictionary<MsoLanguageID, String> langDict = new Dictionary<MsoLanguageID, string>();
                MsoLanguageID doc_lang = ppApp.ActivePresentation.DefaultLanguageID;
                langDict.Add(doc_lang, "lang");
                //Set language in each textbox in each slide
                for (int sld = 1; sld <= ppApp.ActivePresentation.Slides.Count; sld++)
                {
                    foreach (PowerPoint.Shape shp in ppApp.ActivePresentation.Slides[sld].Shapes)
                    {
                        // '---------------- Check if it is a table
                        if (shp.Type == MsoShapeType.msoTable)
                        {
                            for (int r = 1; r <= shp.Table.Rows.Count; r++)
                            {
                                for (int c = 1; c <= shp.Table.Columns.Count; c++)
                                {
                                    shp_lang = shp.Table.Cell(r, c).Shape.TextFrame.TextRange.LanguageID;
                                    if (!langDict.ContainsKey(shp_lang)) { langDict.Add(shp_lang, "lang"); }
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
                                        shp_lang = shp.GroupItems[i].TextFrame.TextRange.LanguageID;
                                        if (!langDict.ContainsKey(shp_lang)) { langDict.Add(shp_lang, "lang"); }
                                    }
                                }
                            }
                        }
                        //'-------------- Check if it is a simple shape
                        if (shp.HasTextFrame == MsoTriState.msoTrue)
                        {
                            shp_lang = shp.TextFrame.TextRange.LanguageID;
                            if (!langDict.ContainsKey(shp_lang)) { langDict.Add(shp_lang, "lang"); }
                        }

                    }
                    not_lang = shp_lang = ActivePPT.Slides[sld].NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.LanguageID;
                    if (!langDict.ContainsKey(not_lang)) { langDict.Add(not_lang, "lang"); }
                }
                if (langDict.Count == 1)
                {
                    List<MsoLanguageID> dLang = new List<MsoLanguageID>();
                    dLang = langDict.Keys.ToList();
                    if (dLang[0] == MsoLanguageID.msoLanguageIDEnglishUS) { language = "US English"; }
                    if (dLang[0] == MsoLanguageID.msoLanguageIDEnglishUK) { language = "UK English"; }
                }
                else { language = "Mixed (US/UK)"; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "DocumentLanguage");
            }
            return (language);
        }

        public List<int> QuickSort()
        {
            List<int> selSlides = new List<int>();
            return (selSlides);
        }
        public void SetNamesUnique(int sldNum)
        {
            int shpCount = ActivePPT.Slides[sldNum].Shapes.Count;
            try
            {
                for (int i = 1; i <= shpCount; i++)
                {
                    PowerPoint.Shape osh = ActivePPT.Slides[sldNum].Shapes[i];
                    if (osh.Type == MsoShapeType.msoGroup)
                    {
                        for (int x = 1; x <= osh.GroupItems.Count; x++)
                        {
                            string shpName = osh.GroupItems[x].Name;
                            ChangeNames(sldNum, shpName, x);
                        }
                    }
                    else
                    {
                        string shpName = osh.Name;
                        ChangeNames(sldNum, shpName, i);
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "SetNamesUnique");
            }
        }
        public void ChangeNames(int sldNum, string shpName, int index)
        {

            try
            {
                for (int j = 1; j <= ActivePPT.Slides[sldNum].Shapes.Count; j++)
                {
                    PowerPoint.Shape osh1 = ActivePPT.Slides[sldNum].Shapes[j];
                    if (osh1.Type == MsoShapeType.msoGroup)
                    {
                        for (int y = 1; y <= osh1.GroupItems.Count; y++)
                        {
                            string tName = osh1.GroupItems[y].Name;
                            if (tName == shpName && y != index) { osh1.GroupItems[y].Name = osh1.GroupItems[y].Name + "_C"; }
                        }
                    }
                    else
                    {
                        string tName = osh1.Name;
                        if (tName == shpName && j != index) { osh1.Name = osh1.Name + "_C"; }
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "ChangeNames");
            }
        }
    }
}
