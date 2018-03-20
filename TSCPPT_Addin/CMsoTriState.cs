using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Data;
using System;

namespace TSCPPT_Addin
{
    class CMsoTriState
    {
        char splitChar = ',';
        MsoTriState objValue;
        MsoVerticalAnchor objValue1;
        PowerPoint.PpAutoSize objValue2;
        PowerPoint.PpParagraphAlignment objValue3;
        PowerPoint.PpChangeCase objValue4;
        PowerPoint.PpBulletType objValue5;
        MsoTextOrientation objValue6;
        public List<int> get_RGBValue(string rgbValue)
        {
            List<string> rgbIndex = new List<string>();
            List<int> finalrgbIndex = new List<int>();
            try
            {
                rgbIndex = rgbValue.Split(splitChar).ToList();
                foreach (string item in rgbIndex)
                {
                    finalrgbIndex.Add(Convert.ToInt32(item));
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "get_RGBValue");
            }
            return (finalrgbIndex);
        }

        public MsoTriState getMsoTriState(int dbVal)
        {
            //int dbVal = Convert.ToInt32(dt.Rows[0]["FillVisible"]);
            try
            {
                if (dbVal == 1) { objValue = MsoTriState.msoCTrue; }
                else if (dbVal == 0) { objValue = MsoTriState.msoFalse; }
                else if (dbVal == -2) { objValue = MsoTriState.msoTriStateMixed; }
                else if (dbVal == -3) { objValue = MsoTriState.msoTriStateToggle; }
                else if (dbVal == -1) { objValue = MsoTriState.msoTrue; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "getMsoTriState");
            }
            return objValue;
        }
       
        public MsoVerticalAnchor getVerticalAnchor(int dbVal)
        {

            try
            {
                if (dbVal == 4) { objValue1 = MsoVerticalAnchor.msoAnchorBottom; }
                else if (dbVal == 5) { objValue1 = MsoVerticalAnchor.msoAnchorBottomBaseLine; }
                else if (dbVal == 3) { objValue1 = MsoVerticalAnchor.msoAnchorMiddle; }
                else if (dbVal == 1) { objValue1 = MsoVerticalAnchor.msoAnchorTop; }
                else if (dbVal == 2) { objValue1 = MsoVerticalAnchor.msoAnchorTopBaseline; }
                else if (dbVal == -2) { objValue1 = MsoVerticalAnchor.msoVerticalAnchorMixed; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "getVerticalAnchor");
            }
            return objValue1;
        }
        public PowerPoint.PpAutoSize TxtAutoSize(int dbVal)
        {
            try
            {
                if (dbVal == -2) { objValue2 = PowerPoint.PpAutoSize.ppAutoSizeMixed; }
                else if (dbVal == 0) { objValue2 = PowerPoint.PpAutoSize.ppAutoSizeNone; }
                else if (dbVal == 1) { objValue2 = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "TxtAutoSize");
            }
            return objValue2;
        }
        
        public PowerPoint.PpParagraphAlignment ParagraphFormatAlignment(int dbVal)
        {
            try
            {
                if (dbVal == 2) { objValue3 = PowerPoint.PpParagraphAlignment.ppAlignCenter; }
                else if (dbVal == 5) { objValue3 = PowerPoint.PpParagraphAlignment.ppAlignDistribute; }
                else if (dbVal == 4) { objValue3 = PowerPoint.PpParagraphAlignment.ppAlignJustify; }
                else if (dbVal == 7) { objValue3 = PowerPoint.PpParagraphAlignment.ppAlignJustifyLow; }
                else if (dbVal == 1) { objValue3 = PowerPoint.PpParagraphAlignment.ppAlignLeft; }
                else if (dbVal == -2) { objValue3 = PowerPoint.PpParagraphAlignment.ppAlignmentMixed; }
                else if (dbVal == 3) { objValue3 = PowerPoint.PpParagraphAlignment.ppAlignRight; }
                else if (dbVal == 6) { objValue3 = PowerPoint.PpParagraphAlignment.ppAlignThaiDistribute; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "getMsoTriState");
            }
            return objValue3;
        }
        

        public PowerPoint.PpChangeCase txtChangeCase(int dbVal)
        {
            try
            {
                if (dbVal == 2) { objValue4 = PowerPoint.PpChangeCase.ppCaseLower; }
                else if (dbVal == 1) { objValue4 = PowerPoint.PpChangeCase.ppCaseSentence; }
                else if (dbVal == 4) { objValue4 = PowerPoint.PpChangeCase.ppCaseTitle; }
                else if (dbVal == 5) { objValue4 = PowerPoint.PpChangeCase.ppCaseToggle; }
                else if (dbVal == 3) { objValue4 = PowerPoint.PpChangeCase.ppCaseUpper; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "txtChangeCase");
            }
            return objValue4;
        }
        public PowerPoint.PpBulletType getPpBulletType(int dbVal)
        {
            try
            {
                if (dbVal == 1) { objValue5 = PowerPoint.PpBulletType.ppBulletUnnumbered; }
                else if (dbVal == -2) { objValue5 = PowerPoint.PpBulletType.ppBulletMixed; }
                else if (dbVal == 0) { objValue5 = PowerPoint.PpBulletType.ppBulletNone; }
                else if (dbVal == 2) { objValue5 = PowerPoint.PpBulletType.ppBulletNumbered; }
                else if (dbVal == 3) { objValue5 = PowerPoint.PpBulletType.ppBulletPicture; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "getPpBulletType");
            }
            return objValue5;
        }
        public MsoTextOrientation getOrientation(int dVal)
        {
            try
            {
                if (dVal == 3) { objValue6 = MsoTextOrientation.msoTextOrientationDownward; }
                else if (dVal == 3) { objValue6 = MsoTextOrientation.msoTextOrientationDownward; }
                else if (dVal == 1) { objValue6 = MsoTextOrientation.msoTextOrientationHorizontal; }
                else if (dVal == 6) { objValue6 = MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast; }
                else if (dVal == -2) { objValue6 = MsoTextOrientation.msoTextOrientationMixed; }
                else if (dVal == 2) { objValue6 = MsoTextOrientation.msoTextOrientationUpward; }
                else if (dVal == 5) { objValue6 = MsoTextOrientation.msoTextOrientationVertical; }
                else if (dVal == 4) { objValue6 = MsoTextOrientation.msoTextOrientationVerticalFarEast; }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "getOrientation");
            }
            return objValue6;
        }
    }

}
