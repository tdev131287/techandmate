using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;

namespace TSCPPT_Addin
{
    public partial class frmtable : Form
    {
        PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        PowerPoint.Presentation ActivePPT = Globals.ThisAddIn.Application.ActivePresentation;
        public frmtable()
        {
            InitializeComponent();
        }
        //- Click on this button insert table in Current slide
        private void button2_Click(object sender, EventArgs e)
        {
            int lnRow = Convert.ToInt32(nud_row.Value);
            int lnCol = Convert.ToInt32(nud_Col.Value);
            string themeType = cmbTheme.Text;
            PowerPoint.Table tbl;
            tbl = ActivePPT.Slides[1].Shapes.AddTable(lnRow, lnCol).Table;
            
            switch (themeType)
            {
                case "Table Theme1": { tbl.ApplyStyle("5C22544A-7EE6-4342-B048-85BDC9FD1C3A"); break; }
                case "Table Theme2": { tbl.ApplyStyle("21E4AEA4-8DFA-4A89-87EB-49C32662AFE0"); break; }
                case "Table Theme3": { tbl.ApplyStyle("F5AB1C69-6EDB-4FF4-983F-18BD219EF322"); break; }
                case "Table Theme4": { tbl.ApplyStyle("93296810-A885-4BE3-A3E7-6D5BEEA58F35"); break; }
                case "Table Theme5": { tbl.ApplyStyle("5C22544A-7EE6-4342-B048-85BDC9FD1C3A"); break; }
                case "Table Theme6": { tbl.ApplyStyle("21E4AEA4-8DFA-4A89-87EB-49C32662AFE0"); break; }
                case "Table Theme7": { tbl.ApplyStyle("F5AB1C69-6EDB-4FF4-983F-18BD219EF322"); break; }
                case "Table Theme8": { tbl.ApplyStyle("93296810-A885-4BE3-A3E7-6D5BEEA58F35"); break; }
            }
            for (int tr = 1; tr <= lnRow; tr++)
            {
                for (int tc = 1; tc <= lnCol; tc++)
                {
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderRight].ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderLeft].ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                    //-- Table border none --
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderRight].Visible = MsoTriState.msoFalse;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = 0;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderTop].Visible = MsoTriState.msoFalse;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = 0;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderLeft].Visible = MsoTriState.msoFalse;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = 0;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].Visible = MsoTriState.msoFalse;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 0;

                    tbl.Rows[tr].Cells[tc].Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    //tbl.Rows[tr].Cells[tc].Shape.TextFrame.MarginLeft = 3;
                    //tbl.Rows[tr].Cells[tc].Shape.TextFrame.MarginRight = 3;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = 1;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = 1;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = 1;
                    tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 1;
                    //-- Set Text font Size
                    tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    if (tr != 1) { tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.Font.Size = 12; }
                }
            }
            for (int tc = 1; tc <= lnCol; tc++)
            {
                tbl.Rows[1].Cells[tc].Shape.TextFrame.TextRange.Text = "Heading";
                tbl.Rows[1].Cells[tc].Shape.TextFrame.TextRange.Font.Size = 14;
            }
            for (int tc = 1; tc <= lnCol; tc++)
            {
                tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(102,114,0).ToArgb();
                tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = (float)1.5;
            }
            //tbl.ApplyStyle("5C22544A-7EE6-4342-B048-85BDC9FD1C3A");
            this.Close();
        }

        private void frmtable_Load(object sender, EventArgs e)
        {
            //- Insert Table theme 
            for(int x = 1; x <= 8; x++) { cmbTheme.Items.Add("Table Theme" + x); }
            cmbTheme.SelectedIndex = 0;
            // - Insert Table Rows/Columns Count
        }

        private void cmbTheme_SelectedIndexChanged(object sender, EventArgs e)
        {
            string imgName = cmbTheme.Text.Replace("Table ","");
            string tblpath = PPTAttribute.Tabletheme + imgName+".png";
            pictureBox1.ImageLocation = tblpath;
            

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
