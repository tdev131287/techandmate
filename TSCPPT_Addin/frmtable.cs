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
            int slIndex = ppApp.ActiveWindow.Selection.SlideRange.SlideIndex;
            int tCase =0;
            string themeType = PPTAttribute.tableType[PPTAttribute.tableType.Count - 1];
            PowerPoint.Table tbl;
            tbl = ActivePPT.Slides[slIndex].Shapes.AddTable(lnRow, lnCol).Table;
            
            switch (themeType)
            {
                case "Atable1": { tbl.ApplyStyle("5C22544A-7EE6-4342-B048-85BDC9FD1C3A"); tCase = 1; break; }
                case "Atable2": { tbl.ApplyStyle("21E4AEA4-8DFA-4A89-87EB-49C32662AFE0"); tCase = 2; break; }
                case "Atable3": { tbl.ApplyStyle("F5AB1C69-6EDB-4FF4-983F-18BD219EF322"); tCase = 3; break; }
                case "Atable4": { tbl.ApplyStyle("93296810-A885-4BE3-A3E7-6D5BEEA58F35"); tCase = 4; break; }
                case "Atable5": { tbl.ApplyStyle("5C22544A-7EE6-4342-B048-85BDC9FD1C3A"); tCase = 5; break; }
                case "Atable6": { tbl.ApplyStyle("21E4AEA4-8DFA-4A89-87EB-49C32662AFE0"); tCase = 6; break; }
                case "Atable7": { tbl.ApplyStyle("F5AB1C69-6EDB-4FF4-983F-18BD219EF322"); tCase = 7; break; }
                case "Atable8": { tbl.ApplyStyle("93296810-A885-4BE3-A3E7-6D5BEEA58F35"); tCase = 8; break; }
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


                    if (tr != 1)
                    {
                        tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.Font.Size = 12;
                        tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.Color.FromArgb(57, 42, 30).ToArgb();
                        tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore =6 ;
                        tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0;
                        tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = (float)0.9;
                    }
                    else
                    {
                        tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0;
                        tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0;
                        //tbl.Rows[tr].Cells[tc].Shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 0;
                    }
                    if((tr != 1) && (tCase==5 || tCase == 6|| tCase == 7|| tCase == 8))
                    {
                        tbl.Rows[tr].Cells[tc].Shape.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255).ToArgb();
                        tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderRight].ForeColor.RGB = System.Drawing.Color.FromArgb(227, 227, 227).ToArgb();
                        tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = System.Drawing.Color.FromArgb(227, 227, 227).ToArgb();
                        tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderLeft].ForeColor.RGB = System.Drawing.Color.FromArgb(227, 227, 227).ToArgb();
                        tbl.Rows[tr].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(227, 227, 227).ToArgb();
                    }
                }
            }
            for (int tc = 1; tc <= lnCol; tc++)
            {
                tbl.Rows[1].Cells[tc].Shape.TextFrame.TextRange.Text = "Heading";
                tbl.Rows[1].Cells[tc].Shape.TextFrame.TextRange.Font.Size = 14;
            }

            // Remove outer a Border Color --
            for (int tr = 1; tr <= lnRow; tr++)
            {
                tbl.Rows[tr].Cells[1].Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = 0;
                tbl.Rows[tr].Cells[1].Borders[PowerPoint.PpBorderType.ppBorderLeft].ForeColor.RGB = System.Drawing.Color.FromArgb(255,255,255).ToArgb();
                tbl.Rows[tr].Cells[lnCol].Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = 0;
                tbl.Rows[tr].Cells[lnCol].Borders[PowerPoint.PpBorderType.ppBorderRight].ForeColor.RGB = System.Drawing.Color.FromArgb(255,255,255).ToArgb();
            }
            for (int tc = 1; tc <= lnCol; tc++)
            {
                
                tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = (float)1.5;
                switch (tCase)
                {
                    case 1: { tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(102, 114, 0).ToArgb(); break; }
                    case 2: { tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(155, 174, 0).ToArgb(); break; }
                    case 3: { tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(97, 85, 75).ToArgb(); break; }
                    case 4: { tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(78, 204, 124).ToArgb(); break; }
                    case 5: { tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(102, 114, 0).ToArgb(); break; }
                    case 6: { tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(155, 174, 0).ToArgb(); break; }
                    case 7: { tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(97, 85, 75).ToArgb(); break; }
                    case 8: { tbl.Rows[lnRow].Cells[tc].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.Color.FromArgb(78, 204, 124).ToArgb(); break; }

                }
            }
            //tbl.ApplyStyle("5C22544A-7EE6-4342-B048-85BDC9FD1C3A");
            this.Close();
        }

        private void frmtable_Load(object sender, EventArgs e)
        {
            this.Top = 100;
            this.Left = 20;
        }

        private void cmbTheme_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmtable_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void frmtable_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }
    }
}
