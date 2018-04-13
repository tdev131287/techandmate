using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TSCPPT_Addin
{
    public partial class frmReplacewith : Form
    {
        public frmReplacewith()
        {
            InitializeComponent();
        }

        private void frmReplacewith_Load(object sender, EventArgs e)
        {
            char splitChar = ',';
            List<string> errList = new List<string>();
            List<string> replaceWith = new List<string>();
            char splitchar = '|';
            StreamReader sr = new StreamReader(PPTAttribute.supportfile);
            string line = sr.ReadLine();
            errList = line.Split(splitchar).ToList();
            lbltype.Text = "Editorial Review : Slide =" + errList[0];
            txtError.Text = errList[1];
            replaceWith=errList[2].Split(splitChar).ToList();
            foreach (string rItem in replaceWith)
            {
                cmbReplace.Items.Add(rItem);
            }
            cmbReplace.SelectedIndex = 0;
            sr.Close();

            //- set the form position 
            this.Top = 50;
            this.Left = 10;
        }

        private void btn_Correct_Click(object sender, EventArgs e)
        {
            int Cindex = cmbReplace.SelectedIndex;
            PPTAttribute.ErrIndex = Cindex;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //PPTAttribute.reviewExitFlag = true;
            PPTAttribute.discardFlag = true;
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PPTAttribute.reviewExitFlag = true;
            this.Close();
        }
    }
}
