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
    public partial class frmErrorlist : Form
    {
        bool chkflag = false;
        public frmErrorlist()
        {
            InitializeComponent();
        }

        private void frmErrorlist_Load(object sender, EventArgs e)
        {
            List<string> errList = new List<string>();
            char splitchar = '|';
            StreamReader sr = new StreamReader(PPTAttribute.supportfile);
            string line = sr.ReadLine();
            errList = line.Split(splitchar).ToList();
            lblErrortype.Text = "Format Review : Slide =" + errList[0];
            lbl_Error.Text = errList[1];
            chk_shpname.Text = errList[2];
            sr.Close();

            //- set the form position 
            this.Top = 50;
            this.Left = 10;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Formatshapes fobj = new Formatshapes();
            int sldNum;
            string shpname;
            sldNum = Convert.ToInt32(lblErrortype.Text.Replace("Format Review : Slide =", ""));
            shpname = chk_shpname.Text;
            fobj.CorrectFormat_Selected(sldNum, shpname);
            this.Close();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                this.Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void frmErrorlist_FormClosed(object sender, FormClosedEventArgs e)
        {
            //MessageBox.Show("Call Me");
           
        }

        private void btn_Exit_Click(object sender, EventArgs e)
        {
            PPTAttribute.exitFlag = true;
            this.Close();
        }
    }
}
