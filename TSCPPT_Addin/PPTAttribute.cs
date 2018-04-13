using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace TSCPPT_Addin
{
    public static class PPTAttribute
    {
         
        public static string basePath = System.AppDomain.CurrentDomain.BaseDirectory;
        public static string dbPath = basePath + @"AppData\Mapping\PPT_Specification.xlsx";
        public static string WordList = basePath + @"AppData\Mapping\WordList.xlsx";
        public static string mPPTPath = basePath + @"AppData\Template\Template_Automation.pptx";
        public static string standardppt = basePath + @"AppData\Template\Standard.pptx";
        public static string logfile = basePath + @"AppData\Logfile\";
        public static string supportfile = basePath + @"AppData\Mapping\support.txt";
        public static string txtbox= basePath+ @"AppData\Template\txtplaceholder.pptx";
       
        public static string PiCon = basePath + @"AppData\Image\Icon\";
        public static string themeColor = basePath + @"AppData\Image\ThemeColor\";
        public static string Tabletheme = basePath + @"AppData\Image\Tables\";
        //Standard
        public static PowerPoint.Application ppApp = Globals.ThisAddIn.Application;
        public static PowerPoint.Presentation ActivePPT;
        public static string msgTitle = "The Smart Cube";
        public static string Error;
        public static bool exitFlag = false;
        public static int ErrIndex;
        public static bool reviewExitFlag = false;
        public static bool discardFlag = false;
        static public void UserTracker(Office.IRibbonControl rib)
        {
            String strUserName;
            String strWholeText;
            try
            {
                string path = logfile + "UserLog_" + DateTime.Now.ToString("ddMMyyyy") + ".txt";
                strUserName = Environment.UserName;

                strWholeText = strUserName + "|" + getButtonDiscription(rib.Id) + "|" + DateTime.Now.ToString();

                if (!File.Exists(path))
                {
                    //File.Create(path);
                    using (StreamWriter sw = File.CreateText(path))
                    {
                        sw.WriteLine(strWholeText);
                        sw.Close();
                    }

                }
                else if (File.Exists(path))
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine(strWholeText);
                        sw.Close();
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "UserTracker");
                //MessageBox.Show("Error in errorlog After Apply Theme");
            }
        }
        public static void ErrorLog(string ErrDis, string rib)
        {
            String strUserName;
            String strWholeText;
            try
            {
                string path = logfile + "ErrorLog_" + DateTime.Now.ToString("ddMMyyyy") + ".txt";
                strUserName = Environment.UserName;
                if (rib.Length <= 12) { return; }
                if (rib.Substring(0, 12) == "customButton")
                {
                    strWholeText = strUserName + "|" + getButtonDiscription(rib) + "|" + DateTime.Now.ToString() + "|" + ErrDis;
                }
                else
                {
                    strWholeText = strUserName + "|" + rib + "|" + DateTime.Now.ToString() + "|" + ErrDis;
                }
                if (!File.Exists(path))
                {
                    //File.Create(path);
                    using (StreamWriter sw = File.CreateText(path))
                    {
                        sw.WriteLine(strWholeText);
                        sw.Close();
                    }

                }
                else if (File.Exists(path))
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine(strWholeText);
                        sw.Close();
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "ErrorLog");
            }
        }

        public static string getButtonDiscription(string ribID)
        {
            string btndiscription=null;
            DataTable dt = new DataTable();
            string cnText = @"Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";Data Source=" + PPTAttribute.dbPath;
            string qText = "select * from [btnDiscription$A1:B100] where button_id='" + ribID + "'";
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter(qText, cnText);
                da.Fill(dt);
                btndiscription = Convert.ToString(dt.Rows[0]["btnDiscription"]);
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "getButtonDiscription");
            }
            return (btndiscription);
        }
        public static  void saveSpacification(StringBuilder str,string shpName)
        {

            try
            {
                string path = logfile + "Spacification_" + shpName + ".txt";
                if (!File.Exists(path))
                {
                    using (StreamWriter sw = File.CreateText(path))
                    {
                        sw.WriteLine(str);
                        sw.Close();
                    }

                }
                else if (File.Exists(path))
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine(str);
                        sw.Close();
                    }
                }
            }
            catch(Exception err)
            {
                string errtext = err.Message;
                PPTAttribute.ErrorLog(errtext, "saveSpacification");
            }
        }

       
        public static void SQLConnection()
        {
            string connetionString = null;
            SqlConnection cnn;
            char splitChar = '|';
            connetionString = "Data Source=172.24.2.115; Initial Catalog=PPTAddin; User ID=SMAUser; Password=SMA@2017;";
            //connetionString = "Data Source=172.22.0.16; Initial Catalog=PPTAddin; User ID=SMAUser; Password=ADMIN@1234;";
            DirectoryInfo dir = new DirectoryInfo(logfile);
            cnn = new SqlConnection(connetionString);
            try
            {   
                
                foreach (FileInfo flInfo in dir.GetFiles())
                {
                    string fPath = logfile +@"\"+flInfo.Name;
                    string fname = flInfo.Name.Substring(0,flInfo.Name.IndexOf("_"));
                    string lname= flInfo.Name.Substring(flInfo.Name.IndexOf("_")+1, (Convert.ToInt32(flInfo.Name.Length) - flInfo.Name.IndexOf("_"))-5);
                    string cDate = DateTime.Now.ToString("dd")+ DateTime.Now.ToString("MM")+ DateTime.Now.ToString("yyyy");  
                    if (fname == "UserLog" && lname!= cDate)
                    {
                        
                        //MessageBox.Show(flInfo.Name);
                        foreach (string line in File.ReadLines(logfile+ flInfo.Name))
                        {
                            //MessageBox.Show(line);
                            List<String> dbValue = line.Split(splitChar).ToList();
                            SqlCommand cmd = new SqlCommand("sp_insert", cnn);
                            cmd.CommandType = CommandType.StoredProcedure;
                            int count = 1;
                            //string username = "Devendra";
                            //string task = "Load Theme";
                            //DateTime time = DateTime.Now;
                            cmd.Parameters.AddWithValue("@Sr_No", count);
                            cmd.Parameters.AddWithValue("@UName", dbValue[0]);
                            cmd.Parameters.AddWithValue("@PPTAction", dbValue[1]);
                            cmd.Parameters.AddWithValue("@ActionTime", Convert.ToDateTime(dbValue[2]));
                            cnn.Open();
                            int i = cmd.ExecuteNonQuery();
                            cnn.Close();
                            cmd.Dispose();
                        }
                        File.Delete(fPath);
                    }
                }
                //if (i != 0)
                //{
                //    MessageBox.Show(i + "Data Saved");
                //}
                //cnn.Open();
                
                //cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Can not open connection ! ");
            }
        }

        
    }
}
