using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Timers;
using System.Configuration;

namespace ZTECDRConverterService
{
    public partial class Service1 : ServiceBase
    {
        
        System.Timers.Timer timer = new System.Timers.Timer();
        //private System.Threading.Thread _thread;
        
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                Thread MyThread = new Thread(new ThreadStart(ReadTxtFile));
                WriteToFile("Service is started at " + DateTime.Now);
                MyThread.Start();
                timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
                timer.Interval = 50000; //number in miliseconds 5000 = approx 1 min, 120000 = 2mins  //3600000 = 1hr, 1800000= 30mins
                //ReadTxtFile();
                timer.Enabled = true;
                //method to read .txt file to db
                //ReadTxtFile();
                // Create the thread object that will do the service's work.
               // _thread = new System.Threading.Thread(ReadTxtFile);

                // Start the thread.
               // _thread.Start();

                
                
                //using (StreamReader sr = new StreamReader(File.Open("C:\\Tr500.txt", FileMode.Open)))
                //    {
                //        using (SqlConnection txtbaglan = new SqlConnection(_dbCon))
                //        {
                //            txtbaglan.Open();
                //            string line = "";
                //            while ((line = sr.ReadLine()) != "")
                //            {
                //                string[] parts = line.Split(new string[] { "," }, StringSplitOptions.None);
                //                string cmdTxt = String.Format("INSERT INTO pdks(kod,il) VALUES ('{0}','{1}')", parts[0], parts[1]);//", parts[0], parts[1]);
                //                using (SqlCommand cmddd = new SqlCommand(cmdTxt, txtbaglan))
                //                {
                //                    cmddd.ExecuteNonQuery();
                //                }
                //            }
                //        }
                //    }
                // }
            }
            catch (Exception ex)
            {

                EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
            }
        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Service is recall at " + DateTime.Now);
        }
        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
        private void ReadTxtFile()
        {
            try
            {
                string _dbCon = ConfigurationManager.ConnectionStrings["myCon"].ConnectionString;
                using (SqlConnection _con = new SqlConnection(_dbCon))
                {

                    string source = "C:\\ZTECDRConverter\\OriginalFile";
                    // string Destination = "C:\\ZTECDRConverter\\Dumb";
                    string filename = string.Empty;
                    if (!(Directory.Exists(source)))
                        return;
                    string[] Templateexcelfile = Directory.GetFiles(source);
                    foreach (string file in Templateexcelfile)
                    {
                        if (Templateexcelfile[0].Contains("UN"))
                        {
                            filename = System.IO.Path.GetFileName(file);
                            _con.Open();
                            SqlCommand cmd = new SqlCommand();
                            cmd.Connection = _con;
                            cmd.CommandText = "Select Filenames from tblFiles where Filenames='" + filename.ToString().Trim() + "'";
                            SqlDataReader sdr = cmd.ExecuteReader();
                            if (sdr.Read())
                            {

                            }
                            else
                            {
                                _con.Close();
                                _con.Open();
                                using (StreamReader sr = new StreamReader(File.Open("C:\\ZTECDRConverter\\OriginalFile\\" + filename.ToString().Trim() + "", FileMode.Open)))
                                {
                                    string line = "";
                                    while ((line = sr.ReadLine()) != null)
                                    {
                                        string[] parts = line.Split(new string[] { " " }, StringSplitOptions.None);
                                        string cmdTxt = string.Format("INSERT INTO tblTestService(id,col1,col2,col3) VALUES ('{0}','{1}','{2}','{3}')", parts[0], parts[1], parts[2], parts[3]);//", parts[0], parts[1]);

                                        using (SqlCommand cmddd = new SqlCommand(cmdTxt, _con))
                                        {
                                            cmddd.ExecuteNonQuery();
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
            }
            
        }

        //copy file from original location to dumb folder
        //public static void Copy_Excelfile_And_Paste_at_anotherloaction()
        //{
        //    try
        //    {
        //        string source = "C:\\ZTECDRConverter\\OriginalFile";
        //        string Destination = "C:\\ZTECDRConverter\\Dumb";
        //        string filename = string.Empty;
        //        if (!(Directory.Exists(Destination) && Directory.Exists(source)))
        //            return;
        //        string[] Templateexcelfile = Directory.GetFiles(source);
        //        foreach (string file in Templateexcelfile)
        //        {
        //            if (Templateexcelfile[0].Contains("UN"))
        //            {
        //                filename = System.IO.Path.GetFileName(file);
        //                //Destination = System.IO.Path.Combine(Destination, filename.Replace(".xlsx", DateTime.Now.ToString("yyyyMMdd")) + ".xlsx");
        //                System.IO.File.Copy(file, Destination, true);
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //       //Create_ErrorFile(ex);
        //    }

        //}
    }
}
