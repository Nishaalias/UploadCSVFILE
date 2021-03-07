using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using MySql.Data.MySqlClient;
using System.Configuration;
namespace Uploadcsvfile
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
        private string ConnectionString;
        public Service1()
        {
            InitializeComponent();
            ConnectionString = "Database=" + ConfigurationManager.AppSettings["Database"] + ";Data Source=" + ConfigurationManager.AppSettings["Data Source"].ToString() + ";UserId=" + ConfigurationManager.AppSettings["UserId"].ToString() + ";Password=" + ConfigurationManager.AppSettings["Password"] + ";Allow User Variables=True;Allow Zero Datetime=true;";

        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            ServiceUploadethod();
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 9000; //number in milisecinds  
            timer.Enabled = true;
        }
        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            ServiceUploadethod();
        }

        public void ServiceUploadethod()
        {
            try
            {

                DataTable dt = GetExcelsheetData();
                InsertExcelsheet(dt);
            }
            catch (Exception ex)
            {
                WriteToFile(ex.Message + ex.StackTrace + " Error in ServiceUploadMethod");

            }
        }

        public void InsertExcelsheet(DataTable dt)
        {
            string Command = string.Empty;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Command = "INSERT INTO `product`( region,country,itemtype,saleschannel,orderpriority,orderdate,orderid,shipdate,unitssold,Unitprice,Unitcost, Totalrevenue,Totalcost,Totalprofit) VALUES (" + dt.Rows[i][0].ToString() + ",'" + dt.Rows[i][1].ToString() + "','" + dt.Rows[i][2].ToString() + "','" + dt.Rows[i][3].ToString() + "','" + dt.Rows[i][4].ToString() + "','" + dt.Rows[i][5].ToString() + "','" + dt.Rows[i][6].ToString() + "' ,  '" + dt.Rows[i][7].ToString() + "' ," + dt.Rows[i][8].ToString() + "," + dt.Rows[i][9].ToString() + ", " + dt.Rows[i][10].ToString() + "," + dt.Rows[i][11].ToString() + "," + dt.Rows[i][12].ToString() + "," + dt.Rows[i][13].ToString() + ");";

                using (MySqlConnection mConnection = new MySqlConnection(ConnectionString))
                {
                    mConnection.Open();
                    using (MySqlCommand myCmd = new MySqlCommand(Command, mConnection))
                    {
                        myCmd.ExecuteNonQuery();

                    }
                }
            }

        }

       


        public DataTable GetExcelsheetData()
        {
            DataTable dt = new DataTable();

            string filepath=System.AppDomain.CurrentDomain.BaseDirectory + "10000 Sales Records.csv";
            using (StreamReader sr = new StreamReader(filepath))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }

                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    if (rows.Length > 1)
                    {
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i].Trim();
                        }
                        dt.Rows.Add(dr);
                    }
                }

            }


            return dt;
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




    }
}
