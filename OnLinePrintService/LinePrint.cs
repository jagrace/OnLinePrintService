using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Collections;
using System.IO;                        //filestream-getting database name
using System.Net;
using System.Net.Sockets;
using System.Threading;


namespace OnLinePrintService
{
    public partial class LinePrint : ServiceBase
    {
        private bool isStopped = false;
        private bool loggingEnabled;
        private string logName;
        private EventLog log;
        SqlConnection SQLConn;

        public void Start()
        {
            ServiceSettings settings = new ServiceSettings();
            string Server = settings.Server;
            string Database = settings.Database;
            string User = settings.User;
            string Pass = settings.Pass;
            string LineNumber = settings.LineNumber;

            string ConnectionString = "Data Source=" + Server + ";";
            ConnectionString += "user id=" + User + ";";
            ConnectionString += "password=" + Pass + ";";
            ConnectionString += "Initial Catalog=" + Database + ";MultipleActiveResultSets=true";
            SQLConn = new SqlConnection();
            SQLConn.ConnectionString = ConnectionString;
            SQLConn.Open();


            while (!this.isStopped)
            {
                try
                {
                    Thread.Sleep(1000);
                    MainLine(SQLConn, LineNumber);
                }
                catch (Exception Ex)
                {
                    if (SQLConn != null)
                    {
                        SQLConn.Close();
                        SQLConn.Dispose();
                    }
                    this.Log(Ex);
                }
            }
            SQLConn.Close();
            SQLConn.Dispose();
            this.Log(this.ServiceName + " Sql C&D");
        }
        public static void MainLine(SqlConnection SQLConnection, string prt_line)
        {
            Socket m_socClient;
            int port = 9100;
            string prt_address = "";
            string format = "";
            string item = "";
            string item10 = "";
            string img = "";
            string fam_desc = "";
            string fam_descraw = "";
            string prod_desc = "";
            string pcs = "";
            string length_ctn = "";
            string finish = "";
            string lsize = "";
            string color = "";
            string trademark = "";
            string fam_size = "";
            string metric_carton = "";
            string shift = "";
            string user = "";
            string plant = "";
            string batch = "";
            string workCenter = "";
            string monthDayYear = "";
            string peakIcon = "";
            string station = "";
            string ID = "";

            SqlDataReader prtreader = null;
            SqlCommand prtcommand = new SqlCommand("select * from toPrint" + prt_line, SQLConnection);
            prtreader = prtcommand.ExecuteReader();
            prtreader.Read();
            if (prtreader.HasRows == true)
            {
                ID = (prtreader["ID"].ToString());
                prt_address = (prtreader["printer"].ToString());
                item = (prtreader["item"].ToString());
                item10 = item.Substring(0, 10);
                img = (prtreader["image"].ToString());
                fam_descraw = (prtreader["fam_desc"].ToString());
                prod_desc = (prtreader["itm_desc"].ToString());
                pcs = (prtreader["pcs"].ToString());
                length_ctn = (prtreader["len_ctn"].ToString());
                finish = (prtreader["finish"].ToString());
                lsize = (prtreader["size"].ToString());
                color = (prtreader["color"].ToString());
                trademark = (prtreader["tmk"].ToString());
                fam_size = (prtreader["fam_size"].ToString());
                metric_carton = (prtreader["metric_ctn"].ToString());
                shift = (prtreader["shift"].ToString());
                user = (prtreader["user_num"].ToString());
                monthDayYear = (prtreader["date"].ToString());
                plant = (prtreader["plant"].ToString());
                batch = (prtreader["batch"].ToString());
                workCenter = (prtreader["line"].ToString());
                peakIcon = (prtreader["peak"].ToString());
                station = (prtreader["station"].ToString());
                char[] zeros = { '0' };
                pcs = pcs.TrimStart(zeros);
                length_ctn = length_ctn.TrimStart(zeros);
                metric_carton = metric_carton.TrimStart(zeros);
                //string tmrk = "¬FS¬FT,¬GSR,40¬FDA¬FS¬FT,¬A0R,41,47¬CI0¬FD";
                string tmrk = "^FS^FT,^GSR,40^FDA^FS^FT,^A0R,41,47^CI0^FD";
                fam_desc = fam_descraw.Replace("@", tmrk);
                format = "ZO" + lsize + "NNN" + peakIcon;
                prtreader.Close();
                prtreader.Dispose();
                prtcommand.Dispose();
                // bool jeff;
                bool socDis2 = true;
            retry:
                try
                {
                    m_socClient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                    char blanks = (' ');
                    String szIPSelected = prt_address.TrimEnd(blanks);
                    int alPort = port;
                    System.Net.IPAddress remoteIPAddress = System.Net.IPAddress.Parse(szIPSelected);
                    System.Net.IPEndPoint remoteEndPoint = new System.Net.IPEndPoint(remoteIPAddress, alPort);
                    m_socClient.Connect(remoteEndPoint);
                    SqlDataReader fmtreader = null;
                    SqlCommand fmtcommand = new SqlCommand("select * from LABEL_FORMATS where format='" + format + "'", SQLConnection);
                    string a = "";
                    string objData = "";
                    // Debugging Command - defines string for text file 
                    //string result="";

                    fmtreader = fmtcommand.ExecuteReader();
                    while (fmtreader.Read())
                    {
                        a = TrimString(fmtreader["lbldata"].ToString());
                        a = fmtreader["lbldata"].ToString();

                        // Use for Peel Off Printers
                        a = a.Replace("¬XA¬MMT¬FS¬XZ", "¬XA¬MMP¬FS¬XZ");

                        a = a.Replace("¬", "^");
                        a = a.Replace("~", "}");
                        a = a.Replace("}CC¬", "~CC^~CT}");
                        a = a.Replace(">", shift);
                        a = a.Replace("!18", user);
                        a = a.Replace("!16", monthDayYear);
                        a = a.Replace("!15", item10);//product_bar);
                        a = a.Replace("!11", metric_carton);
                        a = a.Replace("!10", plant);
                        a = a.Replace("!9", batch);
                        a = a.Replace("!8", fam_size);
                        a = a.Replace("!7", finish);
                        a = a.Replace("!6", length_ctn);
                        a = a.Replace("!5", pcs);
                        a = a.Replace("!4", prod_desc);
                        a = a.Replace("!3", fam_desc);
                        a = a.Replace("!2", img);
                        a = a.Replace("!1", workCenter);
                        a = a.Replace("( )", "(" + station + ")");
                        //                      objData += a;
                        objData = a;
                        byte[] byData = System.Text.Encoding.ASCII.GetBytes(objData.ToString());
                        m_socClient.Send(byData);

                        //Debugging Command - sets string for text file write 
                        //result = result + System.Text.Encoding.UTF8.GetString(byData);

                    }
                    //Debugging Command - Writes text file
                    //System.IO.File.WriteAllText(@"C:\ADC\LabelTest.txt", result);

                    fmtreader.Close();
                    fmtreader.Dispose();
                    fmtcommand.Dispose();
                    //objData += "^XA^IDR:*.ZPL^XZ"; // +objData;

                    //                    objData = "^XA^ZC2^JUS^XZ"  +objData;
                    //                    byte[] byData = System.Text.Encoding.ASCII.GetBytes(objData.ToString());
                    //                    m_socClient.Send(byData);
                    //m_socClient.Close();
                    m_socClient.Disconnect(socDis2);
                    SqlCommand dltcommand = new SqlCommand("delete from toPrint" + prt_line + " where ID=" + ID, SQLConnection);
                    int x = dltcommand.ExecuteNonQuery();
                    dltcommand.Dispose();
                }
                catch
                {
                    goto retry;
                }
            }
            else
            {
                prtreader.Close();
                prtreader.Dispose();
                prtcommand.Dispose();
            }
        }

        public static string TrimString(string str)
        {
            try
            {
                string pattern = @"^[ \t]+|[ \t]+$";
                Regex reg = new Regex(pattern, RegexOptions.IgnoreCase);
                str = reg.Replace(str, "");
                return str;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public LinePrint()
        {
            InitializeComponent();
            this.logName = ConfigurationSettings.AppSettings["LogName"];
            if (ConfigurationSettings.AppSettings["LoggingEnabled"] == "true")
                this.loggingEnabled = true;
            else
                this.loggingEnabled = false;
            if (this.loggingEnabled)
            {
                this.log = new EventLog();
                this.log.Source = this.logName;
            }
        }

        protected override void OnStart(string[] args)
        {
            this.isStopped = false;
            this.Log(this.ServiceName + " Started");
            Thread t = new Thread(new ThreadStart(this.Start));
            t.Start();
        }

        protected override void OnStop()
        {
            this.Log(this.ServiceName + " Stopped");
            this.isStopped = true;
        }






        private void Log(string message, EventLogEntryType type)
        {
            if (this.loggingEnabled)
                this.log.WriteEntry(message, type);
        }

        private void Log(string message)
        {
            this.Log(message, EventLogEntryType.Information);
        }

        private void Log(Exception e)
        {
            this.Log(e.Message + "\n" + e.StackTrace, EventLogEntryType.Error);
        }
    } //end class

}