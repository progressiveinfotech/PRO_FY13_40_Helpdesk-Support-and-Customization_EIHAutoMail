using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Data.SqlClient;
using System.Timers;
using System.Net.Mail;
using System.IO;
using System.Runtime.InteropServices;


namespace WinServ
{
    public partial class HelpDeskMailReader : ServiceBase
    {
        //Timer timer = new Timer();
        public HelpDeskMailReader()
        {
            InitializeComponent();
        }
        public void read()
        {
            Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace ns = outlook.GetNamespace("Mapi");
            object _missing = Type.Missing;
            ns.Logon(_missing, _missing, false, true);
            Microsoft.Office.Interop.Outlook.MAPIFolder inbox = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            string foldername = inbox.Name;
            int iMailCount = inbox.Items.Count;
            TraceService("Inside read method");
           
            for(int iCount =1 ;iCount <= iMailCount;iCount++)
            {
                Object mail1 = inbox.Items[iCount];
                if (((mail1 as Microsoft.Office.Interop.Outlook.MailItem) != null) && ((mail1 as Microsoft.Office.Interop.Outlook.MailItem).UnRead == true))
                {
                    TraceService("Inside unread mail");
                    Microsoft.Office.Interop.Outlook.MailItem mail = (Microsoft.Office.Interop.Outlook.MailItem)inbox.Items[iCount];
                    SqlConnection con;
                    string connection = @"server=CSM-DEV\SQL2008;database=TerexBest;uid=sa;pwd=rimc@123";
                    TraceService("Connection with database done1");
                    con = new SqlConnection(connection);
                    SqlCommand cmd;
                    cmd = new SqlCommand();
                    con.Open();
                        
                    cmd.Connection = con;
                    TraceService("Connection assigned to sql command");

                    string subject = mail.Subject.ToString();
                    TraceService("mail subject written"+subject);
                    string body = mail.Body.ToString();
                    cmd.Parameters.Add("@subject", SqlDbType.NVarChar).Value = subject;

                    TraceService(subject); //writing subject

                    cmd.Parameters.Add("@body", SqlDbType.NVarChar).Value = body;

                    TraceService(body); //writing subject

                    cmd.Parameters.Add("@recievedtime", SqlDbType.DateTime).Value = mail.ReceivedTime;

                    TraceService(mail.ReceivedTime.ToString()); //writing subject

                    cmd.Parameters.Add("@mailfrom", SqlDbType.NVarChar).Value = mail.SenderEmailAddress;

                    TraceService(mail.SenderEmailAddress); //writing subject
                    TraceService("Before Inventory saved");

                    cmd.CommandText = "insert into storemail(subject,body,createddatetime,mailfrom,isActive) values(@subject,@body,@recievedtime,@mailfrom,1)";
                    TraceService("Inventory saved");
                    try
                    {
                        cmd.ExecuteNonQuery();
                        mail.Delete();
                        iMailCount = iMailCount - 1;
                    }
                    catch (SqlException ex)
                    {
                        ex.ToString();
                    }
                        con.Close();
                        cmd.Dispose();
                        con.Dispose();
                    
               }
            }

            GC.Collect();
            inbox = null;
            ns = null;
            outlook = null;

        }
        protected override void OnStart(string[] args)
        {
            TraceService("start service");
            read();
            TraceService("out of on start");
            //timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            //timer.Interval = 300000;  //60,000=1 minute
            //timer.Enabled = true;
        }
        protected override void OnStop()
        {
            //timer.Enabled = false;
            TraceService("stopping service");
        }
        //private void OnElapsedTime(object source, ElapsedEventArgs e)
        //{
        //    TraceService("Another entry at " + DateTime.Now);
        //    startservice();
        //    read();
        //}
        private void TraceService(string content)
        {
            FileStream fs = new FileStream(@"d:\ScheduledService.txt", FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.BaseStream.Seek(0, SeekOrigin.End);
            sw.WriteLine(content);
            sw.Flush();
            sw.Close();
        }
       
    }
}
