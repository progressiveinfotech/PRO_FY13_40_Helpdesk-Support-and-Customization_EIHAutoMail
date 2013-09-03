
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Net.Mail;
using Microsoft.Office.Interop.Outlook;
using System.Web.Security;
using System.Configuration;
//using Microsoft.Exchange.WebServices.Data;
using System.Runtime.InteropServices;


namespace HelpDeskMailReader
{
    public class Program
    {
        static void Main(string[] args)
        {
            readMails rmails = new readMails();
            rmails.read();
        }
    }

   public class readMails
    {
       //public void ReadExchangeServer()
       //{
       //    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
       //    //NetWorkCredentialMailId = "lalit.joshi@progressive.in";
       //    //NetWorkCredentialPassword = "pipl?123";
       //    LogMessage("A");
       //    service.Credentials = new WebCredentials("v-helpdesk.central@oberoigroup.com", "password@123");
       //    service.AutodiscoverUrl("v-helpdesk.central@oberoigroup.com");
       //    LogMessage("B");    
       //    FindItemsResults<Item> findresults = service.FindItems(WellKnownFolderName.Inbox, new ItemView(10));
       //    LogMessage("C");    
       //    foreach (Item item in findresults.Items)
       //    {
       //        //Console.WriteLine(item.Subject);
       //        LogMessage(item.Subject);     
       //    }

       //}
        //Customer_mst objCustomer = new Customer_mst();
        //BLLCollection<Customer_mst> colCust = new BLLCollection<Customer_mst>();
        //CustomerToSiteMapping objCustToSite = new CustomerToSiteMapping();
        //Organization_mst objOrganization = new Organization_mst();
          
        SqlConnection con;

        //public string connection = @"Data Source=CSM-DEV\SQL2008;Initial Catalog=EIH_live;Persist Security Info=True;User ID=sa;Password=rimc@123;User Instance=False; Connect Timeout=110000";
        //public string smtphost = "10.1.0.12";
        //public string NetWorkCredentialMailId = "csm.admin@progessive.in";
        //public string NetWorkCredentialPassword = "pipl?1234";
        //public string SMTPPort = "25";
        //public string FromMailId = "csm.admin@progessive.in";

        //public string connection = @"Data Source=CCORPSRVHD02\SQLEXPRESS_NEW;Initial Catalog=Best;Persist Security Info=True;User ID=sa;Password=pass@123;User Instance=False; Connect Timeout=110000";
        //public string smtphost = "132.122.6.60";
        //public string NetWorkCredentialMailId = "v-helpdesk.central@oberoigroup.com";
        //public string NetWorkCredentialPassword = "password@123";
        //public string SMTPPort = "25";
        //public string FromMailId = "v-helpdesk.central@oberoigroup.com";

        public string connection = ConfigurationManager.ConnectionStrings["CSM_DB"].ConnectionString;
        public string smtphost = ConfigurationManager.AppSettings["SMTPHost"];
        public string NetWorkCredentialMailId = ConfigurationManager.AppSettings["NetWorkCredentialMailId"];
        public string NetWorkCredentialPassword = ConfigurationManager.AppSettings["NetWorkCredentialPassword"];
        public string SMTPPort = ConfigurationManager.AppSettings["SMTPPort"];
        public string FromMailId = ConfigurationManager.AppSettings["FromMailId"];

        public void LogMessage(string sMsg)
        {
            File.AppendAllText(@"C:\ScheduledService.log", sMsg + Environment.NewLine);
        }

        private string GetEmailAddressFromExchange(string Name)
        {
            string emailAddress = string.Empty;
            try
            {
                Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.NameSpace oNS = outlook.GetNamespace("Mapi");
                Recipient recip = oNS.CreateRecipient(Name);
                recip.Resolve();
                ExchangeUser exUser = recip.AddressEntry.GetExchangeUser();
                emailAddress = exUser.PrimarySmtpAddress;
                Marshal.ReleaseComObject(exUser);
                Marshal.ReleaseComObject(recip);
                Marshal.ReleaseComObject(oNS);
                LogMessage("Mail Id is:" + emailAddress);
            }
            catch (System.Exception ex) { };
            return emailAddress;

        }

        public void read()
        {
            try
            {
                LogMessage("Read Started");
                string _SLAId = "0";
                //Microsoft.Office.Interop.Outlook.AddressEntry SMTPAddress;
                Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.NameSpace ns = outlook.GetNamespace("Mapi");
                int i = ns.Folders.Count; 
                object _missing = Type.Missing;
                ns.Logon(_missing, _missing, false, true);
                //Microsoft.Office.Interop.Outlook.MAPIFolder inbox = ns.Folders[j];
                Microsoft.Office.Interop.Outlook.MAPIFolder inbox = ns.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
                string foldername = inbox.Name;
                int iMailCount = inbox.Items.Count;
                LogMessage("No Of MailCount " + iMailCount.ToString());
                for (int iCount = 1; iCount <= iMailCount; )
                {
                    LogMessage("Reading Mail =" + iCount.ToString());
                    Object mail1 = inbox.Items[iCount];
                    if (((mail1 as Microsoft.Office.Interop.Outlook.MailItem) != null) && ((mail1 as Microsoft.Office.Interop.Outlook.MailItem).UnRead == true))
                    {
                        LogMessage("Reading Mail =" + iCount.ToString() + " Unread Mail");
                        Microsoft.Office.Interop.Outlook.MailItem mail = (Microsoft.Office.Interop.Outlook.MailItem)inbox.Items[iCount];
                        //LogMessage("A");
                        //LogMessage("Sender Mail type:"+ mail.SenderEmailType);
                        con = new SqlConnection(connection);
                        LogMessage(connection);
                    
                        SqlCommand cmd;
                        SqlCommand cmd1;
                        SqlCommand cmd2;
                        cmd = new SqlCommand();
                        cmd1 = new SqlCommand();
                        cmd2 = new SqlCommand();
                        //LogMessage("before connection open");
                        con.Open();
                        LogMessage("Connection Open");
                        cmd.Connection = con;
                        cmd1.Connection = con;
                        cmd2.Connection = con;
                        string subject = "";
                        string body = "";
                        if (mail.Subject != null)
                        {
                            subject = mail.Subject.ToString();
                        }
                        if (mail.Body != null)
                        {
                            body = mail.Body.ToString();
                        }
                        string mailto = mail.To;
                        string mfrm = mail.SentOnBehalfOfName;
                        LogMessage("Mail To:"+ mailto);
                        LogMessage("Mail From:" + mfrm);
                        string senderemailadd = ((Microsoft.Office.Interop.Outlook.MailItem)inbox.Items[iMailCount]).SenderEmailAddress;
                        //LogMessage("Address :" + senderemailadd);
                        
                        //Microsoft.Office.Interop.Outlook.AddressEntry currentUser ;
                        //if (currentUser.Type == "EX")
                        //{
                        //    Microsoft.Office.Interop.Outlook.ExchangeUser manager =
                        //        currentUser.GetExchangeUser().GetExchangeUserManager();
                        //     Add recipient using display name, alias, or smtp address
                        //    mail.Recipients.Add(manager.PrimarySmtpAddress);
                        //    mail.Recipients.ResolveAll();
                        //    mail.Attachments.Add(@"c:\sales reports\fy06q4.xlsx",
                        //        Outlook.OlAttachmentType.olByValue, Type.Missing,
                        //        Type.Missing);
                        //    mail.Send();
                        //}

                        // Begin Gulshan
                        //String TestEmail = "";
                        //TestEmail = mail.SenderEmailAddress + " " + mail.Recipients.ToString() + " " + mail.SenderName;
                        //string t1 = mail.Recipients.ToString() + " " + mail.SenderName;
                        //string t2 = mail.SenderName;
                        // End Gulshan

                        //string[] s = t2.Split(' ');
                        //string mailid = s[0] + "." + s[1];
                        string loc = getlocation(mailto);
                        LogMessage("Location :'" + loc + "'");
                        ////////////////// vIshal
                        //if (mail.SenderEmailType == "EX")
                        //{
                        //    //Issue a reply on the mail message to create a recipient object that is the sender address.
                        //    Microsoft.Office.Interop.Outlook._MailItem Temp = ((Microsoft.Office.Interop.Outlook._MailItem)mail).Reply();
                        //    //Use the recipient object to access the smtp address of the exchange user
                        //    SMTPAddress = Temp.Recipients[1].AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                        //    Temp.Delete();
                        //}
                        //else
                        //{
                        //    SMTPAddress = mail.SenderEmailAddress;
                        //}
                        //Microsoft.Office.Interop.Outlook.AddressEntry currentUser ;
                        //if (currentUser.Type == "EX")
                        //{
                        //    Microsoft.Office.Interop.Outlook.ExchangeUser manager =
                        //        currentUser.GetExchangeUser().GetExchangeUserManager();
                        //     Add recipient using display name, alias, or smtp address
                        //    mail.Recipients.Add(manager.PrimarySmtpAddress);
                        //    mail.Recipients.ResolveAll();
                        //    mail.Attachments.Add(@"c:\sales reports\fy06q4.xlsx",
                        //        Outlook.OlAttachmentType.olByValue, Type.Missing,
                        //        Type.Missing);
                        //    mail.Send();
                        //}
                        ////////////////////
                        LogMessage("Sender Email Addres='" + mail.SenderEmailAddress + "'");
                        //Commented by lalit on 15july 2013
                       //string frmid = GetValidEmailAddress(mail.SenderEmailAddress, mfrm, loc);
                        //end
                        string frmid = GetEmailAddressFromExchange(mail.SenderEmailAddress);
                        if (frmid == "")
                        {
                            mail.Delete();
                            return;
                        }
                        LogMessage("GetValidEmailAddress='" + frmid + "'");
                        string contactnum = "0";
                        string m1 = mail.CC;
                        string location = "0";
                        string mailid = "";
                        string techid = "";
                        //////////////////////////////////////////////////send mail//////////////////////////////////////////////////////
                        //////////////////////////////////commented on 12.06.13
                        ///////////////////////////////////////end
                        //cmd.Parameters.Add("@subject", SqlDbType.NVarChar).Value = subject;
                        //cmd.Parameters.Add("@body", SqlDbType.NVarChar).Value = body;
                        //cmd.Parameters.Add("@recievedtime", SqlDbType.DateTime).Value = mail.ReceivedTime;
                        //cmd.Parameters.Add("@mailfrom", SqlDbType.NVarChar).Value = GetValidEmailAddress(mail.SenderEmailAddress);
                        //cmd.CommandText = "insert into storemail(subject,body,createddatetime,mailfrom,isActive) values(@subject,@body,@recievedtime,@mailfrom,1)";
                        ///////////////////////////////////////////////////change done by MEENAKSHI 24 aug 2012//////////////////////////////////////////////////////
                        ///////////////////////////////////////////////////insert into table incident_mst////////////////////////////////////////////////
                        cmd.Parameters.Add("@title", SqlDbType.NVarChar).Value = subject;
                        cmd.Parameters.Add("@description", SqlDbType.NVarChar).Value = body;
                        cmd.Parameters.Add("@createdatetime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                        cmd.Parameters.Add("@reporteddatetime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                        /////////////////////////////////////////////////////insert into table incidentstates
                        cmd1.Parameters.Add("@assignedtime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                        ///////////////////////////////////////////////////////insert into table IncidentHistory///////////////////////////////////////////////////
                        cmd2.Parameters.Add("@operationtime", SqlDbType.DateTime).Value = DateTime.Now.ToString();
                        cmd2.Parameters.Add("@description", SqlDbType.NVarChar).Value = body;
                        //LogMessage("Command Parameter Defined");
                        //if (body.Length <= 20000)
                        //{
                        ///////////////////////////////////////commented & change by meenakshi 28 august 2013
                        //if (mailto == "Helpdesk Corporate" || mailto == "Helpdesk.Corporate@oberoigroup.com")
                        //{
                        //    location = "1";   contactnum = "9811801977"; mailid = "Helpdesk.Corporate@oberoigroup.com"; techid = "44";
                        //}
                        //else if (mailto == "Helpdesk NarimanPointMumbai" || mailto == "Helpdesk.NarimanPointMumbai@tridenthotels.com")
                        //{
                        //    techid = "24"; location = "2"; contactnum = "+912266326582"; mailid = "Helpdesk.NarimanPointMumbai@tridenthotels.com";
                        //}
                        //else if (mailto == "Helpdesk Bandrakurla" || mailto == "Helpdesk.Bandrakurla@tridenthotels.com")
                        //{
                        //    techid = "32";location = "3";contactnum = "+919930455789";mailid = "Helpdesk.Bandrakurla@tridenthotels.com";
                        //}
                        //else if (mailto == "Helpdesk Bangalore" || mailto == "Helpdesk.Bangalore@oberoihotels.com")
                        //{
                        //    techid = "1078";location = "4";contactnum = "+918041358518"; mailid = "Helpdesk.Bangalore@oberoihotels.com";
                        //}
                        //else if (mailto == "Helpdesk Kolkata" || mailto == "Helpdesk.Kolkata@oberoihotels.com")
                        //{
                        //    techid = "1079";location = "5";contactnum = "+919831519900";mailid = "Helpdesk.Kolkata@oberoihotels.com";
                        //}
                        //else if (mailto == "Helpdesk Hyderabad" || mailto == "Helpdesk.Hyderabad@tridenthotels.com")
                        //{
                        //    techid = "1125";location = "6";contactnum = "9811801977"; mailid = "Helpdesk.Hyderabad@tridenthotels.com";
                        //}
                        //else if (mailto == "Helpdesk Delhi" || mailto == "Helpdesk.Delhi@oberoihotels.com")
                        //{
                        //    techid = "1076";location = "7"; contactnum = "9811801977"; mailid = "Helpdesk.Delhi@oberoihotels.com";
                        //}
                        //else if (mailto == "Helpdesk Chennai" || mailto == "Helpdesk.Chennai@tridenthotels.com")
                        //{
                        //    techid = "1077"; location = "8"; contactnum = "+919962218422";mailid = "Helpdesk.Chennai@tridentshotels.com";
                        //}
                        //else if (mailto == "Helpdesk Maidens" || mailto == "Helpdesk.Maidens@maidenshotel.com")
                        //{
                        //    techid = "1076";location = "9"; contactnum = "+911123890505"; mailid = "Helpdesk.Maidens@maidenshotel.com";
                        //}
                        if (mailto.Contains("Helpdesk Corporate") || mailto.Contains("Helpdesk.Corporate@oberoigroup.com"))
                        {
                            location = "1"; contactnum = "9811801977"; mailid = "Helpdesk.Corporate@oberoigroup.com"; techid = "17";
                        }
                        else if (mailto.Contains("Helpdesk NarimanPointMumbai") || mailto.Contains("Helpdesk.NarimanPointMumbai@tridenthotels.com"))
                        {
                            techid = "22"; location = "2"; contactnum = "+912266326582"; mailid = "Helpdesk.NarimanPointMumbai@tridenthotels.com";
                        }
                        else if (mailto.Contains("Helpdesk Bandrakurla") || mailto.Contains("Helpdesk.Bandrakurla@tridenthotels.com"))
                        {
                            techid = "37"; location = "3"; contactnum = "+919930455789"; mailid = "Helpdesk.Bandrakurla@tridenthotels.com";
                        }
                        else if (mailto.Contains("Helpdesk Bangalore") || mailto.Contains("Helpdesk.Bangalore@oberoihotels.com"))
                        {
                            techid = "1078"; location = "4"; contactnum = "+918041358518"; mailid = "Helpdesk.Bangalore@oberoihotels.com";
                        }
                        else if (mailto.Contains("Helpdesk Kolkata") || mailto.Contains("Helpdesk.Kolkata@oberoihotels.com"))
                        {
                            techid = "1079"; location = "5"; contactnum = "+919831519900"; mailid = "Helpdesk.Kolkata@oberoihotels.com";
                        }
                        else if (mailto.Contains("Helpdesk Hyderabad") || mailto.Contains("Helpdesk.Hyderabad@tridenthotels.com"))
                        {
                            techid = "4467"; location = "6"; contactnum = "9811801977"; mailid = "Helpdesk.Hyderabad@tridenthotels.com";
                        }
                        else if (mailto.Contains("Helpdesk Delhi") || mailto.Contains("Helpdesk.Delhi@oberoihotels.com"))
                        {
                            techid = "1076"; location = "7"; contactnum = "9811801977"; mailid = "Helpdesk.Delhi@oberoihotels.com";
                        }
                        else if (mailto.Contains("Helpdesk Chennai") || mailto.Contains("Helpdesk.Chennai@tridenthotels.com"))
                        {
                            techid = "1077"; location = "8"; contactnum = "+919962218422"; mailid = "Helpdesk.Chennai@tridentshotels.com";
                        }
                        else if (mailto.Contains("Helpdesk Maidens") || mailto.Contains("Helpdesk.Maidens@maidenshotel.com"))
                        {
                            techid = "1076"; location = "9"; contactnum = "+911123890505"; mailid = "Helpdesk.Maidens@maidenshotel.com";
                        }

                        else
                        {
                            location = "1"; contactnum = "9811801977"; mailid = "Helpdesk.Corporate@oberoigroup.com";techid = "16";
                        }
                        string[] EmailID = frmid.Split('@');
                        string Username = EmailID[0];
                        LogMessage("UserName:"+ Username);
                        if (!IsUserExist(Username))
                        {
                            LogMessage("UserDoesNotExist.");
                            UserCreate(Username, "pass@123", "", "", "1", frmid, "User", "", "", "", "", location, "");
                        }
                        string Uid = GetUserId_ByUserName(Username);
                        LogMessage("User Id :'" + Uid + "'");

                        _SLAId = GetSLAId_BySite_Priority(location, "1");
                        LogMessage("SLA Id-"+ _SLAId);
                        cmd.CommandText = @"insert into Incident_mst(Title, Timespentonreq, Slaid, Siteid, Requesterid, 
                            Modeid, Description, Deptid, Createdbyid, Createdatetime,Extension,Completedtime, ExternalTicketNo, 
                            VendorId,active,AMCCall ,Reporteddatetime)values(@Title, 0,'" + _SLAId + "', '" + location
                        + "','" + Uid + "', 5,@Description, 0, 2, @Createdatetime,0,null, null, 0,1, 0 ,@Reporteddatetime)";
                        int countoperation = cmd.ExecuteNonQuery();
                        ////////////////////////////////////////////////////////retrieve incidentid
                        string sQuery = ("select top 1 incidentid from incident_mst order by incidentid desc");
                        SqlConnection sc = new SqlConnection(connection);
                        sc.Open();
                        SqlCommand cmdn = new SqlCommand(sQuery, sc);
                        SqlDataReader dr = cmdn.ExecuteReader();
                        dr.Read();
                        int incidntid = Convert.ToInt32(dr["incidentid"].ToString());
                        cmd1.CommandText = @"insert into IncidentStates( Technicianid, Subcategoryid, Statusid, Requesttypeid,
                        Reqapproval, Reopened, Priorityid, Isescalated, Incidentid, Impactid, Hasattachment, Closecomments, 
                        Closeaccepted, Categoryid, AssignedTime )values('" + techid + "', 1, 12, 2, 0, 0, 1, 0,'"
                        + incidntid + "',0, 0, null, null, 1, @assignedtime )";

                        string HistroyDiscr = "";
                        if (body.Length>500)
                        {
                           HistroyDiscr = body.Substring(0, 500);
                           HistroyDiscr= HistroyDiscr.Replace("'","");
                        }
                        //LogMessage("HistoryDescr"+ HistroyDiscr);
                        cmd2.CommandText = @"insert into IncidentHistory(Operationtime, Operationownerid, Operation, 
                            Incidentid, Description )values(@Operationtime, 2, 'create','" + incidntid + "','" + HistroyDiscr + "' )";
         
                        if (countoperation == 1)
                        {
                            LogMessage("Data Inserted Incident Mst- Yes");
                            countoperation= cmd1.ExecuteNonQuery();
                            if (countoperation == 1)
                            {
                                LogMessage("Data Inserted Incident Status- Yes");
                                countoperation = cmd2.ExecuteNonQuery();
                                if (countoperation == 1)
                                {
                                    LogMessage("Data Inserted Incident History- Yes");
                                }
                            }
                        }
                        con.Close();
                        cmd.Dispose();
                        con.Dispose();
                        LogMessage("Incident ID created:"+incidntid);
                        ////////////////Vishal 
                        //Description length 200 Added by lalit
                        LogMessage("Body Length :'" + body.Length + "'");
                        if (body.Length > 200)
                        {
                            body = body.Substring(0, 200);
                            body = body + "(.....More Details are in the Mail)";
                        }
                        //End
                        location = GetSiteName_ById(location);
                        string strBody = "Dear " + mfrm + ",<br/><br/> Thank you for contacting IT Service desk, please find below the new Ticket Id details for your future reference.<br/><br/><b>Incident Details : </b> <br/><br/><b>Ticket Id&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</b> " + incidntid;
                        strBody = strBody + "<br/><b>Site&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</b> " + location;
                        strBody = strBody + "<br/><b>Logged Date & Time&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</b> " + mail.ReceivedTime;
                        strBody = strBody + "<br/><b>Description &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:" + body;
                        strBody = strBody + "<br/><b>Priority&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</b> Normal <br/><b>Username&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</b>" + mfrm;
                        strBody = strBody + "<br/><b>Email Address&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:</b> " + mailid;
                        if (mailto == "Helpdesk Corporate")
                        {
                            strBody = strBody + "<br/><br/>For any other support, kindly get in touch with us at +911123890505, Extn:2180 / 2181  or at +911123906180/81.<br/><br/> <b>In case your queries unresolved or not satisfactory  kindly reach to below escalation matrix:</b><br/><br/><b>1st Level : Arshad.Ali@oberoigroup.com, Ext: 2259 or +919999969434</b><br/><b>2nd Level : Ashish.Khanna@oberoigroup.com, Ext: 2178 or +919871164411</b><br/><b>3rd Level : Rajesh.Chopra@oberoigroup.com, Ext: 2175 or +919810079140</b><br/>";
                            strBody = strBody + "<br/><br/><b>Yours sincerely,</b><br/><br/> <b> EIH Central Helpdesk </b>";
                        }
                        else if (mailto == "Helpdesk NarimanPointMumbai")
                        {
                            strBody = strBody + "<br/><br/>For any other support, kindly get in touch with us at " + contactnum + "and 6587 or +91 9820831745 (24/7)<br/><br/> <b>In case your queries unresolved or not satisfactory  kindly reach to below escalation matrix:</b><br/><br/><b>Name & Contact Details</b><br/><br/><b>Maruti Aldar : 6587</b><br/><b>Ranjana Kamle : 6587</b><br/><b>Brijesh Singh: 6587</b><br/>";
                            strBody = strBody + "<b>Daulat Choudhary : 6581</b><br/><b>Mr. Noel Alvares : 6580</b><br/><br/><b>Yours sincerely,</b><br/><br/> <b> Trident NarimanPoint IT Support Desk, Ext. 6587 </b>";
                        }
                        else if (mailto == "Helpdesk Bandrakurla")
                        {
                            strBody = strBody + "<br/><br/>For any other support kindly get in touch with us at Extn:7411 / 7412,  " + contactnum + ", +917875551291 or at +912266727411/12<br/><br/> <b>In case your queries unresolved or not satisfactory  kindly reach to below escalation matrix:</b><br/><br/><b>Tier 1:  Darshan Trivedi – (+91 9819404011)</b><br/><b>Tier 2: Sudhir Nate (+91 9930455714), Jignesh Patel (+91 9987228227)</b><br/><b>Tier 3: Apurv Sharma (+91 9930455781)</b>";
                            strBody = strBody + "<br/><br/><b>Yours sincerely,</b><br/><br/> <b> Trident BandraKurla IT Support, Ext. 7411 / 7412 </b>";
                        }
                        else if (mailto == "Helpdesk Bangalore")
                        {
                            strBody = strBody + "<br/><br/>For any other support kindly get in touch with us at  " + contactnum + ", Extn:8516 / 8517 or email at itsupport.tobl@oberoihotels.com<br/><br/> <b>In case your queries unresolved or not satisfactory  kindly reach to below escalation matrix:</b><br/><br/><b>Chethan Kumar. J</b><br/><b>Chethankumar.j@oberoihotels.com</b><br/><b>System Administrator</b>";
                            strBody = strBody + "<br/><b>91-9916033777</b><br/><br/><b>Yunus Khatib</b><br/><b>Yunus.khatib@oberoihotels.com</b><br/><b>Systems Manager</b><br/><b>91-9886058585</b><br/><br/><b>Yours sincerely,</b><br/><br/> <b> The Oberoi Bangalore IT Support, Ext. 8546 / 8517s </b>";
                        }
                        else if (mailto == "Helpdesk Kolkata")
                        {
                            strBody = strBody + "<br/><br/>For any other support kindly get in touch with us at Ext: +91 33 2249 2323 2323 or at " + contactnum + "/ arindam.banerjee@oberoihotels.com<br/><br/> <b>In case your queries unresolved or not satisfactory  kindly reach to below escalation matrix:</b><br/><br/><b>Arindam Banerjee </b><br/><b>Systems Manager</b><br/><b>The Oberoi Grand</b>";
                            strBody = strBody + "<br/><b>Email: arindam.banerjee@oberoihotels.com</b><br/><br/><b>Yours sincerely,</b><br/><br/> <b> The Oberoi Grand Kolkata IT Support </b>";
                        }
                        else if (mailto == "Helpdesk Hyderabad")
                        {
                            ///////////////////////////////commented and added by Nagaraju 24 july 2013
                           
                            //strBody = strBody + "<br/><br/> <b>For any other support kindly get in touch with the </b><br/><br/><b>IT Support team on Extension 6252,Direct Number +91 40 6603 6252,Mobile Number + 91  88860 48662</b><br/><br/><b>Telephone Support staff  on  Extension 6253, Direct Number +91 40 6603 6253, Mobile Number +91 88860 48663</b><br/><br/><b>If your queries are unresolved, please call us at the following extension / mobile numbers:</b>";
                            //strBody = strBody + "<br/><b>IT Supervisor on extension 6251 (or) +91 40 6603 6251</b><br/><b>IT Manager on extension 6250   (or) +91 40 6603 6250 </b>";                           
                            //strBody = strBody + "<br/><br/><b>Yours sincerely,</b><br/><br/> <b> Trident Hyderabad IT Support </b>";
                            strBody = strBody + "<br/><br/> For any other support kindly get in touch with the <br/><br/>IT Support team &nbsp&nbsp&nbsp&nbsp on Extension 6252,Direct Number +91 40 6603 6252,Mobile Number + 91  88860 48662<br/><br/>Telephone Support staff &nbsp&nbsp&nbsp&nbsp on  Extension 6253, Direct Number +91 40 6603 6253, Mobile Number +91 88860 48663<br/><br/>If your queries are unresolved, please call us at the following extension / mobile numbers:";
                            strBody = strBody + "<br/><br/>IT Supervisor &nbsp&nbsp&nbsp&nbsp&nbsp on extension 6251 (or) +91 40 6603 6251<br/>IT Manager &nbsp&nbsp&nbsp&nbsp on extension 6250   (or) +91 40 6603 6250 ";
                            strBody = strBody + "<br/><br/>Yours sincerely,<br/><br/>  Trident Hyderabad IT Support ";
                        }
                        else if (mailto == "Helpdesk Delhi")
                        {
                            strBody = strBody + "<br/><br/>For any other support, kindly get in touch with us at +911123890505, Extn:2180 / 2181  or at +911123906180/81.<br/><br/> <b>In case your queries unresolved or not satisfactory  kindly reach to below escalation matrix:</b><br/><br/><b>1st Level : Arshad.Ali@oberoigroup.com, Ext: 2259 or +919999969434</b><br/><b>2nd Level : Ashish.Khanna@oberoigroup.com, Ext: 2178 or +919871164411</b><br/><b>3rd Level : Rajesh.Chopra@oberoigroup.com, Ext: 2175 or +919810079140</b><br/>";
                            strBody = strBody + "<br/><br/><b>Yours sincerely,</b><br/><br/> <b> EIH Central Helpdesk </b>";
                        }
                        else if (mailto == "Helpdesk Chennai")
                        {
                            strBody = strBody + "<br/><br/>For any other support kindly get in touch with us at Extn:8271 or at " + contactnum + "/ itsupport.ttch@tridenthotels.com<br/><br/> <b>In case your queries unresolved or not satisfactory  kindly reach to below escalation matrix:</b><br/><br/><b>Satheesh Kumar.S,</b><br/><b>Extn: 8270</b><br/><b>Mobile: +919884398638</b><br/><b>Email: satheesh.kumar@tridenthotels.com</b>";
                            strBody = strBody + "<br/><br/><b>Yours sincerely,</b><br/><br/> <b> Trident Chennai IT Support, Ext. 8271</b>";
                        }
                        else if (mailto == "Helpdesk Maidens")
                        {
                            strBody = strBody + "<br/><br/>For any other support kindly get in touch with us at " + contactnum + ", Extn:2370 or at +911123906180/81<br/><br/> <b>In case your queries unresolved or not satisfactory  kindly reach to below escalation matrix:</b><br/><br/><b>Pawan Suman & Contact Details +91 7838651733 with email address pawan.suman@maidenshotel.com</b>";
                            strBody = strBody + "<br/><br/><b>Yours sincerely,</b><br/><br/> <b> Maidens Hotel IT Support, Ext, 2370 </b>";
                        }
                        else
                        {
                            strBody = strBody + "<br/><br/>For any other support kindly get in touch with us at " + contactnum + ".<br/><br/> <b>This is an auto generated mail. Please do not reply.</b><br/><br/><b>Yours sincerely,</b><br/><br/> <b> EIH Central Helpdesk </b>";
                        }
                        MailMessage myMessage = new MailMessage();
                        myMessage.From = new MailAddress(FromMailId);
                        myMessage.To.Add(frmid);
                        myMessage.Subject = subject;
                        myMessage.IsBodyHtml = true;
                        myMessage.Body = strBody;
                        SmtpClient mySmtpClient = new SmtpClient();

                        System.Net.NetworkCredential myCredential = new System.Net.NetworkCredential(NetWorkCredentialMailId, NetWorkCredentialPassword);
                        mySmtpClient.Host = smtphost;
                        mySmtpClient.Credentials = myCredential;

                        mySmtpClient.ServicePoint.MaxIdleTime = 1;
                        mySmtpClient.Port = Convert.ToInt32(SMTPPort);
                        mySmtpClient.Send(myMessage);
                        myMessage.Dispose();
                        LogMessage("Mail Sent to User");
                        //dr.Close();
                        /////////////////////////////////////end BY MEENAKHI 24 aug 2012///////////////////////////////////////////////////////////////////////////
                        mail.Delete();
                        LogMessage("MailDeleted");
                        iMailCount = iMailCount - 1;
                        LogMessage("iMailCount = " + iMailCount.ToString());
                        ////////////////////////////////////////end
                      }
                    else
                    {
                        iCount = iCount + 1;
                    }
                }
                GC.Collect();
                inbox = null;
                ns = null;
                outlook = null;
                LogMessage("Execution Successfully Complete");
            }
            catch (SqlException ex)
            {
                ex.ToString();
                LogMessage("Read Mails fails");
                LogMessage(ex.ToString());

            }
        }
       //////////////////////////////////////////change done by meenakshi 28 august 2013
        //public string getlocation(string mailto)
        //{
        //    string location = "0";
        //    if (mailto == "Helpdesk Corporate" || mailto == "Helpdesk.Corporate@oberoigroup.com") { location = "1"; }
        //    else if (mailto == "Helpdesk NarimanPointMumbai" || mailto == "Helpdesk.NarimanPointMumbai@tridenthotels.com") { location = "2"; }
        //    else if (mailto == "Helpdesk Bandrakurla" || mailto == "Helpdesk.Bandrakurla@tridenthotels.com") { location = "3"; }
        //    else if (mailto == "Helpdesk Bangalore" || mailto == "Helpdesk.Bangalore@oberoihotels.com") { location = "4"; }
        //    else if (mailto == "Helpdesk Kolkata" || mailto == "Helpdesk.Kolkata@oberoihotels.com") { location = "5"; }
        //    else if (mailto == "Helpdesk Hyderabad" || mailto == "Helpdesk.Hyderabad@tridenthotels.com") { location = "6"; }
        //    else if (mailto == "Helpdesk Delhi" || mailto == "Helpdesk.Delhi@oberoihotels.com") { location = "7"; }
        //    else if (mailto == "Helpdesk Chennai" || mailto == "Helpdesk.Chennai@tridenthotels.com") { location = "8"; }
        //    else if (mailto == "Helpdesk Maidens" || mailto == "Helpdesk.Maidens@maidenshotel.com") { location = "9"; }
        //    else { location = "1";}
        //    return location;
        //}
        public string getlocation(string mailto)
        {
            string location = "0";
            if (mailto.Contains("Helpdesk Corporate") || mailto.Contains("Helpdesk.Corporate@oberoigroup.com")) { location = "1"; }
            else if (mailto.Contains("Helpdesk NarimanPointMumbai") || mailto.Contains("Helpdesk.NarimanPointMumbai@tridenthotels.com")) { location = "2"; }
            else if (mailto.Contains("Helpdesk Bandrakurla") || mailto.Contains("Helpdesk.Bandrakurla@tridenthotels.com")) { location = "3"; }
            else if (mailto.Contains("Helpdesk Bangalore") || mailto.Contains("Helpdesk.Bangalore@oberoihotels.com")) { location = "4"; }
            else if (mailto.Contains("Helpdesk Kolkata") || mailto.Contains("Helpdesk.Kolkata@oberoihotels.com")) { location = "5"; }
            else if (mailto.Contains("Helpdesk Hyderabad") || mailto.Contains("Helpdesk.Hyderabad@tridenthotels.com")) { location = "6"; }
            else if (mailto.Contains("Helpdesk Delhi") || mailto.Contains("Helpdesk.Delhi@oberoihotels.com")) { location = "7"; }
            else if (mailto.Contains("Helpdesk Chennai") || mailto.Contains("Helpdesk.Chennai@tridenthotels.com")) { location = "8"; }
            else if (mailto.Contains("Helpdesk Maidens") || mailto.Contains("Helpdesk.Maidens@maidenshotel.com")) { location = "9"; }
            else { location = "1"; }
            return location;
        }
       /// <summary>
       /// /////////////////////////////////////////end
      
        public string GetValidEmailAddress(string sEmail, string mfr, string sloc)
        { //changedone by meenakshi
            LogMessage("sMail:" + sEmail);
            string sName = "";
            if (sEmail.Contains(",") || sEmail.Contains("["))
            {
                int l = sEmail.Length;
                int index = sEmail.IndexOf("[");
                int finalindex = sEmail.IndexOf("]");
                int len = finalindex - index;
                string s = sEmail.Substring(index + 1, len - 1);
                //EventLog.WriteEntry("Oberoigroup Email ID After formating 1- ", sEmail);
                LogMessage("if (sEmail.Contains(',') || sEmail.Contains('['))");
                return s;
            }
            else if (sEmail.Contains("<"))
            {
                int indx = sEmail.IndexOf("<");
                int finalindx = sEmail.IndexOf(">");
                int ln = finalindx - indx;
                string m = sEmail.Substring(indx + 1, ln - 1);
                //EventLog.WriteEntry("Oberoigroup Email ID After formating 2- ", sEmail);
                LogMessage("if (sEmail.Contains(' < '))");
                return m;
            }
            else //end
            {
                if (sEmail.Contains("@"))
                {
                    return sEmail;
                    LogMessage("Mail contains @ now mail id is:"+ sEmail);
                }
                if (sEmail.Contains("SERVER.ADMINTEAM"))  
                { 
                    sEmail = "V-server.administrators@oberoigroup.com";
                    LogMessage("(sEmail.Contains('SERVER.ADMINTEAM'))");
                    return sEmail;
                }
                if (mfr.Contains("Corporate Network Administrator"))
                { 
                    sEmail = "V-Network.administrators@oberoigroup.com";
                    LogMessage("(mfr.Contains('Corporate Network Administrator'))");
                    return sEmail;  
                }

                if (mfr.Contains("(Vendor)"))
                {
                    string[] S = mfr.Split(' ');
                    sName = "V-" + S[0] + "." + S[1];
                    LogMessage("(mfr.Contains('(Vendor)'))");
                }
                else
                    if (sEmail.Contains("smtp:"))
                    {
                        int iPos = sEmail.IndexOf("smtp:");
                        sName = sEmail.Substring(iPos + ("smtp:").Length, sEmail.Length - (iPos + ("smtp:").Length));
                        LogMessage("(sEmail.Contains('smtp:'))");
                    }
                    else 
                    if (sEmail.Contains("RECIPIENTS/CN="))
                    {
                        int iPos = sEmail.IndexOf("RECIPIENTS/CN=");
                        sName = sEmail.Substring(iPos + ("RECIPIENTS/CN=").Length, sEmail.Length - (iPos + ("RECIPIENTS/CN=").Length));
                        LogMessage("(sEmail.Contains('RECIPIENTS/CN='))");
                    }
                }
                if (sloc == "3")  {  sEmail = sName + "@tridenthotels.com";  }
                else if (sloc == "6")  { sEmail = sName + "@tridenthotels.com";  }
                else if (sloc == "8")  { sEmail = sName + "@tridenthotels.com";  }
                else if (sloc == "7") {  sEmail = sName + "@oberoihotels.com";   }
                else if (sloc == "2") {  sEmail = sName + "@oberoihotels.com";   }
                else if (sloc == "5") {  sEmail = sName + "@oberoihotels.com";   }
                else if (sloc == "4") {  sEmail = sName + "@oberoihotels.com";   }
                else if (sloc == "9") {  sEmail = sName + "@maidenshotel.com";   }
                else if (sloc == "1") {  sEmail = sName + "@oberoigroup.com";    }
                EventLog.WriteEntry("Oberoigroup Email ID After formating 3- ", sEmail);
                return sEmail;
            }
         
        //Added by lalit 3 July 2013
        protected bool IsUserExist(string Email)
        {
            try
            {
                string[] mail = Email.Split('@');
                string Username = mail[0];
                string _Query = @"select * from UserLogin_mst where username='" + Username + "' and enable=1";
                SqlConnection sc = new SqlConnection(connection);
                sc.Open();
                SqlCommand cmdn = new SqlCommand(_Query, sc);
                SqlDataReader dr = cmdn.ExecuteReader();
                if (dr.HasRows)
                {
                    sc.Close();
                   return true;
                }
                else
                {
                    sc.Close();
                    return false;
                }
            }
            catch (System.Exception ex)
            {
                EventLog.WriteEntry("Function- IsUserExist ",ex.Message);
            };
            return false;
        }
        protected string GetUserId_ByUserName(string Uname)
        {
            try
            {
                string _Query = @"select * from UserLogin_mst where username='" + Uname + "' and enable=1";
                SqlConnection sc = new SqlConnection(connection);
                sc.Open();
                SqlCommand cmdn = new SqlCommand(_Query, sc);
                SqlDataReader dr = cmdn.ExecuteReader();
                while (dr.Read())
                {
                    if (dr.HasRows)
                    {
                        return dr["userid"].ToString();
                    }
                }
                dr.Close();
                sc.Close();
            }
            catch (System.Exception ex)
            {
                EventLog.WriteEntry("Function- GetUserId_ByUserName ", ex.Message);
            };
            return "";
        }

       //Function to Get SLA on the basis of Site and Priority.........
        protected string GetSLAId_BySite_Priority(string SiteId,string Priority)
        {
            try
            {
                string _Query = @"select SLA_mst.slaid from SLA_mst  Inner join SLA_Priority_mst on SLA_mst.SLAid=SLA_Priority_mst.Slaid 
                                  where SLA_mst.siteid='"+ SiteId +"' and SLA_Priority_mst.priorityid='"+ Priority +"'";
                SqlConnection sc = new SqlConnection(connection);
                sc.Open();
                SqlCommand cmdn = new SqlCommand(_Query, sc);
                SqlDataReader dr = cmdn.ExecuteReader();
                while (dr.Read())
                {
                    if (dr.HasRows)
                    {
                        return dr["slaid"].ToString();
                    }
                }
                dr.Close();
                sc.Close();
            }
            catch (System.Exception ex)
            {
                EventLog.WriteEntry("Function- GetSLAId_BySite_Priority ", ex.Message);
            };
            return "0";
        }

        protected string GetSiteName_ById(string SId)
        {
            try
            {
                string _Query = @"select sitename from Site_mst where siteid='"+ SId +"'";
                SqlConnection sc = new SqlConnection(connection);
                sc.Open();
                SqlCommand cmdn = new SqlCommand(_Query, sc);
                SqlDataReader dr = cmdn.ExecuteReader();
                while (dr.Read())
                {
                    if (dr.HasRows)
                    {
                        return dr["sitename"].ToString();
                    }
                }
                dr.Close();
                sc.Close();
            }
            catch (System.Exception ex)
            {
                EventLog.WriteEntry("Function- GetSiteName_ById ", ex.Message);
            };
            return "";
        }
       //End
        protected string GetOrgId()
        {
            try
            {
                string _Query = @"select top 1  orgid,orgname,createdatetime,description from organization_mst";
                SqlConnection sc = new SqlConnection(connection);
                sc.Open();
                SqlCommand cmdn = new SqlCommand(_Query, sc);
                SqlDataReader dr = cmdn.ExecuteReader();
                while (dr.Read())
                {
                    if (dr.HasRows)
                    {
                        return dr["orgid"].ToString();
                    }
                }
                dr.Close();
                sc.Close();
            }
            catch (System.Exception ex)
            {
                EventLog.WriteEntry("Function- GetOrgId", ex.Message);
            };
            return "";
        }
       protected int IsUserExist(string username,string uid)
        {
            try
            {
                string _Query = @"select userid from userlogin_mst where [username]=rtrim(ltrim('" + username + "'))and [orgid]=rtrim(ltrim('" + uid + "'))";
                SqlConnection sc = new SqlConnection(connection);
                sc.Open();
                SqlCommand cmdn = new SqlCommand(_Query, sc);
                SqlDataReader dr = cmdn.ExecuteReader();
                while (dr.Read())
                {
                    if (dr.HasRows)
                    {
                        return Convert.ToInt32(dr["userid"]);
                    }
                }
                sc.Close();
                dr.Close();
            }
            catch (System.Exception ex)
            {
                EventLog.WriteEntry("Function- IsUserExist", ex.Message);
            };
            return 0;
        }
       protected int InsertInIncidentMst(string OrgId,string Uname,string Pwd,string roleid,string createddatetime,bool enable,bool AdEnable)
       {
           try
           {
               string _Query = @"insert into UserLogin_mst(orgid,username,password,roleid,createdatetime,enable,ADEnable)
                                 values('" + OrgId + "','" + Uname + "','" + Pwd + "','" + roleid + "','" + createddatetime + "','" + enable + "','" + AdEnable + "')";
               SqlConnection sc = new SqlConnection(connection);
               sc.Open();
               SqlCommand cmdn = new SqlCommand(_Query, sc);
               int i = cmdn.ExecuteNonQuery();
               sc.Close();
               return i;
           }
           catch (System.Exception ex)
           {
               EventLog.WriteEntry("Function- InsertInIncidentMst", ex.Message);
           };
           return 0;
       }

       protected int UserToSiteMaping(string Uid,string SiteId)
       {
         try
           {
               string _Query = @"INSERT INTO [UserToSiteMapping]([userid],[siteid]) VALUES('"+Uid+"','"+SiteId +"');";
               SqlConnection sc = new SqlConnection(connection);
               sc.Open();
               SqlCommand cmdn = new SqlCommand(_Query, sc);
               int i = cmdn.ExecuteNonQuery();
               sc.Close();
               return i;
           }
           catch (System.Exception ex)
           {
               EventLog.WriteEntry("Function- InsertInIncidentStatus", ex.Message);
           };
           return 0;
       }


       protected int UserToMailMaping(string Uid, string emailId)
       {
           try
           {
               string _Query = @"INSERT INTO [UserEmail]([userid],[emailid],[Active]) VALUES('" + Uid + "','" + emailId + "',1);";
               SqlConnection sc = new SqlConnection(connection);
               sc.Open();
               SqlCommand cmdn = new SqlCommand(_Query, sc);
               int i = cmdn.ExecuteNonQuery();
               sc.Close();
               return i;
           }
           catch (System.Exception ex)
           {
               EventLog.WriteEntry("Function- UserToMailMaping", ex.Message);
           };
           return 0;
       }


       protected int InsertInLoginStatus(string Uid, string Fname, string Lname, string Email, string Mobile, 
                      string Landline, string deptname,string emailid,string description,string siteid,string deptid)
       {
           try
           {
               string _Query = @"Insert Into ContactInfo_mst(userid,firstname,lastname,mobile,landline,deptname,emailid,description,siteid,deptid)
                                  Values('" + Uid + "','" + Fname + "','" + Lname + "','" + Mobile + "','" + Landline + "','" + deptname
                                  + "','" + emailid + "','" + description + "','" + siteid + "','" + deptid + "')";
               SqlConnection sc = new SqlConnection(connection);
               sc.Open();
               SqlCommand cmdn = new SqlCommand(_Query, sc);
               int i = cmdn.ExecuteNonQuery();
               sc.Close();
               return i;
           }
           catch (System.Exception ex)
           {
               EventLog.WriteEntry("Function- InsertInIncidentStatus", ex.Message);
           };
           return 0;
       }

       public string UserCreate(string UName, string Password, string Company, string city, string roleid,
                              string UserEmailId, string RoleName, string Description, string EmployeeId, string LandLineNo,
                              string MobileNo, string Location, string DepartmentId)
        {
            string OrgId = GetOrgId();
            //LogMessage("OrgId='"+OrgId+"'"); 
            int Flag;
            string varRoleName;
            bool FlagMembership;
            // Use Asp.Net Membership Validator Control Membership.ValidateUser to check User Exist in aspnet Database 
            FlagMembership = Membership.ValidateUser(UName, Password);
            Flag = IsUserExist(UName, OrgId);
            if (Flag == 0 && FlagMembership == false)
            {
               // LogMessage("Flag == 0 && FlagMembership == false"); 
                int status;
                status = InsertInIncidentMst(OrgId, UName, Password, roleid, DateTime.Now.ToString(), true, false);
                if (status != 1)
                {
                    LogMessage("Function- Add User- Insert in UserLoginMst','User is not Add in UserLogin Table");
                    //EventLog.WriteEntry("Function- Add User- Insert in UserLoginMst", "User is not Add in UserLogin Table");
                }
                if (status == 1)
                {
                    LogMessage("UserLoginMst user entered"); 
                    string varEmail="";
                    if (UserEmailId != "")
                    {
                        varEmail = UserEmailId.Trim();
                    }
                    varRoleName = RoleName.Trim();
                    MembershipCreateStatus Mstatus = default(MembershipCreateStatus);
                    Membership.CreateUser(UName.Trim(), Password.Trim(), varEmail, "Project Name", "Helpdesk", true, out Mstatus);
                    Roles.AddUserToRole(UName.Trim(), varRoleName);
                    int userid;
                    userid = IsUserExist(UName, OrgId);
                    UserToSiteMapping objusertositemapping = new UserToSiteMapping();
                    if (userid != 0)
                    {
                        int k= InsertInLoginStatus(userid.ToString(), UName.Trim(), UName.Trim(), UserEmailId, MobileNo, LandLineNo, DepartmentId,
                        UserEmailId, Description, Location,DepartmentId);
                        LogMessage("Contact Info-Mst user created:'"+ k +"'"); 

                        if (k != 1)
                        {
                            LogMessage("Function- Add User- Insert in UserLoginMst','User is not Add in UserLogin Table"); 
                            //EventLog.WriteEntry("Function- Add User- Insert in UserLoginMst", "User is not Add in UserLogin Table");
                        }
                        int i=UserToSiteMaping(userid.ToString(), Location);
                        LogMessage("User To Site Mapping:'" + i + "'"); 
                        if (i != 1)
                        {
                            EventLog.WriteEntry("Function- Add User- Asset Mapping","User To Site Maaping is not done");
                        }
                        i = UserToMailMaping(userid.ToString(), UserEmailId);
                        return "Created";
                     }
                    else   {  return "Error";  }

                }
                else {  return "Error";  }
            }
            else
            {
                return "Already Exist";
            }
        }
    }
}

