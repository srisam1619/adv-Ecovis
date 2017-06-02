using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Configuration;
using System.Data;
using System.IO;
using System.Web;
using System.Globalization;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace PersonalTaxFilingReminder
{
    public class SendEmail
    {
        clsLog oLog = new clsLog();
        string sErrDesc = string.Empty;

        string sFromMailId = ConfigurationManager.AppSettings["FromMailId"];
        string sFromMailIdPassword = ConfigurationManager.AppSettings["FromMailIdPassword"];
        string sSMTPHost = ConfigurationManager.AppSettings["SMTPHost"];
        int iSMTPPort = Convert.ToInt32(ConfigurationManager.AppSettings["SMTPPort"]);
        int iSMTPConnTimeout = Convert.ToInt32(ConfigurationManager.AppSettings["SMTPConnTimeout"]);
        string sSubject = ConfigurationManager.AppSettings["Subject"];
        string sServer = ConfigurationManager.AppSettings["Server"];
        string sDatabase = ConfigurationManager.AppSettings["Database"];
        string sUid = ConfigurationManager.AppSettings["uid"];
        string sPwd = ConfigurationManager.AppSettings["pwd"];
        string ConnectionString = ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        string sErrorEmailId = ConfigurationManager.AppSettings["ErrorEmailId"];
        string sErrorSubject = ConfigurationManager.AppSettings["ErrorSubject"];

        public DataSet GetCardCode()
        {
            DataSet oDataset = new DataSet();
            string sFuncName = string.Empty;
            string sProcName = string.Empty;

            try
            {
                sFuncName = "GetCardCode()";
                sProcName = "AB_PTAXFILING_SP001_GetCardCodes";
                oLog.WriteToDebugLogFile("Starting Function ", sFuncName);

                oLog.WriteToDebugLogFile("Calling Run_StoredProcedure() " + sProcName, sFuncName);
                oDataset = SqlHelper.ExecuteDataSet(ConnectionString, CommandType.StoredProcedure, sProcName);

                oLog.WriteToDebugLogFile("Completed With SUCCESS ", sFuncName);
                return oDataset;
            }
            catch (Exception Ex)
            {
                sErrDesc = Ex.Message.ToString();
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
                throw Ex;
            }
        }

        public DataSet CheckDuplicateMailSending(string sCode, string sName, string sPTAXFILINGDate, string Mail)
        {
            string sFuncName = string.Empty;
            DataSet oDataset = new DataSet();
            string sProcName = string.Empty;
            try
            {
                sFuncName = "CheckDuplicateMailSending";
                //string sInsertQuery = "insert into [@EMAILLOG_SOA](" + sColumnNames + ") values(" + sParameter + ")";
                sProcName = "AB_PTAXFILING_SP002_CheckDuplicateMailSending";
                oLog.WriteToDebugLogFile("Starting Function ", sFuncName);

                oLog.WriteToDebugLogFile("Calling Run_StoredProcedure() " + sProcName, sFuncName);
                oDataset = SqlHelper.ExecuteDataSet(ConnectionString, CommandType.StoredProcedure, sProcName, Data.CreateParameter("@CardCode", sCode),
                    Data.CreateParameter("@CardName", sName), Data.CreateParameter("@PTAXFILINGDate", sPTAXFILINGDate), Data.CreateParameter("@Email", Mail));

                oLog.WriteToDebugLogFile("Completed With SUCCESS ", sFuncName);
            }
            catch (Exception Ex)
            {
                sErrDesc = Ex.Message.ToString();
                //sResult = sErrDesc;
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
            }
            return oDataset;
        }

        public string ComposeBody(string sContPrsnName, string sCompanyName, int ToIdentifyTemplate)
        {
            string sFuncName = string.Empty;
            string stextBody = string.Empty;
            try
            {
                sFuncName = "ComposeBody";
                string sBodyDetail = string.Empty;
                string sBodyDetail1 = string.Empty;
                string sTableFormat = string.Empty;

                var sbMail = new StringBuilder();

                if (ToIdentifyTemplate == 1)
                {
                    string sTemplatePath = AppDomain.CurrentDomain.BaseDirectory;
                    int index = sTemplatePath.IndexOf("\\bin");
                    if (index > 0)
                        sTemplatePath = sTemplatePath.Substring(0, index) + "\\Email Template\\Reminder1.htm";

                    using (var sReader = new StreamReader(sTemplatePath))
                    {
                        sbMail.Append(sReader.ReadToEnd());

                        sbMail.Replace("{ContactPersonName}", System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(sContPrsnName.ToLower()));
                    }
                }
                else if (ToIdentifyTemplate == 2)
                {
                    string sTemplatePath = AppDomain.CurrentDomain.BaseDirectory;
                    int index = sTemplatePath.IndexOf("\\bin");
                    if (index > 0)
                        sTemplatePath = sTemplatePath.Substring(0, index) + "\\Email Template\\Reminder2.htm";

                    using (var sReader = new StreamReader(sTemplatePath))
                    {
                        sbMail.Append(sReader.ReadToEnd());

                        sbMail.Replace("{ContactPersonName}", System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(sContPrsnName.ToLower()));
                    }
                }
                
                stextBody = sbMail.ToString();
            }
            catch (Exception ex)
            {
                sErrDesc = ex.Message.ToString();
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
            }
            return stextBody;
        }

        public string InsertEmailLog(string sCode, string sName, string sPTAXFILINGDate, string sDateSent, string Mail)
        {
            string sResult = string.Empty;
            string sFuncName = string.Empty;
            DataSet oDataset = new DataSet();
            string sProcName = string.Empty;
            try
            {
                sFuncName = "InsertEmailLog";
                //string sInsertQuery = "insert into [@EMAILLOG_SOA](" + sColumnNames + ") values(" + sParameter + ")";
                sProcName = "AB_PTAXFILING_SP003_InsertPTAXFILINGLog";
                oLog.WriteToDebugLogFile("Starting Function ", sFuncName);

                oLog.WriteToDebugLogFile("Calling Run_StoredProcedure() " + sProcName, sFuncName);
                oDataset = SqlHelper.ExecuteDataSet(ConnectionString, CommandType.StoredProcedure, sProcName, Data.CreateParameter("@CardCode", sCode),
                    Data.CreateParameter("@CardName", sName), Data.CreateParameter("@PTAXFILINGDate", sPTAXFILINGDate), Data.CreateParameter("@DateSent", sDateSent),
                    Data.CreateParameter("@Email", Mail));

                oLog.WriteToDebugLogFile("Completed With SUCCESS ", sFuncName);
                sResult = "SUCCESS";
            }
            catch (Exception Ex)
            {
                sErrDesc = Ex.Message.ToString();
                sResult = sErrDesc;
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
            }
            return sResult;
        }

        public string SendAutomatedEmail(string sEmailTo, string sCntctPrsnName, string sCompanyName, int ToIdentifyTemplate, string sSuffix, ref string sErrDesc)
        {
            string functionReturnValue = string.Empty;

            string sFuncName = "SendAutomatedEmail";

            try
            {
                oLog.WriteToDebugLogFile("Sarting function", sFuncName);
                oLog.WriteToDebugLogFile("Setting SMTP properties", sFuncName);
                SmtpClient smtpClient = new SmtpClient(sSMTPHost, iSMTPPort);

                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new System.Net.NetworkCredential(sFromMailId, sFromMailIdPassword);

                //smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                // smtpClient.EnableSsl = True
                smtpClient.EnableSsl = true;

                oLog.WriteToDebugLogFile("Calling Function CreateDefaultMailMessage()", sFuncName);

                MailMessage message = CreateDefaultMailMessage(sFromMailId, sEmailTo, sCntctPrsnName, sCompanyName, ToIdentifyTemplate, sSuffix, ref sErrDesc);
                object userState = message;

                oLog.WriteToDebugLogFile("Sending Email Message", sFuncName);

                oLog.WriteToDebugLogFile("Sending Email Messages to : " + sEmailTo, sFuncName);

                ServicePointManager.ServerCertificateValidationCallback = delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                { return true; };

                smtpClient.Send(message);

                message.Dispose();

                functionReturnValue = "SUCCESS";

                oLog.WriteToDebugLogFile("Function completed with Success", sFuncName);

            }
            catch (Exception ex)
            {
                functionReturnValue = ex.Message;
                sErrDesc = ex.Message;

                oLog.WriteToDebugLogFile("Function completed with Error", sFuncName);
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToErrorLogFile("Failed sending email to : " + " " + sEmailTo, sFuncName);

            }
            finally
            {
            }
            return functionReturnValue;

        }

        public string SendEmailOnErrorCase(DataTable dt, ref string sErrDesc)
        {
            string functionReturnValue = string.Empty;

            string sFuncName = "SendEmailOnErrorCase";

            try
            {
                oLog.WriteToDebugLogFile("Sarting function", sFuncName);
                oLog.WriteToDebugLogFile("Setting SMTP properties", sFuncName);
                SmtpClient smtpClient = new SmtpClient(sSMTPHost, iSMTPPort);

                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new System.Net.NetworkCredential(sFromMailId, sFromMailIdPassword);

                smtpClient.EnableSsl = true;

                MailMessage message = new MailMessage();

                message.SubjectEncoding = System.Text.Encoding.UTF8;
                message.To.Add(new MailAddress(sErrorEmailId));
                message.From = new MailAddress(sFromMailId);
                message.Subject = sErrorSubject;
                message.IsBodyHtml = true;
                message.Body = ComposeErrorBody(dt);
                oLog.WriteToDebugLogFile("Sending Email Message", sFuncName);

                oLog.WriteToDebugLogFile("Sending Email Messages to : " + sErrorEmailId, sFuncName);

                smtpClient.Send(message);

                functionReturnValue = "SUCCESS";

                oLog.WriteToDebugLogFile("Function completed with Success", sFuncName);

            }
            catch (Exception ex)
            {
                functionReturnValue = ex.Message;
                sErrDesc = ex.Message;

                oLog.WriteToDebugLogFile("Function completed with Error", sFuncName);
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToErrorLogFile("Failed sending email to : " + " " + sErrorEmailId, sFuncName);

            }
            finally
            {
            }
            return functionReturnValue;

        }

        public string ComposeErrorBody(DataTable dtBody)
        {
            string sBodyDetail = string.Empty;
            string sBodyDetail1 = string.Empty;
            string sTableFormat = string.Empty;
            foreach (DataRow item in dtBody.Rows)
            {
                sBodyDetail = "<tr><td>&nbsp;" + item["CardCode"] + "</td><td>&nbsp;" + item["CardName"].ToString() + " </td> " +
                    " <td>&nbsp;" + item["PTAXFILINGDate"].ToString() + " </td><td>&nbsp;" + item["Email"].ToString() + " </td> " +
                    " <td>&nbsp;" + item["ErrorMsg"].ToString() + " </td></tr>";
                sBodyDetail1 = sBodyDetail1 + sBodyDetail;
            }
            sTableFormat = "<table border = '1' cellspacing = 0 cellpadding = 0 style='font-size:10.0pt;font-family:Arial;width: 85%;'> " +
                                "<tr><td><strong style='color: blue; background-color: transparent;'>&nbsp;Customer Code&nbsp;</strong></td> " +
                                "<td><strong style='color: blue; background-color: transparent;'>&nbsp;Customer Name &nbsp; </strong></td> " +
                                "<td><strong style='color: blue; background-color: transparent;'>&nbsp;Personal Tax Filing Date &nbsp;</strong></td> " +
                                "<td><strong style='color: blue; background-color: transparent;'>&nbsp;Email Id &nbsp;</strong></td> " +
                                "<td><strong style='color: blue; background-color: transparent;'>&nbsp;Error Message &nbsp;</strong></td></tr> " +
                                sBodyDetail1 + " </table> ";

            string stextBody = "<p style='font-size:10.0pt;font-family:Arial;'>Dear Admin,<br /><br /> Below is a list of the Customers that have not been Send out the Personal Tax Filing Reminder. Please check. " +
                                ".<br/><br /> " + sTableFormat + "<br/> Thank you. " +
                                "</p>";

            return stextBody;
        }

        private object smtp_SendCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            string smessage = string.Empty;
            if (((e.Error != null)))
            {
                smessage = e.Error.Message;
            }
            return smessage;
        }

        private MailMessage CreateDefaultMailMessage(string MailFrom, string MailTo, string sCntctPrsnName, string sCompanyName, int ToIdentifyTemplate, string sSuffix, ref string sErrDesc)
        {
            MailMessage functionReturnValue = default(MailMessage);

            MailMessage message = new MailMessage();
            string sUploadFile = string.Empty;
            string sFuncName = "CreateDefaultMailMessage";
            string sEmailAddress = string.Empty;
            string sAttachements = string.Empty;
            string sFileName = string.Empty;

            try
            {

                oLog.WriteToDebugLogFile("Sarting function", sFuncName);


                oLog.WriteToDebugLogFile("Assigning Email Properties..", sFuncName);

                message.From = new MailAddress(MailFrom);

                message.To.Add(new MailAddress(MailTo));

                oLog.WriteToDebugLogFile("Adding From Email Address" + ":  " + MailTo, sFuncName);

                message.SubjectEncoding = System.Text.Encoding.UTF8;

                message.Subject = sSubject + " - " + ToIdentifyTemplate + sSuffix + " " + "reminder";
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.Body = ComposeBody(sCntctPrsnName, sCompanyName, ToIdentifyTemplate);
                message.IsBodyHtml = true;

                return message;

            }
            catch (Exception ex)
            {
                sErrDesc = ex.Message;

                oLog.WriteToDebugLogFile("Function completed with Error", sFuncName);
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                functionReturnValue = null;

            }
            finally
            {
            }
            return functionReturnValue;

        }
    }
}
