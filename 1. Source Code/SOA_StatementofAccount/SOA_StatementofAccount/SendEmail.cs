using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Configuration;
using System.Data;
using System.IO;
using System.Web;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Globalization;

namespace SOA_StatementofAccount
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
                sProcName = "AB_SOA_SP002_GetCardCodes";
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

        public DataSet CheckDuplicateMailSending(string sCode, string sName, string sSOADate, string Mail)
        {
            string sFuncName = string.Empty;
            DataSet oDataset = new DataSet();
            string sProcName = string.Empty;
            try
            {
                sFuncName = "CheckDuplicateMailSending";
                //string sInsertQuery = "insert into [@EMAILLOG_SOA](" + sColumnNames + ") values(" + sParameter + ")";
                sProcName = "AB_SOA_SP004_CheckDuplicateMailSending";
                oLog.WriteToDebugLogFile("Starting Function ", sFuncName);

                oLog.WriteToDebugLogFile("Calling Run_StoredProcedure() " + sProcName, sFuncName);
                oDataset = SqlHelper.ExecuteDataSet(ConnectionString, CommandType.StoredProcedure, sProcName, Data.CreateParameter("@CardCode", sCode),
                    Data.CreateParameter("@CardName", sName), Data.CreateParameter("@SOADate", sSOADate), Data.CreateParameter("@Email", Mail));

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

        public DataSet GetSOADetails(string sCode, string sSOADate)
        {
            string sFuncName = string.Empty;
            DataSet oDataset = new DataSet();
            string sProcName = string.Empty;
            try
            {
                sFuncName = "GetSOADetails";
                sProcName = "AB_SOA_SP001";
                oLog.WriteToDebugLogFile("Starting Function ", sFuncName);

                oLog.WriteToDebugLogFile("Calling Run_StoredProcedure() " + sProcName, sFuncName);
                oDataset = SqlHelper.ExecuteDataSet(ConnectionString, CommandType.StoredProcedure, sProcName, Data.CreateParameter("@BPFrom", sCode),
                    Data.CreateParameter("@BPTo", sCode), Data.CreateParameter("@DateTo", sSOADate));

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

        public string ComposeBody(string sContPrsnName, string sCntctPrsnCode, string sStatementDate, string sCompanyName)
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

                string sTemplatePath = AppDomain.CurrentDomain.BaseDirectory;
                int index = sTemplatePath.IndexOf("\\bin");
                if (index > 0)
                    sTemplatePath = sTemplatePath.Substring(0, index) + "\\Email Template\\EmailContent.htm";

                using (var sReader = new StreamReader(sTemplatePath))
                {
                    sbMail.Append(sReader.ReadToEnd());

                    sbMail.Replace("{ContactPersonName}", System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(sContPrsnName.ToLower()));
                    sbMail.Replace("{DateofStatement}", sStatementDate);
                    sbMail.Replace("{CompanyName}", System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(sCompanyName.ToLower()));

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

        public string CreatePDF(string sContPrsnName, string sCntctPrsnCode, string sStatementDate, string sFileName)
        {
            string sFuncName = string.Empty;
            string AttachFile = string.Empty;
            try
            {
                sFuncName = "CreatePDF";
                ReportDocument cryRpt = new ReportDocument();
                CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
                CrystalDecisions.Shared.TableLogOnInfo oLogonInfo = new CrystalDecisions.Shared.TableLogOnInfo();
                int intCounter = 0;

                string Basedirectory = AppDomain.CurrentDomain.BaseDirectory;
                int baseIndex = Basedirectory.IndexOf("\\bin");
                if (baseIndex > 0)
                    Basedirectory = Basedirectory.Substring(0, baseIndex);

                string directory = AppDomain.CurrentDomain.BaseDirectory;
                int index = directory.IndexOf("\\bin");
                if (index > 0)
                    directory = directory.Substring(0, index) + "\\PDF";

                AttachFile = AppDomain.CurrentDomain.BaseDirectory;
                int index1 = AttachFile.IndexOf("\\bin");
                if (index1 > 0)
                    AttachFile = AttachFile.Substring(0, index1) + "\\PDF\\" + sFileName;

                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                cryRpt.Load(Basedirectory + "\\Report\\AB__SOA_RP003.rpt");

                ParameterValues crParameterValues = new ParameterValues();
                ParameterDiscreteValue crParameterDiscreteValue = new ParameterDiscreteValue();

                oLogonInfo.ConnectionInfo.ServerName = sServer;
                oLogonInfo.ConnectionInfo.DatabaseName = sDatabase;
                oLogonInfo.ConnectionInfo.UserID = sUid;
                oLogonInfo.ConnectionInfo.Password = sPwd;

                for (intCounter = 0; intCounter <= cryRpt.Database.Tables.Count - 1; intCounter++)
                {
                    cryRpt.Database.Tables[intCounter].ApplyLogOnInfo(oLogonInfo);
                }


                cryRpt.SetParameterValue("BPFrom@SELECT * FROM OCRD WHERE CARDTYPE='C'", sCntctPrsnCode);
                cryRpt.SetParameterValue("BPTo@SELECT * FROM OCRD WHERE CARDTYPE='C'", sCntctPrsnCode);
                cryRpt.SetParameterValue("@DateTo", sStatementDate);

                ExportOptions CrExportOptions = default(ExportOptions);
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                ExcelFormatOptions CrExcelFormat = new ExcelFormatOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                ExcelFormatOptions CrExcelTypeOptions = new ExcelFormatOptions();


                CrDiskFileDestinationOptions.DiskFileName = AttachFile;
                CrExportOptions = cryRpt.ExportOptions;
                var _with1 = CrExportOptions;
                _with1.ExportDestinationType = ExportDestinationType.DiskFile;
                _with1.ExportFormatType = ExportFormatType.PortableDocFormat;
                _with1.DestinationOptions = CrDiskFileDestinationOptions;
                _with1.FormatOptions = CrFormatTypeOptions;
                cryRpt.Export();

                oLog.WriteToDebugLogFile("PDF Created successfully for Customer Code : " + sCntctPrsnCode, sFuncName);
            }
            catch (Exception ex)
            {
                oLog.WriteToDebugLogFile("Function completed with Error", sFuncName);
                oLog.WriteToErrorLogFile(ex.Message, sFuncName);
            }
            return AttachFile;
        }

        public string InsertEmailLog(string sCode, string sName, string sSOADate, string sDateSent, string Mail)
        {
            string sResult = string.Empty;
            string sFuncName = string.Empty;
            DataSet oDataset = new DataSet();
            string sProcName = string.Empty;
            try
            {
                sFuncName = "InsertEmailLog";
                //string sInsertQuery = "insert into [@EMAILLOG_SOA](" + sColumnNames + ") values(" + sParameter + ")";
                sProcName = "AB_SOA_SP003_InsertSOALog";
                oLog.WriteToDebugLogFile("Starting Function ", sFuncName);

                oLog.WriteToDebugLogFile("Calling Run_StoredProcedure() " + sProcName, sFuncName);
                oDataset = SqlHelper.ExecuteDataSet(ConnectionString, CommandType.StoredProcedure, sProcName, Data.CreateParameter("@CardCode", sCode),
                    Data.CreateParameter("@CardName", sName), Data.CreateParameter("@SOADate", sSOADate), Data.CreateParameter("@DateSent", sDateSent),
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

        public string SendAutomatedEmail(string sEmailTo, string sCntctPrsnName, string sCntctPrsnCode, string sStatementDate, string sFileName, string sCompanyName, ref string sErrDesc)
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

                MailMessage message = CreateDefaultMailMessage(sFromMailId, sEmailTo, sCntctPrsnName, sCntctPrsnCode, sStatementDate, sCompanyName, ref sErrDesc);
                object userState = message;

                string filePath = CreatePDF(sCntctPrsnName, sCntctPrsnCode, sStatementDate, sFileName);
                //message.Attachments = 
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(filePath);
                message.Attachments.Add(attachment);

                oLog.WriteToDebugLogFile("Sending Email Message", sFuncName);

                oLog.WriteToDebugLogFile("Sending Email Messages to : " + sEmailTo, sFuncName);

                smtpClient.Send(message);

                message.Dispose();

                oLog.WriteToDebugLogFile("After Sending Mail , Before Deleting the attachment", sFuncName);
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    oLog.WriteToDebugLogFile("After Sending Mail , after Deleting the attachment", sFuncName);
                }
                else
                {
                    oLog.WriteToDebugLogFile("After Sending Mail , attachment file is not there", sFuncName);
                }


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
                    " <td>&nbsp;" + item["SOADate"].ToString() + " </td><td>&nbsp;" + item["Email"].ToString() + " </td> " +
                    " <td>&nbsp;" + item["ErrorMsg"].ToString() + " </td></tr>";
                sBodyDetail1 = sBodyDetail1 + sBodyDetail;
            }
            sTableFormat = "<table border = '1' cellspacing = 0 cellpadding = 0 style='font-size:10.0pt;font-family:Arial;width: 85%;'> " +
                                "<tr><td><strong style='color: blue; background-color: transparent;'>&nbsp;Customer Code&nbsp;</strong></td> " +
                                "<td><strong style='color: blue; background-color: transparent;'>&nbsp;Customer Name &nbsp; </strong></td> " +
                                "<td><strong style='color: blue; background-color: transparent;'>&nbsp;SOA Date &nbsp;</strong></td> " +
                                "<td><strong style='color: blue; background-color: transparent;'>&nbsp;Email Id &nbsp;</strong></td> " +
                                "<td><strong style='color: blue; background-color: transparent;'>&nbsp;Error Message &nbsp;</strong></td></tr> " +
                                sBodyDetail1 + " </table> ";

            string stextBody = "<p style='font-size:10.0pt;font-family:Arial;'>Dear Admin,<br /><br /> Below is a list of the Customers that have not been Send out the SOA. Please check. " +
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

        private MailMessage CreateDefaultMailMessage(string MailFrom, string MailTo, string sCntctPrsnName, string sCntctPrsnCode, string sStatementDate, string sCompanyName, ref string sErrDesc)
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

                string sFormattedDate = string.Empty;
                DateTime dt;
                if (DateTime.TryParseExact(sStatementDate, "MM/dd/yyyy hh:mm:ss",
                                           CultureInfo.InvariantCulture, DateTimeStyles.None,
                                           out dt))
                {
                    //sFormattedDate = dt.ToString("dd MMMM yyyy", CultureInfo.InvariantCulture);
                    sFormattedDate = string.Format(new MyCustomDateProvider(), "{0}", dt);
                }
                else
                {
                    // Handle failure
                }

                message.Subject = sSubject + " - " + sFormattedDate;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.Body = ComposeBody(sCntctPrsnName, sCntctPrsnCode, sFormattedDate, sCompanyName);
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
