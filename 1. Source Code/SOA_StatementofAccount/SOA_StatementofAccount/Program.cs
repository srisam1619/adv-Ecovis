using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

namespace SOA_StatementofAccount
{
    class Program
    {
        public static void Main(string[] args)
        {
            string sFuncName = string.Empty;
            clsLog oLog = new clsLog();
            SendEmail oSendMail = new SendEmail();
            string sErrDesc = string.Empty;
            //11/28/2016 11:42:01
            try
            {
                sFuncName = "Main Program";
                oLog.WriteToDebugLogFile("Starting Program", sFuncName);
                Console.WriteLine("Starting Program");
                oLog.WriteToDebugLogFile("Before Getting the Customer details", sFuncName);
                DataSet ds = oSendMail.GetCardCode();
                oLog.WriteToDebugLogFile("After Getting the Customer details", sFuncName);
                DataTable dt = ErrorTable();
                if (ds != null && ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        // Start - Get the Previous month last date based on current date
                        DateTime now = DateTime.Now;
                        DateTime lastDayLastMonth = new DateTime(now.Year, now.Month, 1);
                        lastDayLastMonth = lastDayLastMonth.AddDays(-1);
                        // End - Get the Previous month last date based on current date

                        string sSOADateTime = lastDayLastMonth.ToString("MM/dd/yyyy hh:mm:ss");
                        string sSOADate = lastDayLastMonth.Date.ToString("MM/dd/yyyy");

                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            // Check frequency range
                            string sFreqResult = CheckAging(dr["Code"].ToString(), sSOADate);
                            if (sFreqResult == "SUCCESS")
                            {

                                //Check to avoid duplicate mail sending
                                DataSet dsDuplicateCheck = oSendMail.CheckDuplicateMailSending(dr["Code"].ToString(), dr["Name"].ToString(), sSOADate, dr["Mail"].ToString());

                                if (dsDuplicateCheck.Tables[0].Rows[0]["Status"].ToString() == "Yes")
                                {
                                    oLog.WriteToDebugLogFile("Sending Mail to : " + dr["Mail"].ToString(), sFuncName);
                                    Console.WriteLine("Sending Mail to : " + dr["Mail"].ToString());

                                    string sFileName = dr["Name"].ToString() + "_" + DateTime.Now.ToString("ddMMyyyyhhmmss") + DateTime.Now.Millisecond + ".pdf";
                                    string lReturn = oSendMail.SendAutomatedEmail(dr["Mail"].ToString(), dr["Name"].ToString(), dr["Code"].ToString(), sSOADateTime, sFileName, dr["CompanyName"].ToString(), ref sErrDesc);
                                    if (lReturn == "SUCCESS")
                                    {
                                        oLog.WriteToDebugLogFile("Mail Sent Successfully to : " + dr["Mail"].ToString(), sFuncName);
                                        Console.WriteLine("Mail Sent Successfully to : " + dr["Mail"].ToString());
                                        oLog.WriteToDebugLogFile("Before Inserting to Email Log for : " + dr["Mail"].ToString(), sFuncName);
                                        string sResult = oSendMail.InsertEmailLog(dr["Code"].ToString(), dr["Name"].ToString(), sSOADate, DateTime.Now.ToString(), dr["Mail"].ToString());
                                        oLog.WriteToDebugLogFile("After Inserting to Email Log for : " + dr["Mail"].ToString(), sFuncName);
                                    }
                                    else
                                    {
                                        oLog.WriteToDebugLogFile("Failed to send Mail for: " + dr["Mail"].ToString(), sFuncName);
                                        Console.WriteLine("Failed to send Mail for : " + dr["Mail"].ToString());
                                        dt.Rows.Add(dr["Code"].ToString(), dr["Name"].ToString(), sSOADate, dr["Mail"].ToString(), lReturn.ToString());
                                    }
                                }
                                else
                                {
                                    oLog.WriteToDebugLogFile("Email Previously Sent for : " + dr["Mail"].ToString(), sFuncName);
                                    Console.WriteLine("Email Previously Sent for : " + dr["Mail"].ToString());
                                }
                            }
                            else
                            {
                                oLog.WriteToDebugLogFile("Can't able to send email because aging doesnt have value : ", sFuncName);
                                Console.WriteLine("Can't able to send email because aging doesnt have value");
                            }
                        }
                    }
                }

                //send Email if any email fails on sending
                if (dt.Rows.Count > 0)
                {
                    oLog.WriteToDebugLogFile("Before Sending Email to admin on error method SendEmailOnErrorCase() : ", sFuncName);
                    oSendMail.SendEmailOnErrorCase(dt, ref sErrDesc);
                    oLog.WriteToDebugLogFile("After Sending Email to admin on error method SendEmailOnErrorCase() ", sFuncName);
                }
                Console.WriteLine("Ending Program");
                oLog.WriteToDebugLogFile("Ending Program", sFuncName);

            }
            catch (Exception Ex)
            {
                sErrDesc = Ex.Message.ToString();
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
            }

        }

        static string CheckAging(string sBPCode, string sSOADate)
        {
            string sErrDesc = string.Empty;
            string sFuncName = string.Empty;
            string sResult = string.Empty;
            clsLog oLog = new clsLog();
            SendEmail oMail = new SendEmail();
            try
            {
                DataSet ds = oMail.GetSOADetails(sBPCode, sSOADate);
                if (ds != null && ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        decimal dBracket1 = ds.Tables[0].AsEnumerable().Sum(s => s.Field<decimal>("Bracket1"));
                        decimal dBracket1FC = ds.Tables[0].AsEnumerable().Sum(s => s.Field<decimal>("Bracket1FC"));
                        decimal dFinal1 = dBracket1FC == 0 ? dBracket1 : dBracket1FC;

                        decimal dBracket2 = ds.Tables[0].AsEnumerable().Sum(s => s.Field<decimal>("Bracket2"));
                        decimal dBracket2FC = ds.Tables[0].AsEnumerable().Sum(s => s.Field<decimal>("Bracket2FC"));
                        decimal dFinal2 = dBracket2FC == 0 ? dBracket2 : dBracket2FC;

                        decimal dBracket3 = ds.Tables[0].AsEnumerable().Sum(s => s.Field<decimal>("Bracket3"));
                        decimal dBracket3FC = ds.Tables[0].AsEnumerable().Sum(s => s.Field<decimal>("Bracket3FC"));
                        decimal dFinal3 = dBracket3FC == 0 ? dBracket3 : dBracket3FC;

                        if (dFinal1 > 0 || dFinal2 > 0 || dFinal3 > 0)
                        {
                            sResult = "SUCCESS";
                            oLog.WriteToDebugLogFile("Can able to send email", sFuncName);
                        }
                        else
                        {
                            sResult = "FAILURE";
                            oLog.WriteToDebugLogFile("Can't able to send email because aging doesnt have value", sFuncName);
                        }
                    }
                }
            }
            catch (Exception Ex)
            {
                sErrDesc = Ex.Message.ToString();
                oLog.WriteToErrorLogFile(sErrDesc, sFuncName);
                oLog.WriteToDebugLogFile("Completed With ERROR  ", sFuncName);
                sResult = sErrDesc;
            }
            return sResult;
        }

        static DataTable ErrorTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("CardCode", typeof(string));
            table.Columns.Add("CardName", typeof(string));
            table.Columns.Add("SOADate", typeof(string));
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("ErrorMsg", typeof(string));

            //table.Rows.Add("1", "1", "1", "1", "1");
            return table;
        }
    }
}
