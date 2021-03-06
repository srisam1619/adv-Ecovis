﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Configuration;

namespace SOA_ECIFilingReminder
{
    public class Program
    {
        static void Main(string[] args)
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
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            int ToIdentifyTemplate = 0;
                            string sECIDate = string.Empty;
                            string sSuffix = string.Empty;

                            //string sFinancialYearEnd = ConfigurationManager.AppSettings["FinancialYearEnd"];
                            string sFinancialYearEnd = dr["FinancialYearEnd"].ToString();

                            if (DateTime.Now.Date <= Convert.ToDateTime(sFinancialYearEnd).AddDays(10).Date)
                            {
                                sECIDate = Convert.ToDateTime(sFinancialYearEnd).AddDays(10).ToString("MM/dd/yyyy");
                                ToIdentifyTemplate = 1;
                                sSuffix = "st";
                            }
                            else if (DateTime.Now.Date > Convert.ToDateTime(sFinancialYearEnd).AddDays(10).Date && DateTime.Now.Date <= Convert.ToDateTime(sFinancialYearEnd).AddMonths(2).AddDays(1).Date)
                            {
                                sECIDate = Convert.ToDateTime(sFinancialYearEnd).AddMonths(2).AddDays(1).ToString("MM/dd/yyyy");
                                ToIdentifyTemplate = 2;
                                sSuffix = "nd";
                            }

                            string sSendingEmailDate = DateTime.Now.Date.ToString("MM/dd/yyyy");

                            if (sECIDate == sSendingEmailDate)
                            {

                                //Check to avoid duplicate mail sending
                                DataSet dsDuplicateCheck = oSendMail.CheckDuplicateMailSending(dr["Code"].ToString(), dr["Name"].ToString(), sECIDate, dr["Mail"].ToString());

                                if (dsDuplicateCheck.Tables[0].Rows[0]["Status"].ToString() == "Yes")
                                {
                                    oLog.WriteToDebugLogFile("Sending Mail to : " + dr["Mail"].ToString(), sFuncName);
                                    Console.WriteLine("Sending Mail to : " + dr["Mail"].ToString());

                                    string lReturn = oSendMail.SendAutomatedEmail(dr["Mail"].ToString(), dr["Name"].ToString(), dr["CompanyName"].ToString(), ToIdentifyTemplate, sSuffix, dr["FinancialYearEnd"].ToString(), ref sErrDesc);
                                    if (lReturn == "SUCCESS")
                                    {
                                        oLog.WriteToDebugLogFile("Mail Sent Successfully to : " + dr["Mail"].ToString(), sFuncName);
                                        Console.WriteLine("Mail Sent Successfully to : " + dr["Mail"].ToString());
                                        oLog.WriteToDebugLogFile("Before Inserting to Email Log for : " + dr["Mail"].ToString(), sFuncName);
                                        string sResult = oSendMail.InsertEmailLog(dr["Code"].ToString(), dr["Name"].ToString(), sECIDate, DateTime.Now.ToString(), dr["Mail"].ToString());
                                        oLog.WriteToDebugLogFile("After Inserting to Email Log for : " + dr["Mail"].ToString(), sFuncName);
                                    }
                                    else
                                    {
                                        oLog.WriteToDebugLogFile("Failed to send Mail for: " + dr["Mail"].ToString(), sFuncName);
                                        Console.WriteLine("Failed to send Mail for : " + dr["Mail"].ToString());
                                        dt.Rows.Add(dr["Code"].ToString(), dr["Name"].ToString(), sECIDate, dr["Mail"].ToString(), lReturn.ToString());
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
                                oLog.WriteToDebugLogFile("Current date " + sSendingEmailDate + " doesn't match the reminder date " + sECIDate + " to send email", sFuncName);
                                Console.WriteLine("Current date " + sSendingEmailDate + " doesn't match the reminder date " + sECIDate + " to send email");
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

        static DataTable ErrorTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("CardCode", typeof(string));
            table.Columns.Add("CardName", typeof(string));
            table.Columns.Add("ECIDate", typeof(string));
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("ErrorMsg", typeof(string));

            //table.Rows.Add("1", "1", "1", "1", "1");
            return table;
        }
    }
}
