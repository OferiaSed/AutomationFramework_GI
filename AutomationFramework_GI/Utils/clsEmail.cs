﻿using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutomationFramework_GI.Utils
{
    public class clsEmail
    {
        public string strServer;
        public string strFromEmail;
        public string strPassword;
        public string strToEmail;
        public string strSubject;
        public string strContentBody;
        private string[] arrSTP = new string[2];

        public clsEmail() 
        {
            strServer = "";
            strFromEmail = "";
            strPassword = "";
            strToEmail = "";
            strSubject = "";
            strContentBody = "";
        }


        public void fnSendSimpleEmail() 
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(strFromEmail);
            mail.To.Add(strToEmail);
            mail.Subject = strSubject;
            mail.Body = strContentBody;
            mail.IsBodyHtml = true;
            fnGetServerName(strServer);
            if (strServer != "" && arrSTP[0] != "invalid" && strFromEmail != "" && strPassword != "") 
            {
                SmtpClient smtp = new SmtpClient(arrSTP[0], Convert.ToInt32(arrSTP[1]));
                smtp.Credentials = new NetworkCredential(strFromEmail, strPassword);
                smtp.EnableSsl = true;
                smtp.Send(mail);
            }
        }

        private void fnGetServerName(string pstrServer) 
        {
            switch (strServer.ToUpper()) 
            {
                case "POPGMAIL":
                    arrSTP[0] = "pop.gmail.com";
                    arrSTP[1] = "995";
                    break;
                case "GMAIL":
                    arrSTP[0] = "smtp.gmail.com";
                    arrSTP[1] = "587";
                    break;
                case "OUTLOOK":
                    arrSTP[0] = "smtp.live.com";
                    arrSTP[1] = "587";
                    break;
                case "HOTMAIL":
                    arrSTP[0] = "smtp.live.com";
                    arrSTP[1] = "465";
                    break;
                case "OFFICE365":
                    arrSTP[0] = "smtp.office365.com";
                    arrSTP[1] = "587";
                    break;
                default:
                    arrSTP[0] = "invalid";
                    arrSTP[1] = "invalid";
                    break;
            }
        }

        public string fnReadTwoFactorToken()
        {
            int intTimeAttemp = 0;
            string strToken = "";
            fnGetServerName(strServer);
            //Perform unitl 2Factor is received for 2 minutes
            do
            {
                Pop3Client client = new Pop3Client();
                client.Connect(arrSTP[0], Convert.ToInt32(arrSTP[1]), true);
                client.Authenticate(strFromEmail, strPassword, AuthenticationMethod.UsernameAndPassword);
                int intEmailcount = client.GetMessageCount();
                for (int intRow = intEmailcount; intRow >= 1; intRow--)
                {
                    Message message = client.GetMessage(intRow);
                    MessagePart messagepart = message.FindFirstPlainTextVersion();
                    if (messagepart != null)
                    {
                        //Get Token as Plan Text
                        string strTemp = messagepart.GetBodyAsText();
                        if (strTemp.Contains("code is:"))
                        {
                            string[] arrSeparators = { "code is:", "\r\nThe code" };
                            string[] arrToken = strTemp.Split(arrSeparators, System.StringSplitOptions.RemoveEmptyEntries);
                            strToken = arrToken[1];
                            client.DeleteMessage(intRow);
                            break;
                        }
                    }
                    else
                    {
                        //Get Token as HTML
                        messagepart = message.FindFirstHtmlVersion();
                        if (messagepart != null)
                        {
                            string strTemp = messagepart.GetBodyAsText();
                            if (strTemp.Contains("code is:"))
                            {
                                string[] arrSeparators = { "code is:", "<br/>" };
                                string[] arrToken = strTemp.Split(arrSeparators, System.StringSplitOptions.RemoveEmptyEntries);
                                strToken = arrToken[1];
                                client.DeleteMessage(intRow);
                                break;
                            }
                        }
                    }
                }
                intTimeAttemp++;
                if (strToken == "") { Thread.Sleep(TimeSpan.FromSeconds(10)); }
                client.Disconnect();
            }
            while (intTimeAttemp < 12 && strToken == "");
            return strToken;
        }

        public string fnReadForgotPassword()
        {
            int intTimeAttemp = 0;
            string strURL = "";
            fnGetServerName(strServer);
            do
            {
                Pop3Client client = new Pop3Client();
                client.Connect(arrSTP[0], Convert.ToInt32(arrSTP[1]), true);
                client.Authenticate(strFromEmail, strPassword, AuthenticationMethod.UsernameAndPassword);
                int intEmailcount = client.GetMessageCount();
                for (int intRow = intEmailcount; intRow >= 1; intRow--)
                {
                    Message message = client.GetMessage(intRow);
                    MessagePart messagepart = message.FindFirstPlainTextVersion();
                    if (messagepart != null)
                    {
                        //Get Token as Plan Text
                        string strTemp = messagepart.GetBodyAsText();
                        if (strTemp.Contains("reset your password"))
                        {
                            string[] arrSeparators = { "your browser: ", "---" };
                            string[] arrToken = strTemp.Split(arrSeparators, System.StringSplitOptions.RemoveEmptyEntries);
                            strURL = arrToken[1].Replace("\n", "").Replace("\r", "");
                            //client.DeleteMessage(intRow);
                            break;
                        }
                    }
                    else
                    {
                        //Get Token as HTML
                        messagepart = message.FindFirstHtmlVersion();
                        if (messagepart != null)
                        {
                            string strTemp = messagepart.GetBodyAsText();
                            if (strTemp.Contains("reset your password"))
                            {
                                string[] arrSeparators = { "your browser: ", "---" };
                                string[] arrToken = strTemp.Split(arrSeparators, System.StringSplitOptions.RemoveEmptyEntries);
                                strURL = arrToken[1].Replace("\n", "").Replace("\r", "");
                                //client.DeleteMessage(intRow);
                                break;
                            }
                        }
                    }
                }
                intTimeAttemp++;
                if (strURL == "") { Thread.Sleep(TimeSpan.FromSeconds(10)); }
                client.Disconnect();
            }
            while (intTimeAttemp < 12 && strURL == "") ;
            return strURL;
        }

        public bool fnReadSimpleEmail(string pstrText)
        {
            int intTimeAttemp = 0;
            bool bFound = false;
            fnGetServerName(strServer);
            do
            {
                Pop3Client client = new Pop3Client();
                client.Connect(arrSTP[0], Convert.ToInt32(arrSTP[1]), true);
                client.Authenticate(strFromEmail, strPassword, AuthenticationMethod.UsernameAndPassword);
                int intEmailcount = client.GetMessageCount();
                for (int intRow = intEmailcount; intRow >= 1; intRow--)
                {
                    Message message = client.GetMessage(intRow);
                    MessagePart messagepart = message.FindFirstPlainTextVersion();
                    if (messagepart != null)
                    {
                        //Get Token as Plan Text
                        string strTemp = messagepart.GetBodyAsText();
                        if (strTemp.Contains(pstrText))
                        {
                            bFound = true;
                            //client.DeleteMessage(intRow);
                            break;
                        }
                    }
                    else
                    {
                        //Get Token as HTML
                        messagepart = message.FindFirstHtmlVersion();
                        if (messagepart != null)
                        {
                            string strTemp = messagepart.GetBodyAsText();
                            if (strTemp.Contains("reset your password"))
                            {
                                bFound = true;
                                //client.DeleteMessage(intRow);
                                break;
                            }
                        }
                    }
                }
                intTimeAttemp++;
                if (!bFound) { Thread.Sleep(TimeSpan.FromSeconds(10)); }
                client.Disconnect();
            }
            while (intTimeAttemp < 12 && !bFound);
            return bFound;
        }

    }
}