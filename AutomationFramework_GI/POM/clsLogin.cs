using AutomationFramework;
using AutomationFramework_GI.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutomationFramework_GI.POM
{
    class clsLogin
    {
        private clsWebElements clsWE = new clsWebElements();
        private clsMegaIntake mega = new clsMegaIntake();

        //Page Object Methods
        public bool fnLoginData(string pstrSetNo)
        {
            bool blResult = true;
            clsData objData = new clsData();
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    clsReportResult.fnLog("Login Function", "The Login Functions Starts.", "Info", false);
                    clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, true);
                    if (clsWE.fnElementExist("Accept Cookies Message", "//button[@id='cookie-accept']", true)) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                    clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", false);
                    clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", objData.fnGetValue("User", ""), false);
                    clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", false);
                    clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", objData.fnGetValue("Password", ""), true);
                    clsWE.fnClick(clsWE.fnGetWe("//button[text()='BEGIN']"), "Begin", false);
                    blResult = clsWE.fnElementExist("Login Label", "//span[contains(text(), 'You are currently logged into')]", false, false);
                    if (blResult)
                    { clsReportResult.fnLog("Login Page", "Login Page was successfully.", "Info", true); }
                    else
                    { clsReportResult.fnLog("Login Page", "Login Page was not successfully.", "Info", true, true); }
                }
            }
            return blResult;
        }

        public bool fnTwoFactorsVerification(string pstrSetNo)
        {
            bool blResult = true;
            clsData objData = new clsData();
            clsReportResult.fnLog("Two Factor Authentication", "<<<<<<<<<< Two Factor Authentication Function Starts. >>>>>>>>>>", "Info", false);
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Enter Credentails
                    fnLogOffSession();
                    fnEnterCredentails(objData);
                    if (mega.IsElementPresent("//h3[text()='Multifactor Authentication']"))
                    {
                        //Select method & send code
                        mega.fnSelectDropDownWElm("Authentication Method", "//input[@class='select-dropdown form-control']", "Email", false, false);
                        clsWE.fnClick(clsWE.fnGetWe("//button[text()='Send Code']"), "Send Code", false);
                        //Get Email Token
                        string strToken = fnReadEmailConfirmation(objData.fnGetValue("EmailAcc", ""), objData.fnGetValue("PassAcc", ""), "code is:", "code is:", "\r\nThe code", "code is:", "<br/>");
                        //Verify if token was received
                        if (strToken != "")
                        {
                            clsReportResult.fnLog("Two Factor Authentication", "The 2FA email was received with value: " + strToken, "Pass", false, false);
                            mega.fnEnterTextWElm("Enter Code", "//input[@id='code']", strToken, true, false);
                            clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit", false);
                        }
                        else
                        {
                            clsReportResult.fnLog("Two Factor Value", "The 2FA email was not received after 2 minutes.", "Fail", false, false);
                            blResult = false;
                        }
                        //Verify Login Page
                        if (clsWE.fnElementExist("Login Label", "//span[contains(text(), 'You are currently logged into')]", false, false))
                        { clsReportResult.fnLog("Two Factor Authentication", "The Login with 2FA was done as expected.", "Info", false, false); }
                        else
                        { 
                            clsReportResult.fnLog("Two Factor Authentication", "The Login with 2FA was not completed as expected.", "Fail", true, true);
                            blResult = false;
                        }
                    }
                    else
                    {
                        //Verify Error Messages
                        if (clsWE.fnElementExist("Login  - Error Message", "//div[@class='validation-summary-errors']", false, false))
                        {
                            { clsReportResult.fnLog("Two Factor Authentication", "Some errors were found after privide 2FA credentials.", "Fail", true, true); }
                            blResult = false;
                        }
                    }
                }
            }
            if (blResult)
            { clsReportResult.fnLog("Two Factor Authentication", "The Two Factor Authentication was executed successfully.", "Pass", true, false); }
            else
            { clsReportResult.fnLog("Two Factor Authentication", "Two Factor Authentication was not executed successfully.", "Fail", false, true); }
            return blResult;
        }

        public bool fnForgotPasswordVerification(string pstrSetNo)
        {
            bool blResult = true;
            clsData objData = new clsData();
            clsReportResult.fnLog("Forgot Password", "<<<<<<<<<< Forfot Password Function Starts. >>>>>>>>>>", "Info", false);
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Verify if "Forgot Password" link exist
                    fnLogOffSession();
                    clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
                    //if (clsWE.fnElementExist("Accept Cookies Message", "//button[@id='cookie-accept']", true)) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                    if (mega.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                    if (clsWE.fnElementExist("Forgot Password Link", "//a[text()='Forgot Your Password?']", false))
                    {
                        //verify that link is working and error messages displayed.
                        clsWE.fnClick(clsWE.fnGetWe("//a[text()='Forgot Your Password?']"), "Forgot Password", false);
                        if (mega.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                        if (clsWE.fnElementExist("Forgot Password Page", "//h3[text()='Forgot Password']", false))
                        {
                            //Verify Required Fields
                            clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit Button", false);
                            if (clsWE.fnElementExist("The Username/Captcha field is required.", "//div[@class='invalid-feedback']", false))
                            {
                                //Enter username/captcha
                                clsReportResult.fnLog("Forgot Password", "The required field messages are displayed as expected.", "info", true, true);
                                mega.fnEnterTextWElm("Username*", "//input[@id='uname']", objData.fnGetValue("User", ""), true, false);
                                do { Thread.Sleep(TimeSpan.FromSeconds(10)); }
                                while (clsWE.fnGetAttribute(clsWE.fnGetWe("//input[@id='captcha-input']"), "Captcha", "value", false, false) == "");
                                clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit", false);
                                //Verify that email is received to change the password
                                //string strURLReset = fnForgotPasswordURL();
                                string strURLReset = fnReadEmailConfirmation(objData.fnGetValue("EmailAcc", ""), objData.fnGetValue("PassAcc", ""), "reset your password", "your browser: ", "---", "your browser: ", "</span>");
                                if (strURLReset != "")
                                {
                                    clsWebBrowser.objDriver.Navigate().GoToUrl(strURLReset);
                                    //Enter New Password and save it
                                    string strNewPass = RandomString(8);
                                    mega.fnEnterTextWElm("New Password", "//input[@id='new-pwd']", strNewPass, false, false);
                                    mega.fnEnterTextWElm("Confirm New Password", "//input[@id='new-pwd-v']", strNewPass, false, false);
                                    clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit", true);
                                    //Verify that password was changes successfully
                                    if (clsWE.fnElementExist("Your password has been successfully changed", "//div[contains(text(), 'Your password has been successfully')]", false))
                                    {
                                        //Save the claim
                                        clsData objSaveData = new clsData();
                                        objSaveData.fnSaveValue(ConfigurationManager.AppSettings["FilePath"], "LogInData", "Password", intRow, strNewPass);

                                        //Verify email confirmation for password reset
                                        if (fnReadTextEmail(objData.fnGetValue("EmailAcc", ""), objData.fnGetValue("PassAcc", ""), "password for your account was changed"))
                                        {
                                            clsReportResult.fnLog("Forgot Password", "The Password Change Confirmation email was received as expected.", "info", false, false);
                                        }
                                        else 
                                        {
                                            clsReportResult.fnLog("Forgot Password", "The Password Change Confirmation email was not received.", "Fail", false, true);
                                            blResult = false;
                                        }
                                    }
                                    else
                                    {
                                        clsReportResult.fnLog("Forgot Password", "The new password was not entered successfully.", "Fail", true, true);
                                        blResult = false;
                                    }
                                }
                                else 
                                {
                                    clsReportResult.fnLog("Forgot Password", "The email/url was not received to reset the password.", "Fail", false, true);
                                    blResult = false;
                                }
                            }
                            else 
                            {
                                clsReportResult.fnLog("Forgot Password", "The invalid messages for Username/Captcha are not displayed.", "Fail", true, true);
                                blResult = false;
                            }
                        }
                        else 
                        {
                            clsReportResult.fnLog("Forgot Password", "The Forgot Password Page was not loaded after click on the link.", "Fail", true, true);
                            blResult = false;
                        }
                    }
                    else 
                    {
                        clsReportResult.fnLog("Forgot Password", "The forgot password link was not found in hte page.", "Fail", true, true);
                        blResult = false;
                    }
                }
            }
            if (blResult)
            { clsReportResult.fnLog("Forgot Password", "The Forgot Password Function was executed successfully", "Pass", false, false); }
            else
            { clsReportResult.fnLog("Forgot Password", "The Forgot Password Function was executed not successfully", "Fail", false, false); }
            return blResult;
        }

        public bool fnForgotUsernameVerification(string pstrSetNo)
        {
            bool blResult = true;
            clsData objData = new clsData();
            string strUsername = "";
            clsReportResult.fnLog("Forgot Username", "<<<<<<<<<< Two Forgot Username Function Starts. >>>>>>>>>>", "Info", false);
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Verify if "Forgot Password" link exist
                    fnLogOffSession();
                    clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
                    //if (clsWE.fnElementExist("Accept Cookies Message", "//button[@id='cookie-accept']", true)) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                    if (mega.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                    if (clsWE.fnElementExist("Forgot Password Link", "//a[text()='Forgot Username?']", false))
                    {
                        //verify that link is working and error messages displayed.
                        clsWE.fnClick(clsWE.fnGetWe("//a[text()='Forgot Username?']"), "Forgot Username", false);
                        clsWE.fnPageLoad(clsWE.fnGetWe("//h3[text()='Forgot Username']"), "Forgot Username Page", false, false);
                        if (mega.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                        if (clsWE.fnElementExist("Forgot Username Page", "//h3[text()='Forgot Username']", false))
                        {
                            clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit Button", false);
                            if (clsWE.fnElementExist("The Email/Captcha field is required.", "//div[@class='invalid-feedback']", false))
                            {
                                //Enter username/captcha
                                clsReportResult.fnLog("Forgot Username", "The required field messages are displayed as expected.", "info", true, false);
                                mega.fnEnterTextWElm("Email", "//input[@id='uname']", objData.fnGetValue("EmailAcc", ""), false, false);
                                do { Thread.Sleep(TimeSpan.FromSeconds(10)); }
                                while (clsWE.fnGetAttribute(clsWE.fnGetWe("//input[@id='captcha-input']"), "Captcha", "value", false, false) == "");
                                clsReportResult.fnLog("Forgot Username", "The email/captcha was entered", "Info", true, false);
                                clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit", true);
                                if (!mega.IsElementPresent("//div[@class='invalid-feedback']"))
                                {
                                    //Verify that email is received with username
                                    strUsername = fnReadEmailConfirmation(objData.fnGetValue("EmailAcc", ""), objData.fnGetValue("PassAcc", ""), "username(s) associated", "are:\r\n\r\n", "\r\n\r\n", "are:</br> <ul><li>", "</li></ul>.");
                                    if (strUsername != "")
                                    {
                                        clsReportResult.fnLog("Forgot Username", "The Username email request was received as expected. The user is: " + strUsername, "Pass", false, false);
                                    }
                                    else
                                    {
                                        clsReportResult.fnLog("Forgot Username", "The Username email request was not received and cannot be validated.", "Fail", false, true);
                                        blResult = false;
                                    }
                                }
                                else 
                                {
                                    clsReportResult.fnLog("Forgot Username", "Some issues were found after procide email/captcha and test cannot continue", "Fail", true, true);
                                    blResult = false;
                                }
                            }
                            else 
                            {
                                clsReportResult.fnLog("Forgot Username", "The invalid messages for Username/Captcha are not displayed.", "Fail", true, true);
                                blResult = false;
                            }
                        }
                        else 
                        {
                            clsReportResult.fnLog("Forgot Username", "The Forgot Username Page was not loaded after click on the link.", "Fail", true, true);
                            blResult = false;
                        }
                    }
                    else 
                    {
                        clsReportResult.fnLog("Forgot Username", "The forgot username link was not found in hte page.", "Fail", true, true);
                        blResult = false;
                    }
                }
            }
            if (blResult)
            { clsReportResult.fnLog("Forgot Username", "The Forgot Username function was executed successfully.", "Pass", false, false);  }
            else
            { clsReportResult.fnLog("Forgot Username", "The Forgot Username function was not executed successfully.", "Pass", false, false);  }
            return blResult;
        }


        //Functions or methods section
        private string RandomString(int plength)
        {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLmnopqrstuvwxyz0123456789";
            const string special_chars = "!$%&";
            string strRnd = new string(Enumerable.Repeat(chars, plength).Select(s => s[random.Next(s.Length)]).ToArray());
            string strRnd2 = new string(Enumerable.Repeat(special_chars, 2).Select(s => s[random.Next(s.Length)]).ToArray());
            return strRnd + strRnd2;
        }

        public bool fnReadTextEmail(string pstrUser, string pstrPassword, string pstrVal)
        {
            clsEmail email = new clsEmail();
            email.strFromEmail = pstrUser;
            email.strPassword = pstrPassword;
            email.strServer = "popgmail";
            return email.fnReadSimpleEmail(pstrUser, pstrPassword, pstrVal);
        }

        private void fnEnterCredentails(clsData pobjData)
        {
            //Wait to load Page and verify if Cookies button appears
            clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
            //if (clsWE.fnElementExist("Accept Cookies Message", "//button[@id='cookie-accept']", true)) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
            if (mega.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
            //Enter Credentials
            clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", false);
            clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", pobjData.fnGetValue("User", ""), false);
            clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", false);
            clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", pobjData.fnGetValue("Password", ""), true);
            //Click on Begin button
            clsWE.fnClick(clsWE.fnGetWe("//button[text()='BEGIN']"), "Begin", false);
        }

        private string fnReadEmailConfirmation(string pstrEmail, string pstrPassword, string pstrContainsText, string pstrStartWithPlainText, string pstrEndwithPlainText, string pstrStartWithHtml, string pstrEndwithHtml)
        {
            clsEmail email = new clsEmail();
            email.strFromEmail = pstrEmail;
            email.strPassword = pstrPassword;
            email.strServer = "popgmail";
            string strValue = email.fnReadEmailText(pstrContainsText, pstrStartWithPlainText, pstrEndwithPlainText, pstrStartWithHtml, pstrEndwithHtml);
            return strValue;
        }

        public bool fnLogOffSession() 
        {
            bool blResult = true;
            //Verify is there is an active session
            if (mega.IsElementPresent("//span[text()='You are currently logged into ']")) 
            {
                clsWE.fnClick(clsWE.fnGetWe("//li[@class='nav-item m-menuitem-show']/a[@id='topmenu-logout']"), "Logout Link", false, false);
                clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
                clsReportResult.fnLog("Logout Session", "An active session was terminated", "Info", false, false);
            }
            return blResult;
        }

        private bool Template(string pstrSetNo) 
        {
            bool blResult = true;
            clsData objData = new clsData();
            clsReportResult.fnLog("Two Factor Authentication", "Two Factor Authentication Function Starts.", "Info", false);
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    
                }
            }
            return blResult;
        }

    }
}
