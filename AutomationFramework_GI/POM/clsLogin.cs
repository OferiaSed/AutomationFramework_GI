using AutomationFramework;
using AutomationFramework_GI.Utils;
using OpenQA.Selenium;
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
        private clsMegaIntake clsMG = new clsMegaIntake();

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
                    //clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", false);
                    //clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", objData.fnGetValue("User", ""), false);
                    //clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", false);
                    //clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", objData.fnGetValue("Password", ""), true);
                    //clsWE.fnClick(clsWE.fnGetWe("//button[text()='BEGIN']"), "Begin", false);
                    fnEnterCredentails_New(objData.fnGetValue("User", ""), objData.fnGetValue("Password", ""));

                    blResult = clsWE.fnElementExist("Login Label", "//span[contains(text(), 'You are currently logged into')]", false, false);
                    if (blResult)
                    { clsReportResult.fnLog("Login Page", "Login Page was successfully.", "Pass", true); }
                    else
                    { clsReportResult.fnLog("Login Page", "Login Page was not successfully.", "Fail", true, true); }
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
                    fnEnterCredentails(objData.fnGetValue("User", ""), objData.fnGetValue("Password", ""));
                    if (clsMG.IsElementPresent("//h3[text()='Multifactor Authentication']"))
                    {
                        //Select method & send code
                        clsMG.fnSelectDropDownWElm("Authentication Method", "//input[@class='select-dropdown form-control']", "Email", false, false);
                        clsWE.fnClick(clsWE.fnGetWe("//button[text()='Send Code']"), "Send Code", false);
                        //Get Email Token
                        string strToken = fnReadEmailConfirmation(objData.fnGetValue("EmailAcc", ""), objData.fnGetValue("PassAcc", ""), "code is:", "code is:", "\r\nThe code", "code is:", "<br/>");
                        //Verify if token was received
                        if (strToken != "")
                        {
                            clsReportResult.fnLog("Two Factor Authentication", "The 2FA email was received with value: " + strToken, "Pass", false, false);
                            clsMG.fnEnterTextWElm("Enter Code", "//input[@id='code']", strToken, true, false);
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
                    if (clsMG.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                    if (clsWE.fnElementExist("Forgot Password Link", "//a[text()='Forgot Your Password?']", false))
                    {
                        //verify that link is working and error messages displayed.
                        clsWE.fnClick(clsWE.fnGetWe("//a[text()='Forgot Your Password?']"), "Forgot Password", false);
                        if (clsMG.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                        if (clsWE.fnElementExist("Forgot Password Page", "//h3[text()='Forgot Password']", false))
                        {
                            //Verify Required Fields
                            clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit Button", false);
                            if (clsWE.fnElementExist("The Username/Captcha field is required.", "//div[@class='invalid-feedback']", false))
                            {
                                //Enter username/captcha
                                clsReportResult.fnLog("Forgot Password", "The required field messages are displayed as expected.", "info", true, true);
                                clsMG.fnEnterTextWElm("Username*", "//input[@id='uname']", objData.fnGetValue("User", ""), true, false);
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
                                    clsMG.fnEnterTextWElm("New Password", "//input[@id='new-pwd']", strNewPass, false, false);
                                    clsMG.fnEnterTextWElm("Confirm New Password", "//input[@id='new-pwd-v']", strNewPass, false, false);
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
                    if (clsMG.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                    if (clsWE.fnElementExist("Forgot Password Link", "//a[text()='Forgot Username?']", false))
                    {
                        //verify that link is working and error messages displayed.
                        clsWE.fnClick(clsWE.fnGetWe("//a[text()='Forgot Username?']"), "Forgot Username", false);
                        clsWE.fnPageLoad(clsWE.fnGetWe("//h3[text()='Forgot Username']"), "Forgot Username Page", false, false);
                        if (clsMG.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                        if (clsWE.fnElementExist("Forgot Username Page", "//h3[text()='Forgot Username']", false))
                        {
                            clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit Button", false);
                            if (clsWE.fnElementExist("The Email/Captcha field is required.", "//div[@class='invalid-feedback']", false))
                            {
                                //Enter username/captcha
                                clsReportResult.fnLog("Forgot Username", "The required field messages are displayed as expected.", "info", true, false);
                                clsMG.fnEnterTextWElm("Email", "//input[@id='uname']", objData.fnGetValue("EmailAcc", ""), false, false);
                                do { Thread.Sleep(TimeSpan.FromSeconds(10)); }
                                while (clsWE.fnGetAttribute(clsWE.fnGetWe("//input[@id='captcha-input']"), "Captcha", "value", false, false) == "");
                                clsReportResult.fnLog("Forgot Username", "The email/captcha was entered", "Info", true, false);
                                clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit", true);
                                if (!clsMG.IsElementPresent("//div[@class='invalid-feedback']"))
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

        public bool fnExpiredUserRestriction(string pstrSetNo)
        {
            bool blResult = true;
            clsData objData = new clsData();
            clsReportResult.fnLog("Restriction for Expired User", "<<<<<<<<<< Restriction for Expired User Function Starts. >>>>>>>>>>", "Info", false);
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Close active session and verify if cookies button exist
                    fnLogOffSession();
                    clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
                    if (clsMG.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
                    //Enter Credentials 4 time with invalid password
                    int intAttemps = 1;
                    do
                    {
                        clsReportResult.fnLog("Restriction for Expired User", "Entering Invalid Credentials attemp #" + intAttemps, "Info", false);
                        fnEnterCredentails_New(objData.fnGetValue("User", ""), "PSF$%&SDFSDF");
                        clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
                        intAttemps++;
                    }
                    while (intAttemps <= 4);
                    //Enter Valid Credentials
                    clsReportResult.fnLog("Restriction for Expired User", "Entering Valid Credentials.", "Info", false);
                    fnEnterCredentails_New(objData.fnGetValue("User", ""), objData.fnGetValue("Password", ""));
                    //verify that security questions are triggered.
                    if (clsWE.fnElementExist("Account Unlock Page", "//h3[text()='Account Unlock']", true))
                    {
                        clsReportResult.fnLog("Restriction for Expired User", "The Account Unlock Page was displayed as expected.", "Info", false);
                        IList<IWebElement> lsQuestions = clsWebBrowser.objDriver.FindElements(By.XPath("//p/strong"));
                        for (intAttemps = 1; intAttemps <= 4; intAttemps++) 
                        {
                            clsReportResult.fnLog("Restriction for Expired User", "Entering Invalid Securtity Answer attemp #" + intAttemps, "Info", false);
                            for (int intQuestion = 1; intQuestion <= lsQuestions.Count; intQuestion++)
                            {
                                clsMG.fnCleanAndEnterText("Security Question #" + intQuestion, "(//input[@type='password'])[" + intQuestion + "]", RandomString(8), false, false, "", false);
                            }
                            clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Begin", false);
                            clsWE.fnPageLoad(clsWE.fnGetWe("//div[@class='alert alert-danger']"), "Alert Mesage", false, false);
                            Thread.Sleep(2000);
                        }

                        if (clsWE.fnElementExist("Account Disabled Message", "//div[contains(text(), 'Your account has been disabled after too many failed reset attempts.')]", false))
                        {
                            clsReportResult.fnLog("Restriction for Expired User", "The Account Disabled message was displayed as expected.", "Info", true);
                            clsWebBrowser.objDriver.Navigate().GoToUrl(clsMG.fnGetURLEnv((objData.fnGetValue("Environment", ""))));
                            //Enable and unlock user

                        }
                        else 
                        {
                            clsReportResult.fnLog("Restriction for Expired User", "The Account Disabled message was not displayed after provide invalid security questions.", "Fail", true);
                            blResult = false;
                        }
                    }
                    else 
                    {
                        clsReportResult.fnLog("Restriction for Expired User", "The Account Unlock Page was not displayed after enter valid credetials.", "Fail", false);
                        blResult = false;
                    }
                }
            }
            if (blResult) 
            { clsReportResult.fnLog("Restriction for Expired User", "The Restriction for Expired User Function was executed successfully.", "Pass", false); }
            else
            { clsReportResult.fnLog("Restriction for Expired User", "The Restriction for Expired User Function was not executed successfully.", "Fail", false); }
            return blResult;
        }

        public bool fnTimeoutSessionVerification(string pstrSetNo)
        {
            bool blResult = true;
            bool bFound = false;
            int intCount = 0;
            clsData objData = new clsData();
            clsReportResult.fnLog("Timeout session", "<<<<<<<<<< The Login Timeout session function Starts >>>>>>>>>>", "Info", false);
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    fnLogOffSession();
                    clsLogin login = new clsLogin();
                    if (login.fnLoginData(pstrSetNo))
                    {
                        //Do Until Timeout is displayed
                        do
                        {
                            intCount++;
                            clsReportResult.fnLog("Timeout session", "Waiting timeout label (" + intCount + ") minute(s).", "Info", false, false);
                            Thread.Sleep(TimeSpan.FromMinutes(1));
                            if (clsMG.IsElementPresent("//div[@id='modalSessionNotification' and contains(@style, 'display: block;')]//p[contains(text(), 'Your Session is about to expire ')]")) { bFound = true; }
                        }
                        while (!bFound && intCount <= 20); ;
                        //Report Log
                        if (bFound)
                        {
                            clsReportResult.fnLog("Timeout session", "The timeout session label was displayed successfully for user role: " + objData.fnGetValue("Role", "") + " after " + intCount + " minutes.", "Pass", true, false);
                        }
                        else
                        {
                            clsReportResult.fnLog("Timeout session", "The timeout session label was not displayed for user role: " + objData.fnGetValue("Role", "") + " after " + intCount + " minutes.", "Fail", true, false);
                            blResult = false;
                        }
                    }
                }
            }
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

        private bool fnReadTextEmail(string pstrUser, string pstrPassword, string pstrVal)
        {
            clsEmail email = new clsEmail();
            email.strFromEmail = pstrUser;
            email.strPassword = pstrPassword;
            email.strServer = "popgmail";
            return email.fnReadSimpleEmail(pstrUser, pstrPassword, pstrVal);
        }

        private void fnEnterCredentails(string pstrUser, string pstrPassword)
        {
            //Wait to load Page and verify if Cookies button appears
            clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
            if (clsMG.IsElementPresent("//button[@id='cookie-accept']")) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
            //Enter Credentials
            clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", false);
            clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", pstrUser, false);
            clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", false);
            clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", pstrPassword, true);
            //Click on Begin button
            clsWE.fnClick(clsWE.fnGetWe("//button[text()='BEGIN']"), "Begin", false);
        }

        private void fnEnterCredentails_New(string pstrUser, string pstrPassword)
        {
            //Enter Credentials and click on Begin
            clsMG.fnCleanAndEnterText("Username", "//input[@id='orangeForm-name']", pstrUser, false, false, "", false);
            clsMG.fnCleanAndEnterText("Password", "//input[@id='orangeForm-pass']", pstrPassword, false, false, "", false);
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
            if (clsMG.IsElementPresent("//span[text()='You are currently logged into ']")) 
            {
                clsWE.fnClick(clsWE.fnGetWe("//li[@class='nav-item m-menuitem-show']/a[@id='topmenu-logout']"), "Logout Link", false, false);
                clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
                clsReportResult.fnLog("Logout Session", "An active session was terminated", "Info", false, false);
            }
            return blResult;
        }

        public bool fnEnableAndUnlockUser(string pstrSetNo) 
        {
            bool blResult = true;
            clsData objData = new clsData();
            clsReportResult.fnLog("Enable / Unclock User", "Enable / Unclock User Function Starts.", "Info", false);
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Verify if an active account exist and log off
                    fnLogOffSession();
                    if (fnLoginData(pstrSetNo)) 
                    { 

                    }
                    else 
                    {
                        blResult = false;
                    }

                }
            }
            if (blResult)
            { clsReportResult.fnLog("Enable / Unclock User", "Enable / Unclock User Function was executed successfully.", "Pass", false); }
            else
            { clsReportResult.fnLog("Enable / Unclock User", "Enable / Unclock User Function was not executed successfully.", "false", false); }

            return blResult;
        }


        private void fnSecurityQuestion() 
        {
            Dictionary<string, string> dicQuestion = new Dictionary<string, string>();
            dicQuestion.Add("In what city did you meet your spouse/significant other?", "test1");
            dicQuestion.Add("In what city does your nearest sibling live?", "test2");
            dicQuestion.Add("In what city or town did your mother and father meet?", "test3");
            dicQuestion.Add("In what city or town was your first job?", "test4");
            dicQuestion.Add("In what city were you born?", "test5");

            IList<IWebElement> lsQuestions = clsWebBrowser.objDriver.FindElements(By.XPath("//p/strong"));
            int intIdx = 0;
            foreach (IWebElement element in lsQuestions)
            {
                intIdx++;
                string strValue;
                strValue = element.GetAttribute("innerText");
                if (dicQuestion.ContainsKey(strValue)) 
                {
                    clsMG.fnCleanAndEnterText("Security Question #" + intIdx, "(//input[@type='password'])["+ intIdx + "]", dicQuestion[strValue], false, false, "", false);
                    clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Begin", false);
                    if (clsMG.IsElementPresent("//div[@class='invalid-feedback']")) 
                    {
                    }
                
                }
            }
            
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
