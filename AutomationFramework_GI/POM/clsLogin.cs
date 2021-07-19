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
            clsReportResult.fnLog("Two Factor Authentication", "Two Factor Authentication Function Starts.", "Info", false);
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Enter Credentails
                    fnEnterCredentails(objData);
                    if (mega.IsElementPresent("//h3[text()='Multifactor Authentication']"))
                    {
                        //Select method & send code
                        mega.fnSelectDropDownWElm("Authentication Method", "//input[@class='select-dropdown form-control']", "Email", false, false);
                        clsWE.fnClick(clsWE.fnGetWe("//button[text()='Send Code']"), "Send Code", false);
                        //Get Email Token
                        Thread.Sleep(TimeSpan.FromSeconds(30));
                        string strToken = fnGet2FactorToken();
                        //Verify if token was received
                        if (strToken != "")
                        {
                            clsReportResult.fnLog("Two Factor Value", "The 2Factor email was received with value: " + strToken, "Pass", false, false);
                            mega.fnEnterTextWElm("Enter Code", "//input[@id='code']", strToken, true, false);
                            clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit", false);
                        }
                        else 
                        {
                            clsReportResult.fnLog("Two Factor Value", "The 2Factor email was not received after 2 minutes.", "Fail", false, false);
                            blResult = false;
                        }
                        //Verify Login Page
                        blResult = clsWE.fnElementExist("Login Label", "//span[contains(text(), 'You are currently logged into')]", false, false);
                        if (blResult)
                        { clsReportResult.fnLog("Two Factor Authentication", "The Login with 2FA was done successfully.", "Pass", true); }
                        else
                        { clsReportResult.fnLog("Two Factor Authentication", "The Login with 2FA was not completed successfully.", "Fail", true, true); }
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
            return blResult;
        }

        private void fnEnterCredentails(clsData pobjData) 
        {
            //Wait to load Page and verify if Cookies button appears
            clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
            if (clsWE.fnElementExist("Accept Cookies Message", "//button[@id='cookie-accept']", true)) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
            //Enter Credentials
            clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", false);
            clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-name']"), "Username", pobjData.fnGetValue("User", ""), false);
            clsWE.fnClick(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", false);
            clsWE.fnSendKeys(clsWE.fnGetWe("//input[@id='orangeForm-pass']"), "Password", pobjData.fnGetValue("Password", ""), true);
            //Click on Begin button
            clsWE.fnClick(clsWE.fnGetWe("//button[text()='BEGIN']"), "Begin", false);
        }

        private string fnGet2FactorToken() 
        {
            clsEmail email = new clsEmail();
            email.strFromEmail = "sedgautoemail@gmail.com";
            email.strPassword = "P4ssw0rd!123";
            email.strServer = "popgmail";
            string strValue = email.fnReadTwoFactorToken();
            return strValue;
        }

        private string fnForgotPasswordURL()
        {
            clsEmail email = new clsEmail();
            email.strFromEmail = "intakeautoemail@gmail.com";
            email.strPassword = "P4ssw0rd!123";
            email.strServer = "popgmail";
            string strValue = email.fnReadForgotPassword();
            return strValue;
        }


        public bool fnForgotPasswordVerification(string pstrSetNo)
        {
            bool blResult = true;
            clsData objData = new clsData();
            clsReportResult.fnLog("Forgot Password", "Forfot Password Function Starts.", "Info", false);
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInData");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Verify if "Forgot Password" link exist
                    clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='BEGIN']"), "Login", false, false);
                    if (clsWE.fnElementExist("Accept Cookies Message", "//button[@id='cookie-accept']", true)) { clsWE.fnClick(clsWE.fnGetWe("//button[@id='cookie-accept']"), "Accept Cookies Button", false); }
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
                                do { Thread.Sleep(TimeSpan.FromSeconds(30)); }
                                while (clsWE.fnGetAttribute(clsWE.fnGetWe("//input[@id='captcha-input']"), "Captcha", "value", false, false) == "");
                                clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit", false);
                                //Verify that email is received to change the password
                                string strURLReset = fnForgotPasswordURL();
                                if (strURLReset != "")
                                {
                                    clsWebBrowser browser = new clsWebBrowser();
                                    browser.fnNavigateToUrl(strURLReset);
                                    //Enter New Password and save it
                                    string strNewPass = RandomString(8);
                                    mega.fnEnterTextWElm("New Password", "//input[@id='new-pwd']", strNewPass, false, false);
                                    mega.fnEnterTextWElm("Confirm New Password", "//input[@id='new-pwd-v']", strNewPass, false, false);
                                    clsWE.fnClick(clsWE.fnGetWe("//input[@id='show-password-text']"), "Show Password", false);
                                    clsWE.fnClick(clsWE.fnGetWe("//button[text()='Submit']"), "Submit", true);
                                    //Verify that password was changes successfully
                                    if (clsWE.fnElementExist("Password Message", "//div[contains(text(), 'Your password has been successfully')]", false))
                                    {
                                        //Save the claim
                                        clsData objSaveData = new clsData();
                                        objSaveData.fnSaveValue(ConfigurationManager.AppSettings["FilePath"], "LogInData", "Password", intRow, strNewPass);

                                        //Verify email confirmation for password reset
                                        if (fnReadTextEmail(objData.fnGetValue("EmailAcc", ""), objData.fnGetValue("PassAcC", ""), "password for your account was "))
                                        {
                                            clsReportResult.fnLog("Forgot Password", "The  Password Change Confirmation email was received successfully.", "Pass", true, true);
                                        }
                                        else 
                                        {
                                            clsReportResult.fnLog("Forgot Password", "The  Password Change Confirmation email was not received.", "Fail", false, true);
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
            return blResult;
        }

        public string RandomString(int plength)
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
            return email.fnReadSimpleEmail(pstrVal);
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
