using AutomationFramework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Tesseract;

namespace AutomationFramework_GI.POM
{
    class clsMegaIntake
    {
        //Object Declaration
        private clsWebElements clsWE = new clsWebElements();
        private string[] garrStates = { "Alaska", "Alabama", "Arkansas", "Arizona", "California", "TestState" };
        private string[] garrAbbre = { "AK", "AL", "AR", "AZ", "CA", "TS" };

        //string[] arrStates = { "Alaska", "Alabama", "Arkansas", "Arizona", "California", "Colorado", "Connecticut", "District of Columbia", "Delaware", "Florida", "Georgia", "Hawaii", "Iowa", "Idaho", "Illinois", "Indiana", "Kansas", "Kentucky", "Louisiana", "Massachusetts", "Maryland", "Maine", "Michigan", "Minnesota", "Missouri", "Mississippi", "Montana", "North Carolina", "North Dakota", "Nebraska", "New Hampshire", "New Jersey", "New Mexico", "Nevada", "New York", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Puerto Rico", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Virginia", "Vermont", "Washington", "Wisconsin", "West Virginia", "Wyoming" };
        //string[] arrAbbre = { "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DC", "DE", "FL", "GA", "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA", "MD", "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE", "NH", "NJ", "NM", "NV", "NY", "OH", "OK", "OR", "PA", "PR", "RI", "SC", "SD", "TN", "TX", "UT", "VA", "VT", "WA", "WI", "WV", "WY" };



        //Common Functions

        public bool IsElementPresent(string pstrWebElement)
        {
            try
            {
                IWebElement objWebEdit = clsWebBrowser.objDriver.FindElement(By.XPath(pstrWebElement));
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }

        public bool fnEnterTextWE_(string pstrScreen, string pstrElement, string pstrValue)
        {
            bool blResult = true;
            if (pstrValue != "" && pstrElement != "")
            {
                try
                {
                    IWebElement objWebEdit = clsWebBrowser.objDriver.FindElement(By.XPath(pstrElement));
                    objWebEdit.Click();
                    Thread.Sleep(1000);
                    objWebEdit.SendKeys(Keys.Home);
                    objWebEdit.SendKeys(pstrValue);
                    Thread.Sleep(1000);
                    objWebEdit = clsWebBrowser.objDriver.FindElement(By.XPath("//span[@data-bind='text: headerClientName']"));
                    Thread.Sleep(1500);
                    objWebEdit.Click();
                    Thread.Sleep(1500);
                    blResult = true;
                }
                catch (Exception objException)
                {
                    Console.WriteLine("SendKeys is not working for: " + pstrElement + " an exception was found: " + objException.Message);
                    clsReportResult.fnLog("SendKeysFail", "SendKeys is not working for: " + pstrElement + " with value: " + pstrValue, "Fail", true, true);
                }

                if (blResult)
                { clsReportResult.fnLog("SendKeysPass", "SendKeys for: " + pstrElement + " with value: " + pstrValue, "Pass", false, true); }
                else
                { clsReportResult.fnLog("SendKeysFail", "SendKeys is not working for: " + pstrElement + " with value: " + pstrValue, "Fail", true, true); }

            }
            return blResult;
        }
        
        public bool fnEnterTextWElm(string pstrElement, string pstrWebElement, string pstrValue, bool pblScreenShot = false, bool pblHardStop = false, string pstrHardStopMsg = "")
        {
            bool blResult = true;
            if (pstrValue != "" && pstrWebElement != "")
            {
                try
                {
                    //clsReportResult.fnLog("SendKeys", "Step - Sendkeys: " + pstrValue + " to field: " + pstrWebElement, "Info", false);
                    IWebElement objWebEdit = clsWebBrowser.objDriver.FindElement(By.XPath(pstrWebElement));
                    objWebEdit.Click();
                    //Thread.Sleep(1000);
                    objWebEdit.SendKeys(Keys.Home);
                    objWebEdit.SendKeys(pstrValue);
                    Thread.Sleep(1000);

                    if (IsElementPresent("//span[@data-bind='text: headerClientName']"))
                    {
                        objWebEdit = clsWebBrowser.objDriver.FindElement(By.XPath("//span[@data-bind='text: headerClientName']"));
                        Thread.Sleep(1000);
                        objWebEdit.Click();
                        Thread.Sleep(1000);
                    }
                    blResult = true;
                }
                catch (Exception objException)
                {
                    blResult = false;
                    clsWebElements.fnExceptionHandling(objException);
                    //Console.WriteLine("SendKeys is not working for: " + pstrWebElement + " an exception was found: " + objException.Message);
                    //clsReportResult.fnLog("SendKeysFail", "SendKeys is not working for: " + pstrElement + " with value: " + pstrValue + " and locator: " + pstrWebElement, "Fail", pblScreenShot, pblHardStop); 
                }
                if (blResult)
                { clsReportResult.fnLog("SendKeysPass", "SendKeys for: " + pstrElement + " with value: " + pstrValue, "Pass", pblScreenShot, pblHardStop); }
                else
                {
                    blResult = false;
                    clsReportResult.fnLog("SendKeysFail", "SendKeys is not working for: " + pstrElement + " with value: " + pstrValue + " and locator: " + pstrWebElement, "Fail", pblScreenShot, pblHardStop);
                }
            }
            return blResult;
        }

        public bool fnSelectDropDownWE_(string pstrScreen, string pstrElement, string pstrDropdownValue)
        {
            clsWebElements clsWE = new clsWebElements();
            bool blResult = false;

            try
            {
                if (pstrElement != "" && pstrDropdownValue != "")
                {
                    IWebElement objDropDown = clsWebBrowser.objDriver.FindElement(By.XPath(pstrElement));
                    objDropDown.Click();
                    Thread.Sleep(1500);

                    //IWebElement objDropDownContent = clsWebBrowser.objDriver.FindElement(By.XPath("//ul[@class='dropdown-content select-dropdown w-100 active']"));
                    //IList<IWebElement> objOptions = objDropDownContent.FindElements(By.XPath("//span[@class='filtrable']"));

                    IWebElement objDropDownContent = clsWebBrowser.objDriver.FindElement(By.XPath("//span[@class='select2-results']"));
                    IList<IWebElement> objOptions = objDropDownContent.FindElements(By.XPath("//li[@role='treeitem']"));

                    foreach (IWebElement objOption in objOptions)
                    {
                        string pstrDropdownText = (objOption.GetAttribute("innerText"));
                        if (pstrDropdownText == pstrDropdownValue)
                        {
                            objOption.Click();
                            blResult = true;
                            break;
                        }
                    }
                    objDropDown = clsWebBrowser.objDriver.FindElement(By.XPath("//span[@data-bind='text: headerClientName']"));
                    Thread.Sleep(1500);
                    objDropDown.Click();
                    Thread.Sleep(1500);
                }
            }
            catch (Exception objException)
            {
                clsReportResult.fnLog("DropdownSelectFail", "Dropdown exception" + objException.Message + " with value: " + pstrDropdownValue, "Fail", true);
            }

            if (blResult)
                clsReportResult.fnLog("DropdownSelectPass", "Dropdown selected for: " + pstrElement + " with value: " + pstrDropdownValue, "Pass", false);
            else
                clsReportResult.fnLog("DropdownSelectFail", "Dropdown is not selected for: " + pstrElement + " with value: " + pstrDropdownValue, "Fail", true);

            return blResult;
        }

        public bool fnSelectDropDownWElm(string pstrElement, string pstrWebElement, string pstrValue, bool pblScreenShot = false, bool pblHardStop = false, string pstrHardStopMsg = "")
        {
            clsWebElements clsWE = new clsWebElements();
            bool blResult = false;
            IWebElement objDropDownContent;
            IList<IWebElement> objOptions;

            try
            {
                if (pstrElement != "" && pstrValue != "")
                {
                    clsReportResult.fnLog("SelectDropdown", "Step - Select Dropdown: " + pstrElement + " With Value: " + pstrValue, "Info", false);
                    IWebElement objDropDown = clsWebBrowser.objDriver.FindElement(By.XPath(pstrWebElement));
                    objDropDown.Click();
                    Thread.Sleep(1500);

                    if (IsElementPresent("//span[@class='select2-results']"))
                    {
                        //Common Dropdown
                        objDropDownContent = clsWebBrowser.objDriver.FindElement(By.XPath("//span[@class='select2-results']"));
                        objOptions = objDropDownContent.FindElements(By.XPath("//li[@role='treeitem']"));
                    }
                    else 
                    {
                        //two Factor Dropdown
                        objDropDownContent = clsWebBrowser.objDriver.FindElement(By.XPath("//ul[@class='dropdown-content select-dropdown w-100 active']"));
                        objOptions = objDropDownContent.FindElements(By.XPath("//span[@class='filtrable']"));
                    }

                    foreach (IWebElement objOption in objOptions)
                    {
                        string pstrDropdownText = (objOption.GetAttribute("innerText"));
                        if (pstrDropdownText == pstrValue)
                        {
                            objOption.Click();
                            blResult = true;
                            break;
                        }
                    }

                    if (IsElementPresent("//span[@data-bind='text: headerClientName']"))
                    {
                        objDropDown = clsWebBrowser.objDriver.FindElement(By.XPath("//span[@data-bind='text: headerClientName']"));
                        Thread.Sleep(1500);
                        objDropDown.Click();
                        Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception objException)
            {
                blResult = false;
                clsWebElements.fnExceptionHandling(objException);
                //Console.WriteLine("DropDown is not working for: " + pstrWebElement + " an exception was found: " + objException.Message);
                //clsReportResult.fnLog("DropdownSelectFail", "DropDown is not working for: " + pstrWebElement + " with value: " + pstrValue, "Fail", pblScreenShot, pblHardStop);
            }
            if (blResult)
            {
                clsReportResult.fnLog("SelectListPass", "Select Dropdown for element: " + pstrElement + " was done successfully.", "Pass", pblScreenShot);
            }
            else
            {
                blResult = false;
                clsReportResult.fnLog("SelectListFail", "Select Dropdown for element: " + pstrElement + " has failed with value: " + pstrValue + " and locator " + pstrWebElement, "Fail", pblScreenShot);
                //clsReportResult.fnLog("DropdownSelectFail", "DropDown is not working for: " + pstrElement + " with value: " + pstrValue + " and locator: " + pstrWebElement, "Fail", pblScreenShot, pblHardStop); 
            }

            return blResult;
        }
        public void fnWaitWEUntilAppears(string pstrStepName, string pstrLocator, int pintTime)
        {
            do { Thread.Sleep(pintTime); }
            while (!clsWE.fnElementExistNoReport(pstrStepName, pstrLocator, false));
        }

        public void fnReadCaptcha() 
        {
            var objCaptcha = clsWebBrowser.objDriver.FindElement(By.XPath("//img[@id='ForgotUserNameCaptcha_CaptchaImage']"));
            Point location = objCaptcha.Location;

            var screenshot = (clsWebBrowser.objDriver as ChromeDriver).GetScreenshot();
            using (MemoryStream stream = new MemoryStream(screenshot.AsByteArray)) 
            {
                using (Bitmap bitmap = new Bitmap(stream)) 
                {
                    RectangleF part = new RectangleF(location.X, location.Y, objCaptcha.Size.Width, objCaptcha.Size.Height);
                    using (Bitmap bn = bitmap.Clone(part, bitmap.PixelFormat)) 
                    {
                        bn.Save(@"\\memfp02\share\Any\4th_Automation\VisualStudio\Captcha\" + "CaptchaImg.Png");
                    }
                }
            }

            //Read Text from Image
            string strLangLoc = @"\\memfp02\share\Any\4th_Automation\VisualStudio\Tesseract\Lang";
            using( var engine = new TesseractEngine(strLangLoc, "eng", EngineMode.Default)) 
            {
                Page ocrPage = engine.Process(Pix.LoadFromFile(@"\\memfp02\share\Any\4th_Automation\VisualStudio\Captcha\" + "CaptchaImg.Png"), PageSegMode.AutoOnly);
                var captchaText = ocrPage.GetText();
            }



        }


        //GI Functions
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

        public string fnGetURLEnv(string pstrEnv)
        {
            string URL = "";
            switch (pstrEnv.ToUpper())
            {
                case "QA":
                    URL = ConfigurationManager.AppSettings["UrlQA"];
                    break;
                case "UAT":
                    URL = ConfigurationManager.AppSettings["UrlUAT"];
                    break;
                case "E2E":
                    URL = ConfigurationManager.AppSettings["UrlE2E"];
                    break;
                default:
                    URL = "";
                    break;
            }
            return URL;
        }

        public bool fnCreateSubmitIntake(string pstrSetNo)
        {
            bool blResult = true;
            clsData objData = new clsData();
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "EventInfo");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Start Intake Page
                    clsWE.fnPageLoad(clsWE.fnGetWe("//a[contains(text(), 'START INTAKE')]"), "Start Intake", true, true);
                    clsWE.fnClick(clsWE.fnGetWe("//a[contains(text(), 'START INTAKE')]"), "Start Intake", false);
                    clsWE.fnPageLoad(clsWE.fnGetWe("//span[contains(text(), 'Select Client')]"), "Select Intake", true, true);
                    clsWE.fnClick(clsWE.fnGetWe("//button[@id='selectClient_']"), "Select Client", false);
                    clsWE.fnPageLoad(clsWE.fnGetWe("//div[contains(@style, 'display: block;')]//table[@id='clientSelectorTable_']"), "Select Intake", false, true);
                    clsWE.fnPageLoad(clsWE.fnGetWe("//input[@placeholder='Client Number or Name']"), "Client Number", false, true);
                    clsWE.fnClick(clsWE.fnGetWe("//input[@placeholder='Client Number or Name']"), "Client Number", false);
                    clsWE.fnSendKeys(clsWE.fnGetWe("//input[@placeholder='Client Number or Name']"), "Client Number", objData.fnGetValue("Contract", ""));
                    //Verify if client exist
                    if (!clsWE.fnElementExistNoReport("Search Client", "//td[text()='No matching records found']", false, false))
                    {
                        //Filter Script Name
                        clsWE.fnClick(clsWE.fnGetWe("(//a[text()='Select'])[1]"), "Select", false);
                        clsWE.fnClick(clsWE.fnGetWe("//input[@placeholder='Filter Results']"), "Filter Results", false);
                        clsWE.fnSendKeys(clsWE.fnGetWe("//input[@placeholder='Filter Results']"), "Filter Results", objData.fnGetValue("IntakeType", ""));
                        if (!clsWE.fnElementExist("Start New Intake", "//td[text()='No matching records found']", false, false))
                        {
                            clsWE.fnClick(clsWE.fnGetWe("//tr[@role='row' and td[button[contains(text(), 'Start Intake')]]]//button"), "Start Intake", false);
                            if (fnDuplicatedClaims(objData.fnGetValue("DuplicatedSet", "")))
                            {
                                //Intake Script Caller Information
                                fnWaitWEUntilAppears("Wait to load Panel", "//div[@id='list-example']", 1000);
                                fnWaitWEUntilAppears("Wait to next", "//button[text()='Next']", 1000);
                                //fnEnterTextWE("Caller FN", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='First Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("FirstName"));
                                //fnEnterTextWE("Caller LN", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='Last Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("LastName"));
                                //fnEnterTextWE("Caller Phone", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='Work Phone Number']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("WorkPhoneNumber"));

                                fnEnterTextWElm("Caller FN", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='First Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("FirstName", ""), false, false);
                                fnEnterTextWElm("Caller LN", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='Last Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("LastName", ""), false, false);
                                fnEnterTextWElm("Caller Phoner", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='Work Phone Number']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("WorkPhoneNumber", ""), false, false);



                                //Intake Script Client Location
                                clsWE.fnClick(clsWE.fnGetWe("//span[contains(@data-bind, 'headerClientName')]"), "Click", false);
                                //fnSelectDropDownWE("Is the same location?", "//div[contains(@question-key, 'EMPLOYER_INFORMATION')]//div[@class='row' and div[span[text()='Is This The Loss Location?']]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("IsThisTheLossLocation"));
                                fnSelectDropDownWElm("Is the same location?", "//div[contains(@question-key, 'EMPLOYER_INFORMATION')]//div[@class='row' and div[span[text()='Is This The Loss Location?']]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("IsThisTheLossLocation"), false, false);
                                //Intake Script Employee Information
                                //fnEnterTextWE("Emp First Name", "//div[contains(@question-key, 'EMPLOYEE_INFORMATION')]//div[@class='row' and div[span[text()='First Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("EmpFirstName"));
                                //fnEnterTextWE("Emp Last Name", "//div[contains(@question-key, 'EMPLOYEE_INFORMATION')]//div[@class='row' and div[span[text()='Last Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("EmpLastName"));

                                fnEnterTextWElm("Emp First Name", "//div[contains(@question-key, 'EMPLOYEE_INFORMATION')]//div[@class='row' and div[span[text()='First Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("EmpFirstName", ""), false, false);
                                fnEnterTextWElm("Emp Last Name", "//div[contains(@question-key, 'EMPLOYEE_INFORMATION')]//div[@class='row' and div[span[text()='Last Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("EmpLastName", ""), false, false);


                                //Intake Script Employment Information
                                //fnSelectDropDownWE("Did Employee Miss Work?", "//div[contains(@question-key, 'EMPLOYMENT_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Did Employee Miss Work Beyond Their Normal Shift?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("DidEmployeeMissWork"));
                                fnSelectDropDownWElm("Did Employee Miss Work?", "//div[contains(@question-key, 'EMPLOYMENT_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Did Employee Miss Work Beyond Their Normal Shift?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("DidEmployeeMissWork"), false, false);


                                //Intake Script Incident Information
                                //fnEnterTextWE("Loss Description", "//div[contains(@question-key, 'INCIDENT_INFORMATION')]//div[@class='row' and div[span[text()='Loss Description']]]//textarea", objData.fnGetValue("LossDesc"));
                                fnEnterTextWElm("Loss Description", "//div[contains(@question-key, 'INCIDENT_INFORMATION')]//div[@class='row' and div[span[text()='Loss Description']]]//textarea", objData.fnGetValue("LossDesc", ""), false, false);


                                //Intake Script Contact Information
                                clsWE.fnClick(clsWE.fnGetWe("//span[contains(@data-bind, 'headerClientName')]"), "Click", false);
                                //fnSelectDropDownWE("Is Contact Same As Caller?", "//div[contains(@question-key, 'CONTACT_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Is Contact Same As Caller?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("IsContactSameAsCaller"));
                                fnSelectDropDownWElm("Is Contact Same As Caller?", "//div[contains(@question-key, 'CONTACT_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Is Contact Same As Caller?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("IsContactSameAsCaller"), false, false);

                                Thread.Sleep(5000);

                                //CommentsRemarks
                                //fnEnterTextWE("Comments Remarks", "//div[contains(@question-key, 'COMMENTS_REMARKS')]//div[@class='row' and div[span[text()='Comments / Remarks']]]//following-sibling::textarea", "Claim created by automation script");
                                //fnEnterTextWE("Internal Comments", "//div[contains(@question-key, 'COMMENTS_REMARKS')]//div[@class='row' and div[span[text()='Internal Comments']]]//following-sibling::textarea", "Internal comment test");

                                fnEnterTextWElm("Comments Remarks", "//div[contains(@question-key, 'COMMENTS_REMARKS')]//div[@class='row' and div[span[text()='Comments / Remarks']]]//following-sibling::textarea", "Claim created by automation script", false, false);
                                fnEnterTextWElm("Internal Comments", "//div[contains(@question-key, 'COMMENTS_REMARKS')]//div[@class='row' and div[span[text()='Internal Comments']]]//following-sibling::textarea", "Internal comment test", false, false);


                                if (clsWE.fnElementExist("Loss Time", "//div[@id='list-example']//span[text()='Lost Time Information']", true, false))
                                {
                                    //Click Menu
                                    clsWE.fnClick(clsWE.fnGetWe("//div[@id='list-example']//span[text()='Lost Time Information']"), "Click Menu");
                                    Thread.Sleep(2000);
                                    //Lost Time Information
                                    //fnSelectDropDownWE("Employee Returned To Work", "//div[contains(@question-key, 'LOST_TIME_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Employee Returned To Work?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("EmployeeReturnedToWork"));
                                    fnSelectDropDownWElm("Employee Returned To Work", "//div[contains(@question-key, 'LOST_TIME_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Employee Returned To Work?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("EmployeeReturnedToWork"), false, false);

                                }

                                //Click on NEXT
                                clsWE.fnClick(clsWE.fnGetWe("//button[@type='submit' and text()='Next']"), "Click Next");
                                Thread.Sleep(5000);
                                if (!clsWE.fnElementExist("Error", "//*[@data-bind='text:ValidationMessage']", true, false))
                                {
                                    if (!clsWE.fnElementExist("Closing Script", "//span[text()='Closing Script']", true, false))
                                    {
                                        clsWE.fnClick(clsWE.fnGetWe("//button[@type='submit' and text()='Next']"), "Click Next");
                                        clsWE.fnPageLoad(clsWE.fnGetWe("//span[text()='Closing Script']"), "Closing Script", true, true);

                                    }
                                }

                                //Click on Submit
                                clsWE.fnPageLoad(clsWE.fnGetWe("//button[@id='top-submit']"), "Submit", true, true);
                                clsWE.fnClick(clsWE.fnGetWe("//button[@id='top-submit']"), "Submit");
                                Thread.Sleep(10000);
                                clsWE.fnPageLoad(clsWE.fnGetWe("//span[text()='Thank you']"), "Thank You", true, true);

                                string strClaimNo = "";
                                strClaimNo = clsWE.fnGetAttribute(clsWE.fnGetWe("//span[contains(@data-bind, 'VendorIncidentNumber')]"), "Confirmation Number", "innerText");

                                //Save the claim
                                clsData objSaveData = new clsData();
                                objSaveData.fnSaveValue(ConfigurationManager.AppSettings["FilePath"], "EventInfo", "ClaimNumber", intRow, strClaimNo);

                                //Verify if the claim is created
                                //if (fnGetQuery("3", strClaimNo))
                                if (fnGetQuery("1", "402103288B80001"))
                                { clsReportResult.fnLog("DB Execution", "The case/claim was created successfully and found in the vioone.claim table.", "Info", false); }
                                else
                                { clsReportResult.fnLog("DB Execution", "The case/claim was created successfully but was notfound in the vioone.claim table.", "Info", false); }

                            }
                        }
                        else
                        {
                            clsReportResult.fnLog("Start New Intake", "No active intakes were found in Select Intake screen.", "Fail", false, true);
                            blResult = false;
                        }
                    }
                    else
                    {
                        clsReportResult.fnLog("Select Client", "No clients were found in Search Client screen.", "Fail", false, true);
                        blResult = false;
                    }
                }
            }
            return blResult;
        }

        private bool fnDuplicatedClaims(string pstrSetNo)
        {
            bool blResult = true;
            clsData objData = new clsData();
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "DuplicatedClaim");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Duplicated Claim Check Page
                    clsWE.fnPageLoad(clsWE.fnGetWe("//span[text()='Duplicate Claim Check']"), "Intake", false, true);
                    //fnEnterTextWE("Loss Date", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Loss Date']]]//input[@class='form-control']", objData.fnGetValue("LossDate", ""));
                    //fnEnterTextWE("Loss Time", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Loss Time']]]//input[@class='form-control']", objData.fnGetValue("LossTime", ""));

                    fnEnterTextWElm("Loss Date", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Loss Date']]]//input[@class='form-control']", objData.fnGetValue("LossDate", ""), false, false);
                    fnEnterTextWElm("Loss Time", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Loss Time']]]//input[@class='form-control']", objData.fnGetValue("LossTime", ""), false, false);


                    //Thread.Sleep(1000);
                    //fnSelectDropDownWE("Reporter Type", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Reporter Type']]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("ReporterType", ""));
                    //Thread.Sleep(1000);
                    //fnSelectDropDownWE("Reported By", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Reported By']]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("ReportedBy", ""));

                    Thread.Sleep(1000);
                    fnSelectDropDownWElm("Reporter Type", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Reporter Type']]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("ReporterType"), false, false);
                    Thread.Sleep(1000);
                    fnSelectDropDownWElm("Reported By", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Reported By']]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("ReportedBy"), false, false);




                    //fnSelectDropDownWE("Reported By", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Reported By']]]//following-sibling::input[starts-with(@class, 'select-dropdown')]", objData.fnGetValue("ReportedBy", ""));
                    //fnSelectDropDownWE("Reporter Type", "//div[contains(@question-key, 'DUPLICATE_CLAIM')]//div[@class='row' and div[span[text()='Reporter Type']]]//following-sibling::input[starts-with(@class, 'select-dropdown')]", objData.fnGetValue("ReporterType", ""));

                    //Employee-Location Lookup
                    //if (clsWE.fnElementExist("Employee Lookup", "//button[@id='btnJurisLocation_LOCATION_LOOKUP']")) { blResult = fnLocationLookup(objData.fnGetValue("LocLookupSet", "")); }
                    if (clsWE.fnElementExistNoReport("Location Lookup", "//button[@id='btnJurisLocation_LOCATION_LOOKUP']", false)) { blResult = fnLocationLookup(objData.fnGetValue("LocLookupSet", "")); }
                    //Go to next screen
                    clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='Next']"), "Next", false, false);
                    clsWE.fnClick(clsWE.fnGetWe("//button[text()='Next']"), "Next", false);
                    //Verify if error message exist
                    if (clsWE.fnElementNotExist("Next Verification", "//div[@class='col-md-8 secondary-red']", false))
                    {
                        //Previous Intakes Exists
                        if (clsWE.fnElementExistNoReport("Previously Intakes", "//span[@class='h4 font-weight-light' and text()='Previously Started Intakes']", false))
                        {
                            clsWE.fnPageLoad(clsWE.fnGetWe("//span[@class='h4 font-weight-light' and text()='Previously Started Intakes']"), "Previous Intakes", false, true);
                            fnWaitWEUntilAppears("Start New Intake", "(//button[contains(text(), 'Start New Intake')])[1]", 1000);
                            clsWE.fnClick(clsWE.fnGetWe("(//button[contains(text(), 'Start New Intake')])[1]"), "Start New Intake", false);
                        }
                    }
                    else
                    {
                        clsReportResult.fnLog("Duplicated Claims", "Some errors were found in Duplicated Claims Screen.", "Fail", false, true);
                        blResult = false;
                    }
                }
            }
            return blResult;
        }

        private bool fnLocationLookup(string pstrSetNo)
        {
            //Location Lookup
            bool blResult = true;
            clsWE.fnClick(clsWE.fnGetWe("//button[@id='btnJurisLocation_LOCATION_LOOKUP']"), "Location Lookup Button", false);
            Thread.Sleep(3000);
            if (!clsWE.fnElementExistNoReport("Location Lookup", "//*[@class='fixed-sn white-skin modal-open']", false))
            {
                clsWE.fnClick(clsWE.fnGetWe("//button[@id='btnJurisLocation_LOCATION_LOOKUP']"), "Location Lookup Button", false);
                clsWE.fnPageLoad(clsWE.fnGetWe("//*[@class='fixed-sn white-skin modal-open']"), "Wait Search", false, false);
            }

            //fnWaitWEUntilAppears("Location Lookup is opened", "//*[@class='fixed-sn white-skin modal-open']", 1000);
            clsWE.fnPageLoad(clsWE.fnGetWe("//button[text()='Search']"), "Wait Search", false, true);
            clsWE.fnPageLoad(clsWE.fnGetWe("//button[@id='btn_close_juris' and text()='Close']"), "Wait Close", false, true);
            clsWE.fnPageLoad(clsWE.fnGetWe("//div[@id='jurisLocationSearchModal_LOCATION_LOOKUP' and not(contains(@style, 'display: none'))]"), "Intake", false, true);

            //if Set == 0, select default location, else search specific location
            if (pstrSetNo == "0")
            {
                if (clsWE.fnElementExistNoReport("Location Record", "(//table[@id='jurisLocationResults_LOCATION_LOOKUP']//a[@class='btn-floating btn-sm select-button'])[1]", false))
                {
                    clsWE.fnClick(clsWE.fnGetWe("(//table[@id='jurisLocationResults_LOCATION_LOOKUP']//a[@class='btn-floating btn-sm select-button'])[1]"), "Select Location Row", false);
                    Thread.Sleep(1000);
                    fnWaitWEUntilAppears("Location Lookup is opened", "//*[@class='fixed-sn white-skin']", 1000);
                }
                else
                {
                    clsReportResult.fnLog("Location Lookup", "No record found in Location Lookup screen.", "Fail", true);
                    blResult = false;
                }
            }
            else if (pstrSetNo != "")
            {
                int intRow = 2;
                clsData objData = new clsData();
                objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LocationLookup");
                objData.CurrentRow = intRow;
                for (intRow = 2; intRow <= objData.RowCount; intRow++)
                {
                    if (objData.fnGetValue("Set", "") == pstrSetNo)
                    {
                        objData.CurrentRow = intRow;
                        //fnEnterTextWE("Account Number", "//input[@id='search-accountNumber']", objData.fnGetValue("AccountNumber", ""));
                        //fnEnterTextWE("Unit Number", "//input[@id='search-unitNumber']", objData.fnGetValue("UnitNumber", ""));

                        fnEnterTextWElm("Account Number", "//input[@id='search-accountNumber']", objData.fnGetValue("AccountNumber", ""), false, false);
                        fnEnterTextWElm("Unit Number", "//input[@id='search-unitNumber']", objData.fnGetValue("UnitNumber", ""), false, false);


                        clsWE.fnClick(clsWE.fnGetWe("//button[text()='Search']"), "Search");
                        Thread.Sleep(3000);
                        clsWE.fnPageLoad(clsWE.fnGetWe("//*[@id='jurisLocationResults_LOCATION_LOOKUP']"), "Location Lookup Filter", true, true);
                        if (clsWE.fnElementExist("Location Record", "(//table[@id='jurisLocationResults_LOCATION_LOOKUP']//td[button[text()='Select']]//button)[1]"))
                        {
                            clsWE.fnClick(clsWE.fnGetWe("(//table[@id='jurisLocationResults_LOCATION_LOOKUP']//td[button[text()='Select']]//button)[1]"), "Select Location Row");
                            Thread.Sleep(1000);
                            fnWaitWEUntilAppears("Location Lookup is opened", "//*[@class='fixed-sn white-skin']", 1000);
                        }
                        else
                        {
                            clsReportResult.fnLog("Location Lookup", "No record found in Location Lookup screen.", "Fail", true);
                            blResult = false;
                        }
                    }
                }
            }
            //Close Dialog in case not record was found
            if (clsWE.fnElementExistNoReport("Location Dialog", "//*[@class='fixed-sn white-skin modal-open']", false))
            {
                clsWE.fnClick(clsWE.fnGetWe("//div[@id='jurisLocationSearchModal_LOCATION_LOOKUP']//button[@id='btn_close_juris']"), "Close Location", false);
                fnWaitWEUntilAppears("Location Lookup is opened", "//*[@class='fixed-sn white-skin']", 1000);
            }
            return blResult;
        }

        public bool fnGetQuery(string pstrSetNo, string pstrClaimNo)
        {
            bool bStatus = false;
            clsData objData = new clsData();
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInDataBase");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    clsDB objDBOR = new clsDB();
                    objDBOR.fnOpenConnection(objDBOR.GetConnectionString(objData.fnGetValue("Host", ""), objData.fnGetValue("Port", ""), objData.fnGetValue("Service", ""), objData.fnGetValue("User", ""), objData.fnGetValue("Password", "")));
                    DataTable datatable = new DataTable();
                    string strQuery = "SELECT * FROM viaone.claim WHERE file_num = '" + pstrClaimNo + "'";
                    datatable = objDBOR.fnDataSet(strQuery);
                    if (datatable != null)
                    {
                        foreach (DataRow row in datatable.Rows)
                        {
                            Console.WriteLine("Event_Num: {0}: ", Convert.ToString(row["EVENT_NUM"]));
                            Console.WriteLine("File_Num: {0}: ", Convert.ToString(row["FILE_NUM"]));
                            Console.WriteLine("Output_File_Num: {0}: ", Convert.ToString(row["OUTPUT_FILE_NUM"]));
                        }
                        bStatus = true;
                    }
                    objDBOR.fnCloseConnection();
                }
            }
            return bStatus;
        }

        public bool fnBranchOfficeVerification(string pstrSetNo)
        {
            bool blResult = true;
            clsData objData = new clsData();
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "EventInfo");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Start Intake Page
                    clsWE.fnPageLoad(clsWE.fnGetWe("//a[contains(text(), 'START INTAKE')]"), "Start Intake", false, true);
                    clsWE.fnClick(clsWE.fnGetWe("//a[contains(text(), 'START INTAKE')]"), "Start Intake", false);
                    clsWE.fnPageLoad(clsWE.fnGetWe("//span[contains(text(), 'Select Client')]"), "Select Intake", false, true);
                    clsWE.fnClick(clsWE.fnGetWe("//button[@id='selectClient_']"), "Select Client", false);
                    clsWE.fnPageLoad(clsWE.fnGetWe("//div[contains(@style, 'display: block;')]//table[@id='clientSelectorTable_']"), "Select Intake", true, true);
                    clsWE.fnPageLoad(clsWE.fnGetWe("//input[@placeholder='Client Number or Name']"), "Client Number", false, true);
                    clsWE.fnClick(clsWE.fnGetWe("//input[@placeholder='Client Number or Name']"), "Client Number", false);
                    Thread.Sleep(1000);
                    clsWE.fnSendKeys(clsWE.fnGetWe("//input[@placeholder='Client Number or Name']"), "Client Number", objData.fnGetValue("Contract", ""));
                    Thread.Sleep(5000);
                    //Verify if client exist
                    if (clsWE.fnElementNotExist("Search Client", "//td[text()='No matching records found']", false, false))
                    {
                        //Filter Script Name
                        clsWE.fnClick(clsWE.fnGetWe("(//a[text()='Select'])[1]"), "Select", false);
                        clsWE.fnClick(clsWE.fnGetWe("//input[@placeholder='Filter Results']"), "Filter Results", false);
                        clsWE.fnSendKeys(clsWE.fnGetWe("//input[@placeholder='Filter Results']"), "Filter Results", objData.fnGetValue("IntakeType", ""));
                        Thread.Sleep(1000);
                        if (clsWE.fnElementNotExist("Start New Intake", "//td[text()='No matching records found']", false, false))
                        {
                            clsWE.fnClick(clsWE.fnGetWe("//tr[@role='row' and td[button[contains(text(), 'Start Intake')]]]//button"), "Start Intake", false);
                            if (fnDuplicatedClaims(objData.fnGetValue("DuplicatedSet", "")))
                            {
                                //Intake Script Caller Information
                                fnWaitWEUntilAppears("Wait to load Panel", "//div[@id='list-example']", 1000);
                                fnWaitWEUntilAppears("Wait to next", "//button[text()='Next']", 1000);
                                //fnEnterTextWE("Caller FN", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='First Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("FirstName"));
                                //fnEnterTextWE("Caller LN", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='Last Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("LastName"));
                                //fnEnterTextWE("Caller Phone", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='Work Phone Number']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("WorkPhoneNumber"));

                                fnEnterTextWElm("Caller FN", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='First Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("FirstName", ""), false, false);
                                fnEnterTextWElm("Caller LN", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='Last Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("LastName", ""), false, false);
                                fnEnterTextWElm("Caller Phoner", "//div[contains(@question-key, 'CALLER_INFORMATION')]//div[@class='row' and div[span[text()='Work Phone Number']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("WorkPhoneNumber", ""), false, false);



                                //Intake Script Client Location
                                clsWE.fnClick(clsWE.fnGetWe("//span[contains(@data-bind, 'headerClientName')]"), "Click", false);
                                //fnSelectDropDownWE("Is the same location?", "//div[contains(@question-key, 'EMPLOYER_INFORMATION')]//div[@class='row' and div[span[text()='Is This The Loss Location?']]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("IsThisTheLossLocation"));
                                fnSelectDropDownWElm("Is the same location?", "//div[contains(@question-key, 'EMPLOYER_INFORMATION')]//div[@class='row' and div[span[text()='Is This The Loss Location?']]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("IsThisTheLossLocation"), false, false);



                                //Intake Script Employee Information
                                //fnEnterTextWE("Emp First Name", "//div[contains(@question-key, 'EMPLOYEE_INFORMATION')]//div[@class='row' and div[span[text()='First Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("EmpFirstName"));
                                //fnEnterTextWE("Emp Last Name", "//div[contains(@question-key, 'EMPLOYEE_INFORMATION')]//div[@class='row' and div[span[text()='Last Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("EmpLastName"));

                                fnEnterTextWElm("Emp First Name", "//div[contains(@question-key, 'EMPLOYEE_INFORMATION')]//div[@class='row' and div[span[text()='First Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("EmpFirstName", ""), false, false);
                                fnEnterTextWElm("Emp Last Name", "//div[contains(@question-key, 'EMPLOYEE_INFORMATION')]//div[@class='row' and div[span[text()='Last Name']]]//following-sibling::input[starts-with(@class, 'form-control')]", objData.fnGetValue("EmpLastName", ""), false, false);


                                //Intake Script Employment Information
                                //fnSelectDropDownWE("Did Employee Miss Work?", "//div[contains(@question-key, 'EMPLOYMENT_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Did Employee Miss Work Beyond Their Normal Shift?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("DidEmployeeMissWork"));
                                fnSelectDropDownWElm("Did Employee Miss Work?", "//div[contains(@question-key, 'EMPLOYMENT_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Did Employee Miss Work Beyond Their Normal Shift?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("DidEmployeeMissWork"), false, false);


                                //Intake Script Incident Information
                                //fnEnterTextWE("Loss Description", "//div[contains(@question-key, 'INCIDENT_INFORMATION')]//div[@class='row' and div[span[text()='Loss Description']]]//textarea", objData.fnGetValue("LossDesc"));
                                fnEnterTextWElm("Loss Description", "//div[contains(@question-key, 'INCIDENT_INFORMATION')]//div[@class='row' and div[span[text()='Loss Description']]]//textarea", objData.fnGetValue("LossDesc", ""), false, false);


                                //Intake Script Contact Information
                                clsWE.fnClick(clsWE.fnGetWe("//span[contains(@data-bind, 'headerClientName')]"), "Click", false);
                                //fnSelectDropDownWE("Is Contact Same As Caller?", "//div[contains(@question-key, 'CONTACT_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Is Contact Same As Caller?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("IsContactSameAsCaller"));
                                fnSelectDropDownWElm("Is Contact Same As Caller?", "//div[contains(@question-key, 'CONTACT_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Is Contact Same As Caller?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("IsContactSameAsCaller"), false, false);


                                Thread.Sleep(5000);

                                //CommentsRemarks
                                //fnEnterTextWE("Comments Remarks", "//div[contains(@question-key, 'COMMENTS_REMARKS')]//div[@class='row' and div[span[text()='Comments / Remarks']]]//following-sibling::textarea", "Claim created by automation script");
                                //fnEnterTextWE("Internal Comments", "//div[contains(@question-key, 'COMMENTS_REMARKS')]//div[@class='row' and div[span[text()='Internal Comments']]]//following-sibling::textarea", "Internal comment test");

                                fnEnterTextWElm("Comments Remarks", "//div[contains(@question-key, 'COMMENTS_REMARKS')]//div[@class='row' and div[span[text()='Comments / Remarks']]]//following-sibling::textarea", "Claim created by automation script", false, false);
                                fnEnterTextWElm("Internal Comments", "//div[contains(@question-key, 'COMMENTS_REMARKS')]//div[@class='row' and div[span[text()='Internal Comments']]]//following-sibling::textarea", "Internal comment test", false, false);


                                Thread.Sleep(5000);

                                if (clsWE.fnElementExistNoReport("Loss Time", "//div[@id='list-example']//span[text()='Lost Time Information']", false, false))
                                {
                                    //Click Menu
                                    clsWE.fnClick(clsWE.fnGetWe("//div[@id='list-example']//span[text()='Lost Time Information']"), "Click Menu", false);
                                    Thread.Sleep(2000);
                                    //Lost Time Information
                                    //fnSelectDropDownWE("Employee Returned To Work", "//div[contains(@question-key, 'LOST_TIME_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Employee Returned To Work?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("EmployeeReturnedToWork"));
                                    fnSelectDropDownWElm("Employee Returned To Work", "//div[contains(@question-key, 'LOST_TIME_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Employee Returned To Work?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("EmployeeReturnedToWork"), false, false);

                                }

                                //Click on NEXT
                                clsWE.fnClick(clsWE.fnGetWe("//button[@type='submit' and text()='Next']"), "Click Next", false);
                                Thread.Sleep(5000);
                                if (!clsWE.fnElementExistNoReport("Error", "//*[@data-bind='text:ValidationMessage']", false, false))
                                {
                                    if (!clsWE.fnElementExistNoReport("Closing Script", "//span[text()='Closing Script']", false, false))
                                    {
                                        clsWE.fnClick(clsWE.fnGetWe("//button[@type='submit' and text()='Next']"), "Click Next", false);
                                        clsWE.fnPageLoad(clsWE.fnGetWe("//span[text()='Closing Script']"), "Closing Script", false, true);

                                    }
                                }


                                //Create Spreadsheet
                                string strFileCreated = fnCreateOfficeFormat(ConfigurationManager.AppSettings["BranchPath"], objData.fnGetValue("Contract") + "_BranchOffice.xlsx");

                                //Create State Dictinary
                                Dictionary<string, string> dicStates = new Dictionary<string, string>(); //Dic for States Names
                                Dictionary<string, int> dicStateIndex = new Dictionary<string, int>(); //Dic for States Index
                                Dictionary<string, int> dicDBState = new Dictionary<string, int>(); // Dic for DB State
                                Dictionary<string, string> dicSetupFile = new Dictionary<string, string>(); //Dic for Setup file


                                //Create Dictionary
                                clsData objSaveData = new clsData();
                                for (int intState = 0; intState <= garrStates.Length - 1; intState++)
                                {
                                    dicStates.Add(garrStates[intState], garrAbbre[intState]);
                                }

                                for (int intIndex = 1; intIndex <= garrAbbre.Length; intIndex++)
                                {
                                    dicStateIndex.Add(garrAbbre[intIndex - 1], intIndex + 1);
                                }
                                dicSetupFile = fnSetupfileBranch(ConfigurationManager.AppSettings["SetupfilePath"] + objData.fnGetValue("SetupFile"), objData.fnGetValue("SetupFileSheet"));

                                //Select each state
                                bool blEditState = true;
                                int intCount = 2;
                                foreach (string state in garrStates)
                                {
                                    string strBranchUI = "";
                                    string strDBBranch = "";
                                    //Click on Edit button
                                    clsWE.fnClick(clsWE.fnGetWe("//a[text()='Edit']"), "Click Edit", false);
                                    Thread.Sleep(3000);

                                    if (blEditState)
                                    {
                                        //Correct Benefit State? = NO
                                        //fnSelectDropDownWE("Correct Benefit State?", "//div[contains(@question-key, 'BENEFIT_STATE.CLAIM_BENEFIT_STATE_CORRECT_FLG')]//div[@class='row' and div[span[contains(text(), 'Correct Benefit State?')]]]//span[@class='select2-selection select2-selection--single']", "No");
                                        fnSelectDropDownWElm("Correct Benefit State?", "//div[contains(@question-key, 'BENEFIT_STATE.CLAIM_BENEFIT_STATE_CORRECT_FLG')]//div[@class='row' and div[span[contains(text(), 'Correct Benefit State?')]]]//span[@class='select2-selection select2-selection--single']", "No", false, false);
                                        clsWE.fnPageLoad(clsWE.fnGetWe("//button[@type='submit' and text()='Next']"), "Next Button", false, true);
                                        clsWE.fnPageLoad(clsWE.fnGetWe("//div[contains(@question-key, 'BENEFIT_STATE.CLAIM_CHANGE_BENEFIT_STATE_REASON')]//div[@class='row' and div[span[text()='Reason To Change Benefit State']]]//following-sibling::textarea"), "Reason Benefit State", false, false);
                                        //fnEnterTextWE("Reason To Change Benefit State", "//div[contains(@question-key, 'BENEFIT_STATE.CLAIM_CHANGE_BENEFIT_STATE_REASON')]//div[@class='row' and div[span[text()='Reason To Change Benefit State']]]//following-sibling::textarea", "Claim created by automation script");
                                        fnEnterTextWElm("Reason To Change Benefit State", "//div[contains(@question-key, 'BENEFIT_STATE.CLAIM_CHANGE_BENEFIT_STATE_REASON')]//div[@class='row' and div[span[text()='Reason To Change Benefit State']]]//following-sibling::textarea", "Claim created by automation script", false, false);
                                        blEditState = false;
                                        Thread.Sleep(3000);
                                    }

                                    //Select new state
                                    Thread.Sleep(2000);
                                    bool bSelected = true;
                                    //bSelected = fnSelectDropDownWE("Correct Benefit State?", "//div[contains(@question-key, 'BENEFIT_STATE.CLAIM_BENEFIT_STATE')]//div[@class='row' and div[span[contains(text(), 'Select Benefit State')]]]//span[@class='select2-selection select2-selection--single']", state);
                                    clsReportResult.fnLog("Branch Office Review", "******  The BO Review for state " + state + " will start  ******", "Info", false, false);
                                    bSelected = fnSelectDropDownWElm("Correct Benefit State?", "//div[contains(@question-key, 'BENEFIT_STATE.CLAIM_BENEFIT_STATE')]//div[@class='row' and div[span[contains(text(), 'Select Benefit State')]]]//span[@class='select2-selection select2-selection--single']", state, true, false);
                                    clsWE.fnPageLoad(clsWE.fnGetWe("//div[contains(@question-key, 'BENEFIT_STATE.CLAIM_BENEFIT_STATE')]//div[@class='row' and div[span[contains(text(), 'Select Benefit State')]]]//span[@class='select2-selection select2-selection--single']"), "Reason Benefit State", false, true);


                                    //clsWE.fnPageLoad(clsWE.fnGetWe("//div[contains(@question-key, 'BENEFIT_STATE.CLAIM_CHANGE_BENEFIT_STATE_REASON')]//div[@class='row' and div[span[text()='Reason To Change Benefit State']]]//following-sibling::textarea"), "Reason Benefit State", false, false);
                                    //clsReportResult.fnLog("Select New State", "A New State: "+ state +" was selected.", "Info", true);

                                    //Misouri Questions
                                    if (state == "Missouri")
                                    {
                                        //fnSelectDropDownWE("Is There An Endorsement For Missouri Employer First $500 Medical Payment?", "//div[contains(@question-key, 'STD_WOR_MO')]//div[@class='row' and div[span[contains(text(), 'Is There An Endorsement For Missouri Employer')]]]//span[@class='select2-selection select2-selection--single']", "No");
                                        //fnSelectDropDownWE("Should This Claim Be Processed As An Employer Self-Pay? ", "//div[contains(@question-key, 'STD_WOR_MO')]//div[@class='row' and div[span[contains(text(), 'Should This Claim Be Processed As An Employer')]]]//span[@class='select2-selection select2-selection--single']", "No");

                                        fnSelectDropDownWElm("Is There An Endorsement For Missouri Employer First $500 Medical Payment?", "//div[contains(@question-key, 'STD_WOR_MO')]//div[@class='row' and div[span[contains(text(), 'Is There An Endorsement For Missouri Employer')]]]//span[@class='select2-selection select2-selection--single']", "No", false, false);
                                        fnSelectDropDownWElm("Should This Claim Be Processed As An Employer Self-Pay?", "//div[contains(@question-key, 'STD_WOR_MO')]//div[@class='row' and div[span[contains(text(), 'Should This Claim Be Processed As An Employer')]]]//span[@class='select2-selection select2-selection--single']", "No", false, false);
                                    }

                                    //California Questions
                                    if (state == "California")
                                    {
                                        if (clsWE.fnElementExistNoReport("Loss Time", "//div[@id='list-example']//span[text()='Lost Time Information']", false, false))
                                        {
                                            //Click Menu
                                            clsWE.fnClick(clsWE.fnGetWe("//div[@id='list-example']//span[text()='Lost Time Information']"), "Click Menu", false);
                                            Thread.Sleep(2000);
                                            //Lost Time Information
                                            //fnSelectDropDownWE("Employee Returned To Work", "//div[contains(@question-key, 'LOST_TIME_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Employee Returned To Work?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("EmployeeReturnedToWork"));
                                            fnSelectDropDownWElm("Employee Returned To Work", "//div[contains(@question-key, 'LOST_TIME_INFORMATION')]//div[@class='row' and div[span[contains(text(), 'Employee Returned To Work?')]]]//span[@class='select2-selection select2-selection--single']", objData.fnGetValue("EmployeeReturnedToWork"), false, false);

                                        }
                                        Thread.Sleep(5000);
                                    }

                                    //Click Next
                                    clsWE.fnClick(clsWE.fnGetWe("//button[@type='submit' and text()='Next']"), "Click Next", false);
                                    Thread.Sleep(3000);
                                    /*if (!clsWE.fnElementExistNoReport("Error", "//*[@data-bind='text:ValidationMessage']", false, false))
                                    {
                                        if (!clsWE.fnElementExistNoReport("Closing Script", "//span[text()='Closing Script']", false, false))
                                        {
                                            clsWE.fnClick(clsWE.fnGetWe("//button[@type='submit' and text()='Next']"), "Click Next", false);
                                        }
                                    }*/
                                    clsWE.fnPageLoad(clsWE.fnGetWe("//span[text()='Closing Script']"), "Closing Script", false, true);

                                    //Get Branch Office
                                    if (bSelected)
                                    {
                                        //Get Attribute
                                        if (clsWE.fnElementExistNoReport("Branch Information", "//div[not(contains(@style, 'display: none;')) and contains(@data-bind, 'Answer.OfficeName')]//span[contains(@data-bind, 'OfficeNumber')]", false, false))
                                        {
                                            strBranchUI = clsWE.fnGetAttribute(clsWE.fnGetWe("//div[not(contains(@style, 'display: none;')) and contains(@data-bind, 'Answer.OfficeName')]//span[contains(@data-bind, 'OfficeNumber')]"), "Branch Information", "innerText", false).Replace("(", "").Replace(")", "");
                                            strDBBranch = fnExecuteQuery(objData.fnGetValue("Contract"), dicStates[state]);
                                        }
                                        else
                                        {
                                            //strBranch = "none";
                                            strBranchUI = "";
                                            strDBBranch = fnExecuteQuery(objData.fnGetValue("Contract"), dicStates[state]);
                                        }


                                        //Logic to verify the branch
                                        if (dicSetupFile[dicStates[state]] == strBranchUI)
                                        {
                                            //Setup File == UI
                                            clsReportResult.fnLog("Branch Office Review", "The BO for state: " + state + " matches as expected, the SetupFile and UI have the same office: " + strBranchUI, "Pass", true, false);
                                            objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "Comments", intCount, "Pass");
                                        }
                                        else if (dicSetupFile[dicStates[state]] == "" && strBranchUI != dicSetupFile[dicStates[state]])
                                        {
                                            //Setup File != UI
                                            clsReportResult.fnLog("Branch Office Review", "The BO for state: " + state + " is empty in the SetupFile: " + dicSetupFile[dicStates[state]] + " but display the following office in the UI: " + strBranchUI, "Info", true, false);
                                            objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "Comments", intCount, "Warning: Setup File is null but UI display a branch");
                                        }
                                        else
                                        {
                                            clsReportResult.fnLog("Branch Office Review", "The BO for state: " + state + " does not match, the SetupFile display: " + dicSetupFile[dicStates[state]] + " and the UI displays: " + strBranchUI, "Fail", true, false);
                                            objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "Comments", intCount, "Fail: The branches does not match with Setup File vs UI");
                                        }

                                        //Save Data
                                        objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "UI", intCount, strBranchUI);
                                        objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "DBViaone", intCount, strDBBranch);
                                        objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "SetupFile", intCount, (dicSetupFile[dicStates[state]]));
                                    }
                                    else
                                    {
                                        strDBBranch = fnExecuteQuery(objData.fnGetValue("Contract"), dicStates[state]);
                                        objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "UI", intCount, "Not found in UI");

                                        if (strDBBranch != "")
                                        {
                                            objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "DBViaone", intCount, strDBBranch);
                                        }
                                        else
                                        {
                                            objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "DBViaone", intCount, "Not found in DB");
                                        }


                                        if (dicSetupFile.ContainsKey(dicStates[state]))
                                        {
                                            objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "SetupFile", intCount, (dicSetupFile[dicStates[state]]));
                                        }
                                        else
                                        {
                                            objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "SetupFile", intCount, "Not found in Setup File");
                                        }

                                        clsReportResult.fnLog("Branch Office Review", "The BO for state: " + state + " cannot be validated because UI not display the state.", "Fail", false, false);
                                        objSaveData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + objData.fnGetValue("Contract") + "_BranchOffice.xlsx", "Sheet1", "Comments", intCount, "Fail: The BO review cannot be completed.");
                                    }



                                    intCount = intCount + 1;
                                }

                                //Get DB Branch Office
                                /*if (strFileCreated != "") 
                                {
                                    fnAddBranchesViaone(objData.fnGetValue("DBSet", ""), objData.fnGetValue("Contract", ""), dicStateIndex);
                                }*/

                            }
                        }
                        else
                        {
                            clsReportResult.fnLog("Start New Intake", "No active intakes were found in Select Intake screen.", "Fail", false, true);
                            blResult = false;
                        }
                    }
                    else
                    {
                        clsReportResult.fnLog("Select Client", "No clients were found in Search Client screen.", "Fail", false, true);
                        blResult = false;
                    }
                }
            }
            return blResult;
        }

        public static string fnGetLocationDB(string pstrClient, string pstrState)
        {
            clsDB objDBOR = new clsDB();
            objDBOR.fnOpenConnection(objDBOR.GetConnectionString("lltcsed1dvq-scan", "1521", "viaoner", "oferia", "P@ssw0rd#02"));
            DataTable datatable = new DataTable();
            string strQuery = "select * from viaone.cont_st_off where cont_num = '" + pstrClient + "' and deleted = 'N' and data_set = 'WC' and state = '" + pstrState + "'";
            datatable = objDBOR.fnDataSet(strQuery);
            if (datatable != null && datatable.Rows.Count > 0)
            {
                return Convert.ToString(datatable.Rows[0].Field<Int32>("EX_OFFICE"));
            }
            else
            {
                return "No Data";
            }
        }

        public string fnCreateOfficeFormat(string pstrFilePath, string pstrFileName)
        {
            try
            {
                string[] arrHeader = { "State", "Abbr", "SetupFile", "UI", "DBViaone", "Comments" };
                SLStyle headerStyle = new SLStyle();
                headerStyle.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);
                headerStyle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                headerStyle.Fill.SetPattern(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid, System.Drawing.Color.FromArgb(255, 247, 247), System.Drawing.Color.FromArgb(255, 247, 247));
                headerStyle.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Gray);
                headerStyle.SetRightBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Gray);
                headerStyle.SetLeftBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Gray);
                headerStyle.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Gray);
                headerStyle.Font.FontSize = 12;
                headerStyle.Font.Bold = true;

                SLStyle borderSTyle = new SLStyle();
                borderSTyle.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);
                borderSTyle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left);
                borderSTyle.SetBottomBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Gray);
                borderSTyle.SetRightBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Gray);
                borderSTyle.SetLeftBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Gray);
                borderSTyle.SetTopBorder(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, System.Drawing.Color.Gray);

                SLDocument document = new SLDocument();
                //Set Headers
                for (int intHeader = 0; intHeader <= arrHeader.Length - 1; intHeader++)
                {
                    document.SetCellValue(1, intHeader + 1, arrHeader[intHeader]);
                    document.SetCellStyle(1, intHeader + 1, headerStyle);
                    if (intHeader == 0)
                    {
                        document.SetColumnWidth(intHeader + 1, 35);
                    }
                    else
                    {
                        document.SetColumnWidth(intHeader + 1, 15);
                    }
                }

                //Set Data
                for (int intState = 0; intState <= garrStates.Length - 1; intState++)
                {
                    document.SetCellValue(intState + 2, 1, garrStates[intState]);
                    document.SetCellValue(intState + 2, 2, garrAbbre[intState]);

                    for (int intHeader = 0; intHeader <= arrHeader.Length - 1; intHeader++)
                    {
                        document.SetCellStyle(intState + 2, intHeader + 1, borderSTyle);
                    }
                }

                string strFullPath = string.Format(@"{0}{1}", pstrFilePath, pstrFileName);
                document.SaveAs(strFullPath);
                return strFullPath;
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred procesing the file: {0}", e.Message);
                return "";
            }
        }

        public bool fnAddBranchesViaone(string pstrSetNo, string pstrContract, Dictionary<string, int> pdicStates)
        {
            bool bStatus = false;
            clsData objData = new clsData();
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "LogInDataBase");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    clsDB objDBOR = new clsDB();
                    objDBOR.fnOpenConnection(objDBOR.GetConnectionString(objData.fnGetValue("Host", ""), objData.fnGetValue("Port", ""), objData.fnGetValue("Service", ""), objData.fnGetValue("User", ""), objData.fnGetValue("Password", "")));
                    DataTable datatable = new DataTable();
                    string strQuery = "select * from viaone.cont_st_off where cont_num = '" + pstrContract + "' and deleted = 'N'and data_set = 'WC'";
                    datatable = objDBOR.fnDataSet(strQuery);
                    Thread.Sleep(3000);
                    if (datatable != null && datatable.Rows.Count > 0)
                    {
                        foreach (DataRow row in datatable.Rows)
                        {
                            if (pdicStates.ContainsKey(Convert.ToString(row["STATE"])))
                            {
                                objData.fnSaveValue(ConfigurationManager.AppSettings["BranchPath"] + pstrContract + "_BranchOffice.xlsx", "Sheet1", "DBViaone", pdicStates[Convert.ToString(row["STATE"])], Convert.ToString(row["EX_OFFICE"]));


                            }
                        }
                        bStatus = true;
                    }
                    objDBOR.fnCloseConnection();
                }
            }
            return bStatus;
        }

        public string fnExecuteQuery(string pstrContract, string pstrSatate)
        {
            string strResult = "";
            clsDB objDBOR = new clsDB();
            objDBOR.fnOpenConnection(objDBOR.GetConnectionString(ConfigurationManager.AppSettings["Host"], ConfigurationManager.AppSettings["Port"], ConfigurationManager.AppSettings["Service"], ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Password"]));
            DataTable datatable = new DataTable();
            string strQuery = "select * from viaone.cont_st_off where cont_num = '" + pstrContract + "' and deleted = 'N' and data_set = 'WC' and state = '" + pstrSatate + "'";
            datatable = objDBOR.fnDataSet(strQuery);
            Thread.Sleep(3000);
            if (datatable != null && datatable.Rows.Count > 0)
            {
                DataRow row = datatable.Rows[0];
                strResult = Convert.ToString(row["EX_OFFICE"]);
            }
            else
            {
                strResult = "";
            }
            objDBOR.fnCloseConnection();

            return strResult;
        }

        public Dictionary<string, string> fnSetupfileBranch(string pstrFilePath, string pstrSheet)
        {
            Dictionary<string, string> dicSFStates = new Dictionary<string, string>();
            clsData objData = new clsData();
            objData.fnLoadFile(pstrFilePath, pstrSheet);
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (!dicSFStates.ContainsKey(objData.fnGetValue("Abbv", ""))) { dicSFStates.Add(objData.fnGetValue("Abbv", ""), objData.fnGetValue("JURISFROI", "")); }
            }
            return dicSFStates;
        }


        public bool fnCreateIntake(string pstrSetNo)
        {
            bool blResult = true;
            clsReportResult.fnLog("fnCreateIntake", "Create Intake Functions Starts.", "Info", false);
            clsData objData = new clsData();
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], "EventInfo");
            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Set", "") == pstrSetNo)
                {
                    //Go to New Intake Page
                    clsWE.fnPageLoad(clsWE.fnGetWe("//nav[@id='EnvironmentBar']"), "Dashoard Page", false, false);
                    fnHamburgerMenu("New Intake");
                    //Open Select Client Popup
                    clsWE.fnPageLoad(clsWE.fnGetWe("//button[@id='selectClient_']"), "Select Client Page", false, true);
                    clsWE.fnClick(clsWE.fnGetWe("//button[@id='selectClient_']"), "Select Client Button", false);
                    clsWE.fnPageLoad(clsWE.fnGetWe("//div[contains(@style, 'display: block;')]//table[@id='clientSelectorTable_']"), "Select Client Popup.", false, true);
                    clsWE.fnClick(clsWE.fnGetWe("//input[@placeholder='Client Number or Name']"), "Client Number Button", false);
                    clsWE.fnSendKeys(clsWE.fnGetWe("//input[@placeholder='Client Number or Name']"), "Client Number", objData.fnGetValue("Contract", ""), true);
                    //Verify that table exist
                    if (clsWE.fnElementNotExist("Start New Intake", "//td[text()='No matching records found']", false, false))
                    {
                        //Filter by Script Name
                        //Filter Script Name
                        clsWE.fnClick(clsWE.fnGetWe("(//a[text()='Select'])[1]"), "Select", false);
                        clsWE.fnClick(clsWE.fnGetWe("//input[@placeholder='Filter Results']"), "Filter Results", false);
                        clsWE.fnSendKeys(clsWE.fnGetWe("//input[@placeholder='Filter Results']"), "Filter Results", objData.fnGetValue("IntakeType", ""));
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        clsReportResult.fnLog("Select Client Popup", "The Client No: was not found in the popup.", "Fail", false, true);
                        blResult = false;
                    }
                }
            }
            return blResult;
        }

        public void fnHamburgerMenu(string sptrMenu)
        {
            clsReportResult.fnLog("Hamburger Menu", "Selecting a Hamburger Menu Option.", "Info", false);
            clsWE.fnClick(clsWE.fnGetWe("//div[@class='float-left']//i"), "Hamburger Button", false);
            Thread.Sleep(1000);
            clsWE.fnClick(clsWE.fnGetWe("//a[contains(text(), '" + sptrMenu + "')]"), sptrMenu + " Link", false, true);
        }

    }
}
