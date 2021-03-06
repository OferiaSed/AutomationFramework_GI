using AutomationFramework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
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

        public void WaitWEUntilAppears(string pstrStepName, string pstrLocator, int pintTime)
        {
            do { Thread.Sleep(pintTime); }
            while (!clsWE.fnElementExistNoReport(pstrStepName, pstrLocator, false));
        }

        public bool fnCleanAndEnterText(string pstrElement, string pstrWebElement, string pstrValue, bool pblScreenShot = false, bool pblHardStop = false, string pstrHardStopMsg = "", bool bWaitHeader = true)
        {
            bool blResult = true;
            if (pstrValue != "" && pstrWebElement != "")
            {
                try
                {
                    clsReportResult.fnLog("SendKeys", "Step - Sendkeys on " + pstrElement, "info", false, false);
                    IWebElement objWebEdit = clsWebBrowser.objDriver.FindElement(By.XPath(pstrWebElement));
                    Actions action = new Actions(clsWebBrowser.objDriver);
                    objWebEdit.Click();
                    action.KeyDown(Keys.Control).SendKeys(Keys.Home).Perform();
                    objWebEdit.SendKeys(Keys.Delete);
                    objWebEdit.SendKeys(pstrValue);
                    Thread.Sleep(TimeSpan.FromMilliseconds(500));
                    //Thread.Sleep(1000);

                    if (bWaitHeader)
                    {
                        if (IsElementPresent("//span[@data-bind='text: headerClientName']"))
                        {
                            objWebEdit = clsWebBrowser.objDriver.FindElement(By.XPath("//span[@data-bind='text: headerClientName']"));
                            Thread.Sleep(TimeSpan.FromMilliseconds(500));
                            objWebEdit.Click();
                            Thread.Sleep(TimeSpan.FromMilliseconds(500));
                        }
                    }
                    blResult = true;
                }
                catch (Exception objException)
                {
                    blResult = false;
                    clsWebElements.fnExceptionHandling(objException);
                }
                if (blResult)
                {
                    //Step - Click on Submit Button
                    clsReportResult.fnLog("SendKeys", "The SendKeys for: " + pstrElement + " with value: " + pstrValue + " was done successfully.", "Pass", pblScreenShot, pblHardStop);
                }
                else
                {
                    blResult = false;
                    clsReportResult.fnLog("SendKeys", "The SendKeys for: " + pstrElement + " with value: " + pstrValue + " has failed.", "Fail", true, pblHardStop, pstrHardStopMsg);
                }
            }
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
                    Thread.Sleep(1000);
                    //Thread.Sleep(1500);

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
                        Thread.Sleep(1000);
                        //Thread.Sleep(1500);
                        objDropDown.Click();
                        //Thread.Sleep(1000);
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

        public void fnHamburgerMenu(string pstrMenu)
        {
            clsReportResult.fnLog("Hamburger Menu", "Selecting a Menu Option: " + pstrMenu.ToString(), "Info", false);
            //Verify if menu is collapsed
            if (IsElementPresent("//div[@id='slide-out' and not(contains(@style, 'translateX(0px)'))]")) 
            { clsWE.fnClick(clsWE.fnGetWe("//div[@class='float-left']//i"), "Hamburger Button", false); }
            //Select Menu Item
            if (!pstrMenu.Contains(";"))
            {
                Thread.Sleep(TimeSpan.FromSeconds(2));
                if (IsElementPresent("//a[contains(text(), '" + pstrMenu + "')]"))
                {
                    clsWE.fnClick(clsWE.fnGetWe("//a[contains(text(), '" + pstrMenu + "')]"), pstrMenu + " Link", false, false);
                    Thread.Sleep(TimeSpan.FromSeconds(2));
                }
                else
                {
                    clsReportResult.fnLog("Hamburger Menu", "The menu/submenu: "+ pstrMenu + " does not exist.", "Info", false);
                }
            }
            else 
            {
                string[] arrMenu = pstrMenu.Split(';');
                for(int intMenu = 0; intMenu <= arrMenu.Length -1; intMenu++)
                {
                    if (intMenu == 0)
                    {
                        if (IsElementPresent("//li[contains(@data-bind, '" + arrMenu[intMenu].Replace(" ", "") + "')]/a"))
                        {
                            Thread.Sleep(TimeSpan.FromMilliseconds(500));
                            clsWE.fnClick(clsWE.fnGetWe("//li[contains(@data-bind, '" + arrMenu[intMenu].Replace(" ", "") + "')]/a"), arrMenu[intMenu] + " Link", false, false);
                            Thread.Sleep(TimeSpan.FromMilliseconds(500));
                        }
                        else
                        {
                            clsReportResult.fnLog("Hamburger Menu", "The menu/submenu: " + arrMenu[intMenu] + " does not exist.", "Info", false);
                        }
                    }
                    else
                    {
                        if (IsElementPresent("//a[contains(text(), '" + arrMenu[intMenu] + "')]"))
                        {
                            Thread.Sleep(TimeSpan.FromMilliseconds(500));
                            clsWE.fnClick(clsWE.fnGetWe("//a[contains(text(), '" + arrMenu[intMenu] + "')]"), arrMenu[intMenu] + " Link", false, false);
                            Thread.Sleep(TimeSpan.FromMilliseconds(500));
                        }
                        else
                        {
                            clsReportResult.fnLog("Hamburger Menu", "The menu/submenu: " + arrMenu[intMenu] + " does not exist.", "Info", false);
                        }
                    }
                    
                }

            }
        }

        



    }
}
