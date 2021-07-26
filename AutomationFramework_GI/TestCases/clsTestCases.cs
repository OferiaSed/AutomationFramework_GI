using AutomationFramework;
using AutomationFramework_GI.POM;
using AutomationFramework_GI.Utils;
using NUnit.Framework;
using OpenPop.Mime;
using OpenPop.Pop3;
using OpenQA.Selenium;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace AutomationFramework_GI.TestCases
{
    [TestFixture]
    class clsTestCases : clsWebBrowser
    {

        public bool blStop;
        public clsReportResult clsRR = new clsReportResult();
        public clsWebElements clsWE = new clsWebElements();


        [OneTimeSetUp]
        public void BeforeClass()
        {
            blStop = clsReportResult.fnExtentSetup();
            if (!blStop)
                AfterClass();
        }

        [SetUp]
        public void SetupTest()
        {
            clsReportResult.objTest = clsReportResult.objExtent.CreateTest(TestContext.CurrentContext.Test.Name);
            fnOpenBrowser(ConfigurationManager.AppSettings["Browser"]);
        }

        [Test]
        public void Test_OCR() 
        {
            /*
            clsEmail email = new clsEmail();
            email.strFromEmail = "sedgautoemail@gmail.com";
            email.strPassword = "P4ssw0rd!123";
            email.strServer = "popgmail";
            string xTest = email.fnReadTwoFactorToken();
            string ytest = "";
            */


            /*
            clsEmail email = new clsEmail();
            email.strFromEmail = "sedgautoemail@gmail.com";
            email.strPassword = "P4ssw0rd!123";
            email.strToEmail = "omar.feria@sedgwick.com";
            email.strSubject = "Test Subject";
            email.strContentBody = "<h3> Test mail from selenium. </h3>";
            email.strServer = "gmail";
            email.fnSendSimpleEmail();
            */

         

        }

        [Test]
        public void Test_Document()
        {
            string pstrFilePath = @"\\memfp02\share\Any\4th_Automation\Data\ViaOne\GlobalIntake\InputFile\BenefitStates\";
            string pstrFileName = "3505_BranchOffice.xlsx";

            string[] arrAbbre = { "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DC", "DE", "FL", "GA", "HI", "IA", "ID", "IL", "IN", "KS", "KY", "LA", "MA", "MD", "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE", "NH", "NJ", "NM", "NV", "NY", "OH", "OK", "OR", "PA", "PR", "RI", "SC", "SD", "TN", "TX", "UT", "VA", "VT", "WA", "WI", "WV", "WY" };
            string[] arrStates = { "Alaska", "Alabama", "Arkansas", "Arizona", "California", "Colorado", "Connecticut", "District of Columbia", "Delaware", "Florida", "Georgia", "Hawaii", "Iowa", "Idaho", "Illinois", "Indiana", "Kansas", "Kentucky", "Louisiana", "Massachusetts", "Maryland", "Maine", "Michigan", "Minnesota", "Missouri", "Mississippi", "Montana", "North Carolina", "North Dakota", "Nebraska", "New Hampshire", "New Jersey", "New Mexico", "Nevada", "New York", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Puerto Rico", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Virginia", "Vermont", "Washington", "Wisconsin", "West Virginia", "Wyoming" };
            string[] arrHeader = { "State", "Abbr", "SetupFile", "UI GI", "DB Viaone" };
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
            for (int intState = 0; intState <= arrStates.Length - 1; intState++)
            {
                document.SetCellValue(intState + 2, 1, arrStates[intState]);
                document.SetCellValue(intState + 2, 2, arrAbbre[intState]);

                for (int intHeader = 0; intHeader <= arrHeader.Length - 1; intHeader++)
                {
                    document.SetCellStyle(intState + 2, intHeader + 1, borderSTyle);
                }
            }



            string strFullPath = string.Format(@"{0}{1}", pstrFilePath, pstrFileName);
            document.SaveAs(strFullPath);
        }


        [Test]
        public void Test_Locations()
        {
            //Create State Dictinary
            string[] arrStates = { "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming" };
            string[] arrAbbre = { "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY" };
            Dictionary<string, string> dicStates = new Dictionary<string, string>();
            for (int intState = 0; intState <= arrStates.Length - 1; intState++)
            {
                dicStates.Add(arrStates[intState], arrAbbre[intState]);
            }

            //Print Dictonary
            foreach (KeyValuePair<string, string> kvp in dicStates)
            {
                Console.WriteLine("state: {0} --- abrr: {1} --- DB: {2}", kvp.Key, kvp.Value, clsMegaIntake.fnGetLocationDB("9413", kvp.Value.ToString()));
            }
        }



        [Test]
        public void Test_fnPageLoadPass()
        {

            bool blStatus = true;
            clsMegaIntake clsMG = new clsMegaIntake();
            clsData objData = new clsData();
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], ConfigurationManager.AppSettings["Sheet"]);
            fnNavigateToUrl(clsMG.fnGetURLEnv((objData.fnGetValue("Environment", ""))));

            for (int intRow = 2; intRow <= objData.RowCount; intRow++)
            {
                objData.CurrentRow = intRow;
                if (objData.fnGetValue("Run", "") == "1")
                {
                    string[] arrFunctions = objData.fnGetValue("Funcions").Split(';');
                    string[] arrValue = objData.fnGetValue("Values").Split(';');
                    int intCount = -1;
                    foreach (string item in arrFunctions)
                    {
                        intCount = intCount + 1;
                        var TempValue = arrValue[intCount].Split('=')[1];
                        switch (item.ToUpper())
                        {
                            case "LOGIN":
                                blStatus = clsMG.fnLoginData(TempValue);
                                break;
                            case "CREATEINTAKE":
                                blStatus = clsMG.fnCreateSubmitIntake(TempValue);
                                break;
                            case "BRANCHREVIEW":
                                blStatus = clsMG.fnBranchOfficeVerification(TempValue);
                                break;
                            case "INTAKE":
                                blStatus = clsMG.fnCreateIntake(TempValue);
                                break;
                            case "LOCATION":
                                break;
                            default:
                                break;
                        }

                    }
                }
            }
        }

        [Test]
        public void PassResetRestrictions()
        {
            clsMegaIntake clsMG = new clsMegaIntake();
            clsData objData = new clsData();
            objData.fnLoadFile(ConfigurationManager.AppSettings["FilePath"], ConfigurationManager.AppSettings["Sheet"]);
            fnNavigateToUrl(clsMG.fnGetURLEnv((objData.fnGetValue("Environment", ""))));
            clsLogin login = new clsLogin();
            login.fnPassResetRestrictions("10");
        }

        [TearDown]
        public void CloseTest()
        {
           fnCloseBrowser();
           clsReportResult.fnExtentClose();
        }

        [OneTimeTearDown]
        public void AfterClass()
        {
            try
            {
                clsReportResult.objExtent.Flush();

            }
            catch (Exception objException)
            {
                Console.WriteLine("Error: {0}", objException.Message);
            }
        }


    }

}
