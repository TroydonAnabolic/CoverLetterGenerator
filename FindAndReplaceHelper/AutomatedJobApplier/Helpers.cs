using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace FindAndReplaceHelper.AutomatedJobApplier
{
    class Helpers
    {
        public WebDriverWait Wait(IWebDriver driver)
        {
            return new WebDriverWait(driver,
                        System.TimeSpan.FromSeconds(120));
        }

        public StringBuilder ApplyHyphernsToSearchTerm(string argSearchTerm)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var letter in argSearchTerm)
            {
                if (letter == ' ') sb.Append('-');
                else sb.Append(letter);
            }

            return sb;
        }
        public bool CheckIfXPathElementExist(IWebDriver driver, string XPath)
        {
            if (driver.FindElements(By.XPath(XPath)).Count != 0) return true;
            else return false;
        }

        public void ClickOnBodyElement(IWebDriver driver)
        {
            driver.FindElement(By.CssSelector("body")).Click();
        }

        public IWebElement GetBodyElement(IWebDriver driver)
        {
            // if we get an alert when we click away, we accept it, then try to return body again
            try
            {
                return driver.FindElement(By.CssSelector("body"));
            }
            catch (UnhandledAlertException unexpectedAlert)
            {
                Console.WriteLine(unexpectedAlert.Message);

                IAlert alert = driver.SwitchTo().Alert();
                alert.Accept();

                return driver.FindElement(By.CssSelector("body"));
            }
        }

        public IJavaScriptExecutor BypassOverlayUsingJavaScriptExecutorOnBody(Helpers helper, IWebDriver driver)
        {
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            executor.ExecuteScript("arguments[0].click();", helper.GetBodyElement(driver));
            // helper.GetBodyElement(driver).Click();
            return executor;
        }

        public void BypassOverlayUsingJavaScriptExecutor(IJavaScriptExecutor executor, IWebDriver driver, IWebElement element)
        {
            executor.ExecuteScript("arguments[0].click();", element);
        }

        // uses XPath for selection
        public IJavaScriptExecutor ClickBypassOverlay(WebDriverWait wait, Helpers helper, string elementToByPass, IWebDriver driver)
        {
            //Task.Delay(1000).Wait();
            //helper.ClickOnBodyElement(driver); // click on body again to make element interactable // TODO - check if this is cause alert, also check next body click
            Task.Delay(1000).Wait();
            IJavaScriptExecutor executor = helper.BypassOverlayUsingJavaScriptExecutorOnBody(helper, driver);
            if (CheckIfXPathElementExist(driver, elementToByPass)) // add above element if issues
            {
                wait.Until(theDriver =>
                    theDriver.FindElement(By.XPath(elementToByPass)));
                IWebElement elementButton = driver.FindElement(By.XPath(elementToByPass));
                if (elementButton.Text != "Back")
                    helper.BypassOverlayUsingJavaScriptExecutor(executor, driver, elementButton);
            };

            return executor;
        }

        public void SaveFileToAppliedJobs(string jobIdNumber, string saveType)
        {
            string filePath = string.Empty;

            if (saveType.Equals("applied", StringComparison.OrdinalIgnoreCase))
                filePath = @"D:\Troydon\Documents\IT_Project\FindAndReplaceHelper\AutomatedJobApplyer_2\FindAndReplaceHelper\ListOfJobsAppliedFor.txt";
            else if (saveType.Equals("refer", StringComparison.OrdinalIgnoreCase))
                filePath = @"D:\Troydon\Documents\IT_Project\FindAndReplaceHelper\AutomatedJobApplyer_2\FindAndReplaceHelper\ReferredToAdvertWebsiteList.txt";

            using (StreamWriter sw = File.AppendText(filePath))
            {
                sw.Write("\n" + jobIdNumber);
            }
        }
        // 
        public void CoverFileUpload(IWebDriver driver)
        {
            //The first step gets the base directory and the file
            Task.Delay(1000).Wait();
            IWebElement uploadCoverButton = driver.FindElement(By.Id("coverLetterFile"));
            uploadCoverButton.SendKeys(@"D:\Troydon\Documents\JobStuff\JobHunt\cover letters\Customs\SoftwareDeveloperCover.docx");
            Task.Delay(1000).Wait();
            driver.FindElement(By.CssSelector("#bottom-nav > div > div > button")).Click();
        }
    }
}
