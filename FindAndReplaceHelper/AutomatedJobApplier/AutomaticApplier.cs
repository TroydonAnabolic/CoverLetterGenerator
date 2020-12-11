﻿using FindAndReplaceHelper.CoverBuilder;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading.Tasks;

namespace FindAndReplaceHelper.AutomatedJobApplier
{
    class AutomaticApplier
    {
        IWebDriver driver;

        public void BeginApplicationProcess()
        {
            driver = new ChromeDriver("D:\\3rdparty\\chrome");
            driver.Manage().Window.Minimize();
            Helpers helper = new Helpers();
            WebDriverWait wait = helper.Wait(driver);
            Actions action = new Actions(driver);

            // InitiateApplicationProcess(wait);
            BeginApplication(wait, helper);
            EndApplicationProcess();
        }

        // defines what each job search does
        void BeginApplication(WebDriverWait wait, Helpers helper)
        {

            bool continueApp = true;

            while (continueApp)
            {
                Console.WriteLine("Enter job Application Type - 'SEEK' for autopopulation of cover, 'Other' for other type of application\n" +
                    "(Not case sensitive)");
                string coverType = Console.ReadLine(), closeAppWord = "exit";

                if (coverType.Equals(closeAppWord, StringComparison.OrdinalIgnoreCase))
                {
                    continueApp = false;
                }
                else if (coverType.Equals("SEEK", StringComparison.OrdinalIgnoreCase))
                {
                    ApplySeek(wait, helper);
                }
                else if (coverType.Equals("Other", StringComparison.OrdinalIgnoreCase))
                {
                    ApplyOther(wait, helper);
                }
            }

            // for each advert, check if it has a one of the keywords for the search, if so nav to the ad an apply for it
           
        }

        private void ApplyOther(WebDriverWait wait, Helpers helper)
        {
            Console.WriteLine("\n***********************************************************************************");
            Console.WriteLine("********************** WELCOME TO OTHER BUILDER ***********************************");
            Console.WriteLine("***********************************************************************************\n");

            bool continueGeneration = true;

            while (continueGeneration)
            {
                string  exitWord = "exit", jobPositionTitle = string.Empty, companyName = string.Empty;

                if (!jobPositionTitle.Equals(exitWord, StringComparison.OrdinalIgnoreCase) && !companyName.Equals(exitWord, StringComparison.OrdinalIgnoreCase))
                {
                    Console.Write("Position Title: ");
                    jobPositionTitle = Console.ReadLine();

                    Console.Write("\nPosition Title: ");
                    companyName = Console.ReadLine();

                    CoverBuilding coverBuilding = new CoverBuilding();
                    coverBuilding.StartApplication(jobPositionTitle, companyName);
                }
                else if 
                    (jobPositionTitle.Equals(exitWord, StringComparison.OrdinalIgnoreCase) && companyName.Equals(exitWord, StringComparison.OrdinalIgnoreCase))
                        continueGeneration = false;
            }
        }

        private void ApplySeek(WebDriverWait wait, Helpers helper)
        {
            Console.WriteLine("\n***********************************************************************************");
            Console.WriteLine("******************* WELCOME TO SEEK COVER BUILDER *********************************");
            Console.WriteLine("***********************************************************************************\n");

            bool continueGeneration = true;

            while (continueGeneration)
            {
                Console.WriteLine("Enter the link to the job you wish to apply for: ");
                string advertisementLink = Console.ReadLine(), exitWord = "exit";

                if (!advertisementLink.Equals(exitWord, StringComparison.OrdinalIgnoreCase))
                {

                    try
                    {
                        if (!advertisementLink.Contains($"https://www.seek.com.au/job/"))
                            throw new InvalidLinkException(advertisementLink);

                        driver.Url = advertisementLink;


                        if (helper.CheckIfXPathElementExist(driver, Constants.jobTitleXPath))
                        {
                            Task.Delay(1000).Wait();

                            wait.Until(driver =>
                            driver.FindElement(By.XPath(Constants.jobTitleXPath)));
                            IWebElement jobTitleEl = driver.FindElement(By.XPath(Constants.jobTitleXPath));

                            wait.Until(driver =>
                              driver.FindElement(By.XPath(Constants.jobTitleXPath)));
                            IWebElement companyNameEl = driver.FindElement(By.XPath(Constants.companyNameXPath));

                            string jobPositionTitle = jobTitleEl.Text;
                            string companyName = companyNameEl.Text;

                            CoverBuilding coverBuilding = new CoverBuilding();
                            coverBuilding.StartApplication(jobPositionTitle, companyName);
                        }
                    }
                    catch (InvalidLinkException invalidSeekLink)
                    {
                        Console.WriteLine(invalidSeekLink.Message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                else if (advertisementLink.Equals(exitWord, StringComparison.OrdinalIgnoreCase))
                    continueGeneration = false;
            }
        }

        public void EndApplicationProcess()
        {
            driver.Close();
        }
    }
}