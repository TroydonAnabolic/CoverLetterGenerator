using FindAndReplaceHelper.CoverBuilder;
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
            BeginApplication();
            EndApplicationProcess();
        }

        void BeginApplication()
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
                    ApplySeek();
                }
                else if (coverType.Equals("Other", StringComparison.OrdinalIgnoreCase))
                {
                    ApplyOther();
                }
                else Console.WriteLine("\nInvalid input, type 'Seek' or 'Other', then hit Enter.\n");
            }
        }

        private void ApplyOther()
        {
            Console.WriteLine("\n***********************************************************************************");
            Console.WriteLine("********************* WELCOME TO OTHER COVER BUILDER *********************************");
            Console.WriteLine("***********************************************************************************\n");

            bool continueGeneration = true;

            while (continueGeneration)
            {
                string exitWord = "exit", jobPositionTitle = string.Empty, companyName = string.Empty;

                Console.Write("Position Title: ");
                jobPositionTitle = Console.ReadLine();

                if (jobPositionTitle.Equals(exitWord, StringComparison.OrdinalIgnoreCase) || companyName.Equals(exitWord, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine();
                    return;
                }

                Console.Write("\nCompany Name: ");
                companyName = Console.ReadLine();

                if (jobPositionTitle.Equals(exitWord, StringComparison.OrdinalIgnoreCase) || companyName.Equals(exitWord, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine();
                    return;
                }

                CoverBuilding coverBuilding = new CoverBuilding();
                coverBuilding.StartApplication(jobPositionTitle, companyName);

                Console.Beep(37, 3);
            }
        }

        private void ApplySeek()
        {
            Console.WriteLine("\n***********************************************************************************");
            Console.WriteLine("******************* WELCOME TO SEEK COVER BUILDER *********************************");
            Console.WriteLine("***********************************************************************************\n");


            driver = new ChromeDriver("D:\\3rdparty\\chrome");
            driver.Manage().Window.Minimize();
            Helpers helper = new Helpers();
            WebDriverWait wait = helper.Wait(driver);
            Actions action = new Actions(driver);
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
                        string jobPositionTitle = string.Empty, companyName = string.Empty;

                        if (helper.CheckIfXPathElementExist(driver, Constants.jobTitleXPath))
                        {
                            Task.Delay(1000).Wait();

                            wait.Until(driver =>
                            driver.FindElement(By.XPath(Constants.jobTitleXPath)));
                            IWebElement jobTitleEl = driver.FindElement(By.XPath(Constants.jobTitleXPath));
                            jobPositionTitle = jobTitleEl.Text;

                            if (helper.CheckIfXPathElementExist(driver, Constants.companyNameXPath))
                            {
                                wait.Until(driver =>
                                    driver.FindElement(By.XPath(Constants.companyNameXPath)));
                                IWebElement companyNameEl = driver.FindElement(By.XPath(Constants.companyNameXPath));
                                companyName = companyNameEl.Text;
                            }
                            else
                            {
                                wait.Until(driver =>
                                  driver.FindElement(By.XPath(Constants.privateAdvertiserXPath)));
                                IWebElement privateAdvertiserEl = driver.FindElement(By.XPath(Constants.privateAdvertiserXPath));
                                companyName = privateAdvertiserEl.Text;
                            }



                            CoverBuilding coverBuilding = new CoverBuilding();
                            coverBuilding.StartApplication(jobPositionTitle, companyName);

                            Console.Beep(37, 3); // beep to notify when cover is ready - TODO: maybe get aws text to speec to read out job title and company in cover builder to confirm correct details generated
                            // then say application is ready
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
