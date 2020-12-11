using System.Collections.Generic;

namespace FindAndReplaceHelper.AutomatedJobApplier
{
    class Constants
    {
        public const string userName = "troy.incarnate@gmail.com";
        public const string seekHomePage = "https://www.seek.com.au/";
        public const string loginPage = "https://www.seek.com.au/sign-in/";

        public static readonly string[] searchTerms = new string[]
        {
            "junior software developer", "game developer", "junior cloud", "junior qa", "it support",
        };

        public static readonly string[] locationsToWork = new string[]
       {
             "All Sydney NSW", "Queensland QLD",
       };

        public static readonly string[] developerJobTitles = new string[]
        {
            "C#", ".NET", "Developer", "Programmer", "Dev", "Software Engineer"
        };

        public static readonly Dictionary<string, string> specificQuestionXPaths = new Dictionary<string, string>()
        {
            { "//*[@id='question-7821']", "How many years' experience do you have as a javascript developer?" },
        };

        // static xPaths

        public const string jobTitleXPath = "//*[@id='app']/div/div[4]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/article/section/span[2]/span/h1";
        public const string companyNameXPath = "//*[@id='app']/div/div[4]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/article/section/h2/span[2]/span";
       
    }
}
