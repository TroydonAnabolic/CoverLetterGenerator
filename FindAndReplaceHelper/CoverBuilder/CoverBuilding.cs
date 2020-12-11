using FindAndReplaceHelper.AutomatedJobApplier;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace FindAndReplaceHelper.CoverBuilder
{
    class CoverBuilding
    {
        Application word;
        object miss;
        Document docs;

        public void StartApplication(string output, string company)
        {
            //int? count;
            //string relevantSkills = string.Empty;
            //// if none of the desired job titles exist then leave page
            //if (!CheckIfJobTitleOfInterest(output, ref relevantSkills, out count))
            //{
            //    continueApplication = false;
            //    return;
            //}
            //continueApplication = true;

            GetFileApplicationTemplate(out word, out miss, out docs);

            // TODO: if possible execute the shutdown word method to execute if compilation shuts unexpectedly or cancelled to avoid it being left opened.
            //if (Environment.Exit(exitCode)) @"D:\Troydon\Documents\JobStuff\JobHunt\cover letters\Customs\SEEK_Free_cover_letter_template_2018_NZ.docx"

            GenerateCover(ref word, ref miss, ref docs, output, company);
        }

        private static void GenerateCover(ref Application word, ref object miss, ref Document docs, string output, string company)
        {
            try
            {
                bool outputMatch = true;
                // determine outputfile name
                object outputFI = DetermineOutput(ref word, ref miss, ref docs, output, company);

                // if there was a match in desired titles to apply for then generate a cover
                if (outputMatch)
                    docs.SaveAs(ref outputFI, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

                // then close the word document and close MS word process
            }
            catch (COMException ex)
            {
                Console.WriteLine("COM Exception Error, Application is now closing\n" + ex.Message);
                CloseWordApplication(ref word, ref miss, ref docs);
            }
            finally
            {
                CloseWordApplication(ref word, ref miss, ref docs);
            }
        }

        private static object DetermineOutput(ref Application word, ref object miss, ref Document docs, string output, string company)
        {
            object outputFI = new object();
            string inputError = "\nSorry you did not enter a valid input\n", hiringManagersName = "", relevantSkills = "", bonusQuestion = ""; // check here if issues where role title shd go
            int? count = 0;
            
            List<string> outputPath = new List<string>();
            bool outputMatch = SelectOutputMethod(ref word, ref miss, ref docs, ref outputFI, ref output, inputError, ref hiringManagersName, ref relevantSkills, ref company, ref bonusQuestion,
                ref output, ref count);

            return outputFI;
        }

        private static void GetFileApplicationTemplate(out Application word, out object miss, out Document docs)
        {
            byte[] inputBuffer = new byte[1024];
            Stream inputStream = Console.OpenStandardInput(inputBuffer.Length);
            Console.SetIn(new StreamReader(inputStream, Console.InputEncoding, false, inputBuffer.Length));

            // Create new word application
            word = new Application();
            //word.Visible = true; // opens the doc
            // load all the MS Word data
            miss = System.Reflection.Missing.Value;
            // object path = @"D:\//Troydon/Documents/JobStuff/JobHunt/cover letters/Customs/SEEK_Free_cover_letter_template_2018_NZ.docx";
            //  @"D:\Troydon\Documents\JobStuff\JobHunt\cover letters\Customs\SEEK_Free_cover_letter_template_2018_NZ.docx"
            //  Console.WriteLine("Please enter the file path you would like to use for the template.\nE.g.M:\\/Troydon/Documents/Troydon/JobStuff/JobHunt/cover letters/Customs/SEEK_Free_cover_letter_template_2018_NZ.docx");
            Console.WriteLine();
            object path = @"D:\//Troydon/Documents/JobStuff/JobHunt/cover letters/Customs/SEEK_Free_cover_letter_template_2018_NZ.docx" ?? string.Empty;
            object readOnly = true;
            docs = word.Documents.Open(ref path, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
        }

        private static bool SelectOutputMethod(ref Application word, ref object miss, ref Document docs, ref object outputFI, ref string output, string inputError, ref string hiringManagersName,
            ref string relevantSkills, ref string company, ref string bonusQuestion, ref string ouput, ref int? count)
        {
            bool outputMatch = true;
            do
            {
                // if we at count 0 we will ask the first question
                if (count == 0)
                {
                    ExitApp(ref word, ref miss, ref docs, output);
                    // if job title is not on the list, go to the next job title
                    if (!CheckIfJobTitleOfInterest(output, ref relevantSkills, out count)) return false;
                    Console.WriteLine("Relevant skills set");
                }

                string date = DateTime.Now.ToString("dddd, dd MMMM yyyy");
                Console.WriteLine("Date Set");

                if (count == 1)
                {
                    // hiringManagersName = Console.ReadLine();
                    ExitApp(ref word, ref miss, ref docs, hiringManagersName);
                    count = 2;
                }
              //  Console.WriteLine("\nEnter the company (hit 'n' if none)\n");
                if (count == 2)
                {
                    // company = Console.ReadLine();
                    ExitApp(ref word, ref miss, ref docs, company);
                    count = 3;
                    if (company == "Private Advertiser")
                        company = ""; // if there is no name we reassign the value with an empty string
                }
                Console.WriteLine("Job Position Title is" + output + "\n");
                Console.WriteLine("Company name is " + company + "\n");

                string addressingEmployer = "Sir/Madam";

                if (count == 3)
                {
                  //  Console.WriteLine("\nEnter F1 to start, then Enter the role title\n");
                    ExitApp(ref word, ref miss, ref docs, output);
                    count = 4;
                }

                string advertiser = "Seek.com.au";
                if (company != "n") advertiser = company; // if response was not no, we will assign with the name in place of <Company> 


                if (count == 4)
                {
                  //  Console.WriteLine("\nIs there a bonus question?, type 'n' if none\n");

                    // bonusQuestion = Console.ReadLine();
                    ExitApp(ref word, ref miss, ref docs, bonusQuestion);
                    count = 5;
                    // if (bonusQuestion == "n") 
                    bonusQuestion = "";
                }

                miss = PopulateWordVariables(word, miss, docs, hiringManagersName, relevantSkills, company, bonusQuestion, output, date, addressingEmployer, advertiser);

                // TODO:possibly put this all on the top except saveas so we can use it when selecting the top intro part to acustom to job type

                switch (output)
                {
                    case "admin":
                        outputFI = @"D:\Troydon\Documents\JobStuff\JobHunt\cover letters\Customs\GeneralAdminCover.docx";
                        break;
                    case "itadmin":
                        outputFI = @"D:\Troydon\Documents\JobStuff\JobHunt\cover letters\Customs\ItAdminCover.docx";
                        break;
                    case "dev":
                        outputFI = @"D:\Troydon\Documents\JobStuff\JobHunt\cover letters\Customs\SoftwareDeveloperCover.docx";
                        break;
                    case "support":
                        outputFI = @"D:\Troydon\Documents\JobStuff\JobHunt\cover letters\Customs\ITSupportCover.docx";
                        break;
                    case "sales":
                        outputFI = @"D:\Troydon\Documents\JobStuff\JobHunt\cover letters\Customs\ITSalesCover.docx";
                        break;
                    default:
                        outputFI = @"D:\Troydon\Documents\JobStuff\JobHunt\cover letters\Customs\SoftwareDeveloperCover.docx";
                        break;
                }
                // keep looping until either the counter is not one of the readline values, and while the elected output is valid
            } while (outputMatch && (count == 0 || count == 1 || count == 2 || count == 3 || count == 4));

            Console.WriteLine("Cover Letter Generated, file saved to " + outputFI + "\n");
            return outputMatch;
        }

        static bool CheckIfJobTitleOfInterest(string jobTitle, ref string relevantSkills, out int? count)
        {
            bool isOfInterest = true;

            count = 0;

            foreach (string title in Constants.developerJobTitles)
            {
                if (jobTitle.Contains(title))
                {
                    // if contains a developer for e.g and contains a title called senior without junior, then we dont want to apply
                    // incase the job title says both junior and senior positions available
                    if (jobTitle.Contains("Senior") && !jobTitle.Contains("Junior")) return false;
                    
                    // otherwise proceed to say we can apply
                    relevantSkills = "I have relevant skills useful to the role, which I obtained when I completed a Certificate IV in programming, that include basic C# OOP, ASP.NET Core, HTML, CSS, JavaScript, jQuery, SQL DBMS, Unity3D, Windows 10 OS, Microsoft Office 365 including Word and Excel, basic understanding of cloud services such as Azure and AWS. I am now applying these technologies into building my knowledge and skill level in creating beautiful, user friendly software, web and game applications. I am also working on using the programming, scripting and mark-up languages to learn the programming concepts such as algorithms, data structures, writing user-centric functional specifications, writing scalable code, understanding conditional logic, database design, responsive design.";
                    count = 1;
                    return true;
                }
            }
            // set description based on output's keywords
            if (jobTitle == "admin")
            {
                Console.WriteLine("Success");
                relevantSkills = "I have three years’ work experience working in a call centre as a Customer Service Rep with Woolworths Mobile. Some of the duties of the role included managing SAP tickets, whereby I had to administer events such as network incidents, transport and logistics, customer complaints, and input relevant data such as customer information, device details, funds, network incidents and faults. I was also known by the company to perform well with these duties.";
                count = 1;
            }
            else if (jobTitle == "itadmin")
            {
                relevantSkills = "I have relevant skills useful to the role, I have achieved this when I completed a Certificate IV in proramming, this includes RDBMS such as SQL Server, MySQL, SQL, T-SQL. I have three years’ work experience working in a call centre as a Customer Service Rep with Woolworths Mobile. Some of the duties of the role included managing SAP tickets, whereby I had to administer events such as network incidents, transport and logistics, customer complaints, and input relevant data such as customer information, device details, funds, network incidents and faults. I was also known by the company to perform well with these duties.";
                count = 1;
            }
            else if (jobTitle == "support")
            {
                relevantSkills = "I am interested in this position; I believe I have the skills and enthusiasm needed to do well. I am looking to secure a role that involves working with technology. I enjoy working with technology and dealing with computers, and Windows OS. I have knowledge in Office 364 Suite and great troubleshooting skills, which I have gained when working with Woolworths Mobile as a Tech Support Representative for mobile devices including devices such as Android, iOS and OPPO.";
                count = 1;
            }
            else if (jobTitle == "sales")
            {
                relevantSkills = "I am interested in this position; I believe I have the skills and enthusiasm needed to do well. I am looking to secure a role that involves working with technology. I enjoy working with technology and dealing with computers, and Windows OS. I have knowledge in Office 364 Suite and great sales skills, which I have gained when working with Woolworths Mobile as a Tech Support Representative for mobile devices including devices such as Android, and OPPO.";
                count = 1;
            }
            else isOfInterest = false;

            return isOfInterest;
        }

        private static object PopulateWordVariables(Application word, object miss, Document docs, string hiringManagersName, string relevantSkills, string company, string bonusQuestion, string output, string date, string addressingEmployer, string advertiser)
        {
            // Replace text now, loop through this 8 times replacing all the needed text
            for (int i = 0; i < 8; i++)
            {
                Find contentReplace = word.Selection.Find;
                contentReplace.ClearFormatting(); // check if this is hitting
                Range rng = docs.Content;

                // based on the iteration number we set what the text that is to be replaced will be
                switch (i)
                {
                    case 0:
                        contentReplace.Text = "<dd Month YYYY>";
                        break;
                    case 1:
                        contentReplace.Text = "<Hiring manager’s name>";
                        break;
                    case 2:
                        contentReplace.Text = "<Company>";
                        break;
                    case 3:
                        contentReplace.Text = "Sir/Madam";
                        break;
                    case 4:
                        contentReplace.Text = "<insert role title>";
                        break;
                    case 5:
                        contentReplace.Text = "Seek.com.au";
                        break;
                    case 6:
                        contentReplace.Text = "<bonus question>";
                        break;
                    case 7:
                        contentReplace.Text = "<insert relevant skills intro here>"; // might be skipping because it is to enter before other answers? if it does not work then put both out of loop
                        break;
                    default:
                        break;
                }

                contentReplace.Replacement.ClearFormatting();

                // depending on the text we are replacing we assign a different value for the replacement
                switch (contentReplace.Text)
                {
                    case "<dd Month YYYY>":
                        contentReplace.Replacement.Text = date;
                        break;
                    case "<Hiring manager’s name>":
                        contentReplace.Replacement.Text = hiringManagersName;
                        break;
                    case "<Company>":
                        contentReplace.Replacement.Text = company;
                        break;
                    case "Sir/Madam":
                        contentReplace.Replacement.Text = addressingEmployer;
                        break;
                    case "<insert role title>":
                        contentReplace.Replacement.Text = output; // TODO: check if role title will be correct after change
                        break;
                    case "Seek.com.au":
                        contentReplace.Replacement.Text = advertiser;
                        break;
                    case "<bonus question>":
                        contentReplace.Replacement.Text = bonusQuestion;
                        break;
                    case "<insert relevant skills intro here>":
                        // if there is more than 250 characters in the parameter the replacement.text will not work so we implement copy and replace
                        if (relevantSkills.Length >= 250)
                        {
                            contentReplace.Replacement.Text = "^c"; // copy to clipboard action is assigned to replacement
                            contentReplace.Replacement.ClearFormatting();
                            // now we search the whole document for this again
                            rng.Find.Execute("<insert relevant skills intro here>", ref miss, ref miss, ref miss, ref miss, ref miss,
                                ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss); // find this text
                                                                                                                           // now replace the text in the document
                            rng.Text = relevantSkills;
                        }
                        break;
                    default:
                        break;
                }

                // now execute the replace for all values
                object replaceAll = WdReplace.wdReplaceAll;
                contentReplace.Execute(ref miss, ref miss, ref miss, ref miss, ref miss,
                    ref miss, ref miss, ref miss, ref miss, ref miss,
                    ref replaceAll, ref miss, ref miss, ref miss, ref miss);
            }

            docs.Content.Font.Color = WdColor.wdColorBlack;
            return miss;
        }

        private static void ExitApp(ref Application word, ref object miss, ref Document docs, string output)
        {
            if (output == "exitapp")
            {
                Console.WriteLine("Application Shutting down");
                CloseWordApplication(ref word, ref miss, ref docs);
                Environment.Exit(0);
            }
        }
        private static void CloseWordApplication(ref Application word, ref object miss, ref Document docs)
        {
            if (docs != null)
            {
                docs.Close(
                    /* ref object SaveChanges */ ref miss,
                    /* ref object OriginalFormat */ ref miss,
                    /* ref object RouteDocument */ ref miss);
                docs = null;
            }

            if (word != null)
            {
                word.Quit(
                    /* ref object SaveChanges */ ref miss,
                    /* ref object OriginalFormat */ ref miss,
                    /* ref object RouteDocument */ ref miss);
                word = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void ExitApplication(ref Application word, ref object miss, ref Document docs, string appCommand)
        {
            if (appCommand == "exitapp")
            {
                Console.WriteLine("Application Shutting down");
                CloseWordApplication(ref word, ref miss, ref docs);
                Environment.Exit(0);
            }
        }
    }
}
