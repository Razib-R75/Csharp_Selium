using System;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.Extensions;
using OfficeOpenXml;
using System.Linq;

namespace ExcelAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define the path to your Excel file
            string excelFilePath = @"D:\unity\SpaceR\Python_selium\Excel.xlsx";

            // Define a function to get Google search suggestions for a given keyword
            static (string, string) GetGoogleSuggestions(IWebDriver driver, string keyword)
            {
                driver.Navigate().GoToUrl("https://www.google.com");
                var searchBox = driver.FindElement(By.Name("q"));
                searchBox.Clear();
                searchBox.SendKeys(keyword);

                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//ul[@role='listbox']/li[@role='presentation']")));

                var suggestions = driver.FindElements(By.XPath("//ul[@role='listbox']/li[@role='presentation']"));
                var suggestionTexts = suggestions.Select(suggestion => suggestion.Text).ToList();

                if (suggestionTexts.Any())
                {
                    var longestSuggestion = suggestionTexts.OrderByDescending(s => s.Length).First();
                    var shortestSuggestion = suggestionTexts.OrderBy(s => s.Length).First();
                    return (longestSuggestion, shortestSuggestion);
                }
                else
                {
                    return (null, null);
                }
            }

            var excelFileInfo = new FileInfo(excelFilePath);

            using (var package = new ExcelPackage(excelFileInfo))
            {
                // Get the current day
                var currentDay = DateTime.Now.ToString("dddd");

                if (package.Workbook.Worksheets.Any(ws => ws.Name == currentDay))
                {
                    var worksheet = package.Workbook.Worksheets[currentDay];

                    var longestSuggestions = new List<string>();
                    var shortestSuggestions = new List<string>();

                    // Iterate through the rows and get suggestions
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        string keyword = worksheet.Cells[row, 2].GetValue<string>();
                        if (!string.IsNullOrWhiteSpace(keyword))
                        {
                            using (var driver = new FirefoxDriver())
                            {
                                var (longSuggestion, shortSuggestion) = GetGoogleSuggestions(driver, keyword);
                                longestSuggestions.Add(longSuggestion);
                                shortestSuggestions.Add(shortSuggestion);
                            }
                        }
                    }

                    // Populate the Excel worksheet with suggestions
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        worksheet.Cells[row, 3].Value = longestSuggestions[row - 2];
                        worksheet.Cells[row, 4].Value = shortestSuggestions[row - 2];
                    }

                    // Save the updated Excel file
                    package.Save();
                }
            }
        }
    }
}
