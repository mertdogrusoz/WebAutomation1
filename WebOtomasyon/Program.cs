using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace NewsScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var options = new ChromeOptions();
            options.AddArgument("--ignore-certificate-errors"); 
            options.AddArgument("--allow-insecure-localhost"); 


            var driver = new ChromeDriver(options);
            List<String> Titlelist = new List<String>();
            List<String> DescriptionList = new List<String>();

            try
            {
                driver.Navigate().GoToUrl("https://www.hurriyet.com.tr/gundem/");
                driver.Manage().Window.Maximize();

                int page = 1;

                while (page < 5) 
                {
                    var newslink = driver.FindElements(By.CssSelector(".category__list__item"));

                    foreach (var item in newslink)
                    {
                        var NewsTitle = item.FindElement(By.CssSelector("h2"));
                        var NewsDescription = item.FindElement(By.CssSelector("p"));
                        Titlelist.Add(NewsTitle.Text);
                        DescriptionList.Add(NewsDescription.Text);
                    }
                    try
                    {
                        page++; 
                        var nextPageButton = driver.FindElement(By.CssSelector($".paging__btn[data-page-index='{page}']"));
                        nextPageButton.Click();
                        System.Threading.Thread.Sleep(2000);
                       
                    }
                    catch (NoSuchElementException)
                    {
                        Console.WriteLine("Max Haber sayısına ulaşıldı");


                        break;
                       
                    }
                }

                var fileInfo = new FileInfo(@"C:\Users\yigit\Documents\NewsArticlesDeneme2.xlsx");

                using (var package = new ExcelPackage(fileInfo))
                {
                    var workSheet = package.Workbook.Worksheets.Add("HaberlerDeneme2");

                    workSheet.Cells[1, 1].Value = "Başlık";
                    workSheet.Cells[1, 2].Value = "Açıklama";

                    for (int i = 0; i < Titlelist.Count; i++)
                    {
                        workSheet.Cells[i + 2, 1].Value = Titlelist[i];
                        workSheet.Cells[i + 2, 2].Value = DescriptionList[i];
                    }

                    package.Save();
                }

                Console.WriteLine("Haber başlıkları ve açıklamaları Excel dosyasına kaydedildi.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Bir hata oluştu: {ex.Message}");
            }
            finally
            {
                driver.Quit();
            }
        }
    }
}
