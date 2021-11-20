
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;

namespace AddIndexToWordDoc
{
    class Program
    {
        static IConfigurationRoot config;
        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Path.Combine(AppContext.BaseDirectory))
                .AddJsonFile("appsettings.json", optional: true);

            config = builder.Build();

            Console.WriteLine("Hi Liran!");
            Console.WriteLine("Please insert the file name of the report, and press enter");
            string fileName = Console.ReadLine();         
            
            Console.WriteLine("Please insert the number of the investors, and press enter");
            int investorsNum = int.Parse(Console.ReadLine());
            
            Console.WriteLine("How much bookmarks insert to the doc file?");
            int bookmarkCount = int.Parse(Console.ReadLine());
            
            List<string> bookmarksName = new List<string>();
            
            Console.WriteLine("Please enter all the bookmark names you added to the report, split them by enter");
            for(int i=0; i< bookmarkCount; i++)
            {
                bookmarksName.Add(Console.ReadLine());
            }

            string reportsPath = GetValueFromConfiguration("ReportFolder");
            string outputFolderPdf = GetValueFromConfiguration("OutputFolderPdf");
            WordFileHandler wordFileHandler = new (reportsPath, outputFolderPdf);
            foreach(int i in wordFileHandler.OpenWordFile(investorsNum, fileName, bookmarksName))
            {
                Console.WriteLine($"Finish to save pdf {i}");
            }
            Console.WriteLine("Finish all the pdf! press any key to close");
            Console.ReadLine();
        }

        static string GetValueFromConfiguration(string key)
        {
            return config.GetSection(key).Value;
        }

    }
}
