
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Range = Microsoft.Office.Interop.Word.Range;

namespace AddIndexToWordDoc
{
    public class WordFileHandler
    {
        private string _reportsFolder;
        private string _outputFolderPdf;
        public WordFileHandler(string reportsFolder, string outputFolderPdf)
        {
            _reportsFolder = reportsFolder;
            _outputFolderPdf = outputFolderPdf;
        }
        public IEnumerable<int> OpenWordFile(int investorsNum, string fileName, List<string> bookmarksName)
        {
            Application app = new();
            Document doc = app.Documents.Open($"{_reportsFolder}\\{fileName}");
                
            for (int i = 1; i <= investorsNum; i++)
            {

                foreach(var name in bookmarksName)
                {
                    if (doc.Bookmarks.Exists(name))
                    {
                        Bookmark bm = doc.Bookmarks[name];
                        Range range = bm.Range;
                        range.Text = i.ToString();
                        range.Font.SizeBi = 5;
                        range.Font.Size = 5;
                        doc.Bookmarks.Add(name, range);         
                    }
                }          
                string pdfDocName = $"{_outputFolderPdf}\\{i}.pdf";

                doc.ExportAsFixedFormat(pdfDocName, WdExportFormat.wdExportFormatPDF);
                yield return i;
            }

            foreach (var name in bookmarksName)
            {
                if (doc.Bookmarks.Exists(name))
                {
                    Bookmark bm = doc.Bookmarks[name];
                    Range range = bm.Range;
                    range.Text = string.Empty;
                    doc.Bookmarks.Add(name, range);
                }
            }

            doc.Close();
            app.Quit();
        }
    }


}
