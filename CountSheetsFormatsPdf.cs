using System;
using System.Collections.Generic;
using System.Linq;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Word;

// get sheets formats and sheets count from PDF
namespace TransmitLetter
{
    class CountSheetsFormatsPdf
    {
        public void countSheetsFormatsPdf(string Path, FileInfo fn)
        {
            PdfArray mediabox;
            int urx, ury;
            PdfReader reader = new PdfReader(Path);
            int numberOfPages = reader.NumberOfPages;

            List<string> formats = new List<string>();
            for (int i = 1; i <= numberOfPages; i++) {
                PdfDictionary pageDict = reader.GetPageN(i);
                mediabox = pageDict.GetAsArray(PdfName.MEDIABOX);
                urx = mediabox.GetAsNumber(2).IntValue;
                ury = mediabox.GetAsNumber(3).IntValue;

                int[] sides = new int[] { urx, ury };
                int shortSide = sides.Min();
                int longSide = sides.Max();

                if (shortSide < 600 && longSide < 900)
                    formats.Add("A4");
                else if (shortSide < 900 && longSide < 1200)
                    formats.Add("A3");
                else if (shortSide < 1200 && longSide < 1700)
                    formats.Add("A2");
                else if (shortSide < 1700 && longSide < 2400)
                    formats.Add("A1");
                else if (shortSide < 2400 && longSide < 3400)
                    formats.Add("A0");
            }
            var fUniq = formats.Distinct();
            string format = String.Join("/", fUniq);


            fn.countSheets = numberOfPages.ToString();
            fn.formatPages = format;

        }

        public void countSheetsFormatsWord(string Path, FileInfo fn)
        {
            var application = new Application();
            var document = application.Documents.Open(Path);

            var numberOfPages = document.ComputeStatistics(WdStatistic.wdStatisticPages, false);

            fn.countSheets = numberOfPages.ToString();
            fn.formatPages = "A4";

          
            application.Quit(false);
        }






    }
}
