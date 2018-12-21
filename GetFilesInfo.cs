using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

//collect info about pdf files
// name, short name, revision, language, formats, countSheets, native
namespace TransmitLetter
{
  class GetFilesInfo
  {

    //public static bool ContainsString(string source, string toCheck, StringComparison comp)
    //{
    //  return source?.IndexOf(toCheck, comp) >= 0;
    //}


    public List<List<string>> getFiles(string Path)
    {
      string[] files;
      List<List<string>> filesInfo = new List<List<string>>();
      files = Directory.GetFiles(Path, "*", SearchOption.AllDirectories);

      //Проверка правильности наименования 0055-CPC-ARM-4.1.3.20.101-AOV-ZH-0001_A1_ER.pdf
      Regex rgxName = new Regex(@"^\d+-\w+-\w+-\d.\d.\d.\d+.\d+-\w+-\w+-\d+_\w\d_\w+.\w+$");

      string AlertMsg = "Пример правильного наименования файлов: 0055-CPC-ARM-4.1.3.20.101-AOV-ZH-0001_A1_ER.pdf\n\n";
      string msg = "";
      foreach (var f in files)
      {
        string fName = f.Split('\\').Last();
        if (!rgxName.IsMatch(fName) & fName != "Thumbs.db")
        {
          msg += fName + '\n';
        }
      }

      if (msg != "")
      {
        AlertMsg += msg;
        MessageBox.Show(AlertMsg, "Предупреждение");
        Environment.Exit(0);
      }
      else
      {
        foreach (string f in files)
        {
          if (Regex.IsMatch(f, Regex.Escape(".pdf"), RegexOptions.IgnoreCase))
          {
            string _native;
            string native = null;
            string shortName;
            string rev;
            string lang;
            string countSheets;
            string format;
            string pdfName;
            string[] splitPdfName;
            string gcDocN;
            string oldPdfName;

            oldPdfName = f.Split('\\').Last();
            pdfName = Regex.Replace(oldPdfName, ".pdf", ".pdf", RegexOptions.IgnoreCase);
            splitPdfName = pdfName.Split('_');
            shortName = splitPdfName[0];
            rev = splitPdfName[1];
            lang = splitPdfName[2].Split('.')[0];
            gcDocN = String.Format("{0}-{1}-{2}-{3}-{4}.{5}-{6}", shortName.Split('-'));

            _native = files.FirstOrDefault(x => x.Contains(shortName) && !Regex.IsMatch(f, Regex.Escape(".pdf"), RegexOptions.IgnoreCase));
            if (_native != null)
              native = _native.Split('\\').Last();

            CountSheetsFormatsPdf csf = new CountSheetsFormatsPdf();
            List<string> pdfCountAndFormat = csf.countSheetsFormatsPdf(f);
            countSheets = pdfCountAndFormat[0];
            format = pdfCountAndFormat[1];

            List<string> fileInfo = new List<string>{
              pdfName,
              shortName,
              rev,
              lang,
              format,
              countSheets,
              native,
              gcDocN};

            filesInfo.Add(fileInfo);
          }
        }
        
      }
      return filesInfo;
    }
    
  }
}
