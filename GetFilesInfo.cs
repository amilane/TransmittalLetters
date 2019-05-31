using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using iTextSharp.text.pdf;

//collect info about pdf files
// name, short name, revision, language, formats, countSheets, native
namespace TransmitLetter
{
  public class FileInfo
  {
    public string electronicFilename = null;
    public string pdfName = null;
    public string nativeName = null;
    public string shortName = null;
    public string rev = null;
    public string lang = null;
    public string countSheets = null;
    public string formatPages = null;
    public string gcDocN = null;
    public string docNameForCSV = null; //4.1.3.20.101
    public string subPhase = null;
    public string nativeStatus = null;
    public string fileExtension = null;
    public string section = null; // AOV
  }
  class GetFilesInfo
  {

    public List<FileInfo> getFiles(string Path)
    {
      List<FileInfo> filesInfo = new List<FileInfo>();

      var files = Directory.GetFiles(Path, "*", SearchOption.AllDirectories).Where(i => !i.Contains("Thumbs"));

      //Проверка правильности наименования 0055-CPC-ARM-4.1.3.20.101-AOV-ZH-0001_A1_ER.pdf
      //Regex rgxName = new Regex(@"^\d+-\w+-\w+-\d.\d.\d.\d+.\d+-\w+-\w+-\d+_\w\d_\w+.\w+$");
      Regex rgxName = new Regex(@"^\d+(-\w+){2}-\d[.]\d[.]\d[.]\d+[.][A-Z0-9]+(-[A-Z0-9]+){2,3}-\d+_\w+_[A-Z]+[.][0-9A-Za-z]+");

      string AlertMsg = "Пример правильного наименования файлов: 0055-CPC-ARM-4.1.3.20.101-AOV-ZH-0001_A1_ER.pdf, 0055-CPC-ARM-4.1.3.20.101-AOV-ZH-A-0001_A1_ER.pdf\n\n";
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

        // имена файлов
        List<string> fileNames0 = new List<string>();
        foreach (var f in files)
        {
          fileNames0.Add(f.Split('\\').Last());
        }

        //Отправляем имена файлов CKL после их PDF
        var cklFiles = fileNames0.Where(i => i.Contains("_CKL"));
        var fileNames = fileNames0.Where(i => !i.Contains("_CKL")).ToList();

        if (cklFiles.Count() > 0)
        {
          foreach (var c in cklFiles)
          {
            var ind = fileNames.FindLastIndex(i => i.Contains(c.Split('_').First()));
            fileNames.Insert(ind + 1, c);
          }
        }






        // группируем пдфки с нативами по имени

        var groupFiles = fileNames.GroupBy(i => i.Substring(0, i.LastIndexOf('.')));


        foreach (var group in groupFiles)
        {

          FileInfo fn = new FileInfo();

          var splitKey = group.Key.Split('_');
          fn.shortName = splitKey[0];
          fn.rev = splitKey[1];
          fn.lang = splitKey[2].Split('.')[0];
          fn.gcDocN = String.Format("{0}-{1}-{2}-{3}-{4}.{5}-{6}", splitKey[0].Split('-'));
          fn.docNameForCSV = splitKey[0].Split('-')[3]; //4.1.3.20.101
          fn.section = splitKey[0].Split('-')[4]; // AOV
          fn.subPhase = splitKey[0].Split('-')[3].Substring(0, 3);

          string pathToFile;
          //берем ПДФ, если она есть
          var getPdf = group.Where(i => Regex.IsMatch(i, Regex.Escape(".pdf"), RegexOptions.IgnoreCase));
          if (getPdf.Any())
          {
            fn.pdfName = Regex.Replace(getPdf.First(), ".pdf", ".pdf", RegexOptions.IgnoreCase);
            pathToFile = files.Where(i => i.Contains(getPdf.First())).First();
            CountSheetsFormatsPdf csf = new CountSheetsFormatsPdf();
            csf.countSheetsFormatsPdf(pathToFile, fn);

          }

          //берем doc
          var getDoc = group.Where(i => Regex.IsMatch(i, Regex.Escape(".doc"), RegexOptions.IgnoreCase));
          if (getDoc.Any())
          {
            pathToFile = files.Where(i => i.Contains(getDoc.First())).First();
            CountSheetsFormatsPdf csf = new CountSheetsFormatsPdf();
            csf.countSheetsFormatsWord(pathToFile, fn);

          }


          //берем натив
          var getNative = group.Where(i => !Regex.IsMatch(i, Regex.Escape(".pdf"), RegexOptions.IgnoreCase));
          if (getNative.Any())
          {
            fn.nativeName = getNative.First();
            fn.nativeStatus = "Native Format";
          }

          if (fn.pdfName != null)
          {
            fn.electronicFilename = fn.pdfName;
            fn.fileExtension = "pdf";
          }
          else
          {
            fn.electronicFilename = fn.nativeName;
            fn.fileExtension = fn.nativeName.Substring(fn.nativeName.LastIndexOf('.') + 1); ;
          }


          filesInfo.Add(fn);

        }

      }
      return filesInfo;
    }

  }
}
