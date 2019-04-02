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
      Regex rgxName = new Regex(@"^\d+-\w+-\w+-\d.\d.\d.\d+.\d+-\w+-\w+-\d+_\w\d_\w+.\w+$");

      string AlertMsg = "Пример правильного наименования файлов: 0055-CPC-ARM-4.1.3.20.101-AOV-ZH-0001_A1_ER.pdf\n\n";
      string msg = "";
      foreach (var f in files) {
        string fName = f.Split('\\').Last();
        if (!rgxName.IsMatch(fName) & fName != "Thumbs.db") {
          msg += fName + '\n';
        }
      }

      if (msg != "") {
        AlertMsg += msg;
        MessageBox.Show(AlertMsg, "Предупреждение");
        Environment.Exit(0);
      } else {

        // имена файлов
        List<string> fileNames = new List<string>();
        foreach (var f in files) {
          fileNames.Add(f.Split('\\').Last());
        }

        // группируем пдфки с нативами по имени

        var groupFiles = fileNames.GroupBy(i => i.Substring(0, i.LastIndexOf('.')));


        foreach (var group in groupFiles) {

          FileInfo fn = new FileInfo();

          var splitKey = group.Key.Split('_');
          fn.shortName = splitKey[0];
          fn.rev = splitKey[1];
          fn.lang = splitKey[2].Split('.')[0];
          fn.gcDocN = String.Format("{0}-{1}-{2}-{3}-{4}.{5}-{6}", splitKey[0].Split('-'));
          fn.docNameForCSV = splitKey[0].Split('-')[3]; //4.1.3.20.101
          fn.section = splitKey[0].Split('-')[4]; // AOV
          fn.subPhase = splitKey[0].Split('-')[3].Substring(0, 3);

          //берем ПДФ, если она есть
          var getPdf = group.Where(i => Regex.IsMatch(i, Regex.Escape(".pdf"), RegexOptions.IgnoreCase));
          if (getPdf.Any()) {
            fn.pdfName = Regex.Replace(getPdf.First(), ".pdf", ".pdf", RegexOptions.IgnoreCase);
            string pathToPDF = files.Where(i => i.Contains(getPdf.First())).First();
            CountSheetsFormatsPdf csf = new CountSheetsFormatsPdf();
            List<string> pdfCountAndFormat = csf.countSheetsFormatsPdf(pathToPDF);
            fn.countSheets = pdfCountAndFormat[0];
            fn.formatPages = pdfCountAndFormat[1];
          }

          //берем натив
          var getNative = group.Where(i => !Regex.IsMatch(i, Regex.Escape(".pdf"), RegexOptions.IgnoreCase));
          if (getNative.Any()) {
            fn.nativeName = getNative.First();
            fn.nativeStatus = "Native Format";
          }

          if (fn.pdfName != null) {
            fn.electronicFilename = fn.pdfName;
            fn.fileExtension = "pdf";
          } else {
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
