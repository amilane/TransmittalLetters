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

    public List<List<string>> getFiles(string Path)
    {
      string[] _files;
      List<List<string>> filesInfo = new List<List<string>>();
      _files = Directory.GetFiles(Path, "*", SearchOption.AllDirectories);

      var files = _files.Where(i => !i.Contains("Thumbs"));

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

        string native = null;
        string shortName = null;
        string rev = null;
        string lang = null;
        string countSheets = null;
        string format = null;
        string pdfName = null;
        string[] splitKey;
        string gcDocN = null;
        string docNameForCSV = null;
        string oldPdfName = null;
        string subPhase = null;
        string nativeStatus = null;

        foreach (var group in groupFiles) {
          splitKey = group.Key.Split('_');

          shortName = splitKey[0];
          rev = splitKey[1];
          lang = splitKey[2].Split('.')[0];
          gcDocN = String.Format("{0}-{1}-{2}-{3}-{4}.{5}-{6}", shortName.Split('-'));
          docNameForCSV = String.Format("{0}-{1}-{2}-{3}-{4}", shortName.Split('-'));
          subPhase = shortName.Split('-')[3].Substring(0, 3);

          //берем ПДФ, если она есть
          var getPdf = group.Where(i => Regex.IsMatch(i, Regex.Escape(".pdf"), RegexOptions.IgnoreCase));
          if (getPdf.Any()) {
            pdfName = Regex.Replace(getPdf.First(), ".pdf", ".pdf", RegexOptions.IgnoreCase);
            
            string pathToPDF = files.Where(i => i.Contains(getPdf.First())).First();
            CountSheetsFormatsPdf csf = new CountSheetsFormatsPdf();
            List<string> pdfCountAndFormat = csf.countSheetsFormatsPdf(pathToPDF);
            countSheets = pdfCountAndFormat[0];
            format = pdfCountAndFormat[1];
          }

          //берем натив
          var getNative = group.Where(i => !Regex.IsMatch(i, Regex.Escape(".pdf"), RegexOptions.IgnoreCase));
          if (getNative.Any()) {
            native = getNative.First();
            nativeStatus = "Native Format";
          }

          List<string> fileInfo = new List<string>{
            pdfName,
            shortName,
            rev,
            lang,
            format,
            countSheets,
            native,
            gcDocN,
            nativeStatus,
            docNameForCSV,
            subPhase
          };

          filesInfo.Add(fileInfo);

        }

        
      }
      return filesInfo;
    }

  }
}
