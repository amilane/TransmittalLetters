using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//collect info about pdf files
// name, short name, revision, language, formats, countSheets, native
namespace TransmitLetter
{
  class GetFilesInfo
  {
    public List<List<string>> getFiles(string Path)
    {
      string[] files;
      List<List<string>> filesInfo = new List<List<string>>();

      files = Directory.GetFiles(Path, "*", SearchOption.AllDirectories);
      foreach (string f in files)
      {
        if (f.Contains(".pdf"))
        {
          
          string pdfName = f;
          string[] splitPdfName = pdfName.Split('_');

          string shortName = splitPdfName[0];
          string rev = splitPdfName[1];
          string lang = splitPdfName[2].Split('.')[0];

          string native = files.FirstOrDefault(x => x.Contains(shortName) && !x.Contains(".pdf"));

          CountSheetsFormatsPdf csf = new CountSheetsFormatsPdf();
          List<string> pdfCountAndFormat = csf.countSheetsFormatsPdf(f);
          string countSheets = pdfCountAndFormat[0];
          string format = pdfCountAndFormat[1];

          List<string> fileInfo = new List<string>{
            pdfName,
            shortName,
            rev,
            lang,
            format,
            countSheets,
            native };

          filesInfo.Add(fileInfo);
        }
      }
      return filesInfo;
    }







  }
}
