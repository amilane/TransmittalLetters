using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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

          pdfName = f.Split('\\').Last();
          splitPdfName = pdfName.Split('_');
          shortName = splitPdfName[0];
          rev = splitPdfName[1];
          lang = splitPdfName[2].Split('.')[0];
          gcDocN = String.Format("{0}-{1}-{2}-{3}-{4}.{5}-{6}", shortName.Split('-'));

          _native = files.FirstOrDefault(x => x.Contains(shortName) && !x.Contains(".pdf"));
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
      return filesInfo;
    }







  }
}
