using System;
using System.Collections.Generic;
using System.Linq;


namespace TransmitLetter
{
  class DataCrs
  {
    public List<List<string>> dataCrs(List<FileInfo> filesInfo)
    {
      List<List<string>> dataCrs = new List<List<string>>();
      WritterReader writterReader = new WritterReader();

      //путь к VRD
      string pathToVDR = new GetPathsToTemplates().getPathsToTemplates()[0];

      List<List<string>> vdrData = writterReader.Read(pathToVDR, "VDR");

      foreach (FileInfo fn in filesInfo)
      {
        //string excelDocName = Regex.Replace(f[0], ".pdf", "_CRS.xlsx", RegexOptions.IgnoreCase);
        string excelDocName = String.Format("{0}_{1}_{2}_CRS.xlsx", fn.shortName, fn.rev, fn.lang);
        List<string> vdrDataRow = vdrData.FirstOrDefault(x => x[28].Contains(fn.shortName));

        List<string> crsRow = new List<string>();
        if (vdrDataRow != null)
        {
          string docTitle = String.Format("{0} / {1}", vdrDataRow[30], vdrDataRow[29]);
          crsRow.Add(excelDocName); // excel name
          crsRow.Add(vdrDataRow[33]);// status
          crsRow.Add(fn.gcDocN);// doc No
          crsRow.Add(vdrDataRow[15]);// doc class
          crsRow.Add(docTitle);// doc title RU/EN
          crsRow.Add(fn.rev);// rev
        }
        else
        {
          crsRow.Add(excelDocName); // excel name
          crsRow.Add(null);// status
          crsRow.Add(fn.gcDocN);// doc No
          crsRow.Add(null);// doc class
          crsRow.Add(null);// doc title RU/EN
          crsRow.Add(fn.rev);// rev
        }
        dataCrs.Add(crsRow);
      }
      return dataCrs;
    }
  }
}
