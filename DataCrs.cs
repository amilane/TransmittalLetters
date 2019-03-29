using System;
using System.Collections.Generic;
using System.Linq;


namespace TransmitLetter
{
  class DataCrs
  {
    public List<List<string>> dataCrs(List<List<string>> filesInfo)
    {
      List<List<string>> dataCrs = new List<List<string>>();
      WritterReader writterReader = new WritterReader();

      //путь к VRD
      string pathToVDR = new GetPathsToTemplates().getPathsToTemplates()[0];

      List<List<string>> vdrData = writterReader.Read(pathToVDR, "VDR");

      foreach (List<string> f in filesInfo)
      {
        //string excelDocName = Regex.Replace(f[0], ".pdf", "_CRS.xlsx", RegexOptions.IgnoreCase);
        string excelDocName = String.Format("{0}_{1}_{2}_CRS.xlsx", f[1], f[2], f[3]);
        List<string> vdrDataRow = vdrData.FirstOrDefault(x => x[28].Contains(f[1]));

        List<string> crsRow = new List<string>();
        if (vdrDataRow != null)
        {
          string docTitle = String.Format("{0} / {1}", vdrDataRow[30], vdrDataRow[29]);
          crsRow.Add(excelDocName); // excel name
          crsRow.Add(vdrDataRow[33]);// status
          crsRow.Add(f[7]);// doc No
          crsRow.Add(vdrDataRow[15]);// doc class
          crsRow.Add(docTitle);// doc title RU/EN
          crsRow.Add(f[2]);// rev
        }
        else
        {
          crsRow.Add(excelDocName); // excel name
          crsRow.Add(null);// status
          crsRow.Add(f[7]);// doc No
          crsRow.Add(null);// doc class
          crsRow.Add(null);// doc title RU/EN
          crsRow.Add(f[2]);// rev
        }
        dataCrs.Add(crsRow);
      }
      return dataCrs;
    }
  }
}
