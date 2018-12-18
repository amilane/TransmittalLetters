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
      List<List<string>> vdrData = writterReader.Read(@"\\arena\ARMO-GROUP\ОБЪЕКТЫ\В_РАБОТЕ\41XX_AGPZ\30-РД\02-ГИП\TRANSMITTAL VDR.xlsx", "VDR");

      foreach (List<string> f in filesInfo)
      {
        string excelDocName = f[0].Replace(".pdf", "_CRS.xlsx");
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
