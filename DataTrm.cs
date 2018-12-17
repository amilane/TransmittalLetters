using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransmitLetter
{
  class DataTrm
  {
    public List<List<string>> dataTrm(List<List<string>> filesInfo)
    {
      List<List<string>> dataTrm = new List<List<string>>();
      WritterReader writterReader = new WritterReader();
      List<List<string>> vdrData = writterReader.Read(@"\\arena\ARMO-GROUP\ОБЪЕКТЫ\В_РАБОТЕ\41XX_AGPZ\30-РД\02-ГИП\TRANSMITTAL VDR.xlsx", "VDR");

      foreach (List<string> f in filesInfo)
      {
        string packageCode = f[1].Split('-')[4];
        string docTypeCode = f[1].Split('-')[5];
        List<string> vdrDataRow = vdrData.FirstOrDefault(x => x[28].Contains(f[1]));

        List<string> trmRow = new List<string>();
        if (vdrDataRow != null)
        {
          trmRow.Add(vdrDataRow[35]); // PO NO
          trmRow.Add(null);//Tag No
          trmRow.Add(null);//supplier doc number
          trmRow.Add(f[7]);// doc No
          trmRow.Add(f[3]);// language
          trmRow.Add(vdrDataRow[30]);// doc title RU
          trmRow.Add(vdrDataRow[29]);// doc title EN
          trmRow.Add(vdrDataRow[24]);// dics
          trmRow.Add(packageCode);// package Code
          trmRow.Add(vdrDataRow[33]);// status
          trmRow.Add(vdrDataRow[40]);// doc date
          trmRow.Add(vdrDataRow[15]);// doc class
          trmRow.Add(f[2]);// rev
          trmRow.Add(docTypeCode);// doc type code
          trmRow.Add(f[5]);// count of sheets
          trmRow.Add(f[4]);// format
          trmRow.Add("pdf");// file format
          trmRow.Add(f[0]);//filename
        }
        else
        {
          trmRow.Add(null); // PO NO
          trmRow.Add(null);//Tag No
          trmRow.Add(null);//supplier doc number
          trmRow.Add(f[7]);//doc No
          trmRow.Add(f[3]);// language
          trmRow.Add(null);// doc title RU
          trmRow.Add(null);// doc title EN
          trmRow.Add(null);// dics
          trmRow.Add(packageCode);// package Code
          trmRow.Add(null);// status
          trmRow.Add(null);// doc date
          trmRow.Add(null);// doc class
          trmRow.Add(f[2]);// rev
          trmRow.Add(docTypeCode);// doc type code
          trmRow.Add(f[5]);// count of sheets
          trmRow.Add(f[4]);// format
          trmRow.Add("pdf");// file format
          trmRow.Add(f[0]);//filename
        }
        dataTrm.Add(trmRow);
      }
      return dataTrm;
    }
  }
}
