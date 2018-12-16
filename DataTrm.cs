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

      WritterReader writterReader = new WritterReader();
      List<List<string>> vdrData = writterReader.Read(@"D:\#Projects\AGPZ\TRANSMITTAL VDR.xlsx", "VDR");

      foreach (List<string> f in filesInfo)
      {
        string shortName = f[1];
        List<string> vdrDataRow = vdrData.FirstOrDefault(x => x[27] == shortName);

      }

    }
  }
}
