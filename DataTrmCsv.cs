using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TransmitLetter
{
  class DataTrmCsv
  {
    public Dictionary<string, string> statusDict = new Dictionary<string, string>
    {
      ["AFC"] = "AFC - Утверждено для строительства",
      ["IFA"] = "IFA - Выпущено для согласования",
      ["IFC"] = "IFC - Выпущено для строительства",
      ["IFD"] = "IFD - Выпущено для проектирования",
      ["IFH"] = "IFH – Выпущено для HAZOP",
      ["IFI"] = "IFI - Выпущено для информации",
      ["IFP"] = "IFP - Выпущено для закупки",
      ["IFQ"] = "IFQ - Выпущено для предложения цены",
      ["IFR"] = "IFR - Выпущено для рассмотрения",
      ["IFU"] = "IFU - Выпущено для использования",
      ["SD"] = "SD - Заменен",
      ["VD"] = "VD - Аннулирован",
    };

    public List<List<string>> dataTrmCsv(List<List<string>> filesInfo, string transmitNumber)
    {
      List<List<string>> dataTrmCsv = new List<List<string>>();
      WritterReader writterReader = new WritterReader();
      List<List<string>> vdrCsvData =
        writterReader.Read(@"\\arena\ARMO-GROUP\ОБЪЕКТЫ\В_РАБОТЕ\41XX_AGPZ\30-РД\02-ГИП\TRANSMITTAL VDR.xlsx", "VDR CSV");
      List<List<string>> vdrData = writterReader.Read(@"\\arena\ARMO-GROUP\ОБЪЕКТЫ\В_РАБОТЕ\41XX_AGPZ\30-РД\02-ГИП\TRANSMITTAL VDR.xlsx", "VDR");

      foreach (List<string> f in filesInfo)
      {
        string packageCode = f[1].Split('-')[4];
        string docTypeCode = f[1].Split('-')[5];
        List<string> vdrCsvDataRow = vdrCsvData.FirstOrDefault(x => x[0].Contains(f[1]));
        List<string> vdrDataRow = vdrData.FirstOrDefault(x => x[28].Contains(f[1]));

        List<string> trmCsvRow = new List<string>();
        if (vdrCsvDataRow != null)
        {
          trmCsvRow.Add(null); //Row Status
          trmCsvRow.Add(f[1]); //short name
          trmCsvRow.Add(null); //Vendor doc no
          trmCsvRow.Add(vdrDataRow[30]); // doc title RU
          trmCsvRow.Add(vdrDataRow[29]); // doc title EN
          trmCsvRow.Add(docTypeCode); // doc type code
          trmCsvRow.Add(docTypeCode); // doc type code RU
          trmCsvRow.Add(docTypeCode); // doc type code EN
          trmCsvRow.Add(vdrDataRow[40]); // doc date
          trmCsvRow.Add(f[2]); // rev
          trmCsvRow.Add(statusDict[vdrDataRow[33]]); // status
          trmCsvRow.Add("CRECC"); // originator
          trmCsvRow.Add("0055"); // contract number
          trmCsvRow.Add(null); //vendor
          trmCsvRow.Add(vdrDataRow[35]); // PO NO
          trmCsvRow.Add("4 - ГПЗ"); // work phase
          trmCsvRow.Add("4.1"); // work sub-phase
          trmCsvRow.Add(vdrCsvDataRow[1]); // subproject code
          trmCsvRow.Add(vdrCsvDataRow[2]); // operation complex RU
          trmCsvRow.Add(vdrCsvDataRow[3]); // operation complex EN
          trmCsvRow.Add(vdrCsvDataRow[4]); // start complex RU
          trmCsvRow.Add(vdrCsvDataRow[5]); // start complex EN
          trmCsvRow.Add(vdrCsvDataRow[6]); // title object RU
          trmCsvRow.Add(vdrCsvDataRow[7]); // title object EN
          trmCsvRow.Add(vdrCsvDataRow[8]); // title number RU
          trmCsvRow.Add(vdrCsvDataRow[9]); // title number EN
          trmCsvRow.Add(vdrCsvDataRow[10]); // package name RU
          trmCsvRow.Add(vdrCsvDataRow[11]); // package name EN
          trmCsvRow.Add(null); //sequence number
          trmCsvRow.Add(vdrCsvDataRow[13]); // doc class code
          trmCsvRow.Add(vdrCsvDataRow[14]); // discip RU
          trmCsvRow.Add(vdrCsvDataRow[15]); // discip EN
          trmCsvRow.Add(null); //construction contractor
          trmCsvRow.Add(vdrCsvDataRow[17]); // construction type
          trmCsvRow.Add(f[5]); // count of sheets
          trmCsvRow.Add("1"); //sheet number
          trmCsvRow.Add(null); //number of copies
          trmCsvRow.Add(null); //ACRS
          trmCsvRow.Add(null); //status ACRS
          trmCsvRow.Add(null); //PEM
          trmCsvRow.Add(null); //equipment tag
          trmCsvRow.Add(f[3]); // language
          trmCsvRow.Add(f[4]); // format
          trmCsvRow.Add(null); //review code
          trmCsvRow.Add(transmitNumber); // incomming transmittal N
          trmCsvRow.Add(null); //old transmittal N
          trmCsvRow.Add("Latest"); // doc status
          trmCsvRow.Add(null); //storage path
          trmCsvRow.Add(f[0]); //content
          trmCsvRow.Add("Native Format"); //rendition name
          trmCsvRow.Add(f[6]); //rendition file
          trmCsvRow.Add(null); //link to
          trmCsvRow.Add(null); //stamped for construction
        }
        else
        {
          trmCsvRow.Add(null); //Row Status
          trmCsvRow.Add(f[1]); //short name
          trmCsvRow.Add(null); //Vendor doc no
          trmCsvRow.Add(null); // doc title RU
          trmCsvRow.Add(null); // doc title EN
          trmCsvRow.Add(docTypeCode); // doc type code
          trmCsvRow.Add(docTypeCode); // doc type code RU
          trmCsvRow.Add(docTypeCode); // doc type code EN
          trmCsvRow.Add(null); // doc date
          trmCsvRow.Add(f[2]); // rev
          trmCsvRow.Add(null); // status
          trmCsvRow.Add("CRECC"); // originator
          trmCsvRow.Add("0055"); // contract number
          trmCsvRow.Add(null); //vendor
          trmCsvRow.Add(null); // PO NO
          trmCsvRow.Add("4 - ГПЗ"); // work phase
          trmCsvRow.Add("4.1"); // work sub-phase
          trmCsvRow.Add(null); // subproject code
          trmCsvRow.Add(null); // operation complex RU
          trmCsvRow.Add(null); // operation complex EN
          trmCsvRow.Add(null); // start complex RU
          trmCsvRow.Add(null); // start complex EN
          trmCsvRow.Add(null); // title object RU
          trmCsvRow.Add(null); // title object EN
          trmCsvRow.Add(null); // title number RU
          trmCsvRow.Add(null); // title number EN
          trmCsvRow.Add(null); // package name RU
          trmCsvRow.Add(null); // package name EN
          trmCsvRow.Add(null); //sequence number
          trmCsvRow.Add(null); // doc class code
          trmCsvRow.Add(null); // discip RU
          trmCsvRow.Add(null); // discip EN
          trmCsvRow.Add(null); //construction contractor
          trmCsvRow.Add(null); // construction type
          trmCsvRow.Add(f[5]); // count of sheets
          trmCsvRow.Add("1"); //sheet number
          trmCsvRow.Add(null); //number of copies
          trmCsvRow.Add(null); //ACRS
          trmCsvRow.Add(null); //status ACRS
          trmCsvRow.Add(null); //PEM
          trmCsvRow.Add(null); //equipment tag
          trmCsvRow.Add(f[3]); // language
          trmCsvRow.Add(f[4]); // format
          trmCsvRow.Add(null); //review code
          trmCsvRow.Add(transmitNumber); // incomming transmittal N
          trmCsvRow.Add(null); //old transmittal N
          trmCsvRow.Add("Latest"); // doc status
          trmCsvRow.Add(null); //storage path
          trmCsvRow.Add(f[0]); //content
          trmCsvRow.Add("Native Format"); //rendition name
          trmCsvRow.Add(f[6]); //rendition file
          trmCsvRow.Add(null); //link to
          trmCsvRow.Add(null); //stamped for construction
        }

        dataTrmCsv.Add(trmCsvRow);
      }

      return dataTrmCsv;
    }
  }
}
