using System.Collections.Generic;
using System.Linq;


namespace TransmitLetter
{
  class DataTrmCsv
  {
    private readonly Dictionary<string, string> statusDict = DataDictionaries.status;
    private readonly Dictionary<string, string> docTypesDictRu = DataDictionaries.documentTypeRus;
    private readonly Dictionary<string, string> docTypesDictEn = DataDictionaries.documentTypeEn;
    private readonly Dictionary<string, string> sectionDictRu = DataDictionaries.sectionRu;
    private readonly Dictionary<string, string> sectionDictEn = DataDictionaries.sectionEn;



    public List<List<string>> dataTrmCsv(List<FileInfo> filesInfo, string transmitNumber)
    {
      //путь к VDR
      string pathToVDR = new GetPathsToTemplates().getPathsToTemplates()[0];

      List<List<string>> dataTrmCsv = new List<List<string>>();
      WritterReader writterReader = new WritterReader();
      List<List<string>> vdrCsvData =
        writterReader.Read(pathToVDR, "VDR CSV");
      List<List<string>> vdrData = writterReader.Read(pathToVDR, "VDR");

      foreach (FileInfo fn in filesInfo) {
        string docTypeCode = fn.shortName.Split('-')[5];
        string docTypeCodeRu = docTypesDictRu[docTypeCode];
        string docTypeCodeEn = docTypesDictEn[docTypeCode];

        List<string> vdrCsvDataRow = vdrCsvData.FirstOrDefault(x => x[0].Contains(fn.docNameForCSV));
        List<string> vdrDataRow = vdrData.FirstOrDefault(x => x[28].Contains(fn.shortName));

        List<string> trmCsvRow = new List<string>();
        if (vdrCsvDataRow != null) {
          trmCsvRow.Add(null); //Row Status
          trmCsvRow.Add(fn.gcDocN); //doc No
          trmCsvRow.Add(null); //Vendor doc no
          trmCsvRow.Add(vdrDataRow[30]); // doc title RU
          trmCsvRow.Add(vdrDataRow[29]); // doc title EN
          trmCsvRow.Add(docTypeCode); // doc type code
          trmCsvRow.Add(docTypeCodeRu); // doc type code RU
          trmCsvRow.Add(docTypeCodeEn); // doc type code EN
          trmCsvRow.Add(vdrDataRow[40]); // doc date
          trmCsvRow.Add(fn.rev); // rev
          trmCsvRow.Add(statusDict[vdrDataRow[33]]); // status
          trmCsvRow.Add("CPECC"); // originator
          trmCsvRow.Add("0055"); // contract number
          trmCsvRow.Add("Joint-Stock Company\"ARMO - GROUP\""); //vendor
          trmCsvRow.Add(vdrDataRow[35]); // PO NO
          trmCsvRow.Add("4 - ГПЗ"); // work phase
          trmCsvRow.Add(fn.subPhase); // work sub-phase
          trmCsvRow.Add(vdrCsvDataRow[1]); // subproject code
          trmCsvRow.Add(vdrCsvDataRow[2]); // operation complex RU
          trmCsvRow.Add(vdrCsvDataRow[3]); // operation complex EN
          trmCsvRow.Add(vdrCsvDataRow[4]); // start complex RU
          trmCsvRow.Add(vdrCsvDataRow[5]); // start complex EN
          trmCsvRow.Add(vdrCsvDataRow[6]); // title object RU
          trmCsvRow.Add(vdrCsvDataRow[7]); // title object EN
          trmCsvRow.Add(vdrCsvDataRow[8]); // title number RU
          trmCsvRow.Add(vdrCsvDataRow[9]); // title number EN
          trmCsvRow.Add(sectionDictRu[fn.section]); // package name RU
          trmCsvRow.Add(sectionDictEn[fn.section]); // package name EN
          trmCsvRow.Add(null); //sequence number
          trmCsvRow.Add(vdrCsvDataRow[13]); // doc class code
          trmCsvRow.Add(vdrCsvDataRow[14]); // discip RU
          trmCsvRow.Add(vdrCsvDataRow[15]); // discip EN
          trmCsvRow.Add(null); //construction contractor
          trmCsvRow.Add(vdrCsvDataRow[17]); // construction type
          trmCsvRow.Add(fn.countSheets); // count of sheets
          trmCsvRow.Add("1"); //sheet number
          trmCsvRow.Add(null); //number of copies
          trmCsvRow.Add(null); //ACRS
          trmCsvRow.Add(null); //status ACRS
          trmCsvRow.Add(null); //PEM
          trmCsvRow.Add(null); //equipment tag
          trmCsvRow.Add(fn.lang); // language
          trmCsvRow.Add(fn.formatPages); // format
          trmCsvRow.Add(null); //review code
          trmCsvRow.Add(transmitNumber); // incomming transmittal N
          trmCsvRow.Add(null); //old transmittal N
          trmCsvRow.Add("Latest"); // doc status
          trmCsvRow.Add("00.Holding Folder"); //storage path
          trmCsvRow.Add(fn.electronicFilename); //content
          trmCsvRow.Add(fn.nativeStatus); //rendition name
          trmCsvRow.Add(fn.nativeName); //rendition file
          trmCsvRow.Add(null); //link to
          trmCsvRow.Add(null); //stamped for construction
        } else {
          trmCsvRow.Add(null); //Row Status
          trmCsvRow.Add(fn.gcDocN); //doc No
          trmCsvRow.Add(null); //Vendor doc no
          trmCsvRow.Add(null); // doc title RU
          trmCsvRow.Add(null); // doc title EN
          trmCsvRow.Add(docTypeCode); // doc type code
          trmCsvRow.Add(docTypeCodeRu); // doc type code RU
          trmCsvRow.Add(docTypeCodeEn); // doc type code EN
          trmCsvRow.Add(null); // doc date
          trmCsvRow.Add(fn.rev); // rev
          trmCsvRow.Add(null); // status
          trmCsvRow.Add("CPECC"); // originator
          trmCsvRow.Add("0055"); // contract number
          trmCsvRow.Add("Joint-Stock Company\"ARMO - GROUP\""); //vendor
          trmCsvRow.Add(null); // PO NO
          trmCsvRow.Add("4 - ГПЗ"); // work phase
          trmCsvRow.Add(fn.subPhase); // work sub-phase
          trmCsvRow.Add(null); // subproject code
          trmCsvRow.Add(null); // operation complex RU
          trmCsvRow.Add(null); // operation complex EN
          trmCsvRow.Add(null); // start complex RU
          trmCsvRow.Add(null); // start complex EN
          trmCsvRow.Add(null); // title object RU
          trmCsvRow.Add(null); // title object EN
          trmCsvRow.Add(null); // title number RU
          trmCsvRow.Add(null); // title number EN
          trmCsvRow.Add(sectionDictRu[fn.section]); // package name RU
          trmCsvRow.Add(sectionDictEn[fn.section]); // package name EN
          trmCsvRow.Add(null); //sequence number
          trmCsvRow.Add(null); // doc class code
          trmCsvRow.Add(null); // discip RU
          trmCsvRow.Add(null); // discip EN
          trmCsvRow.Add(null); //construction contractor
          trmCsvRow.Add(null); // construction type
          trmCsvRow.Add(fn.countSheets); // count of sheets
          trmCsvRow.Add("1"); //sheet number
          trmCsvRow.Add(null); //number of copies
          trmCsvRow.Add(null); //ACRS
          trmCsvRow.Add(null); //status ACRS
          trmCsvRow.Add(null); //PEM
          trmCsvRow.Add(null); //equipment tag
          trmCsvRow.Add(fn.lang); // language
          trmCsvRow.Add(fn.formatPages); // format
          trmCsvRow.Add(null); //review code
          trmCsvRow.Add(transmitNumber); // incomming transmittal N
          trmCsvRow.Add(null); //old transmittal N
          trmCsvRow.Add("Latest"); // doc status
          trmCsvRow.Add("00.Holding Folder"); //storage path
          trmCsvRow.Add(fn.electronicFilename); //content
          trmCsvRow.Add(fn.nativeStatus); //rendition name
          trmCsvRow.Add(fn.nativeName); //rendition file
          trmCsvRow.Add(null); //link to
          trmCsvRow.Add(null); //stamped for construction
        }

        dataTrmCsv.Add(trmCsvRow);
      }

      return dataTrmCsv;
    }
  }
}
