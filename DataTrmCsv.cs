﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace TransmitLetter
{
  class DataTrmCsv
  {
    private readonly Dictionary<string, string> statusDict = DataDictionaries.status;
    private readonly Dictionary<string, string> docTypesDictRu = DataDictionaries.documentTypeRus;
    private readonly Dictionary<string, string> docTypesDictEn = DataDictionaries.documentTypeEn;


    public List<List<string>> dataTrmCsv(List<FileInfo> filesInfo, string transmitNumber, string _status, string _rev)
    {
      //путь к VDR
      string pathToVDR = new GetPathsToTemplates().getPathsToTemplates()[0];

      List<List<string>> dataTrmCsv = new List<List<string>>();
      WritterReader writterReader = new WritterReader();
      List<List<string>> vdrCsvData =
        writterReader.Read(pathToVDR, "VDR CSV");
      List<List<string>> vdrData = writterReader.Read(pathToVDR, "VDR");

      foreach (FileInfo fn in filesInfo)
      {
        string docTypeCode = fn.shortName.Split('-')[5];


        string docTypeCodeRu = null;
        string docTypeCodeEn = null;

        if (!docTypesDictRu.ContainsKey(docTypeCode) || !docTypesDictEn.ContainsKey(docTypeCode))
        {
          string AlertMsg = String.Format("Неправильное наименование типа документа в имени файла:\n{0} ({1})", fn.electronicFilename, docTypeCode);
          MessageBox.Show(AlertMsg, "Предупреждение");
          Environment.Exit(0);
        }
        else
        {
          docTypeCodeRu = docTypesDictRu[docTypeCode];
          docTypeCodeEn = docTypesDictEn[docTypeCode];
        }

        List<string> vdrCsvDataRow = vdrCsvData.FirstOrDefault(x => x[0].Contains(fn.docNameForCSV));
        List<string> vdrDataRow = vdrData.FirstOrDefault(x => x[28].Contains(fn.shortName));

        //Rev
        string Rev;
        if (_rev != null && _rev.Trim() != "")
        {
          Rev = _rev;
        }
        else
        {
          Rev = fn.rev;
        }

        //Status
        string Status = null;
        if (_status != null && _status.Trim() != "" && statusDict.ContainsKey(_status))
        {
          Status = statusDict[_status];
        }
        else if (vdrDataRow != null)
        {
          Status = statusDict[vdrDataRow[33]];
        }

        string[] sectionDescriptions = DataDictionaries.GetSectionDescriptions(fn.section);
        string sectionRu = sectionDescriptions[0];
        string sectionEn = sectionDescriptions[1];

        string docNumber = null;
        string docTitleRu = null;
        string docTitleEn = null;

        if (fn.electronicFilename.Contains("_CKL"))
        {
          docTitleRu = "Контрольный перечень проверки дублирования";
          docTitleEn = "Duplication Review Checklist";
          docTypeCode = null;
          docTypeCodeRu = null;
          docTypeCodeEn = null;
          Rev = null;
          Status = null;
        }




        List<string> trmCsvRow = new List<string>();
        if (vdrCsvDataRow != null && vdrDataRow != null)
        {

          if (!fn.electronicFilename.Contains("_CKL"))
          {
            docTitleRu = vdrDataRow[30];
            docTitleEn = vdrDataRow[29];
            docNumber = vdrDataRow[27];
          }

          trmCsvRow.Add(null); //Row Status
          trmCsvRow.Add(docNumber); //doc No
          trmCsvRow.Add(null); //Vendor doc no
          trmCsvRow.Add(docTitleRu); // doc title RU
          trmCsvRow.Add(docTitleEn); // doc title EN
          trmCsvRow.Add(docTypeCode); // doc type code
          trmCsvRow.Add(docTypeCodeRu); // doc type code RU
          trmCsvRow.Add(docTypeCodeEn); // doc type code EN
          trmCsvRow.Add(vdrDataRow[40]); // doc date
          trmCsvRow.Add(Rev); // rev
          trmCsvRow.Add(Status); // status
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
          trmCsvRow.Add(sectionRu); // package name RU
          trmCsvRow.Add(sectionEn); // package name EN
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
        }
        else
        {
          trmCsvRow.Add(null); //Row Status
          trmCsvRow.Add(docNumber); //doc No
          trmCsvRow.Add(null); //Vendor doc no
          trmCsvRow.Add(docTitleRu); // doc title RU
          trmCsvRow.Add(docTitleEn); // doc title EN
          trmCsvRow.Add(docTypeCode); // doc type code
          trmCsvRow.Add(docTypeCodeRu); // doc type code RU
          trmCsvRow.Add(docTypeCodeEn); // doc type code EN
          trmCsvRow.Add(null); // doc date
          trmCsvRow.Add(Rev); // rev
          trmCsvRow.Add(Status); // status
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
          trmCsvRow.Add(sectionRu); // package name RU
          trmCsvRow.Add(sectionEn); // package name EN
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
