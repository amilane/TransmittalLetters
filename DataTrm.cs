using System.Collections.Generic;
using System.Linq;

namespace TransmitLetter
{
  class DataTrm
  {
    public List<List<string>> dataTrm(List<FileInfo> filesInfo, string _status, string _rev, string _date)
    {
      //путь к VDR
      string pathToVDR = new GetPathsToTemplates().getPathsToTemplates()[0];

      List<List<string>> dataTrm = new List<List<string>>();
      WritterReader writterReader = new WritterReader();
      List<List<string>> vdrData = writterReader.Read(pathToVDR, "VDR");

      foreach (FileInfo fn in filesInfo) {
        string packageCode = fn.shortName.Split('-')[4];
        string docTypeCode = fn.shortName.Split('-')[5];
        List<string> vdrDataRow = vdrData.FirstOrDefault(x => x[28].Contains(fn.shortName));

        //Rev
        string Rev;
        if (_rev != null && _rev.Trim() != "") {
          Rev = _rev;
        } else {
          Rev = fn.rev;
        }

        //Status
        string Status = null;
        if (_status != null && _status.Trim() != "") {
          Status = _status;
        } else if(vdrDataRow != null) {
          Status = vdrDataRow[33];
        }

        //Date
        string Date = null;
        if (_date != null && _date.Trim() != "") {
          Date = _date;
        } else if(vdrDataRow!=null) {
          Date = vdrDataRow[40];
        }

        List<string> trmRow = new List<string>();
        if (vdrDataRow != null) {

          

          trmRow.Add(vdrDataRow[35]); // PO NO
          trmRow.Add(null);//Tag No
          trmRow.Add(null);//supplier doc number
          trmRow.Add(fn.gcDocN);// doc No
          trmRow.Add(fn.lang);// language
          trmRow.Add(vdrDataRow[30]);// doc title RU
          trmRow.Add(vdrDataRow[29]);// doc title EN
          trmRow.Add(vdrDataRow[24]);// dics
          trmRow.Add(packageCode);// package Code
          trmRow.Add(Status);// status
          trmRow.Add(Date);// doc date
          trmRow.Add(vdrDataRow[15]);// doc class
          trmRow.Add(Rev);// rev
          trmRow.Add(docTypeCode);// doc type code
          trmRow.Add(fn.countSheets);// count of sheets
          trmRow.Add(fn.formatPages);// format
          trmRow.Add(fn.fileExtension);// file format
          trmRow.Add(fn.electronicFilename);//filename
        } else {
          trmRow.Add(null); // PO NO
          trmRow.Add(null);//Tag No
          trmRow.Add(null);//supplier doc number
          trmRow.Add(fn.gcDocN);//doc No
          trmRow.Add(fn.lang);// language
          trmRow.Add(null);// doc title RU
          trmRow.Add(null);// doc title EN
          trmRow.Add(null);// dics
          trmRow.Add(packageCode);// package Code
          trmRow.Add(Status);// status
          trmRow.Add(Date);// doc date
          trmRow.Add(null);// doc class
          trmRow.Add(Rev);// rev
          trmRow.Add(docTypeCode);// doc type code
          trmRow.Add(fn.countSheets);// count of sheets
          trmRow.Add(fn.formatPages);// format
          trmRow.Add(fn.fileExtension);// file format
          trmRow.Add(fn.electronicFilename);//filename
        }
        dataTrm.Add(trmRow);
      }
      return dataTrm;
    }
  }
}
