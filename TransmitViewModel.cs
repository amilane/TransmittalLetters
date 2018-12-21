using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace TransmitLetter
{
  public class TransmitViewModel : INotifyPropertyChanged
  {
   // в Path вытаскиватся текущее значение из TextBox
    public event PropertyChangedEventHandler PropertyChanged;
    protected virtual void OnPropertyChanged([CallerMemberName]string prop = null)
    {
      PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
    }


    private string _Path;
    public string Path {
      get { return _Path; }
      set {
        _Path = value;
        OnPropertyChanged(nameof(Path));
      }
    }

    public ICommand RetrieveParametersValuesCommand {
      get {
        return new DelegateCommand(GenerateParametersAndValues);
      }
    }

    public void GenerateParametersAndValues()
    {
      string transmitNumber = Path.Split('\\').Last();

      GetFilesInfo gf = new GetFilesInfo();
      var info = gf.getFiles(Path);

      DataTrm trm = new DataTrm();
      var dataTrm = trm.dataTrm(info);

      DataTrmCsv trmCsv = new DataTrmCsv();
      var dataTrmCsv = trmCsv.dataTrmCsv(info, transmitNumber);

      DataCrs crs = new DataCrs();
      var dataCrs = crs.dataCrs(info);

      WritterReader wr = new WritterReader();

      string pathToTRM = new GetPathsToTemplates().getPathsToTemplates()[1];
      string pathToCSV = new GetPathsToTemplates().getPathsToTemplates()[2];

      string fileName = String.Format("{0}.xlsx", transmitNumber);
      wr.Write(Path, pathToTRM, "TRM Template 08-Feb-2018",
        8, 1,
        dataTrm,
        fileName,
        transmitNumber);

      string fileCsvName = String.Format("{0}_CSV.xlsx", transmitNumber);
      wr.Write(Path, pathToCSV, "Document Load",
        2, 1,
        dataTrmCsv,
        fileCsvName,
        transmitNumber);

      wr.WriteCrs(Path, dataCrs, transmitNumber);


    }
  }
}
