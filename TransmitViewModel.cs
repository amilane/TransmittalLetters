using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace TransmitLetter
{
  public class TransmitViewModel: INotifyPropertyChanged
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
      //WritterReader wr = new WritterReader();
      //var lol = wr.Read(Path,"UnitList");
      //string kek = lol[1][1];
      //MessageBox.Show(kek);
      GetFilesInfo gf = new GetFilesInfo();
      var info = gf.getFiles(Path);

      DataTrm trm = new DataTrm();
      trm.dataTrm(info);

    }
  }
}
