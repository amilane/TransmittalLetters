using System;
using System.Windows.Input;

namespace TransmitLetter
{
  public class DelegateCommand : ICommand
  {
    private Action mAction = null;
    public event EventHandler CanExecuteChanged = (sender, e) => { };

    public DelegateCommand(Action action)
    {
      mAction = action;
    }

    public bool CanExecute(object parameter)
    {
      return true;
    }

    public void Execute(object parameter)
    {
      mAction();
    }
  }
}

