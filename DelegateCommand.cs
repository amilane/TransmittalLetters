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

    //RelayCommand
    //private Action<object> execute;
    //private Func<object, bool> canExecute;

    //public event EventHandler CanExecuteChanged {
    //  add { CommandManager.RequerySuggested += value; }
    //  remove { CommandManager.RequerySuggested -= value; }
    //}

    //public DelegateCommand(Action<object> execute, Func<object, bool> canExecute = null)
    //{
    //  this.execute = execute;
    //  this.canExecute = canExecute;
    //}

    //public bool CanExecute(object parameter)
    //{
    //  return this.canExecute == null || this.canExecute(parameter);
    //}

    //public void Execute(object parameter)
    //{
    //  this.execute(parameter);
    //}


  }
}

