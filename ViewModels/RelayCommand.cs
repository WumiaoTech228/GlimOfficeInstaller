using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace GOI.ViewModels
{
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<Task> _asyncExecute;
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public RelayCommand(Func<Task> execute, Func<bool> canExecute = null)
        {
            _asyncExecute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object parameter) => _canExecute?.Invoke() ?? true;

        public async void Execute(object parameter)
        {
            if (_execute != null)
            {
                try
                {
                    _execute();
                }
                catch (Exception ex)
                {
                    GOI.Helpers.Logger.Error("Unhandled exception in command execution", ex);
                }
            }
            else if (_asyncExecute != null)
            {
                try
                {
                    await _asyncExecute();
                }
                catch (Exception ex)
                {
                    GOI.Helpers.Logger.Error("Unhandled exception in async command execution", ex);
                }
            }
        }
    }

    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Func<T, Task> _asyncExecute;
        private readonly Func<T, bool> _canExecute;

        public RelayCommand(Action<T> execute, Func<T, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public RelayCommand(Func<T, Task> execute, Func<T, bool> canExecute = null)
        {
            _asyncExecute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object parameter) => _canExecute?.Invoke((T)parameter) ?? true;

        public async void Execute(object parameter)
        {
            T typedParam = (parameter is T) ? (T)parameter : default(T);
            if (_execute != null)
            {
                try
                {
                    _execute(typedParam);
                }
                catch (Exception ex)
                {
                    GOI.Helpers.Logger.Error($"Unhandled exception in command<{typeof(T).Name}> execution", ex);
                }
            }
            else if (_asyncExecute != null)
            {
                try
                {
                    await _asyncExecute(typedParam);
                }
                catch (Exception ex)
                {
                    GOI.Helpers.Logger.Error($"Unhandled exception in async command<{typeof(T).Name}> execution", ex);
                }
            }
        }
    }
}
