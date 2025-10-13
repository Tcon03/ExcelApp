using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApp.Commands
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Input;

    namespace ExcelApp.ViewModel
    {
        public class AsyncCommand : ICommand
        {
            private readonly Func<Task> _execute;
            private readonly Func<bool> _canExecute;
            private bool _isExecuting;

            public AsyncCommand(Func<Task> execute, Func<bool> canExecute = null)
            {
                _execute = execute;
                _canExecute = canExecute ?? (() => true);
            }

            public bool CanExecute(object parameter)
            {
                return !_isExecuting && _canExecute();
            }

            public async void Execute(object parameter)
            {
                _isExecuting = true;
                RaiseCanExecuteChanged();
                try
                {
                    await _execute();
                }
                finally
                {
                    _isExecuting = false;
                    RaiseCanExecuteChanged();
                }
            }

            public event EventHandler CanExecuteChanged;

            public void RaiseCanExecuteChanged()
            {
                CanExecuteChanged?.Invoke(this, EventArgs.Empty);
            }
        }
    }
}
