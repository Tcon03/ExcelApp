using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelAppCR.Commands
{
    public class VfxCommand : ICommand
    {
        private Action<object> _TargetExecuteMethod;
        private Func<bool> _TargetCanExecuteMethod;

        // 2 constructors để khởi tạo lệnh với phương thức thực thi và điều kiện có thể thực thi 
        // 1. Chỉ cần phương thức thực thi, không cần điều kiện 
        // 2. Cần cả phương thức thực thi và điều kiện có thể thực thi
        public VfxCommand(Action<object> executeMethod)
        {
            _TargetExecuteMethod = executeMethod;
            
        }

        public VfxCommand(Action<object> executeMethod, Func<bool> canExecuteMethod)
        {
            _TargetExecuteMethod = executeMethod;
            _TargetCanExecuteMethod = canExecuteMethod;
        }



        public void RaiseCanExecuteChanged()
        {
            if (this != null)
            {
                CanExecuteChanged(this, EventArgs.Empty);
            }
        }

        bool ICommand.CanExecute(object parameter)
        {
            return _TargetCanExecuteMethod != null ? _TargetCanExecuteMethod() : _TargetExecuteMethod != null;
        }

        // Beware - should use weak references if command instance lifetime
        //  is longer than lifetime of UI objects that get hooked up to command

        // Prism commands solve this in their implementation
        public event EventHandler CanExecuteChanged = delegate { };

        void ICommand.Execute(object parameter)
        {
            _TargetExecuteMethod(parameter);
        }

        //{
        //    private readonly Action<object> _executeAction;
        //    private readonly Func<object, bool> _canExecuteAction;
        //    //Constructors
        //    public VfxCommand(Action<object> executeAction)
        //    {
        //        _executeAction = executeAction;
        //        _canExecuteAction = null;
        //    }


        //    public VfxCommand(Action<object> executeAction, Func<object, bool> canExecuteAction = null) // khởi tạo đối tượng với điều kiện có thể thực thi lệnh 
        //    {
        //        _executeAction = executeAction;
        //        _canExecuteAction = canExecuteAction;
        //    }


        //    //Events
        //    public event EventHandler CanExecuteChanged
        //    {
        //        add { CommandManager.RequerySuggested += value; }
        //        remove { CommandManager.RequerySuggested -= value; }
        //    }
        //    //Methods     gọi để xem lệnh có thể chạy không
        //    public bool CanExecute(object parameter) // Check if the command can be executed đoạn code này sẽ kiểm tra xem lệnh có thể thực thi hay không
        //    {
        //        return _canExecuteAction == null || _canExecuteAction(parameter);
        //    }

        //    //gọi khi người dùng click vào lệnh 
        //    public void Execute(object parameter) // Đoạn code này sẽ thực thi lệnh
        //    {
        //        _executeAction(parameter);
        //    }

    }
}
