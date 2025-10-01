using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelAppCR.View
{
    /// <summary>
    /// Interaction logic for ExcelAppCRMain.xaml
    /// </summary>
    public partial class ExcelAppCRMain : Window
    {
        public ExcelAppCRMain()
        {
            InitializeComponent();
        }

        private void btn_Menu(object sender, RoutedEventArgs e)
        {
            Button bt = sender as Button;
            if (bt == null || bt.ContextMenu == null) return;

            bt.ContextMenu.PlacementTarget = bt;
            bt.ContextMenu.Placement = PlacementMode.Bottom;
            bt.ContextMenu.IsOpen = true;
            e.Handled = true;
        }
    }
}
