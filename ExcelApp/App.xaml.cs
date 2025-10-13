using OfficeOpenXml;
using Serilog;
using Serilog.Sinks.Debug;
using System.Configuration;
using System.Data;
using System.Windows;

namespace ExcelAppCR
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            Log.Logger = new LoggerConfiguration().MinimumLevel.Information()
                .WriteTo.Debug()
                .CreateLogger();

            Log.Information("Application Starting Up.........................................");
        }
    }

}
