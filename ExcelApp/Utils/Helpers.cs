using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApp.Utils
{
    public class Helpers
    {
        public static string GetCurrentVersion()
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Serilog.Log.Information("Current application version: {Version}", version);
            return version;
        }
    }
}
