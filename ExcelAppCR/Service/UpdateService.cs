using AutoUpdaterDotNET;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace ExcelAppCR.Service
{
    public static class UpdateService
    {
        private const string appCastUrl = "https://raw.githubusercontent.com/Tcon03/ExcelAppCR/refs/heads/master/Version.xml";
        public static void CheckForUpdates()
        {
            AutoUpdater.ReportErrors = true;
            AutoUpdater.RunUpdateAsAdmin = false;
            AutoUpdater.Synchronous = false;
            AutoUpdater.ClearAppDirectory = true;
            //AutoUpdater.InstallationPath = AppDomain.CurrentDomain.BaseDirectory;
            AutoUpdater.CheckForUpdateEvent += AutoUpdaterOnCheckForUpdateEvent;
            AutoUpdater.Start(appCastUrl);
        }

        private static void AutoUpdaterOnCheckForUpdateEvent(UpdateInfoEventArgs args)
        {
            var _dispatcher = Application.Current.Dispatcher;
            if (args == null)
            {
                _dispatcher.Invoke(() =>
                    MessageBox.Show("Không thể kiểm tra cập nhật. Kiểm tra lại mạng hoặc URL XML.",
                        "Cập nhật", MessageBoxButton.OK, MessageBoxImage.Warning));
                return;
            }
            if (args.Error != null)
            {
                _dispatcher.Invoke(() =>
                    MessageBox.Show(
                        $"Không thể kiểm tra cập nhật. Vui lòng kiểm tra lại kết nối mạng hoặc URL.\nLỗi: {args.Error.Message}",
                        "Lỗi Cập Nhật", MessageBoxButton.OK, MessageBoxImage.Error));
                return;
            }

            // check if it is the latest version then skip
            if (!args.IsUpdateAvailable)
            {
                return;
            }
            _dispatcher.Invoke(() =>
            {
                var result = MessageBox.Show(
                     $"Phát hiện phiên bản mới!\n\n" +
                     $"Phiên bản hiện tại: v{args.InstalledVersion}\n" +
                     $"Phiên bản mới: v{args.CurrentVersion}\n\n" +
                     $"Bạn có muốn cập nhật ngay không?\n",
                     "Thông báo cập nhật",
                     MessageBoxButton.YesNo,
                     MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        if (AutoUpdater.DownloadUpdate(args))
                        {
                            Application.Current.Shutdown();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"Lỗi trong quá trình cập nhật: {ex.Message}",
                            "Lỗi",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                }
            });
        }
    }
}
