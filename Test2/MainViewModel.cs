using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Test2
{
    public class MainViewModel : INotifyPropertyChanged
    {
        #region Private Fields
        private readonly ExcelService _excelService;
        private readonly Timer _memoryTimer;
        private readonly Stopwatch _stopwatch;
        private CancellationTokenSource _cancellationTokenSource;

        private string _inputFileName = "No file selected";
        private string _inputFileInfo = "";
        private string _processingStatus = "Ready";
        private string _processingDetails = "";
        private string _progressText = "Ready to process";
        private double _progressPercentage = 0;
        private string _elapsedTime = "00:00:00";
        private string _statusMessage = "Ready";
        private string _memoryUsage = "0 MB";
        private long _totalRows = 0;
        private bool _isFileLoaded = false;
        private bool _hasProcessedData = false;
        private bool _showNoDataMessage = true;
        private bool _isProcessing = false;

        private ObservableCollection<DataRowView> _previewData;
        private DataTable _currentDataTable;
        private string _currentFilePath;
        #endregion

        #region Public Properties
        public string InputFileName
        {
            get => _inputFileName;
            set => SetProperty(ref _inputFileName, value);
        }

        public string InputFileInfo
        {
            get => _inputFileInfo;
            set => SetProperty(ref _inputFileInfo, value);
        }

        public string ProcessingStatus
        {
            get => _processingStatus;
            set => SetProperty(ref _processingStatus, value);
        }

        public string ProcessingDetails
        {
            get => _processingDetails;
            set => SetProperty(ref _processingDetails, value);
        }

        public string ProgressText
        {
            get => _progressText;
            set => SetProperty(ref _progressText, value);
        }

        public double ProgressPercentage
        {
            get => _progressPercentage;
            set => SetProperty(ref _progressPercentage, value);
        }

        public string ElapsedTime
        {
            get => _elapsedTime;
            set => SetProperty(ref _elapsedTime, value);
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public string MemoryUsage
        {
            get => _memoryUsage;
            set => SetProperty(ref _memoryUsage, value);
        }

        public long TotalRows
        {
            get => _totalRows;
            set => SetProperty(ref _totalRows, value);
        }

        public bool IsFileLoaded
        {
            get => _isFileLoaded;
            set => SetProperty(ref _isFileLoaded, value);
        }

        public bool HasProcessedData
        {
            get => _hasProcessedData;
            set => SetProperty(ref _hasProcessedData, value);
        }

        public bool ShowNoDataMessage
        {
            get => _showNoDataMessage;
            set => SetProperty(ref _showNoDataMessage, value);
        }

        public bool IsProcessing
        {
            get => _isProcessing;
            set => SetProperty(ref _isProcessing, value);
        }

        public ObservableCollection<DataRowView> PreviewData
        {
            get => _previewData;
            set => SetProperty(ref _previewData, value);
        }
        #endregion

        #region Commands
        public ICommand OpenFileCommand { get; }
        public ICommand ProcessDataCommand { get; }
        public ICommand SaveFileCommand { get; }
        public ICommand CancelCommand { get; }
        #endregion

        #region Constructor
        public MainViewModel()
        {
            _excelService = new ExcelService();
            _stopwatch = new Stopwatch();
            _previewData = new ObservableCollection<DataRowView>();

            // Initialize commands
            OpenFileCommand = new RelayCommand(async () => await OpenFileAsync(), () => !IsProcessing);
            ProcessDataCommand = new RelayCommand(async () => await ProcessDataAsync(), () => IsFileLoaded && !IsProcessing);
            SaveFileCommand = new RelayCommand(async () => await SaveFileAsync(), () => HasProcessedData && !IsProcessing);
            CancelCommand = new RelayCommand(CancelOperation, () => IsProcessing);

            // Memory monitoring timer
            _memoryTimer = new Timer(UpdateMemoryUsage, null, TimeSpan.Zero, TimeSpan.FromSeconds(1));

            StatusMessage = "Application ready. Select an Excel file to begin.";
        }
        #endregion

        #region Public Methods
        public async Task OpenFileAsync()
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Title = "Select Excel File",
                    Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                    RestoreDirectory = true
                };

                if (openFileDialog.ShowDialog() != true) return;

                _currentFilePath = openFileDialog.FileName;
                UpdateFileInfo();

                await LoadFilePreviewAsync();
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error opening file: {ex.Message}";
                ProcessingStatus = "Error";
            }
        }

        public async Task ProcessDataAsync()
        {
            if (string.IsNullOrEmpty(_currentFilePath)) return;

            try
            {
                StartProcessing("Processing Excel file...");

                var progress = new Progress<ProcessingProgress>(UpdateProgress);
                _cancellationTokenSource = new CancellationTokenSource();

                // Thực hiện xử lý dữ liệu
                var processedData = await _excelService.ProcessLargeExcelFileAsync(
                    _currentFilePath,
                    progress,
                    _cancellationTokenSource.Token);

                if (processedData != null && !_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    _currentDataTable = processedData;
                    UpdatePreviewData(processedData);
                    HasProcessedData = true;
                    ProcessingStatus = "Processing Complete";
                    StatusMessage = $"Successfully processed {processedData.Rows.Count:N0} rows";
                    ProgressText = "Processing completed successfully";
                    ProgressPercentage = 100;
                }
                else if (_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    ProcessingStatus = "Cancelled";
                    StatusMessage = "Processing cancelled by user";
                }
            }
            catch (OperationCanceledException)
            {
                ProcessingStatus = "Cancelled";
                StatusMessage = "Processing cancelled";
            }
            catch (Exception ex)
            {
                ProcessingStatus = "Error";
                StatusMessage = $"Processing error: {ex.Message}";
            }
            finally
            {
                StopProcessing();
            }
        }

        public async Task SaveFileAsync()
        {
            if (_currentDataTable == null) return;

            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Title = "Save Processed Data",
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                    DefaultExt = "xlsx",
                    FileName = $"Processed_{Path.GetFileNameWithoutExtension(_currentFilePath)}.xlsx"
                };

                if (saveFileDialog.ShowDialog() != true) return;

                StartProcessing("Saving processed data...");

                var progress = new Progress<ProcessingProgress>(UpdateProgress);
                _cancellationTokenSource = new CancellationTokenSource();

                await _excelService.SaveToExcelAsync(
                    _currentDataTable,
                    saveFileDialog.FileName,
                    progress,
                    _cancellationTokenSource.Token);

                if (!_cancellationTokenSource.Token.IsCancellationRequested)
                {
                    ProcessingStatus = "Save Complete";
                    StatusMessage = $"File saved successfully: {Path.GetFileName(saveFileDialog.FileName)}";
                    ProgressText = "File saved successfully";
                    ProgressPercentage = 100;
                }
            }
            catch (OperationCanceledException)
            {
                ProcessingStatus = "Save Cancelled";
                StatusMessage = "Save operation cancelled";
            }
            catch (Exception ex)
            {
                ProcessingStatus = "Save Error";
                StatusMessage = $"Save error: {ex.Message}";
            }
            finally
            {
                StopProcessing();
            }
        }

        public void CancelOperation()
        {
            _cancellationTokenSource?.Cancel();
            ProcessingStatus = "Cancelling...";
            StatusMessage = "Cancelling operation...";
        }
        #endregion

        #region Private Methods
        private async Task LoadFilePreviewAsync()
        {
            try
            {
                StartProcessing("Loading file preview...");

                var progress = new Progress<ProcessingProgress>(UpdateProgress);
                var previewData = await _excelService.GetPreviewDataAsync(_currentFilePath, 1000, progress);

                if (previewData != null)
                {
                    UpdatePreviewData(previewData);
                    IsFileLoaded = true;
                    ProcessingStatus = "File Loaded";
                    StatusMessage = $"Preview loaded: {previewData.Rows.Count:N0} rows displayed";
                    ProgressText = "File preview loaded";
                    ProgressPercentage = 100;
                }
            }
            catch (Exception ex)
            {
                ProcessingStatus = "Load Error";
                StatusMessage = $"Preview load error: {ex.Message}";
            }
            finally
            {
                StopProcessing();
            }
        }

        private void UpdateFileInfo()
        {
            if (string.IsNullOrEmpty(_currentFilePath)) return;

            var fileInfo = new FileInfo(_currentFilePath);
            InputFileName = Path.GetFileName(_currentFilePath);
            InputFileInfo = $"Size: {FormatBytes(fileInfo.Length)} | Modified: {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm}";
        }

        private void UpdatePreviewData(DataTable dataTable)
        {
            if (dataTable == null) return;

            PreviewData.Clear();

            // Chỉ hiển thị tối đa 1000 rows để tránh lag UI
            var rowsToShow = Math.Min(1000, dataTable.Rows.Count);
            var dataView = dataTable.AsDataView();

            for (int i = 0; i < rowsToShow; i++)
            {
                PreviewData.Add(dataView[i]);
            }

            TotalRows = dataTable.Rows.Count;
            ShowNoDataMessage = PreviewData.Count == 0;
        }

        private void StartProcessing(string message)
        {
            IsProcessing = true;
            ProcessingStatus = "Processing...";
            ProgressText = message;
            ProgressPercentage = 0;
            _stopwatch.Restart();
        }

        private void StopProcessing()
        {
            IsProcessing = false;
            _stopwatch.Stop();
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
        }

        private void UpdateProgress(ProcessingProgress progress)
        {
            ProgressPercentage = progress.Percentage;
            ProgressText = progress.Message;
            ProcessingDetails = progress.Details;
            ElapsedTime = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss");

            if (progress.TotalRows > 0)
            {
                TotalRows = progress.TotalRows;
            }
        }

        private void UpdateMemoryUsage(object state)
        {
            var memoryUsed = GC.GetTotalMemory(false);
            MemoryUsage = FormatBytes(memoryUsed);
        }

        private static string FormatBytes(long bytes)
        {
            string[] suffixes = { "B", "KB", "MB", "GB", "TB" };
            int counter = 0;
            decimal number = bytes;

            while (Math.Round(number / 1024) >= 1)
            {
                number /= 1024;
                counter++;
            }

            return $"{number:n1} {suffixes[counter]}";
        }

        protected virtual void SetProperty<T>(ref T backingStore, T value, [CallerMemberName] string propertyName = "")
        {
            if (Equals(backingStore, value)) return;

            backingStore = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        #endregion

        #region IDisposable
        public void Dispose()
        {
            _memoryTimer?.Dispose();
            _cancellationTokenSource?.Dispose();
            _currentDataTable?.Dispose();
        }
        #endregion
    }
}
