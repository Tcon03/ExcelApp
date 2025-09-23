using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
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

        private readonly IExcelReaderService _reader;
        private readonly IProcessingService _processor;
        private readonly IStagingStore _staging;
        private readonly IExcelWriterService _writer;
        private CancellationTokenSource _cts;


        public string InputPath { get; set; }
        public string StatusText { get; set; }
        public double ProgressPercent { get; set; }


        // Paging
        public int PageSize { get; } = 2000;
        public int PageIndex { get; set; }
        public int TotalPages { get; set; }
        public ObservableCollection<RowModel> CurrentPageItems { get; } = new();


        // Commands
        public ICommand BrowseInputCommand { get; }
        public ICommand ImportCommand { get; }
        public ICommand ExportCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand PrevPageCommand { get; }
        public ICommand NextPageCommand { get; }
        public ICommand ApplyFilterCommand { get; }


        public string FilterText { get; set; } = string.Empty;
        public MainViewModel(IExcelReaderService reader, IProcessingService processor,
        IStagingStore staging, IExcelWriterService writer)
        {
            _reader = reader; _processor = processor; _staging = staging; _writer = writer;
            BrowseInputCommand = new RelayCommand(BrowseInput);
            ImportCommand = new RelayCommand(async _ => await ImportAsync(), _ => File.Exists(InputPath));
            ExportCommand = new RelayCommand(async _ => await ExportAsync());
            CancelCommand = new RelayCommand(_ => _cts?.Cancel());
            PrevPageCommand = new RelayCommand(async _ => await LoadPageAsync(PageIndex - 1), _ => PageIndex > 1);
            NextPageCommand = new RelayCommand(async _ => await LoadPageAsync(PageIndex + 1), _ => PageIndex < TotalPages);
            ApplyFilterCommand = new RelayCommand(async _ => await LoadPageAsync(1));
        }


        private void BrowseInput(object _)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|All files|*.*" };
            if (dlg.ShowDialog() == true) InputPath = dlg.FileName; OnPropertyChanged(nameof(InputPath));
        }


        private async Task ImportAsync()
        {
            _cts = new CancellationTokenSource();
            var progress = new Progress<ProgressInfo>(p => { ProgressPercent = p.Percent; StatusText = p.Message; OnPropertyChanged(null); });
            try
            {
                StatusText = "Importing..."; OnPropertyChanged(nameof(StatusText));
                await _staging.InitAsync(_cts.Token);
                await foreach (var row in _reader.ReadRowsAsync(InputPath, _cts.Token, progress))
                {
                    var transformed = _processor.Transform(row);
                    await _staging.StageAsync(transformed, _cts.Token);
                }
                await _staging.FinalizeBatchAsync(_cts.Token);


                var total = await _staging.CountAsync(_cts.Token);
                TotalPages = (int)Math.Ceiling((double)total / PageSize);
                PageIndex = 1; OnPropertyChanged(nameof(TotalPages)); OnPropertyChanged(nameof(PageIndex));
                await LoadPageAsync(1);
                StatusText = $"Imported {total} rows."; OnPropertyChanged(nameof(StatusText));
            }
            catch (OperationCanceledException) { StatusText = "Canceled"; OnPropertyChanged(nameof(StatusText)); }
            catch (Exception ex) { StatusText = $"Error: {ex.Message}"; OnPropertyChanged(nameof(StatusText)); }
            finally { _cts = null; }
        }


        private async Task LoadPageAsync(int page)
        {
            page = Math.Max(1, Math.Min(page, TotalPages));
            var skip = (page - 1) * PageSize;
            var items = await _staging.FetchPageAsync(skip, PageSize, FilterText, CancellationToken.None);
            CurrentPageItems.Clear();
            foreach (var it in items) CurrentPageItems.Add(it);
            PageIndex = page; OnPropertyChanged(nameof(PageIndex));
        }


        private async Task ExportAsync()
        {
            _cts = new CancellationTokenSource();
            var sfd = new Microsoft.Win32.SaveFileDialog { Filter = "Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv" };
            if (sfd.ShowDialog() != true) return;
            var progress = new Progress<ProgressInfo>(p => { ProgressPercent = p.Percent; StatusText = p.Message; OnPropertyChanged(null); });
            try
            {
                if (sfd.FilterIndex == 2)
                    await _writer.WriteCsvAsync(sfd.FileName, _staging, _cts.Token, progress);
                else
                    await _writer.WriteExcelAsync(sfd.FileName, _staging, _cts.Token, progress);
                StatusText = "Export done."; OnPropertyChanged(nameof(StatusText));
            }
            catch (Exception ex) { StatusText = $"Export error: {ex.Message}"; OnPropertyChanged(nameof(StatusText)); }
            finally { _cts = null; }
        }


        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}


