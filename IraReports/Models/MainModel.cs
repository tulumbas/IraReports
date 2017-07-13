using IraReports.Models.Core;
using IraReports.Models.Source;
using IraReports.OXML;
using IraReports.OXMLTemplate;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;

namespace IraReports.Models
{
    class MainModel : INotifyPropertyChanged
    {
        AdAdapter _adapter;

        /// <summary>
        /// Отчеты каналов с хронометражом
        /// </summary>
        public ObservableCollection<string> SourceReportFiles { get; }

        /// <summary>
        /// File path of a catalog with all ads
        /// </summary>
        private string _catalogPath;
        public string CatalogPath
        {
            get
            {
                return _catalogPath;
            }
            set
            {
                if (_catalogPath != value)
                {
                    _catalogPath = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("CatalogPath"));
                }
            }
        }

        private string _feedback;
        public string Feedback
        {
            get
            {
                return _feedback;
            }
            set
            {
                if (_feedback != value)
                {
                    _feedback = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Feedback"));
                }
            }
        }

        public DelegateCommand SelectAdCatalog { get; }
        public DelegateCommand SelectSourceReportFiles { get; }
        public DelegateCommand ExportReports { get; }
        public DelegateCommand RemoveSourceFile { get; }

        public MainModel()
        {
            SelectAdCatalog = new DelegateCommand((x) => SelectAdCatalogAction());
            SelectSourceReportFiles = new DelegateCommand((x) => SelectSourceReportFilesAction());
            ExportReports = new DelegateCommand((x) => ExportReportsAction(), 
                (x) => { return SourceReportFiles.Count > 0 && CatalogPath != null && CatalogPath.Length > 0; });
            SourceReportFiles = new ObservableCollection<string>();
            RemoveSourceFile = new DelegateCommand((file) => RemoveSourceFileAction((string)file));
        }

        #region Поддержка UI
        private static OpenFileDialog CreateExcelOpenDialog(bool multiselect)
        {
            var dlg = new OpenFileDialog();
            dlg.AddExtension = true;
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files|*.xls;*.xlsx";
            dlg.Multiselect = multiselect;
            var result = dlg.ShowDialog();
            if(result ?? false)
            {
                return dlg;
            }
            else
            {
                return null;
            }
        }

        private void SelectAdCatalogAction()
        {
            var dlg = CreateExcelOpenDialog(false);
            if (dlg != null)
            {
                CatalogPath = dlg.FileName;
            }
            ExportReports.RaiseCanExecuteChanged();
        }

        private void SelectSourceReportFilesAction()
        {
            var dlg = CreateExcelOpenDialog(true);
            if (dlg != null)
            {
                AddSourceReportFiles(dlg.FileNames);
            }
        }

        private void AddSourceReportFiles(string[] fileNames)
        {
            foreach (var item in fileNames)
            {
                if(!SourceReportFiles.Contains(item))
                {
                    SourceReportFiles.Add(item);
                }
            }
            ExportReports.RaiseCanExecuteChanged();
        }

        private void RemoveSourceFileAction(string file)
        {
            SourceReportFiles.Remove(file);
        }

        private void AddFeedback(string msg)
        {
            Feedback += msg + "\n";
        }

        public event PropertyChangedEventHandler PropertyChanged;
        #endregion

        private void ExportReportsAction()
        {
            try
            {
                _adapter = new AdAdapter();

                using (var c = new SafeWaitCursor())
                {
                    List<SourceFile> files = new List<SourceFile>();

                    AddFeedback("Загрузка каталога роликов");
                    _adapter.ReadAdReference(CatalogPath);
                    AddFeedback($"Загружено {_adapter.DB.Count} роликов");

                    foreach (var fileName in SourceReportFiles)
                    {
                        if (File.Exists(fileName))
                        {
                            var file = _adapter.ReadCanalReport(fileName);

                            if (file == null)
                            {
                                AddFeedback($"Ошибка при обработке файла '{fileName}'!");
                                throw new ApplicationException($"Ошибка при обработке файла '{fileName}'!");
                            }

                            AddFeedback($"Лист '{file.Canal}' в файле '{fileName}'");
                            AddFeedback($"Прочитано {file.Count} записей на {file.GetTotalSeconds() } секунд");
                            files.Add(file);
                        }
                    }

                    ExtractMonthlyReport(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Фатальная ошибка обработки: {ex.ToString()}", "Ошибка!");
            }
        }

        private static void ExtractMonthlyReport(List<SourceFile> files)
        {
            var path = Path.GetDirectoryName(files.First().FileName) + "\\справки";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            var clients = files.SelectMany(f => f.Records).Select(r => r.Info?.ClientName).Distinct().OrderBy(x => x).ToList();
            foreach (var client in clients)
            {
                if(!string.IsNullOrEmpty(client))
                {
                    var fileName = Path.Combine(path, Utils.SanitizeFileName(client) + ".xlsx");
                    var report = new MonthlyReport(client);
                    foreach (var file in files)
                    {
                        report.AddCanal(file);
                    }
                    report.Save(fileName);
                }
            }
        }


    }
}