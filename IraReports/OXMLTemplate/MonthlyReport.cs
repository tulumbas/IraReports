using ClosedXML.Excel;
using IraReports.Models.Core;
using IraReports.Models.Source;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace IraReports.OXMLTemplate
{
    class MonthlyReport
    {
        class OutputRecord
        {
            [Display(Name = "Дата")]
            public DateTime Date { get; set; }
            [Display(Name = "Время")]
            public string Time { get; set; }
            [Display(Name = "Ролик")]
            public string Code { get; set; }
            [Display(Name = "Хронометраж")]
            public string DurationString { get; set; }

            public OutputRecord(AdRecord r)
            {
                Date = r.Date;
                Time = r.Time;
                Code = r.Code;
                DurationString = r.DurationString;
            }
        }

        string _clientName;
        IXLWorksheet _ws;
        XLWorkbook _wb;
        int _currentRow;
        DateTime _minDate, _maxDate;

        public MonthlyReport(string clientName)
        {
            _clientName = clientName;
            Instantiate();
            _minDate = DateTime.MaxValue;
            _maxDate = DateTime.MinValue;
        }

        private void Instantiate()
        {
            _wb = new XLWorkbook(XLEventTracking.Disabled);
            _ws = _wb.Worksheets.Add(Utils.SanitizeFileName(_clientName));

            _ws.Column(2).Style.NumberFormat.NumberFormatId = 14; // d/m/yyyy
            for (int i = 2; i <= 5; i++)
            {
                _ws.Column(i).Width = 17;
                _ws.Column(i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            // must be after setting columns
            var title = _ws.Cell(2, 2);
            title.Value = "Эфирная справка рекламной компании " + _clientName;
            StyleTitle(title.Style);

            _currentRow = 4;
        }

        private void StyleTitle(IXLStyle style)
        {
            style.Font.SetFontSize(14)
                .Font.SetBold(true)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
                .Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        }

        private void StyleHeaders(IXLStyle style)
        {
            style.Font.SetFontSize(11)
                .Font.SetBold(true)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                .Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        }

        private void StyleCanal(IXLStyle style)
        {
            style.Font.SetFontSize(18)
                .Font.SetBold(true)
                .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
                .Alignment.SetVertical(XLAlignmentVerticalValues.Center);
        }

        private void StyleTable(IXLStyle style)
        {
            style.Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Border.SetInsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XLColor.Black)
                .Border.SetInsideBorderColor(XLColor.Black);
        }

        public void AddCanal(SourceFile file)
        {
            var records = file.Records.Where(r => r.Info?.ClientName == _clientName).ToList();

            if (records.Count > 0)
            {
                var totalTime = records.Sum(r => r.Duration.TotalSeconds);
                var output = records.Select(r => new OutputRecord(r));

                var start = records.Min(r => r.Date);
                var end = records.Max(r => r.Date);

                if (_minDate > start) _minDate = start;
                if (_maxDate < end) _maxDate = end;

                _ws.Cell(_currentRow, 2).Value = "Телеканал " + file.Canal;
                StyleCanal(_ws.Cell(_currentRow, 2).Style);
                _currentRow++;

                var table = _ws.Cell(_currentRow, 2).InsertTable(output.AsEnumerable(), false);
                StyleHeaders(_ws.Range(_currentRow, 2, _currentRow, 5).Style);
                StyleTable(table.Style);

                _currentRow += table.RowCount();
                _ws.Cell(_currentRow, 1).Value = "ИТОГО";
                _ws.Cell(_currentRow, 5).Value = $"{totalTime} сек.";
                StyleHeaders(_ws.LastRowUsed().Style);

                _currentRow += 2;
            }
        }

        public void Save(string fileName)
        {
            // calculate date range
            var dateRange = string.Join(", ", new string[] { _minDate.ToString("MMMM"), _maxDate.ToString("MMMM") }.Distinct().ToArray());
            _ws.Cell(2, 6).Value = dateRange;
            StyleTitle(_ws.Cell(2, 6).Style);
            _wb.SaveAs(fileName);
        }
    }

}
