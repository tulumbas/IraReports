using ClosedXML.Excel;
using IraReports.OXML;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace IraReports.Models.Source
{
    class AdAdapter : IBindByHeadersCXML<AdInfo>, IBindByHeadersCXML<AdRecord>
    {
        static readonly Regex CANALSHEETNAME = new Regex(@"(.+) - (\d{4}-\d{2}-\d{2})", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        OXMLDriver.TableOptions _adRecordOptions, _adInfoOptions;

        public Dictionary<string, AdInfo> DB { get; private set; }

        public AdAdapter()
        {
            _adRecordOptions = OXMLDriver.GetTableOptions();
            _adRecordOptions.MaxColumns = 4;
            _adInfoOptions = OXMLDriver.GetTableOptions();
            _adInfoOptions.MaxColumns = 7;
        }

        public void ReadAdReference(string adFileName)
        {
            DB = new Dictionary<string, AdInfo>();
            using (var xl = new OXMLDriver(adFileName))
            {
                var ads = xl.GetTableDefinedByHeaders<AdInfo>("ролики", this, _adInfoOptions, 2);
                foreach (var ad in ads)
                {
                    if (DB.ContainsKey(ad.AdCode))
                    {
                        DB[ad.AdCode] = ad;
                    }
                    else
                    {
                        DB.Add(ad.AdCode, ad);
                    }
                }
            }
        }

        public SourceFile ReadCanalReport(string fileName)
        {
            SourceFile file = null;

            using (var xl = new OXMLDriver(fileName))
            {
                var sheetName = xl.GetSheetNames().FirstOrDefault();
                if (sheetName != null)
                {
                    var matches = CANALSHEETNAME.Match(sheetName);
                    if (matches.Success && matches.Groups.Count > 2 && matches.Groups[1].Success)
                    {
                        var canal = matches.Groups[1].Value;
                        file = new SourceFile { FileName = fileName, Canal = canal, Date = matches.Groups[1].Value };

                        var test = xl.TestCell(sheetName, "A1", "Дата");
                        var startRow = test ? 1 : 2;
                        var records = xl.GetTableDefinedByHeaders<AdRecord>(sheetName, this, _adRecordOptions, startRow);
                        AdInfo ad;
                        foreach (var record in records)
                        {
                            file.Records.Add(record);
                            if (DB.TryGetValue(record.Code, out ad))
                            {
                                record.Info = ad;
                            }
                        }
                    }
                }
            }
            return file;
        }

        AdInfo IBindByHeadersCXML<AdInfo>.CreateInstance(IXLTableRow row, int rowNumber)
        {
            var start = 0;
            var code = row.Field(start + 1).GetString();
            var client = row.Field(start + 4).GetString();

            if (string.IsNullOrEmpty(code))
            {
                return null;
            }

            var instance = new AdInfo
            {
                AdCode = code,
                ClientName = client,
                AdName = row.Field(start + 2).GetString(),
                Length = row.Field(start + 3).GetString()
            };

            return instance;
        }

        AdRecord IBindByHeadersCXML<AdRecord>.CreateInstance(IXLTableRow row, int rowNumber)
        {
            if (row.Field(0).IsEmpty())
            {
                return null;
            }

            var instance = new AdRecord
            {
                Date = row.Field(0).GetDateTime().Date,
                Time = row.Field(1).GetString(),
                Code = row.Field(2).GetString(),
                DurationString = row.Field(3).GetString()
            };

            return instance;
        }

        void IBindByHeadersCXML<AdInfo>.DefineHeaders(IEnumerable<string> headers)
        {
        }

        void IBindByHeadersCXML<AdRecord>.DefineHeaders(IEnumerable<string> headers)
        {
        }
    }
}
