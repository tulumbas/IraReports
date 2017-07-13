using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace IraReports.OXML
{
    class OXMLDriver: IDisposable
    {
        public class TableOptions
        {
            public int MaxColumns;
            public int MaxRows;
            public int SkipEmptyRows;
        }

        public static readonly TableOptions DEFAULTOPTIONS = GetTableOptions();

        public static TableOptions GetTableOptions()
        {
            return new TableOptions
            {
                MaxColumns = 200,
                MaxRows = 60000,
                SkipEmptyRows = 1
            };
        }

        bool _isDisposed;
        XLWorkbook _wb;
        public string FileName { get; }

        public OXMLDriver(string fileName, bool trackEvents = false)
        {
            FileName = fileName;
            _wb = new XLWorkbook(fileName, trackEvents ? XLEventTracking.Enabled : XLEventTracking.Disabled);
        }

        /// <summary>
        /// List all sheets. Do not caches the result
        /// </summary>
        /// <returns></returns>
        public List<string> GetSheetNames()
        {
            return _wb.Worksheets.Select(ws => ws.Name).ToList();
        }

        public bool TestCell(string sheetName, string address, string value)
        {
            var ws = _wb.Worksheet(sheetName);
            var cell = ws.Cell(address);
            return string.Equals(cell.Value.CastTo<string>(), value, StringComparison.CurrentCultureIgnoreCase);
        }

        public IEnumerable<T> GetTableDefinedByHeaders<T>(string sheetName, IBindByHeadersCXML<T> binder,
                int startRow = 1, int startColumn = 1) where T: class
        {
            return GetTableDefinedByHeaders<T>(sheetName, binder, DEFAULTOPTIONS, startRow, startColumn);
        }


        /// <summary>
        /// Reads a data range for the sheet and returns elements populated by a binder.
        /// Automtically considers all non empty cells in a row starting from 
        /// to be headers <paramref name="startRow"/> and <paramref name="startColumn"/>. 
        /// Reads row till the first row with an emtpy left (first) cell.
        /// </summary>
        /// <typeparam name="T">Type of elements to create</typeparam>
        /// <param name="sheetName">Name of an Excel sheet</param>
        /// <param name="binder">implementation of <see cref="IBindByHeaders{T}"/> to populate elements</param>
        /// <param name="startRow">header row, optional, default is 1</param>
        /// <param name="startColumn">first data column, optional, default is 1</param>
        /// <param name="maxColumns">max columns to map, optional, default is 200</param>
        /// <param name="maxRows">max rows to read, optional, default is 60K</param>
        /// <returns></returns>
        public IEnumerable<T> GetTableDefinedByHeaders<T>(string sheetName, IBindByHeadersCXML<T> binder,
                TableOptions options, int startRow = 1, int startColumn = 1)
            where T : class
        {
            var result = new List<T>();
            var sheet = _wb.Worksheet(sheetName);
            
            // get last populated cell
            var lastCell = sheet.LastCellUsed(false);
            var lastRow = lastCell.Address.RowNumber;
            var lastColumn = lastCell.Address.ColumnNumber;
            if (lastRow < startRow || lastColumn < startColumn)
            {
                // nothing found
                yield break;
            }

            if(lastRow - startRow > options.MaxRows )
            {
                // adjust max data span
                lastRow = startRow + options.MaxRows;
            }

            if(lastColumn - startColumn > options.MaxColumns)
            {
                // adjust max data span
                lastColumn = startColumn + options.MaxColumns;
            }

            // get all potential data
            var range = sheet.Range(startRow, startColumn, lastRow, lastColumn);
            var data = range.AsTable();

            // collect headers info
            var headers = new List<string>();
            foreach (var field in data.Fields)
            {
                headers.Add(field.Name);
            }

            var columns = headers.Count;
            if (columns == 0)
            {
                yield break;
            }

            binder.DefineHeaders(headers);

            var rowCount = data.RowCount();
            int emptyRows = 0;
            for (int row = 1; row <= rowCount; row++)
            {
                var rowData = data.DataRange.Row(row);
                var instance = binder.CreateInstance(rowData, row);
                if (instance == null)
                {
                    if(emptyRows++ >= options.SkipEmptyRows)
                    {
                        yield break;
                    }
                }
                else
                { 
                    emptyRows = 0;
                    yield return instance;
                }                
            }
        }



        #region IDisposable stuff
        public void Dispose()
        {
            // Dispose of unmanaged resources.
            if (!_isDisposed)
            {
                if(_wb != null) _wb.Dispose();
                _isDisposed = true;
            }
            
            // Suppress finalization.
            GC.SuppressFinalize(this);
        }

        ~OXMLDriver()
        {
            Dispose();
        }

        #endregion

    }
}
