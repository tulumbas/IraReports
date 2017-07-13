using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace IraReports.Xl
{
    class XlDriver : IDisposable
    {
        public int ScanningBatchSize = 200;
        public int MaxEmptyRowsToSkip = 1;

        private string _fileName;
        private bool _isDisposed;

        private Application App;
        private Workbook CurrentWorkbook;

        /// <summary>
        /// Opens excel file and caches Application and Workbook objects
        /// </summary>
        /// <param name="fileName"></param>
        public XlDriver(string fileName)
        {
            _fileName = fileName;

            try
            {
                OleMessageFilter.Register();
            }
            catch (Exception e)
            {
                // log.Error("Could not register OleMessageFilter", e);
                throw;
            }

            App = new Application();
            //App.Visible = true;
            App.DisplayAlerts = false; // ignore message/input boxes
            CurrentWorkbook = App.Workbooks.Open(fileName, 0, true);
            System.Threading.Thread.Sleep(2000); // dirty hack in an attempt to stop 'Call was rejected by callee'
        }

        /// <summary>
        /// List all sheets. Do not caches the result
        /// </summary>
        /// <returns></returns>
        public List<string> GetSheetNames()
        {
            var sheetNames = new List<string>();
            var sheets = CurrentWorkbook.Sheets;
            foreach (Worksheet sheet in sheets)
            {
                sheetNames.Add(sheet.Name);
            }
            return sheetNames;
        }

        /// <summary>
        /// returns a worksheet by name and intercepts some of possible errors
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private Worksheet GetSheet(string sheetName)
        {
            try
            {
                var sheet = CurrentWorkbook.Sheets[sheetName];
                return sheet;
            }
            catch (COMException ex)
            {
                switch (ex.ErrorCode)
                {
                    case -2147352565: // Invalid index. (Exception from HRESULT: 0x8002000B (DISP_E_BADINDEX))
                        //log.ErrorFormat("GetSheet: {0}!{1} not found", _fileName, sheetName);
                        throw new SheetNotFoundException(sheetName);

                    default:
                        //log.Error("GetSheet: Unknown COM error", ex);
                        throw;
                }
            }
        }

        public object[,] GetRange(string sheetName, string topLeft, string bottomRight)
        {
            try
            {
                var sheet = GetSheet(sheetName);
                var range = sheet.Range[topLeft, bottomRight];
                return range.Value as object[,];
            }
            catch (Exception ex)
            {
                //log.ErrorFormat("ReadRange: {0}", ex.ToString());
                throw;
            }
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
        public IEnumerable<T> GetTableDefinedByHeaders<T>(string sheetName, IBindByHeaders<T> binder,
                int startRow = 1, int startColumn = 1, int maxColumns = 200, int maxRows = 60000)
            where T : class, new()
        {
            var result = new List<T>();
            var headers = new List<string>();

            var sheet = GetSheet(sheetName);
            
            for (var c = startColumn; c <= maxColumns; c++) // max 200 columns
            {
                var cellValue = sheet.Cells[startRow, c].Value;
                if (cellValue != null)
                {
                    var name = cellValue.ToString();
                    if (!string.IsNullOrEmpty(name))
                    {
                        headers.Add(name);
                        continue;
                    }
                }
                break;
            }

            var columns = headers.Count;
            if (columns == 0)
            {
                yield break;
            }
            var startLetter = Encoding.ASCII.GetChars(new byte[] { (byte)(64 + startColumn) })[0];
            var endLetter = Encoding.ASCII.GetChars(new byte[] { (byte)(64 + startColumn + columns) })[0];

            binder.DefineHeaders(headers);

            bool done = false;
            int rows, row, step = ScanningBatchSize;
            for (rows = startRow + 1; rows <= maxRows; rows += step) // excluding headers
            {
                //var tl = new ExcelLocation(rows, startColumn);
                //var br = new ExcelLocation(rows + step - 1, columns);
                var range = sheet.Range[$"{startLetter}{rows}", $"{endLetter}{rows + step - 1}"];
                var data = range.Value2 as object[,];
                var emptyRows = 0;
                for (row = 1; row <= data.GetUpperBound(0); row++)
                {
                    int dataCheck;
                    if (data[row, 1] == null || data[row, 1].ToString().Trim().Length < 1 ||
                            (int.TryParse(data[row, 1].ToString().Trim(), out dataCheck) && dataCheck == 0))
                    {
                        if (emptyRows++ < MaxEmptyRowsToSkip) continue;

                        done = true;
                        break;
                    }

                    var instance = new T();
                    //result.Add(instance);
                    binder.PopulateInstance(data, row, instance);
                    yield return instance;
                }

                if (done)
                {
                    yield break;
                }
            }
        }

        public static void CreateFile(string fileName, object[,] data)
        {
            Application app = null;
            Workbook wb = null;
            Worksheet ws;

            try
            {
                app = new Application();
                app.DisplayAlerts = false; // ignore message/input boxes
                app.Visible = false;
                wb = app.Workbooks.Add();
                ws = wb.Worksheets.Add();

                var rows = data.GetUpperBound(0);
                var columns = data.GetUpperBound(1);

                var endLetter = Encoding.ASCII.GetChars(new byte[] { (byte)(columns + 64) })[0];
                var range = ws.Range["A1", endLetter + rows.ToString()];
                range.Value = data;
                wb.SaveAs(fileName, 1);
            }
            catch
            {
                throw;
            }
            finally
            {
                if (app != null)
                {
                    if (wb != null)
                    {
                        Marshal.FinalReleaseComObject(wb);
                    }
                    app.DisplayAlerts = false;
                    app.Quit();
                    Marshal.FinalReleaseComObject(app);
                }
            }
        }

        #region Helpers
        public static bool TryGet<T>(object o, out T value)
        {
            value = default(T);
            if (o == null) { return true; }
            else if (o is T) { value = (T)o; return true; }
            else if (o is string && string.IsNullOrWhiteSpace(o as string)) { return true; }
            return false;
        }

        public static int GetInt(object o)
        {
            int val;
            if (!TryGet<int>(o, out val))
            {
                if (int.TryParse(o.ToString(), out val)) return val;
                if (o is double) return (int)Math.Round((double)o);
                if (o is float) return (int)Math.Round((float)o);
                throw new InvalidCastException(string.Format("Invalid cast: '{0}' casted to int", o));
            }
            return val;
        }

        public static double GetDouble(object o)
        {
            double val;
            if (!TryGet<double>(o, out val))
            {
                if (o is int) return (int)o;
                if (o is float) return (float)o;
                if (double.TryParse(o.ToString(), out val)) return val;
                throw new InvalidCastException(string.Format("Invalid cast: '{0}' casted to double", o));
            }
            return val;
        }

        public static DateTime GetDate(object o)
        {
            DateTime dt;
            if (!TryGet<DateTime>(o, out dt))
            {
                if (DateTime.TryParse(o.ToString(), out dt)) return dt;
                throw new InvalidCastException(string.Format("Invalid cast: '{0}' casted to Date", o));
            }
            return dt;
        }

        public static string GetString(object o)
        {
            string s;
            if (!TryGet<string>(o, out s))
            {
                return o.ToString().Trim();
            }
            return s == null ? null : s.Trim();
        }

        public static int Row(string location)
        {
            if (string.IsNullOrEmpty(location)) return 0;
            var skip = location.Count(c => Char.IsLetter(c));
            int row;
            return int.TryParse(location.Substring(skip), out row) ? row : 0;
        }

        public static int Col(string location)
        {
            int col = 0;
            if (string.IsNullOrEmpty(location)) return col;
            foreach (var c in location.Where(x => Char.IsLetter(x)))
            {
                col = 26 * col + (int)(Char.ToUpper(c)) - 64;
            }

            return col;
        }

        #endregion

        #region IDisposable stuff
        public void Dispose()
        {
            // Dispose of unmanaged resources.
            if (!_isDisposed)
            {
                if (App != null)
                {
                    if (CurrentWorkbook != null)
                    {
                        Marshal.FinalReleaseComObject(CurrentWorkbook);
                    }
                    App.DisplayAlerts = false;
                    App.Quit();
                    Marshal.FinalReleaseComObject(App);

                    try
                    {
                        OleMessageFilter.Revoke();
                    }
                    catch { }
                }

                _isDisposed = true;
            }
            // Suppress finalization.
            //GC.SuppressFinalize(this);
        }

        ~XlDriver()
        {
            Dispose();
        }

        #endregion
    }
}
