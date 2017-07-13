using System;
using System.Runtime.Serialization;

namespace IraReports.Xl
{
    [Serializable]
    class SheetNotFoundException : SheetException
    {
        public SheetNotFoundException() : base("Sheet not found")
        {
        }

        public SheetNotFoundException(string sheetName)
            : base(sheetName, $"Sheet {sheetName} not found")
        {
        }
    }
}