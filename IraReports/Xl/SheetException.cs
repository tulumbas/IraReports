using System;
using System.Runtime.Serialization;

namespace IraReports.Xl
{
    [Serializable]
    class SheetException : Exception
    {
        readonly string _sheetName;
        public string SheetName { get { return _sheetName; } }

        public SheetException() { }
        public SheetException(string message) : base(message) { }

        public SheetException(string sheetName, string message) : base(message)
        {
            _sheetName = sheetName;
        }

        public SheetException(string sheetName, string message, Exception inner) : base(message, inner)
        {
            _sheetName = sheetName;
        }

        protected SheetException(SerializationInfo info, StreamingContext context) : base(info, context)
        { }
    }
}