using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IraReports.Xl
{
    /// <summary>
    /// Implements delegated assignment of object properties from a rectangular 
    /// data area in Excel with headers in a first row. 
    /// Object's driver is initialized by calling <see cref="DefineHeaders"/> 
    /// IEnumerable(of string) with column headers from the first row of a rectangular area. 
    /// Object instances are initialized by a driver in a call to PopulateInstance  with values taken from a row.
    /// </summary>
    /// <typeparam name="T">Type of object to initialize</typeparam>
    interface IBindByHeaders<T> where T : class
    {
        int DefineHeaders(IEnumerable<string> headers);
        void PopulateInstance(object[,] data, int row, T instance);
    }
}
