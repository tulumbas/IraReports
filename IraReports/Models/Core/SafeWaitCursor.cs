using System;
using System.Windows.Input;

namespace IraReports.Models.Core
{
    class SafeWaitCursor : IDisposable
    {
        private bool isDisposed;
        private Cursor _previousCursor;

        public SafeWaitCursor()
        {
            _previousCursor = Mouse.OverrideCursor;
            Mouse.OverrideCursor = Cursors.Wait;
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (!isDisposed)
            {
                try
                {
                    Mouse.OverrideCursor = _previousCursor;
                }
                catch { }
            }
            isDisposed = true;
        }
        #endregion
    }
}
