using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace IraReports.Xl
{
    class OleMessageFilter : IOleMessageFilter
    {
        public static void Register()
        {
            IOleMessageFilter newFilter = new OleMessageFilter();
            IOleMessageFilter oldFilter = null;

            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                CoRegisterMessageFilter(newFilter, out oldFilter);
            }
            else
            {
                throw new COMException("Unable to register message filter because the current thread apartment state is not STA.");
            }
        }

        public static void Revoke()
        {
            IOleMessageFilter oldFilter = null;
            CoRegisterMessageFilter(null, out oldFilter);
        }

        int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return (int)SERVERCALL.SERVERCALL_ISHANDLED;
        }

        int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            return (dwRejectType == (int)SERVERCALL.SERVERCALL_RETRYLATER) ? 99 : -1;
        }

        int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return (int)PENDINGMSG.PENDINGMSG_WAITDEFPROCESS;
        }

        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    }

    enum SERVERCALL
    {
        SERVERCALL_ISHANDLED = 0,
        SERVERCALL_REJECTED = 1,
        SERVERCALL_RETRYLATER = 2
    }

    enum PENDINGMSG
    {
        PENDINGMSG_CANCELCALL = 0,
        PENDINGMSG_WAITNOPROCESS = 1,
        PENDINGMSG_WAITDEFPROCESS = 2
    }

    [ComImport(), Guid("00000016-0000-0000-C000-000000000046"),
    InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

}
