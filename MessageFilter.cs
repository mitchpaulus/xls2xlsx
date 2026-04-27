using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Xls2Xlsx;

[GeneratedComInterface]
[Guid("00000016-0000-0000-C000-000000000046")]
internal partial interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

    [PreserveSig]
    int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

[GeneratedComClass]
internal sealed partial class OleMessageFilter : IOleMessageFilter
{
    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        => 0;

    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        if (dwRejectType == 2)
        {
            if (dwTickCount > 30_000) return -1;
            return 100;
        }
        return -1;
    }

    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) => 2;
}

internal static partial class MessageFilterRegistration
{
    [LibraryImport("ole32.dll")]
    private static partial int CoRegisterMessageFilter(IntPtr lpMessageFilter, out IntPtr lplpMessageFilter);

    private static OleMessageFilter? _filter;
    private static IntPtr _filterPtr;

    public static void Register()
    {
        if (_filter is not null) return;
        _filter = new OleMessageFilter();
        var wrappers = new StrategyBasedComWrappers();
        _filterPtr = wrappers.GetOrCreateComInterfaceForObject(_filter, CreateComInterfaceFlags.None);
        int hr = CoRegisterMessageFilter(_filterPtr, out _);
        if (hr != 0)
        {
            Marshal.Release(_filterPtr);
            _filterPtr = IntPtr.Zero;
            _filter = null;
            Marshal.ThrowExceptionForHR(hr);
        }
    }
}
