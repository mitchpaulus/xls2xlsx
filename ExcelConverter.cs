using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace Xls2Xlsx;

internal sealed class ExcelNotInstalledException : Exception
{
    public ExcelNotInstalledException()
        : base("Excel is not installed or not registered for COM automation.") { }
}

internal static class ExcelConverter
{
    private const int xlOpenXMLWorkbook = 51;
    private const int msoAutomationSecurityForceDisable = 3;

    public static void Convert(string inputAbsPath, string outputAbsPath)
    {
        Type? excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType is null) throw new ExcelNotInstalledException();

        object? excel;
        try
        {
            excel = Activator.CreateInstance(excelType);
        }
        catch (COMException)
        {
            throw new ExcelNotInstalledException();
        }
        if (excel is null) throw new ExcelNotInstalledException();

        try
        {
            Set(excel, "Visible", false);
            Set(excel, "DisplayAlerts", false);
            Set(excel, "ScreenUpdating", false);
            Set(excel, "AutomationSecurity", msoAutomationSecurityForceDisable);

            object workbooks = Get(excel, "Workbooks");
            object? workbook = null;
            try
            {
                workbook = Retry(() => Invoke(workbooks, "Open", inputAbsPath, 0));
                Retry(() => Invoke(workbook!, "SaveAs", outputAbsPath, xlOpenXMLWorkbook));
            }
            finally
            {
                if (workbook is not null)
                {
                    try { Invoke(workbook, "Close", false); } catch { }
                    Marshal.FinalReleaseComObject(workbook);
                }
                Marshal.FinalReleaseComObject(workbooks);
            }
        }
        finally
        {
            try { Invoke(excel, "Quit"); } catch { }
            Marshal.FinalReleaseComObject(excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static T Retry<T>(Func<T> action, int attempts = 4, int initialDelayMs = 200)
    {
        int delay = initialDelayMs;
        for (int i = 0; ; i++)
        {
            try { return action(); }
            catch (COMException ex) when (i < attempts - 1 && IsTransient(ex.HResult))
            {
                Thread.Sleep(delay);
                delay *= 2;
            }
        }
    }

    private static void Retry(Action action, int attempts = 4, int initialDelayMs = 200)
        => Retry(() => { action(); return 0; }, attempts, initialDelayMs);

    private static bool IsTransient(int hr) => hr switch
    {
        unchecked((int)0x80010001) => true, // RPC_E_CALL_REJECTED
        unchecked((int)0x8001010A) => true, // RPC_E_SERVERCALL_RETRYLATER
        unchecked((int)0x800AC472) => true, // VBA_E_IGNORE — Excel busy with another OLE action
        _ => false,
    };

    private const BindingFlags NoWrap = BindingFlags.DoNotWrapExceptions;

    private static object Get(object o, string name)
        => o.GetType().InvokeMember(name, BindingFlags.GetProperty | NoWrap, null, o, null)!;

    private static void Set(object o, string name, object v)
        => o.GetType().InvokeMember(name, BindingFlags.SetProperty | NoWrap, null, o, new[] { v });

    private static object Invoke(object o, string name, params object?[] args)
        => o.GetType().InvokeMember(name, BindingFlags.InvokeMethod | NoWrap, null, o, args)!;
}
