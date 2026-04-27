using System;
using System.IO;

namespace Xls2Xlsx;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 0) { PrintUsage(); return 2; }

        string? input = null;
        string? output = null;
        bool force = false;

        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            switch (a)
            {
                case "-h":
                case "--help":
                    PrintUsage();
                    return 0;
                case "-i":
                case "--input":
                    if (++i >= args.Length) { Err($"missing value for {a}"); return 2; }
                    input = args[i];
                    break;
                case "-o":
                case "--output":
                    if (++i >= args.Length) { Err($"missing value for {a}"); return 2; }
                    output = args[i];
                    break;
                case "-f":
                case "--force":
                    force = true;
                    break;
                default:
                    Err($"unknown argument: {a}");
                    PrintUsage();
                    return 2;
            }
        }

        if (input is null || output is null)
        {
            Err("both -i and -o are required");
            PrintUsage();
            return 2;
        }

        string inAbs = Path.GetFullPath(input);
        string outAbs = Path.GetFullPath(output);

        if (!File.Exists(inAbs))
        {
            Err($"input not found: {inAbs}");
            return 3;
        }

        string? outDir = Path.GetDirectoryName(outAbs);
        if (outDir is not null && !Directory.Exists(outDir))
        {
            Err($"output directory does not exist: {outDir}");
            return 4;
        }

        if (outAbs.IndexOfAny(new[] { '[', ']' }) >= 0)
        {
            Err($"output path contains '[' or ']' which Excel cannot save through: {outAbs}");
            return 4;
        }

        if (File.Exists(outAbs))
        {
            if (!force)
            {
                Err($"output exists; pass -f to overwrite: {outAbs}");
                return 4;
            }
            try { File.Delete(outAbs); }
            catch (Exception ex) { Err($"could not delete existing output: {ex.Message}"); return 4; }
        }

        try { MessageFilterRegistration.Register(); }
        catch { /* non-fatal: retry wrapper still handles transient errors */ }

        try
        {
            ExcelConverter.Convert(inAbs, outAbs);
            return 0;
        }
        catch (ExcelNotInstalledException ex)
        {
            Err(ex.Message);
            return 5;
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            Err($"conversion failed: {ex.Message} (HRESULT 0x{ex.HResult:X8})");
            return 6;
        }
        catch (Exception ex)
        {
            Err($"conversion failed: {ex.GetType().Name}: {ex.Message}");
            return 6;
        }
    }

    private static void Err(string msg) => Console.Error.WriteLine(msg);

    private static void PrintUsage()
    {
        Console.Error.WriteLine("""
            xls2xlsx — convert .xls to .xlsx via Excel COM (Windows only)

            Usage: xls2xlsx -i <input.xls> -o <output.xlsx> [-f]

              -i, --input    path to input .xls (required)
              -o, --output   path to output .xlsx (required)
              -f, --force    overwrite output if it exists
              -h, --help     show this message

            Notes:
              Requires Excel installed and registered for COM automation.
              Macros (VBA) are dropped; output format is .xlsx, not .xlsm.
            """);
    }
}
