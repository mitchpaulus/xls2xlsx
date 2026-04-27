# xls2xlsx

Windows-only CLI to convert legacy `.xls` files to `.xlsx` by driving a real Excel installation via COM.
Output fidelity is identical to opening the file in Excel and using *Save As*.

## Build

Requires .NET 10 SDK on Windows.

```cmd
dotnet publish -c Release
```

Output binary: `bin\Release\net10.0\win-x64\publish\xls2xlsx.exe` (single native AOT executable, no .NET runtime needed on target).

## Usage

```cmd
xls2xlsx.exe -i input.xls -o output.xlsx [-f]
```

| Flag | Meaning |
| --- | --- |
| `-i`, `--input`  | Path to `.xls` (required) |
| `-o`, `--output` | Path to `.xlsx` (required) |
| `-f`, `--force`  | Overwrite output if it exists |
| `-h`, `--help`   | Show usage |

Macros (VBA) are dropped because the output format is `.xlsx`, not `.xlsm`.

## Exit codes

| Code | Meaning |
| --- | --- |
| 0 | Success |
| 2 | Bad / missing arguments |
| 3 | Input file not found |
| 4 | Output exists and `-f` not given |
| 5 | Excel not installed / not registered for COM |
| 6 | Conversion failed (see stderr) |

## Sequential batch (cmd)

```cmd
for %f in (*.xls) do xls2xlsx.exe -i "%f" -o "%~nf.xlsx" -f
```

## Parallel batch (PowerShell 7+)

```powershell
Get-ChildItem *.xls | ForEach-Object -Parallel {
    & xls2xlsx.exe -i $_.FullName -o ($_.BaseName + ".xlsx") -f
} -ThrottleLimit 4
```

Recommended throttle limit: 4–8. Each invocation launches its own `EXCEL.EXE` process. An `IOleMessageFilter` plus a small retry loop absorb transient COM "server busy" errors that arise under concurrency.

## Requirements

- Windows
- A licensed, installed Excel (any reasonably recent version)
- .NET 10 SDK to build (not needed on machines that only run the binary)
