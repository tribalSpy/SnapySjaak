# Portable Clocking App

This is a Windows desktop clocking app for QR scans from a USB scanner. It reads your employee CSV, checks the `TBNR`, shows the employee name with the current time, decides `IN` or `OUT`, and saves one CSV file per day.

## Employee file format

Use a CSV with these headers:

```csv
TBNR,type,name
K2VV71,1.F(fulltime),"Aydin, Serkan"
```

Rules:

- The QR code must contain the `TBNR` value.
- The CSV must have `TBNR`, `type`, and `name`.
- If a name contains a comma, wrap it in quotes.

## How the app behaves

- First scan of the day for a person is `IN`.
- Second scan is `OUT`.
- Third scan is `IN` again, and so on.
- The `Manual Corrections` page lets you add a missed start time, finish time, or both.
- Daily records are saved to `data/records/YYYY-MM-DD.csv`.
- You can export any day from the app and open it directly in Excel.

## Run during development

From this folder:

```powershell
dotnet run --project .\InKlokken.csproj
```

## Build a portable version for another PC

Run:

```powershell
.\build_portable.ps1
```

That creates a self-contained Windows build in `dist/InKlokken/`.

Copy that whole `dist/InKlokken/` folder to the other PC. That PC does not need Python or .NET installed.

Files to keep together:

- `InKlokken.exe`
- `employees.csv`
- `data\` if you want to keep existing clock records
