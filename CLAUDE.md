# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

ExcelLabs is a small C# library (NuGet package `Excel.Labs`) for creating `.xlsx` files via DocumentFormat.OpenXml. The publishable library is the single project `Excel.Labs/` (netstandard2.0); everything else in the solution is a demo console app exercising it on different frameworks (.NET Framework 4.7.2, .NET 5–8, .NET 10).

## Commands

```powershell
dotnet build --configuration Release   # build (CI does this on windows-latest, .NET 10 SDK)
dotnet test                            # CI runs this, but there are no test projects
dotnet run --project Excel.Labs.NET10Demo   # run the current demo; writes an .xlsx under bin/.../Files/
```

- Building the full solution requires Windows/.NET Framework reference assemblies (the `Excel.Labs.Demo` project is old-style csproj targeting net472). Building just `Excel.Labs/Excel.Labs.csproj` needs only the .NET SDK.
- `GeneratePackageOnBuild` is on for the library, so every build produces the NuGet package. The package `Version` lives in `Excel.Labs/Excel.Labs.csproj` (currently 3.0.4) and must be bumped for releases; release notes go in the csproj `PackageReleaseNotes` and the README.
- The Debug/Release `DocumentationFile` paths in the library csproj are hardcoded to machine-specific absolute paths — a known wart; XML doc comments are required on public members since docs are generated.

## Architecture

The entire library is three files in `Excel.Labs/Models/`, all in namespace `Doxa.Labs.Excel.Models` (note: namespace does not match the folder/assembly name):

- `ExcelLabs.cs` — static API: `SaveFile` and `SaveFileWithCleanXmlText` (near-duplicate implementations; a change to one usually belongs in both), plus helpers `CleanTextForXml` and `ColumnIndexToColumnLetter`. Cell data type (Number vs String) is inferred from the string value.
- `Cellx.cs` — the one data model: `{RowIndex, ColumnName, Value}`. `RowIndex` starts at 1 (0 throws `ZeroIndexException`). `ColumnName` of `"-1"` (the constructor default) means "auto/ordered column" rather than an explicit cell reference like `A1`.
- `Exceptions.cs` — `ZeroIndexException`, `NullCellListException`.

Cells passed to `SaveFile` are grouped into rows by consecutive `RowIndex` values, so the input list must be ordered by row.

Since OpenXml 3.0.3, `SpreadsheetDocument.Close()` is obsolete — use `Dispose()` (see comments in `ExcelLabs.cs`).

When the public API changes, update the demo projects (at minimum `Excel.Labs.NET10Demo`) and the usage examples in `README.md`, which serves as the package documentation.
