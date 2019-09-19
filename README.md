
<a href="https://github.com/doxa-labs/ExcelLabs/actions"><img alt="GitHub Actions status" src="https://github.com/doxa-labs/ExcelLabs/workflows/CI/badge.svg"></a>
[![Version](https://img.shields.io/nuget/v/Excel.Labs.svg?style=flat-square)](https://www.nuget.org/packages/Excel.Labs)
[![Downloads](https://img.shields.io/nuget/dt/Excel.Labs.svg?style=flat-square)](https://www.nuget.org/packages/Excel.Labs)

## Welcome to Excel Labs

ExcelLabs is an Excel Helper library written in C#. 

### Features

This tool provides a C# based solution to create Excel files without complex queries. This package supports Android, iOS, Linux, macOS and Windows.

### Installation

#### NuGet Package Manager
```C#
PM> Install-Package Excel.Labs
```

#### .NET CLI
```C#
> dotnet add package Excel.Labs
```

### Definition

#### Model
```C#
public class LabsCell
{
   public int RowIndex { get; set; }
   public int ColumnIndex { get; set; }
   public dynamic Value { get; set; }
}
```

#### Usage
```markdown
1. Init ExcelLabs
2. Create a Cell List
3. Add Some Data
4. Call Save Function
```

```C#
// excel filename
string title = "Excel Labs NuGet";
// where do you want to save?
// you can define subfolder too
string path = AppDomain.CurrentDomain.BaseDirectory + @"Files\";

// init
ExcelLabs excel = new ExcelLabs(title, path, Extension.Xls);

// create a cell list
List<LabsCell> cells = new List<LabsCell>();

// define row and column indexes then add your data
cells.Add(new LabsCell(10, 20, "Your Value"));

// add some data to cell list
for (int i = 1; i < 20; i++)
{
   cells.Add(new LabsCell(i, i, i));
}

// call save function with the cell list
excel.Save(cells);
```

### Support or Contact

Please visit http://doxalabs.co.uk

### License

Excel Labs is released under the MIT license.
