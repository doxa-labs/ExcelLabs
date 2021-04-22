[![Sponsorship](https://img.shields.io/badge/funding-github-%23EA4AAA)](https://github.com/sponsors/doxa-labs)
<a href="https://github.com/doxa-labs/ExcelLabs/actions"><img alt="GitHub Actions status" src="https://github.com/doxa-labs/ExcelLabs/workflows/CI/badge.svg"></a>
[![Codacy Badge](https://api.codacy.com/project/badge/Grade/18e1532454fc48549b95726d6c3cde2a)](https://www.codacy.com/manual/fatihyildizhan/ExcelLabs?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=doxa-labs/ExcelLabs&amp;utm_campaign=Badge_Grade)
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
public class Cellx
{
   public int RowIndex { get; set; }
   public string ColumnName { get; set; }
   public string Value { get; set; }
}
```

#### Usage
```markdown
1. Create a Cell List
2. Add Some Data
3. Call SaveFile Function
```

```C#
string title = "Excel Labs NuGet";
string sheetName = "Simple and Fast";
string path = AppDomain.CurrentDomain.BaseDirectory;

// 1. create a cell list
List<Cellx> cells = new List<Cellx>();

// 2. values as an array
List<string> languages = new List<string>() {
    "Java", // A
    "C#", // B
    "Javascript", // C
    "Swift", // D
    "Php", // E
    "Python", // F
    "Go", // G
    "Swift", // H
    "", // I
    "", // J
    "", // K
    "Objective-C", // L
    "C++", // M
    "F#" // N
    };
    
foreach (string lang in languages)
{
    // no column name for ordered columns
    cells.Add(new Cellx(1, lang));
}

// 3. single value with column name
cells.Add(new Cellx(2, "Fortran", "A"));
cells.Add(new Cellx(2, "Cobol", "D"));
cells.Add(new Cellx(2, "Pascal", "I"));

// 4. single value without column name
cells.Add(new Cellx(3, "Visual Studio"));
cells.Add(new Cellx(3, "Webstorm"));
cells.Add(new Cellx(3, "XCode"));
cells.Add(new Cellx(3, "Notepad"));

// call save function
ExcelLabs.SaveFile(title, path, sheetName, cells);
```

#### Screenshot
![labs](https://user-images.githubusercontent.com/602600/108699615-7f46aa80-7516-11eb-9d46-ff845187ae87.jpg)

### Support or Contact

Please visit https://doxalabs.co.uk

### License

Excel Labs is released under the MIT license.
