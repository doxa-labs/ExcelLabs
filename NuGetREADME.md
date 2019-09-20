## Welcome to Excel Labs

ExcelLabs is an Excel Helper library written in C#. Create Excel files Simple and Fast.

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
![Image](http://fatihyildizhan.com/others/excel-labsl-output.jpg)

### Support or Contact

Please visit https://github.com/doxa-labs/ExcelLabs or http://doxalabs.co.uk

### License

Excel Labs is released under the MIT license.
