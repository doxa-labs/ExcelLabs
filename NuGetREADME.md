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

#### Release Notes - v3.0.3
- Fixed issue where temp files were shareable and not deleted on close
- SaveFileWithCleanXmlText, CleanTextForXml and ColumnIndexToColumnLetter functions added

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

Optionals with June 2024 v3.0.3 Update
4. Call XML-safe SaveFileWithCleanXmlText Function
5. Call CleanTextForXml to clean not-allowed XML characters
6. Call ColumnIndexToColumnLetter to Convert integer to Excel Column Letter like 1 to A
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
    "F#", // N
    "2024 June" // O
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
cells.Add(new Cellx(3, "Xcode"));
cells.Add(new Cellx(3, "Notepad"));

// call save function
ExcelLabs.SaveFile(title, path, sheetName, cells);

// call safe save function
ExcelLabs.SaveFileWithCleanXmlText(title, path, sheetName, cells);

// clean not-allowed XML characters
string safeToWriteText = ExcelLabs.CleanTextForXml(title + " safe");
Console.WriteLine("Safe text: " + safeToWriteText);

// convert integer to Excel Column Letter like 1 to A
string excelColumnLetter1 = ExcelLabs.ColumnIndexToColumnLetter(1);
Console.WriteLine("1 to column letter: " + excelColumnLetter1); // A

// convert integer to Excel Column Letter like 1 to G
string excelColumnLetter7 = ExcelLabs.ColumnIndexToColumnLetter(7);
Console.WriteLine("7 to column letter: " + excelColumnLetter7); // G
```

#### Screenshot
![Image](http://fatihyildizhan.com/others/excel-labsl-output.jpg)

### Support or Contact

Please visit https://github.com/doxa-labs/ExcelLabs or http://doxalabs.co.uk

### License

Excel Labs is released under the MIT license.
