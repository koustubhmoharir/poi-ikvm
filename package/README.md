# poi-ikvm-net

Enables the use of **Apache POI 5.2.5** in .NET applications (.NET 9 and above) by using **IKVM** to compile Java to .NET. 

This package allows you to read, write, and manipulate Microsoft Excel files natively within your .NET applications, exposing identical Java namespaces (`org.apache.poi.*`).

## Features

- **.NET Target**: Targets .NET 9. For .NET Framework use poi-ikvm 
- **Up-to-Date POI**: Targets Apache POI **5.2.5** which is the last version to support Java 8 enabling use with IKVM 
- **Self-Contained**: All transpiled Maven dependencies (such as XMLBeans, Commons-Compress, Commons-Math, etc.) are bundled directly inside the package. No external JDK required!
- **Identical Namespaces**: Keeps namespaces identical to original Java Apache POI

## Installation

Install the package via the .NET CLI:

```bash
dotnet add package poi-ikvm-net
```

Or via the Package Manager Console:

```powershell
Install-Package poi-ikvm-net
```

## Quick Start Example (C#)

The following example demonstrates how to create a new Excel workbook, add a sheet, write a value to a cell, and save it to a file:

```csharp
using System;
using org.apache.poi.xssf.usermodel;

try
{
    Console.WriteLine("Creating workbook...");
    using var workbook = new XSSFWorkbook();
    var sheet = workbook.createSheet("Sheet1");
    var row = sheet.createRow(0);
    var cell = row.createCell(0);
    cell.setCellValue("Hello from Modern .NET 9 using Apache POI!");

    Console.WriteLine("Writing file to disk...");
    using var fileOut = new java.io.FileOutputStream("output.xlsx");
    workbook.write(fileOut);
    
    Console.WriteLine("Done!");
}
catch (Exception ex)
{
    Console.WriteLine("Error: " + ex.Message);
}
```

## License

This project is licensed under the MIT License. Apache POI is licensed under the Apache License, Version 2.0.
