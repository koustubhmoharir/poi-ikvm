using System;
using org.apache.poi.xssf.usermodel;

try
{
    Console.WriteLine("Creating workbook using Apache POI 5.2.5 under .NET 9...");
    var workbook = new XSSFWorkbook();
    var sheet = workbook.createSheet("Test Sheet");
    var row = sheet.createRow(0);
    var cell = row.createCell(0);
    cell.setCellValue("Hello from Modern .NET 9!");

    Console.WriteLine("Writing test.xlsx...");
    var fileOut = new java.io.FileOutputStream("test.xlsx");
    workbook.write(fileOut);
    fileOut.close();
    workbook.close();

    Console.WriteLine("Successfully created test.xlsx!");
}
catch (Exception ex)
{
    Console.WriteLine("Error: " + ex.Message);
    Console.WriteLine(ex.StackTrace);
}
