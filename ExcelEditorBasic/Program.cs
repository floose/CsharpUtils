using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelEditor
{ 
    internal class Program
{ 
    
         static void Main(string[] args)
        {
            Console.WriteLine("Excel Reader App done in Csharp! Yay!");
            Console.WriteLine("Choose a name of Excel file:");
            Console.WriteLine("If nome name is given, filename will be 'text.xlsx'");

            string fileName = Console.ReadLine();
            
            if(fileName == string.Empty)
                fileName = "text.xlsx";

            //check if file exists
            if(!File.Exists(fileName))
            {
                Console.WriteLine($"File {fileName} does not exist. Creating a new file.");
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
                {
                    /*
                     * The following code creates a new Excel file with a single worksheet.
                     * The code basically initializes the structure of an Excel file.
                     * Then, the code creates a new workbook part, adds a worksheet part to it.
                     * Then, it creates a new worksheet with empty sheet data.
                     * Then, it creates a new sheets collection and adds a sheet to it.
                     */
                    WorkbookPart workbookPart = document.AddWorkbookPart(); //Container of excel workbook
                    workbookPart.Workbook = new Workbook(); // Create a new workbook
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(); // Container of excel worksheet. Add to the workbook part
                    worksheetPart.Worksheet = new Worksheet(new SheetData());  //in the worksheet part, create a new worksheet with empty sheet data
                    Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                    // Create a new sheet and add it to the sheets collection
                    Sheet sheet = new Sheet() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                    sheets.Append(sheet);
                    workbookPart.Workbook.Save();
                }
            }

        }

    }

}