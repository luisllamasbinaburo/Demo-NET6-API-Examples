using System.Reflection;

//Start Microsoft.Office.Interop.Excel and get Application object.
var applicationExcel = new Microsoft.Office.Interop.Excel.Application();
applicationExcel.Visible = true;

//Get a new workbook.
var workbook = (Microsoft.Office.Interop.Excel._Workbook)(applicationExcel.Workbooks.Add(Missing.Value));
var worksheet = (Microsoft.Office.Interop.Excel._Worksheet)workbook.ActiveSheet;

//Add table headers going cell by cell.
worksheet.Cells[1, 1] = "A";
worksheet.Cells[2, 1] = "B";
worksheet.Cells[3, 1] = "C";
worksheet.Cells[4, 1] = "D";

//Fill B1:B4 with a formula(=RAND() * 100000) and apply format.
var range = worksheet.get_Range("B1", "B4");
range.Formula = "=RAND() * 100000";
range.NumberFormat = "$0.00";

applicationExcel.Visible = true;
applicationExcel.UserControl = true;

Console.ReadLine();