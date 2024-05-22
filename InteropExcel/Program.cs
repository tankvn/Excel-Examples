using System;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace InteropExcel
{
	internal class Program
	{
		static void Main(string[] args)
		{
			// Get the input file path
			var inputFile = "test.xlsx";
			var outputFile = "output.xlsx";

			Application ExcelApp = null;
			Workbook Workbook = null;
			Worksheet Sheet = null;

			try
			{
				ExcelApp = new Application();
				ExcelApp.DisplayAlerts = false;
				ExcelApp.Visible = true;

				if (ExcelApp == null)
					throw new Exception("Cannot find Excel.");

				var path = Directory.GetCurrentDirectory();
				var filePath = Path.Combine(path, inputFile);
				Console.OutputEncoding = Encoding.UTF8;
				Console.WriteLine("Start open file {0}", filePath);

				Workbook = ExcelApp.Workbooks.Open(filePath,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing, Type.Missing);

				if (Workbook == null)
					throw new Exception("Cannot open spreadsheet: " + filePath);

				//Sheet = (Worksheet)Workbook.Worksheets.get_Item("Sheet1");
				//Sheet = (Worksheet)Workbook.Worksheets.get_Item(1);
				//Sheet = (Worksheet)Workbook.Worksheets.Item[1];
				Sheet = (Worksheet)Workbook.ActiveSheet;

				Sheet.Cells[5, 1].Value = "A5";

				Console.ReadLine();

				// Save to a new file
				/*ExcelApp.ActiveWorkbook.SaveAs(
					outputFile,
					XlFileFormat.xlOpenXMLWorkbook,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					XlSaveAsAccessMode.xlNoChange,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing,
					Type.Missing
				);*/
				Workbook.Save();        // Save changes to the existing file
				Workbook.Close();       // Close without saving changes
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message.ToString());
			}
			finally
			{
				// After reading, relaase the excel project
				if (ExcelApp != null)
					ExcelApp.Quit();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
				Sheet = null;
				Workbook = null;
				ExcelApp = null;
			}
		}
	}
}
