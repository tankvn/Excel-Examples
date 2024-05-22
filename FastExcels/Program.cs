using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FastExcel;

namespace FastExcels
{
	internal class Program
	{
		static void Main(string[] args)
		{
			//WriteList();
			//AddRows();
			//Read();
			Update();
		}

		private static void WriteList()
		{
			// Get your template and output file paths
			var templateFile = new FileInfo("test.xlsx");
			var outputFile = new FileInfo("output.xlsx");

			using (var fastExcel = new FastExcel.FastExcel(templateFile, outputFile))
			{
				var objectList = new List<MyObject>();

				for (var rowNumber = 1; rowNumber < 100000; rowNumber++)
				{
					MyObject genericObject = new MyObject();
					genericObject.StringColumn1 = "A string " + rowNumber.ToString();
					genericObject.IntegerColumn2 = 45678854;
					genericObject.DoubleColumn3 = 87.01d;
					genericObject.ObjectColumn4 = DateTime.Now.ToLongTimeString();

					objectList.Add(genericObject);
				}
				fastExcel.Write(objectList, "sheet2", true);
				Console.ReadLine();
			}
		}

		private static void AddRows()
		{
			// Get your template and output file paths
			var templateFile = new FileInfo("test.xlsx");
			var outputFile = new FileInfo("output.xlsx");

			//Create a worksheet with some rows
			var worksheet = new Worksheet();
			var rows = new List<Row>();
			for (var rowNumber = 1; rowNumber < 100000; rowNumber++)
			{
				var cells = new List<Cell>();
				for (var columnNumber = 1; columnNumber < 13; columnNumber++)
				{
					cells.Add(new Cell(columnNumber, columnNumber * DateTime.Now.Millisecond));
				}
				cells.Add(new Cell(13, "Hello" + rowNumber));
				cells.Add(new Cell(14, "Some Text"));

				rows.Add(new Row(rowNumber, cells));
			}
			worksheet.Rows = rows;

			// Create an instance of FastExcel
			using (var fastExcel = new FastExcel.FastExcel(templateFile, outputFile))
			{
				// Write the data
				fastExcel.Write(worksheet, "sheet1");
			}
		}

		private static void Read()
		{
			// Get the input file path
			var inputFile = new FileInfo("test.xlsx");

			//Create a worksheet
			Worksheet worksheet = null;

			// Create an instance of Fast Excel
			using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, true))
			{
				// Read the rows using worksheet name
				worksheet = fastExcel.Read("sheet1");
				Console.WriteLine(string.Format("Worksheet Name:{0}, Index:{1}", worksheet.Name, worksheet.Index));

				// Read the rows using the worksheet index
				// Worksheet indexes are start at 1 not 0
				// This method is slightly faster to find the underlying file (so slight you probably wouldn't notice)
				worksheet = fastExcel.Read(1);

				foreach (var ws in fastExcel.Worksheets)
				{
					Console.WriteLine(string.Format("Worksheet Name:{0}, Index:{1}", ws.Name, ws.Index));

					//To read the rows call read
					ws.Read();
					var rows = ws.Rows.ToArray();
					//Do something with rows
					Console.WriteLine(string.Format("Worksheet Rows:{0}", rows.Count()));
				}
				Console.ReadLine();
			}
		}

		private static void Update()
		{
			// Get the input file path
			var inputFile = new FileInfo("test.xlsx");

			//Create a some rows in a worksheet
			var worksheet = new Worksheet();
			var rows = new List<Row>();

			for (int rowNumber = 1; rowNumber < 100; rowNumber += 1)
			{
				List<Cell> cells = new List<Cell>();
				for (int columnNumber = 1; columnNumber < 13; columnNumber += 2)
				{
					cells.Add(new Cell(columnNumber, rowNumber));
				}
				cells.Add(new Cell(13, "Updated Row"));

				rows.Add(new Row(rowNumber, cells));
			}
			worksheet.Rows = rows;

			// Create an instance of Fast Excel
			using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile))
			{
				// Read the data
				fastExcel.Update(worksheet, "sheet1");
			}
		}

		public class MyObject
		{
			public string StringColumn1
			{
				get; set;
			}
			public int IntegerColumn2
			{
				get; set;
			}
			public double DoubleColumn3
			{
				get; set;
			}
			public string ObjectColumn4
			{
				get; set;
			}
		}
	}
}
