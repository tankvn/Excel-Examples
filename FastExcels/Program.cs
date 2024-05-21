using System;
using System.Collections.Generic;
using System.IO;

namespace FastExcels
{
	internal class Program
	{
		static void Main(string[] args)
		{
			// Get your template and output file paths
			var templateFile = new FileInfo("test.xlsx");
			var outputFile = new FileInfo("output.xlsx");

			using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(templateFile, outputFile))
			{
				List<MyObject> objectList = new List<MyObject>();

				for (int rowNumber = 1; rowNumber < 100000; rowNumber++)
				{
					MyObject genericObject = new MyObject();
					genericObject.StringColumn1 = "A string " + rowNumber.ToString();
					genericObject.IntegerColumn2 = 45678854;
					genericObject.DoubleColumn3 = 87.01d;
					genericObject.ObjectColumn4 = DateTime.Now.ToLongTimeString();

					objectList.Add(genericObject);
				}
				fastExcel.Write(objectList, "sheet3", true);
				Console.ReadLine();
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
