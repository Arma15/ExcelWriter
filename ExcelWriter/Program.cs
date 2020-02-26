﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "C:\\Users\\3D Infotech.3DCA-LY520-12\\Desktop\\Example.xlsx";

            /*
            if (args.Length < 1)
            {
                Console.WriteLine("Error, no arguments passed.");

                return;
            }
            filePath = args[0];

            Console.WriteLine($"Path passed is: {filePath}");

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File in path: {filePath} does not exist.");
                return;
            }
            */

            //create a fileinfo object of an excel file on the disk (file must exist)
            FileInfo file = new FileInfo(filePath);

            //create a new Excel package from the file
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                //Get a WorkSheet by index. Note that EPPlus indexes are base 1, not base 0!
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1];

                /*
                //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
                ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["SomeWorksheet"];

                //If you don't know if a worksheet exists, you could use LINQ,
                //So it doesn't throw an exception, but return null in case it doesn't find it
                ExcelWorksheet anotherWorksheet =
                    excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "SomeWorksheet");
                */

                //Get the content from cells A1 and B1 as string, in two different notations
                string valA1 = firstWorksheet.Cells["A1"].Value.ToString();
                string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();

                //add some data
                firstWorksheet.Cells[4, 1].Value = "Added data in Cell A4";
                firstWorksheet.Cells[4, 2].Value = "Added data in Cell B4";

                //save the changes
                excelPackage.Save();
            }

        }

        

    }
}
