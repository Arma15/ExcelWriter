using log4net;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

[assembly: log4net.Config.XmlConfigurator(Watch=true)]

namespace ExcelWriter
{
    class Program
    {
        private static readonly ILog _log = LogManager.GetLogger("ExcelWriter.log");
        private static readonly int[] _doublePositions = { 6, 7, 8, 9, 10, 11, 12, 13, 14, 19, 20, 21, 22 };
        private static readonly int[] _skipPositions = { 4, 5, 15, 16, 17, 18 };
        private static readonly int _maxColumn = 22;
        private static int _startLineOffset = 15000;

        static void Main(string[] args)
        {
            // Test paths
            string excelFilePath = "C:\\Users\\3D Infotech.3DCA-LY520-12\\Desktop\\Example.xlsx";
            string txtFilePath = "C:\\Users\\3D Infotech.3DCA-LY520-12\\Desktop\\Test.txt";

            // Command line argument section
            #region Command line
            
            if (args.Length < 1)
            {
               _log.Error("No arguments passed.");
                return;
            }

            if (args.Length < 2)
            {
                _log.Error("Only one argument passed.");
                return;
            }

            if (args.Length < 3)
            {
                excelFilePath = args[0];
                txtFilePath = args[1];
            }

            if (args.Length > 3)
            {
                if (Int32.TryParse(args[2], out int num))
                {
                    _startLineOffset = num;
                }
                else
                {
                    _log.Error("Third parameter is not an integer value.");
                }
            }

            if (!File.Exists(args[0]))
            {
                _log.Error($"File does not exist or path invalid: {args[0]}");
                return;
            }

            if (!File.Exists(args[1]))
            {
                _log.Error($"File does not exist or path invalid: {args[1]}");
                return;
            }
            #endregion Command line


            // Read in text file section
            #region Text File

            // Open text file and read all lines
            string[] lines = File.ReadAllLines(txtFilePath);
        
            #endregion

            //create a fileinfo object of an excel file on the disk (file must exist)
            FileInfo file = new FileInfo(excelFilePath);

            //create a new Excel package from the file
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                if (excelPackage.Workbook.Worksheets[1] == null)
                {
                    _log.Error("Worksheet does not exist..");
                    return;
                }

                //Get a WorkSheet by index. Note that EPPlus indexes are base 1, not base 0!
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1];

                // Find first unused row with no date entered
                while (firstWorksheet.Cells[_startLineOffset, 2].Value != null)
                {
                    ++_startLineOffset;
                }

                _log.Info($"Writing to worksheet, starting at row: {_startLineOffset}");

                // Parse String
                for (int textFileLine = 0; textFileLine < lines.Length; ++textFileLine)
                {
                    string[] words = lines[textFileLine].Split(',');
                    for (int wordsIndex = 0, column = 1; wordsIndex < words.Length && column <= _maxColumn; ++wordsIndex, ++column)
                    {
                        while (_skipPositions.Contains(column))
                        {
                            ++column;
                        }

                        if (_doublePositions.Contains(column))
                        {
                            // If it is a position that is suppose to be a double
                            if (Double.TryParse(words[wordsIndex], out double number))
                            {
                                firstWorksheet.Cells[_startLineOffset + textFileLine, column].Value = number;
                                continue;
                            }
                            else
                            {
                                _log.Error($"Index: {wordsIndex} was expected to be a double value, inserted instead as a text value");
                            }
                        }
                        
                        // Default the rest to text values
                        firstWorksheet.Cells[_startLineOffset + textFileLine, column].Value = words[wordsIndex];
                        
                    }
                }

                try
                {
                    //save the changes
                    excelPackage.Save();
                }
                catch (Exception ex)
                {
                    _log.Error($"Exception when saving excel document: {ex.Message.ToString()}");
                }
            }

        }


    }
}
