using log4net;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

[assembly: log4net.Config.XmlConfigurator(Watch=true)]

namespace ExcelWriter
{
    class ExcelWorker
    {
        #region ExcelWorker data members
        private static readonly ILog _log = LogManager.GetLogger("ExcelWriter.log");
        private static readonly int[] _doublePositions = { 6, 7, 8, 9, 10, 11, 12, 13, 14, 19, 20, 21, 22 };
        private static readonly int[] _skipPositions = { 4, 5, 15, 16, 17, 18 };
        private static readonly int _maxColumn = 22;
        private static int _startLineOffset
        {
            get
            {
                return Properties.Settings.Default.LastUsedRow;
            }
            set
            {
                Properties.Settings.Default.LastUsedRow = value;
                Properties.Settings.Default.Save();
            }
        }
        #endregion

        #region Main entry point of program
        static void Main(string[] args)
        {
            /*// Test paths
            string excelFilePath = "C:\\Users\\3D Infotech.3DCA-LY520-12\\Desktop\\Example2.xlsx";
            string txtFilePath = "C:\\Users\\3D Infotech.3DCA-LY520-12\\Desktop\\Test.txt";*/

            // Command line argument section
            #region Command line

            _log.Info($"Number of arguments passed for this execution: {args.Length}");
            for (int i = 0; i < args.Length; ++i)
            {
                _log.Info($"Arg #{i + 1}: {args[i]}");
            }
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

            // Should be at least 2 arguments which will be the paths to the two files
            string excelFilePath = args[0];
            string txtFilePath = args[1];

            // This is to capture the optional third parameter
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
            string[] lines;
            try
            {
                lines = File.ReadAllLines(txtFilePath);
            }
            catch (Exception ex)
            {
                _log.Error($"Exception thrown when reading text file: {ex.Message.ToString()}");
                return;
            }

            #endregion

            //create a fileinfo object of an excel file on the disk (file must exist)
            FileInfo file;
            try
            {
                file = new FileInfo(excelFilePath);
            }
            catch (Exception ex)
            {
                _log.Error($"Exception thrown when creating file object of excel file. {ex.Message.ToString()}");
                return;
            }

            #region Create Excel library object
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

                // Find first unused row with no data entered
                while (firstWorksheet.Cells[_startLineOffset, 2].Value != null)
                {
                    ++_startLineOffset;
                }

                _log.Info($"Writing to worksheet, starting at row: {_startLineOffset}");

                // Parse Strings and input data
                #region 
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
                #endregion

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
            #endregion
        }
        #endregion
    }
}
