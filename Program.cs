using System;
using System.Globalization;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using ClosedXML.Excel;
using System.Linq;
using System.Text;

namespace CsvToExcelConverter
{
    internal class Program
    {
        // Maximum number of rows per Excel sheet
        const int MaxRowsPerSheet = 1048576;

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    Console.WriteLine("No file specified.");
                    Quit();
                    return;
                }

                string csvFilePath = args[0];
                string excelFilePath = Path.ChangeExtension(csvFilePath, ".xlsx");

                // Check if the Excel file exists and generate a new name if necessary
                excelFilePath = GetUniqueFilePath(excelFilePath);

                // Count the total number of lines in the CSV file
                int totalLines = CountLinesInCsv(csvFilePath);

                // Check if the number of lines exceeds the limit
                if (totalLines > MaxRowsPerSheet)
                {
                    Console.WriteLine($"Error: CSV file contains more than {MaxRowsPerSheet} rows. Limit exceeded.");
                    Quit();
                    return;
                }

                // Ask the user if they want to create a table
                Console.Write("Do you want to create a table in Excel? (y/n): ");
                bool createTable = Console.ReadLine().Trim().ToLower() == "y";

                // Determine the total number of steps
                int stepCount = createTable ? 3 : 2;

                bool hasHeaders = false;
                if (createTable)
                {
                    // Display the first line of the CSV to help determine if they are headers
                    string firstLine = GetFirstLineOfCsv(csvFilePath);
                    Console.WriteLine("\nFirst line of CSV:\n");
                    Console.WriteLine(firstLine);
                    Console.WriteLine(); // Empty line at the bottom

                    // Ask if the CSV file contains headers
                    Console.Write("Does your CSV file contain headers? (y/n): ");
                    hasHeaders = Console.ReadLine().Trim().ToLower() == "y";
                }

                Console.WriteLine($"Total number of lines to be processed: {totalLines}");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    int currentRow = 1;
                    int processedLines = 0;

                    Console.WriteLine($"\nStep 1/{stepCount}: Reading data...");

                    using (TextFieldParser parser = new TextFieldParser(csvFilePath, Encoding.Default))
                    {
                        // Automatically detect the separator
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(DetectSeparator(File.ReadAllLines(csvFilePath, Encoding.Default)[0]).ToString());
                        parser.HasFieldsEnclosedInQuotes = true; // Handle values enclosed in quotes

                        // Handle headers
                        if (hasHeaders && !parser.EndOfData)
                        {
                            var headers = parser.ReadFields();
                            for (int j = 0; j < headers.Length; j++)
                            {
                                worksheet.Cell(currentRow, j + 1).Value = headers[j];
                            }
                            currentRow++; // Move to the next line after headers
                        }

                        // Read and process lines one by one
                        while (!parser.EndOfData && currentRow <= MaxRowsPerSheet)
                        {
                            var values = parser.ReadFields();

                            // Check if there is valid data on the line
                            if (values == null || values.Length == 0 || values.All(string.IsNullOrWhiteSpace))
                            {
                                continue; // Skip empty lines
                            }

                            for (int j = 0; j < values.Length; j++)
                            {
                                var cell = worksheet.Cell(currentRow, j + 1);

                                // Try to detect and convert types
                                if (int.TryParse(values[j], out int intValue))
                                {
                                    cell.Value = intValue; // Store as integer
                                }
                                else if (double.TryParse(values[j], NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleValue))
                                {
                                    cell.Value = doubleValue; // Store as floating-point number
                                }
                                else if (DateTime.TryParse(values[j], out DateTime dateValue))
                                {
                                    cell.Value = dateValue; // Store as date
                                }
                                else
                                {
                                    cell.Value = values[j]; // Store as text
                                }
                            }

                            currentRow++; // Move to the next line
                            processedLines++;

                            // Update the progress bar
                            if (processedLines % 1000 == 0 || processedLines == totalLines)
                            {
                                Console.Write($"\rProgress: Reading data: {processedLines * 100 / totalLines}%");
                            }
                        }
                    }

                    // If the user wants to create a table
                    if (createTable)
                    {
                        Console.WriteLine($"\n\nStep 2/{stepCount}: Table creation...");
                        var range = worksheet.RangeUsed();
                        range.CreateTable();
                    }

                    Console.WriteLine($"\n\nStep {stepCount}/{stepCount}: Saving the file...");
                    workbook.SaveAs(excelFilePath);
                }

                System.Diagnostics.Process.Start(excelFilePath);
                Console.WriteLine("\nConversion complete.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Conversion error: " + ex.Message);
                Quit();
            }
        }

        // Method to get a unique file path
        static string GetUniqueFilePath(string filePath)
        {
            int count = 1;
            string fileNameOnly = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);
            string path = Path.GetDirectoryName(filePath);
            string newFullPath = filePath;

            // Add a suffix as long as a file with the same name exists
            while (File.Exists(newFullPath))
            {
                string tempFileName = $"{fileNameOnly}_{count++}";
                newFullPath = Path.Combine(path, tempFileName + extension);
            }

            return newFullPath;
        }

        // Method to get the first line of the CSV file
        static string GetFirstLineOfCsv(string filePath)
        {
            using (var reader = new StreamReader(filePath, Encoding.Default))
            {
                return reader.ReadLine();
            }
        }

        // Method to count the lines in the CSV file
        static int CountLinesInCsv(string filePath)
        {
            int lineCount = 0;
            using (var reader = new StreamReader(filePath))
            {
                while (reader.ReadLine() != null)
                {
                    lineCount++;
                }
            }
            return lineCount;
        }

        // Method to detect the most likely separator
        static char DetectSeparator(string line)
        {
            char[] possibleSeparators = { ',', ';', '\t', '|' };
            return possibleSeparators.OrderByDescending(sep => line.Count(c => c == sep)).First();
        }

        // Method to quit the application
        static void Quit()
        {
            Console.WriteLine("Press a key to exit...");
            Console.ReadKey();
        }
    }
}
