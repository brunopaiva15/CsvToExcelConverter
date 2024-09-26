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
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Aucun fichier spécifié.");
                return;
            }

            string csvFilePath = args[0];
            string excelFilePath = Path.ChangeExtension(csvFilePath, ".xlsx");

            // Demande à l'utilisateur s'il souhaite créer un tableau
            Console.WriteLine("Voulez-vous créer un tableau dans Excel ? (o/n)");
            bool createTable = Console.ReadLine().Trim().ToLower() == "o";

            // Demande si le fichier CSV contient des en-têtes
            Console.WriteLine("Votre fichier CSV contient-il des en-têtes ? (o/n)");
            bool hasHeaders = Console.ReadLine().Trim().ToLower() == "o";

            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    int currentRow = 1;

                    using (TextFieldParser parser = new TextFieldParser(csvFilePath, Encoding.Default))
                    {
                        // Détecter automatiquement le séparateur
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(DetectSeparator(File.ReadAllLines(csvFilePath, Encoding.Default)[0]).ToString());
                        parser.HasFieldsEnclosedInQuotes = true; // Gérer les valeurs entre guillemets

                        // Gérer les en-têtes
                        if (hasHeaders && !parser.EndOfData)
                        {
                            var headers = parser.ReadFields();
                            for (int j = 0; j < headers.Length; j++)
                            {
                                worksheet.Cell(currentRow, j + 1).Value = headers[j];
                            }
                            currentRow++; // Passer à la ligne suivante après les en-têtes
                        }

                        // Lire les lignes restantes
                        while (!parser.EndOfData)
                        {
                            var values = parser.ReadFields();

                            // Vérifier s'il y a des données valides sur la ligne
                            if (values == null || values.Length == 0 || values.All(string.IsNullOrWhiteSpace))
                            {
                                continue; // Ignorer les lignes vides
                            }

                            for (int j = 0; j < values.Length; j++)
                            {
                                var cell = worksheet.Cell(currentRow, j + 1);

                                // Tenter de détecter et de convertir les types
                                if (int.TryParse(values[j], out int intValue))
                                {
                                    cell.Value = intValue; // Stocker en tant qu'entier
                                }
                                else if (double.TryParse(values[j], NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleValue))
                                {
                                    cell.Value = doubleValue; // Stocker en tant que nombre à virgule flottante
                                }
                                else if (DateTime.TryParse(values[j], out DateTime dateValue))
                                {
                                    cell.Value = dateValue; // Stocker en tant que date
                                }
                                else
                                {
                                    cell.Value = values[j]; // Stocker en tant que texte
                                }
                            }

                            currentRow++; // Passer à la ligne suivante
                        }
                    }

                    // Si l'utilisateur souhaite créer un tableau
                    if (createTable)
                    {
                        var range = worksheet.RangeUsed();
                        range.CreateTable();
                    }

                    workbook.SaveAs(excelFilePath);
                }

                System.Diagnostics.Process.Start(excelFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erreur lors de la conversion : " + ex.Message);
            }
        }

        // Méthode pour détecter le séparateur le plus probable
        static char DetectSeparator(string line)
        {
            char[] possibleSeparators = { ',', ';', '\t', '|' };
            return possibleSeparators.OrderByDescending(sep => line.Count(c => c == sep)).First();
        }
    }
}
