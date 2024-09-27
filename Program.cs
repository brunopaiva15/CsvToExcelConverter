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
        // Nombre maximum de lignes par feuille Excel
        const int MaxRowsPerSheet = 1048576;

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    Console.WriteLine("Aucun fichier spécifié.");
                    Quit();
                    return;
                }

                string csvFilePath = args[0];
                string excelFilePath = Path.ChangeExtension(csvFilePath, ".xlsx");

                // Vérifier si le fichier Excel existe et générer un nouveau nom si nécessaire
                excelFilePath = GetUniqueFilePath(excelFilePath);

                // Compter le nombre total de lignes dans le fichier CSV
                int totalLines = CountLinesInCsv(csvFilePath);

                // Vérifier si le nombre de lignes dépasse la limite
                if (totalLines > MaxRowsPerSheet)
                {
                    Console.WriteLine($"Erreur : Le fichier CSV contient plus de {MaxRowsPerSheet} lignes. Limite dépassée.");
                    Quit();
                    return;
                }

                // Demande à l'utilisateur s'il souhaite créer un tableau
                Console.Write("Voulez-vous créer un tableau dans Excel ? (o/n) : ");
                bool createTable = Console.ReadLine().Trim().ToLower() == "o";

                // Déterminer le nombre total d'étapes
                int stepCount = createTable ? 3 : 2;

                bool hasHeaders = false;
                if (createTable)
                {
                    // Afficher la première ligne du CSV pour aider à déterminer si ce sont des en-têtes
                    string firstLine = GetFirstLineOfCsv(csvFilePath);
                    Console.WriteLine("\nPremière ligne du CSV :\n");
                    Console.WriteLine(firstLine);
                    Console.WriteLine(); // Ligne vide en bas

                    // Demande si le fichier CSV contient des en-têtes
                    Console.Write("Votre fichier CSV contient-il des en-têtes ? (o/n) : ");
                    hasHeaders = Console.ReadLine().Trim().ToLower() == "o";
                }

                Console.WriteLine($"Nombre total de lignes à traiter : {totalLines}");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    int currentRow = 1;
                    int processedLines = 0;

                    Console.WriteLine($"\nÉtape 1/{stepCount} : Lecture des données...");

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

                        // Lire et traiter les lignes une par une
                        while (!parser.EndOfData && currentRow <= MaxRowsPerSheet)
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
                            processedLines++;

                            // Mise à jour de la barre de progression
                            if (processedLines % 1000 == 0 || processedLines == totalLines)
                            {
                                Console.Write($"\rProgression : Lecture des données : {processedLines * 100 / totalLines}%");
                            }
                        }
                    }

                    // Si l'utilisateur souhaite créer un tableau
                    if (createTable)
                    {
                        Console.WriteLine($"\n\nÉtape 2/{stepCount} : Création du tableau...");
                        var range = worksheet.RangeUsed();
                        range.CreateTable();
                    }

                    Console.WriteLine($"\n\nÉtape {stepCount}/{stepCount} : Sauvegarde du fichier...");
                    workbook.SaveAs(excelFilePath);
                }

                System.Diagnostics.Process.Start(excelFilePath);
                Console.WriteLine("\nConversion terminée.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erreur lors de la conversion : " + ex.Message);
                Quit();
            }
        }

        // Méthode pour récupérer un chemin de fichier unique
        static string GetUniqueFilePath(string filePath)
        {
            int count = 1;
            string fileNameOnly = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);
            string path = Path.GetDirectoryName(filePath);
            string newFullPath = filePath;

            // Ajouter un suffixe tant qu'un fichier du même nom existe
            while (File.Exists(newFullPath))
            {
                string tempFileName = $"{fileNameOnly}_{count++}";
                newFullPath = Path.Combine(path, tempFileName + extension);
            }

            return newFullPath;
        }

        // Méthode pour récupérer la première ligne du fichier CSV
        static string GetFirstLineOfCsv(string filePath)
        {
            using (var reader = new StreamReader(filePath, Encoding.Default))
            {
                return reader.ReadLine();
            }
        }

        // Méthode pour compter les lignes dans le fichier CSV
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

        // Méthode pour détecter le séparateur le plus probable
        static char DetectSeparator(string line)
        {
            char[] possibleSeparators = { ',', ';', '\t', '|' };
            return possibleSeparators.OrderByDescending(sep => line.Count(c => c == sep)).First();
        }

        // Méthode pour quitter l'application
        static void Quit()
        {
            Console.WriteLine("Appuyer sur une touche pour quitter...");
            Console.ReadKey();
        }
    }
}
