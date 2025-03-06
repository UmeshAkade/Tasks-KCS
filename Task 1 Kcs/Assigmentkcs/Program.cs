using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

// Class to read the configuration file containing input and output file paths
class Config
{
    public string? InputFilePath { get; set; }
    public string? OutputFilePath { get; set; }
}

// Class to handle data masking and generating random PAN and account numbers
class DataMasking
{
    private static Random random = new Random();

    // Generates a random PAN number following the standard format (AAAAA9999A)
    public static string GenerateRandomPAN()
    {
        string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string firstFive = new string(new char[] {
            letters[random.Next(26)], letters[random.Next(26)], letters[random.Next(26)], 
            letters[random.Next(26)], letters[random.Next(26)]
        });
        string fourDigits = random.Next(0, 10000).ToString("D4");
        string lastLetter = letters[random.Next(26)].ToString();
        return firstFive + fourDigits + lastLetter;
    }

    // Generates a random 8-digit account number
    public static string GenerateRandomAccountNumber()
    {
        return random.Next(10000000, 100000000).ToString();
    }
}

class Program
{
    static void Main()
    {
        string configFilePath = "config.json";

        // Check if the configuration file exists
        if (!File.Exists(configFilePath))
        {
            Console.WriteLine("Config file not found.");
            return;
        }

        // Read and deserialize the JSON configuration file
        string jsonConfig = File.ReadAllText(configFilePath);
        var config = JsonSerializer.Deserialize<Config>(jsonConfig);

        // Validate if input and output file paths are provided
        if (config == null || string.IsNullOrEmpty(config.InputFilePath) || string.IsNullOrEmpty(config.OutputFilePath))
        {
            Console.WriteLine("Invalid file paths in config file.");
            return;
        }

        // Dictionaries to store mappings for PAN, account numbers, and names to ensure consistency
        Dictionary<string, string> panMapping = new Dictionary<string, string>();
        Dictionary<string, string> accountMapping = new Dictionary<string, string>();
        Dictionary<string, string> nameMapping = new Dictionary<string, string>();

        // List of sample Indian names for anonymization
        List<string> indianNames = new List<string> { "Amit", "Rahul", "Priya", "Suresh", "Anjali", "Vikram", "Pooja", "Ravi", "Neha", "Raj" };
        Random random = new Random();

        // Open the Excel file using ClosedXML
        using (var workbook = new XLWorkbook(config.InputFilePath))
        {
            var worksheet = workbook.Worksheet(1); // Access the first worksheet
            var headerRow = worksheet.Row(1); // Get the first row (header row)
            int lastColumn = headerRow.LastCellUsed()?.Address.ColumnNumber ?? 0;

            int nameColumn = 0, panColumn = 0, accountColumn = 0;

            // Debug: Print available headers to identify relevant columns
            Console.WriteLine("Available headers:");
            for (int col = 1; col <= lastColumn; col++)
            {
                string header = headerRow.Cell(col).GetString().Trim();
                Console.WriteLine($"Column {col}: {header}");

                string lowerHeader = header.ToLower();

                // Identify columns based on header names
                if (lowerHeader.Contains("name"))
                    nameColumn = col;
                else if (lowerHeader.Contains("pan"))
                    panColumn = col;
                else if (lowerHeader.Contains("account"))
                    accountColumn = col;
            }

            // If any required column is missing, stop processing
            if (nameColumn == 0 || panColumn == 0 || accountColumn == 0)
            {
                Console.WriteLine("One or more required columns were not found.");
                return;
            }

            var lastRowUsed = worksheet.LastRowUsed();
            int lastRow = lastRowUsed != null ? lastRowUsed.RowNumber() : 1;

            // Iterate through all rows (excluding header) and mask sensitive data
            for (int row = 2; row <= lastRow; row++)
            {
                string name = worksheet.Cell(row, nameColumn).GetString().Trim();
                string pan = worksheet.Cell(row, panColumn).GetString().Trim();
                string account = worksheet.Cell(row, accountColumn).GetString().Trim();

                // Mask PAN numbers if they follow the correct format
                if (!string.IsNullOrEmpty(pan))
                {
                    if (Regex.IsMatch(pan, @"^[A-Z]{5}[0-9]{4}[A-Z]$")) // Validate PAN format
                    {
                        if (!panMapping.ContainsKey(pan))
                            panMapping[pan] = DataMasking.GenerateRandomPAN();
                        worksheet.Cell(row, panColumn).Value = panMapping[pan];
                    }
                    else
                    {
                        worksheet.Cell(row, panColumn).Value = "Invalid PAN"; // Mark invalid PANs
                    }
                }

                // Mask account numbers
                if (!string.IsNullOrEmpty(account))
                {
                    if (!accountMapping.ContainsKey(account))
                        accountMapping[account] = DataMasking.GenerateRandomAccountNumber();
                    worksheet.Cell(row, accountColumn).Value = accountMapping[account];
                }

                // Mask names using a predefined list of names
                if (!string.IsNullOrEmpty(name))
                {
                    if (!nameMapping.ContainsKey(name))
                        nameMapping[name] = indianNames[random.Next(indianNames.Count)];
                    worksheet.Cell(row, nameColumn).Value = nameMapping[name];
                }
            }

            // Save the modified data to the output file
            workbook.SaveAs(config.OutputFilePath);
        }
        
        Console.WriteLine("Processed file saved as " + config.OutputFilePath);
    }
}
