using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

// Class to store file paths from the config file
class Config
{
    public string? InputFilePath { get; set; } // Path of the input Excel file
    public string? OutputFilePath { get; set; } // Path where output Excel file will be saved
}

class DataMasking
{
    private static Random random = new Random();

    // Function to create a random PAN (Permanent Account Number) in the format ABCDE1234F
    public static string GenerateRandomPAN()
    {
        string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string numbers = "0123456789";
        return new string(new char[] {
            letters[random.Next(26)], letters[random.Next(26)], letters[random.Next(26)], letters[random.Next(26)], letters[random.Next(26)],
            numbers[random.Next(10)], numbers[random.Next(10)], numbers[random.Next(10)], numbers[random.Next(10)],
            letters[random.Next(26)]
        });
    }

    // Function to create a random 8-digit account number
    public static string GenerateRandomAccountNumber()
    {
        char[] digits = new char[8];
        for (int i = 0; i < 8; i++)
            digits[i] = (char)('0' + random.Next(10)); // Generates a random digit (0-9)
        return new string(digits);
    }
}

class Program
{
    static void Main()
    {
        // Define the config file path
        string configFilePath = "config.json";

        // Check if the config file exists; if not, show an error and stop the program
        if (!File.Exists(configFilePath))
        {
            Console.WriteLine("Config file not found.");
            return;
        }

        // Read the config file and convert it into a Config object
        string jsonConfig = File.ReadAllText(configFilePath);
        var config = JsonSerializer.Deserialize<Config>(jsonConfig);

        // Check if file paths are missing in the config file
        if (config == null || string.IsNullOrEmpty(config.InputFilePath) || string.IsNullOrEmpty(config.OutputFilePath))
        {
            Console.WriteLine("Invalid file paths in config file.");
            return;
        }

        // Dictionaries to store masked values
        Dictionary<string, string> panMapping = new Dictionary<string, string>(); // For PAN numbers
        Dictionary<string, string> accountMapping = new Dictionary<string, string>(); // For account numbers
        Dictionary<string, string> nameMapping = new Dictionary<string, string>(); // For names

        // List of Indian names to replace real names
        List<string> indianNames = new List<string> { "Amit", "Rahul", "Priya", "Suresh", "Anjali", "Vikram", "Pooja", "Ravi", "Neha", "Raj" };
        Random random = new Random();

        // Open the Excel file
        using (var workbook = new XLWorkbook(config.InputFilePath))
        {
            var worksheet = workbook.Worksheet(1); // Select the first worksheet
            var lastRowUsed = worksheet.LastRowUsed(); // Get the last row that has data
            int lastRow = lastRowUsed != null ? lastRowUsed.RowNumber() : 1;

            // Start reading from the second row (assuming the first row has column headers)
            for (int row = 2; row <= lastRow; row++)
            {
                // Read Name, PAN, and Account Number from Excel
                string name = worksheet.Cell(row, 1).GetString().Trim(); 
                string pan = worksheet.Cell(row, 2).GetString().Trim(); 
                string account = worksheet.Cell(row, 3).GetString().Trim(); 

                // Mask PAN number (if valid) or mark as invalid
                if (!panMapping.ContainsKey(pan))
                {
                    panMapping[pan] = Regex.IsMatch(pan, @"^[A-Z]{5}[0-9]{4}[A-Z]$") ? DataMasking.GenerateRandomPAN() : "Invalid PAN";
                }
                worksheet.Cell(row, 2).Value = panMapping[pan]; // Update Excel with masked PAN

                // Replace account number with a random one
                if (!accountMapping.ContainsKey(account))
                {
                    accountMapping[account] = DataMasking.GenerateRandomAccountNumber();
                }
                worksheet.Cell(row, 3).Value = accountMapping[account]; // Update Excel with new account number

                // Replace name with a randomly chosen Indian name
                if (!nameMapping.ContainsKey(name))
                {
                    nameMapping[name] = indianNames[random.Next(indianNames.Count)];
                }
                worksheet.Cell(row, 1).Value = nameMapping[name]; // Update Excel with new name
            }
            
            // Save the modified Excel file
            workbook.SaveAs(config.OutputFilePath);
        }
        Console.WriteLine("Processed file saved as " + config.OutputFilePath);
    }
}
