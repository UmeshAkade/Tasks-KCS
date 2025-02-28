using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;



// Class to store input and output file paths from the config file
class Config
{
    public string? InputFilePath { get; set; }
    public string? OutputFilePath { get; set; }
}



class Program
{
    static void Main()
    {
        // Define the path of the config file
        string configFilePath = "config.json";

        // Check if the config file exists, if not, exit the program
        if (!File.Exists(configFilePath))
        {
            Console.WriteLine("Config file not found.");
            return;
        }

        // Read the config file and convert it into an object
        string jsonConfig = File.ReadAllText(configFilePath);
        var config = JsonSerializer.Deserialize<Config>(jsonConfig);

        // Ensure config is not null before accessing properties
        if (config == null || string.IsNullOrEmpty(config.InputFilePath) || string.IsNullOrEmpty(config.OutputFilePath))
        {
            Console.WriteLine("Invalid file paths in config file.");
            return;
        }

        // Create a dictionary to store PAN numbers and their masked versions
        Dictionary<string, string> panMapping = new Dictionary<string, string>();
        int counter = 1; // This counter helps in generating unique masked PAN numbers

        // Open the Excel file that needs to be processed
        using (var workbook = new XLWorkbook(config.InputFilePath))
        {
            var worksheet = workbook.Worksheet(1); // Get the first sheet of the Excel file
            var lastRowUsed = worksheet.LastRowUsed(); // Find the last used row in the sheet
            int lastRow = lastRowUsed != null ? lastRowUsed.RowNumber() : 1;

            // Loop through all the rows starting from the second row (ignoring headers)
            for (int row = 2; row <= lastRow; row++)
            {
                // Read the PAN number from the second column
                string pan = worksheet.Cell(row, 2).GetString().Trim();

                // Check if the PAN number is in the correct format using a pattern
                if (Regex.IsMatch(pan, @"^[A-Z]{5}[0-9]{4}[A-Z]$"))
                {
                    // If the PAN is valid and not already stored, create a masked version
                    if (!panMapping.ContainsKey(pan))
                    {
                        string newNumber = counter.ToString("D4"); // Generate a 4-digit unique number
                        panMapping[pan] = "XXXXX" + newNumber + "X"; // Masked PAN format
                        counter++; // Increase the counter for the next unique PAN
                    }
                    // Replace the original PAN with the masked version
                    worksheet.Cell(row, 2).Value = panMapping[pan];
                }
                else
                {
                    // If the PAN is not valid, mark it as "Invalid PAN"
                    worksheet.Cell(row, 2).Value = "Invalid PAN";
                }
            }
            // Save the updated Excel file with masked PAN numbers
            workbook.SaveAs(config.OutputFilePath);
        }
        // Print a message to confirm that the file has been processed successfully
        Console.WriteLine("Processed file saved as " + config.OutputFilePath);
    }
}

