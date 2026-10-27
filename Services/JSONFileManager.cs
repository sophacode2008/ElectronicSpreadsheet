using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Cell = ElectronicSpreadsheet.Models.Cell; // Avoid conflict with MAUI Cell

namespace ElectronicSpreadsheet.Services
{
    public class JSONFileManager
    {
        // Method for saving (serializing) the dictionary of cells to a file
        public void Serialize(string filePath, Dictionary<string, Cell> cells)
        {
            // Configuration for beautiful formatting (with indents)
            JsonSerializerOptions options = new JsonSerializerOptions
            {
                WriteIndented = true
            };

            string jsonString = JsonSerializer.Serialize(cells, options);
            File.WriteAllText(filePath, jsonString);
        }

        // Method for loading (deserializing) cells from a file
        public Dictionary<string, Cell> Deserialize(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return new Dictionary<string, Cell>();
            }

            string jsonString = File.ReadAllText(filePath);
            Dictionary<string, Cell> restoredCells = JsonSerializer.Deserialize<Dictionary<string, Cell>>(jsonString);

            return restoredCells ?? new Dictionary<string, Cell>();
        }
    }
}