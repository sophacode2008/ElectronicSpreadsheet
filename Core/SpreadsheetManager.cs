using System;
using System.Collections.Generic;
using ElectronicSpreadsheet.Models;
using Cell = ElectronicSpreadsheet.Models.Cell;

namespace ElectronicSpreadsheet.Core
{
    public class SpreadsheetManager
    {
        private readonly Dictionary<string, Cell> _cells;
        private readonly Calculator _calculator;

        public int CountRow { get; set; }
        public int CountColumn { get; set; }

        public SpreadsheetManager(int rows, int cols)
        {
            CountRow = rows;
            CountColumn = cols;
            _cells = new Dictionary<string, Cell>();
            _calculator = new Calculator();

            InitializeGrid();
        }

        public void InitializeGrid()
        {
            _cells.Clear();
            for (int row = 1; row <= CountRow; row++)
            {
                for (int col = 0; col < CountColumn; col++)
                {
                    string cellName = GetColumnName(col) + row;
                    _cells[cellName] = new Cell(cellName);
                }
            }
        }

        public Cell GetCell(string name)
        {
            if (_cells.TryGetValue(name, out Cell cell))
            {
                return cell;
            }
            return null;
        }

        public void SetCellExpression(string name, string expression)
        {
            if (_cells.ContainsKey(name))
            {
                _cells[name].Expression = expression;
            }
        }

        public double GetCellValue(string name)
        {
            if (_cells.TryGetValue(name, out Cell cell))
            {
                if (cell.HasError)
                {
                    return 0.0;
                }
                return cell.Value;
            }
            return 0.0;
        }

        public void EvaluateAll()
        {
            // REset errors before evaluation
            foreach (Cell cell in _cells.Values)
            {
                cell.HasError = false;
                cell.ErrorMessage = string.Empty;
            }

            foreach (Cell cell in _cells.Values)
            {
                if (string.IsNullOrWhiteSpace(cell.Expression))
                {
                    cell.Value = 0.0;
                    continue;
                }

                try
                {
                    cell.Value = _calculator.Evaluate(cell.Expression, GetCellValue);
                }
                catch (Exception ex)
                {
                    cell.HasError = true;
                    cell.ErrorMessage = ex.Message;
                    cell.Value = 0.0;
                }
            }
        }

        // Conversion of column index to letter (0 -> A, 1 -> B, etc.)
        public static string GetColumnName(int index)
        {
            int dividend = index + 1;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}