using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Maui.Storage;
using ElectronicSpreadsheet.Models;
using ElectronicSpreadsheet.Services;

namespace ElectronicSpreadsheet.Core
{
    // Controller class encapsulating validation, convergent evaluation, resizing, and file I/O
    public class SpreadsheetController
    {
        private const int InitialRowCount = 16;
        private const int InitialColumnCount = 10;
        private const int MinGridDimension = 1;
        private const int MinEvaluationPasses = 4;
        private const double ValueComparisonTolerance = 1e-9;
        private const string SaveFileName = "spreadsheet.json";

        private readonly SpreadsheetManager _manager;
        private readonly JSONFileManager _fileManager;

        public SpreadsheetController()
        {
            _manager = new SpreadsheetManager(InitialRowCount, InitialColumnCount);
            _fileManager = new JSONFileManager();
        }

        public int RowCount
        {
            get
            {
                return _manager.CountRow;
            }
        }

        public int ColumnCount
        {
            get
            {
                return _manager.CountColumn;
            }
        }

        public string GetColumnHeader(int colIndex)
        {
            return SpreadsheetManager.GetColumnName(colIndex);
        }

        public string GetCellExpression(string cellName)
        {
            Cell cell = _manager.GetCell(cellName);
            return cell != null ? cell.Expression : string.Empty;
        }

        // Updates a cell's expression and immediately recalculates the entire dependency chain
        public void UpdateCellExpression(string cellName, string rawInput)
        {
            bool isNameEmpty = string.IsNullOrWhiteSpace(cellName);
            if (isNameEmpty)
            {
                return;
            }

            Cell targetCell = _manager.GetCell(cellName);
            if (targetCell != null)
            {
                targetCell.Expression = NormalizeExpression(rawInput);
                RecalculateAll();
            }
        }

        // Performs multi-pass convergent evaluation so chains of any length update at once
        public void RecalculateAll()
        {
            Dictionary<string, Cell> allCells = GetInternalCells();
            Dictionary<string, string> isolated = new Dictionary<string, string>();
            int activeCellCount = 0;

            foreach (KeyValuePair<string, Cell> pair in allCells)
            {
                pair.Value.Value = 0.0;
                pair.Value.HasError = false;
                pair.Value.ErrorMessage = string.Empty;

                bool hasExpr = !string.IsNullOrWhiteSpace(pair.Value.Expression);
                if (hasExpr)
                {
                    ++activeCellCount;
                }
            }

            foreach (KeyValuePair<string, Cell> pair in allCells)
            {
                bool hasExpr = !string.IsNullOrWhiteSpace(pair.Value.Expression);
                if (hasExpr)
                {
                    string preError = ResolvePreEvaluationError(pair.Key, new HashSet<string>());
                    bool shouldIsolate = !string.IsNullOrEmpty(preError);
                    if (shouldIsolate)
                    {
                        isolated[pair.Key] = pair.Value.Expression;
                    }
                }
            }

            foreach (KeyValuePair<string, string> pair in isolated)
            {
                Cell cell = _manager.GetCell(pair.Key);
                if (cell != null)
                {
                    cell.Expression = string.Empty;
                }
            }

            RunConvergentPasses(allCells, activeCellCount);

            foreach (KeyValuePair<string, string> pair in isolated)
            {
                Cell cell = _manager.GetCell(pair.Key);
                if (cell != null)
                {
                    cell.Expression = pair.Value;
                    cell.HasError = true;
                }
            }
        }

        private void RunConvergentPasses(Dictionary<string, Cell> allCells, int activeCellCount)
        {
            int maxPasses = Math.Max(activeCellCount + 2, MinEvaluationPasses);
            Dictionary<string, double> previousValues = new Dictionary<string, double>();

            for (int pass = 0; pass < maxPasses; ++pass)
            {
                foreach (KeyValuePair<string, Cell> pair in allCells)
                {
                    previousValues[pair.Key] = pair.Value.Value;
                }

                _manager.EvaluateAll();

                bool anyChanged = HasAnyCellValueChanged(allCells, previousValues);
                bool canStopEarly = !anyChanged && pass >= MinEvaluationPasses - 1;
                if (canStopEarly)
                {
                    break;
                }
            }
        }

        private bool HasAnyCellValueChanged(
            Dictionary<string, Cell> allCells,
            Dictionary<string, double> previousValues)
        {
            foreach (KeyValuePair<string, Cell> pair in allCells)
            {
                double oldVal = previousValues[pair.Key];
                double newVal = pair.Value.Value;
                bool nanDiff = double.IsNaN(oldVal) != double.IsNaN(newVal);
                bool infDiff = double.IsInfinity(oldVal) != double.IsInfinity(newVal);
                bool valDiff = Math.Abs(oldVal - newVal) > ValueComparisonTolerance;

                if (nanDiff || infDiff || valDiff)
                {
                    return true;
                }
            }

            return false;
        }

        public bool IsCellInError(string cellName)
        {
            string errorCode = ResolveCellError(cellName, new HashSet<string>());
            return !string.IsNullOrEmpty(errorCode);
        }

        public string GetCellDisplayText(string cellName, bool showFormulas)
        {
            Cell cell = _manager.GetCell(cellName);
            bool isEmpty = cell == null || string.IsNullOrWhiteSpace(cell.Expression);
            if (isEmpty)
            {
                return string.Empty;
            }

            if (showFormulas)
            {
                return cell.Expression;
            }

            string errorCode = ResolveCellError(cellName, new HashSet<string>());
            bool hasError = !string.IsNullOrEmpty(errorCode);
            if (hasError)
            {
                return errorCode;
            }

            return cell.Value.ToString();
        }

        // Checks only structural, syntax, and circular errors prior to running ANTLR evaluation
        private string ResolvePreEvaluationError(string cellName, HashSet<string> visited)
        {
            bool isNameEmpty = string.IsNullOrWhiteSpace(cellName);
            if (isNameEmpty)
            {
                return string.Empty;
            }

            string upperName = cellName.ToUpperInvariant();
            Cell cell = _manager.GetCell(upperName);
            if (cell == null)
            {
                return "#REF!";
            }

            bool isExprEmpty = string.IsNullOrWhiteSpace(cell.Expression);
            if (isExprEmpty)
            {
                return string.Empty;
            }

            bool isNewVisit = visited.Add(upperName);
            if (!isNewVisit)
            {
                return "#REF!";
            }

            string expr = cell.Expression.Trim();
            bool isNa = expr.Equals("#N/A", StringComparison.OrdinalIgnoreCase)
                || expr.Equals("NA()", StringComparison.OrdinalIgnoreCase);
            if (isNa)
            {
                return "#N/A";
            }

            bool isUnknownName = HasUnknownIdentifier(expr);
            if (isUnknownName)
            {
                return "#NAME?";
            }

            bool isSyntaxBad = HasSyntaxError(expr);
            if (isSyntaxBad)
            {
                return "#VALUE!";
            }

            MatchCollection matches = Regex.Matches(expr, @"\b([A-Za-z]+[0-9]+)\b");
            foreach (Match match in matches)
            {
                string refName = match.Value.ToUpperInvariant();
                bool isMissingOrCycle = _manager.GetCell(refName) == null || visited.Contains(refName);
                if (isMissingOrCycle)
                {
                    return "#REF!";
                }

                HashSet<string> nextVisited = new HashSet<string>(visited);
                string refError = ResolvePreEvaluationError(refName, nextVisited);
                bool hasRefError = !string.IsNullOrEmpty(refError);
                if (hasRefError)
                {
                    return refError;
                }
            }

            return string.Empty;
        }

        // Resolves both pre-evaluation errors and runtime division-by-zero errors
        private string ResolveCellError(string cellName, HashSet<string> visited)
        {
            string preError = ResolvePreEvaluationError(cellName, new HashSet<string>());
            bool hasPreError = !string.IsNullOrEmpty(preError);
            if (hasPreError)
            {
                return preError;
            }

            string upperName = cellName.ToUpperInvariant();
            Cell cell = _manager.GetCell(upperName);
            bool isEmpty = cell == null || string.IsNullOrWhiteSpace(cell.Expression);
            if (isEmpty)
            {
                return string.Empty;
            }

            bool isNewVisit = visited.Add(upperName);
            if (!isNewVisit)
            {
                return "#REF!";
            }

            MatchCollection matches = Regex.Matches(cell.Expression, @"\b([A-Za-z]+[0-9]+)\b");
            foreach (Match match in matches)
            {
                string refName = match.Value.ToUpperInvariant();
                HashSet<string> nextVisited = new HashSet<string>(visited);
                string refError = ResolveCellError(refName, nextVisited);
                bool hasRefError = !string.IsNullOrEmpty(refError);
                if (hasRefError)
                {
                    return refError;
                }
            }

            bool isRuntimeDivZero = cell.HasError
                || double.IsInfinity(cell.Value)
                || double.IsNaN(cell.Value)
                || Regex.IsMatch(cell.Expression, @"(\/|\bmod\b|\bdiv\b)\s*0(\.0+)?\b", RegexOptions.IgnoreCase);
            if (isRuntimeDivZero)
            {
                return "#DIV/0!";
            }

            return string.Empty;
        }

        private string NormalizeExpression(string input)
        {
            bool isEmpty = string.IsNullOrWhiteSpace(input);
            if (isEmpty)
            {
                return string.Empty;
            }

            string trimmed = input.Trim();
            bool startsWithEquals = trimmed.StartsWith("=");
            if (startsWithEquals)
            {
                trimmed = trimmed.Substring(1).Trim();
            }

            string upperRefs = Regex.Replace(trimmed, @"\b([A-Za-z]+[0-9]+)\b", m => m.Value.ToUpperInvariant());
            return Regex.Replace(upperRefs, @"\b(mod|div)\b", m => m.Value.ToLowerInvariant(), RegexOptions.IgnoreCase);
        }

        private bool HasUnknownIdentifier(string expression)
        {
            MatchCollection words = Regex.Matches(expression, @"\b[A-Za-z_][A-Za-z0-9_]*\b");
            foreach (Match word in words)
            {
                string lower = word.Value.ToLowerInvariant();
                bool isOperatorKeyword = lower == "mod" || lower == "div";
                if (isOperatorKeyword)
                {
                    continue;
                }

                bool isCellRef = Regex.IsMatch(word.Value, @"^[A-Za-z]+[1-9][0-9]*$");
                int nextIdx = word.Index + word.Length;
                bool isFuncCall = nextIdx < expression.Length && expression[nextIdx] == '(';
                if (!isCellRef || isFuncCall)
                {
                    return true;
                }
            }

            return false;
        }

        private bool HasSyntaxError(string expression)
        {
            bool hasEmptyParens = Regex.IsMatch(expression, @"\(\s*\)");
            if (hasEmptyParens)
            {
                return true;
            }

            string pattern = @"\G\s*(\d+(?:\.\d+)?|\bmod\b|\bdiv\b|[A-Za-z_][A-Za-z0-9_]*|[\+\-\*\/\^\(\)])\s*";
            MatchCollection tokens = Regex.Matches(expression, pattern, RegexOptions.IgnoreCase);
            int matchedLen = 0;
            int parenDepth = 0;
            bool expectOperand = true;

            foreach (Match m in tokens)
            {
                if (m.Index != matchedLen)
                {
                    return true;
                }

                matchedLen += m.Length;
                string tok = m.Groups[1].Value.ToLowerInvariant();
                bool isBinOp = tok == "*" || tok == "/" || tok == "^" || tok == "mod" || tok == "div";

                if (tok == "(")
                {
                    ++parenDepth;
                }
                else if (tok == ")")
                {
                    --parenDepth;
                }

                bool invalidWhenExpectingOperand = expectOperand && (tok == ")" || isBinOp);
                if (parenDepth < 0 || invalidWhenExpectingOperand)
                {
                    return true;
                }

                bool invalidAfterOperand = !expectOperand && tok != ")" && !isBinOp && tok != "+" && tok != "-";
                if (invalidAfterOperand)
                {
                    return true;
                }

                expectOperand = tok == "(" || tok == "+" || tok == "-" || isBinOp;
            }

            return matchedLen != expression.Length || parenDepth != 0 || expectOperand;
        }

        public void AddRow()
        {
            ResizeGrid(_manager.CountRow + 1, _manager.CountColumn);
        }

        public bool DeleteRow()
        {
            bool canDelete = _manager.CountRow > MinGridDimension;
            if (!canDelete)
            {
                return false;
            }

            ResizeGrid(_manager.CountRow - 1, _manager.CountColumn);
            return true;
        }

        public void AddColumn()
        {
            ResizeGrid(_manager.CountRow, _manager.CountColumn + 1);
        }

        public bool DeleteColumn()
        {
            bool canDelete = _manager.CountColumn > MinGridDimension;
            if (!canDelete)
            {
                return false;
            }

            ResizeGrid(_manager.CountRow, _manager.CountColumn - 1);
            return true;
        }

        private void ResizeGrid(int newRows, int newCols)
        {
            Dictionary<string, string> saved = new Dictionary<string, string>();
            foreach (KeyValuePair<string, Cell> pair in GetInternalCells())
            {
                bool hasExpr = !string.IsNullOrWhiteSpace(pair.Value.Expression);
                if (hasExpr)
                {
                    saved[pair.Key] = pair.Value.Expression;
                }
            }

            _manager.CountRow = newRows;
            _manager.CountColumn = newCols;
            _manager.InitializeGrid();

            foreach (KeyValuePair<string, string> pair in saved)
            {
                Cell cell = _manager.GetCell(pair.Key);
                if (cell != null)
                {
                    cell.Expression = pair.Value;
                }
            }

            RecalculateAll();
        }

        public string SaveToFile()
        {
            string filePath = Path.Combine(FileSystem.AppDataDirectory, SaveFileName);
            _fileManager.Serialize(filePath, GetInternalCells());
            return filePath;
        }

        public bool LoadFromFile()
        {
            string filePath = Path.Combine(FileSystem.AppDataDirectory, SaveFileName);
            Dictionary<string, Cell> restored = _fileManager.Deserialize(filePath);
            bool hasData = restored != null && restored.Count > 0;
            if (!hasData)
            {
                return false;
            }

            foreach (KeyValuePair<string, Cell> item in restored)
            {
                Cell cell = _manager.GetCell(item.Key);
                if (cell != null)
                {
                    cell.Expression = NormalizeExpression(item.Value.Expression);
                }
            }

            RecalculateAll();
            return true;
        }

        private Dictionary<string, Cell> GetInternalCells()
        {
            Dictionary<string, Cell> result = new Dictionary<string, Cell>();
            for (int r = 1; r <= _manager.CountRow; ++r)
            {
                for (int c = 0; c < _manager.CountColumn; ++c)
                {
                    string name = SpreadsheetManager.GetColumnName(c) + r;
                    Cell cell = _manager.GetCell(name);
                    if (cell != null)
                    {
                        result[name] = cell;
                    }
                }
            }

            return result;
        }
    }
}