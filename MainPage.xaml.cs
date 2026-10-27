using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Maui.Controls;
using Microsoft.Maui.Graphics;
using Microsoft.Maui.Storage;
using ElectronicSpreadsheet.Core;
using ElectronicSpreadsheet.Models;
using ElectronicSpreadsheet.Services;
using Cell = ElectronicSpreadsheet.Models.Cell;

namespace ElectronicSpreadsheet
{
    // Main page class for managing the electronic spreadsheet UI
    public partial class MainPage : ContentPage
    {
        private const int InitialRowCount = 16;
        private const int InitialColumnCount = 10;
        private const int MinGridDimension = 1;
        private const int EvaluationPasses = 3;
        private const double RowHeaderWidth = 40.0;
        private const double ColumnCellWidth = 85.0;
        private const double GridRowHeight = 32.0;
        private const double CellFontSize = 13.0;
        private const string SaveFileName = "spreadsheet.json";
        private const string HeaderBackColorHex = "#121212";
        private const string CellBackColorHex = "#1C1C1C";

        private readonly SpreadsheetManager _manager;
        private readonly JSONFileManager _fileManager;
        private readonly Dictionary<string, Entry> _entryMap;
        private string _selectedCellName = string.Empty;
        private bool _showFormulas = false;
        private bool _isUpdatingUI = false;

        public MainPage()
        {
            InitializeComponent();

            _manager = new SpreadsheetManager(InitialRowCount, InitialColumnCount);
            _fileManager = new JSONFileManager();
            _entryMap = new Dictionary<string, Entry>();

            CreateGrid();
        }

        // Initializes the spreadsheet grid structure
        private void CreateGrid()
        {
            gridTable.RowDefinitions.Clear();
            gridTable.ColumnDefinitions.Clear();
            gridTable.Children.Clear();
            _entryMap.Clear();

            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }

        // Adds column definitions and A-Z column header labels
        private void AddColumnsAndColumnLabels()
        {
            int colCount = _manager.CountColumn;
            gridTable.ColumnDefinitions.Add(new ColumnDefinition { Width = RowHeaderWidth });
            gridTable.RowDefinitions.Add(new RowDefinition { Height = GridRowHeight });

            BoxView cornerBox = new BoxView();
            cornerBox.Color = Color.FromArgb(HeaderBackColorHex);
            AddViewToGrid(cornerBox, 0, 0);

            for (int c = 0; c < colCount; ++c)
            {
                gridTable.ColumnDefinitions.Add(new ColumnDefinition { Width = ColumnCellWidth });
                string colName = SpreadsheetManager.GetColumnName(c);
                Label headerLabel = CreateHeaderLabel(colName);
                AddViewToGrid(headerLabel, 0, c + 1);
            }
        }

        // Adds row definitions, row number headers, and cell input entries
        private void AddRowsAndCellEntries()
        {
            int rowCount = _manager.CountRow;
            int colCount = _manager.CountColumn;

            for (int r = 0; r < rowCount; ++r)
            {
                int rowNum = r + 1;
                gridTable.RowDefinitions.Add(new RowDefinition { Height = GridRowHeight });

                Label rowHeader = CreateHeaderLabel(rowNum.ToString());
                AddViewToGrid(rowHeader, rowNum, 0);

                for (int c = 0; c < colCount; ++c)
                {
                    string cellName = SpreadsheetManager.GetColumnName(c) + rowNum;
                    Entry cellEntry = CreateCellEntry(cellName);
                    AddViewToGrid(cellEntry, rowNum, c + 1);
                    _entryMap[cellName] = cellEntry;
                }
            }
        }

        private void AddViewToGrid(View view, int row, int col)
        {
            Grid.SetRow(view, row);
            Grid.SetColumn(view, col);
            gridTable.Children.Add(view);
        }

        private Label CreateHeaderLabel(string text)
        {
            Label label = new Label();
            label.Text = text;
            label.FontAttributes = FontAttributes.Bold;
            label.BackgroundColor = Color.FromArgb(HeaderBackColorHex);
            label.TextColor = Colors.White;
            label.HorizontalTextAlignment = TextAlignment.Center;
            label.VerticalTextAlignment = TextAlignment.Center;
            return label;
        }

        private Entry CreateCellEntry(string cellName)
        {
            Cell cell = _manager.GetCell(cellName);
            string errorCode = ResolveCellError(cellName, new HashSet<string>());
            bool hasError = !string.IsNullOrEmpty(errorCode);

            Entry cellEntry = new Entry();
            cellEntry.Text = GetCellDisplayText(cellName, cell);
            cellEntry.TextColor = hasError ? Colors.OrangeRed : Colors.White;
            cellEntry.HorizontalTextAlignment = TextAlignment.Center;
            cellEntry.BackgroundColor = Color.FromArgb(CellBackColorHex);
            cellEntry.FontSize = CellFontSize;

            cellEntry.Focused += (s, e) => OnCellFocused(cellName, cellEntry);
            cellEntry.TextChanged += (s, e) => OnCellTextChanged(cellName, cellEntry);
            cellEntry.Completed += (s, e) => RecalculateAndRefresh();
            cellEntry.Unfocused += (s, e) => RecalculateAndRefresh();

            return cellEntry;
        }

        private void OnCellFocused(string cellName, Entry cellEntry)
        {
            _isUpdatingUI = true;
            _selectedCellName = cellName;
            lblSelectedCell.Text = cellName;
            lblSelectedCell.IsVisible = true;

            Cell targetCell = _manager.GetCell(cellName);
            string expr = targetCell != null ? targetCell.Expression : string.Empty;
            txtFormula.Text = expr;
            cellEntry.Text = expr;
            cellEntry.TextColor = Colors.White;
            _isUpdatingUI = false;
        }

        private void OnCellTextChanged(string cellName, Entry cellEntry)
        {
            bool ignoreChange = _isUpdatingUI || !cellEntry.IsFocused;
            if (ignoreChange)
            {
                return;
            }

            Cell targetCell = _manager.GetCell(cellName);
            if (targetCell != null)
            {
                string currentText = cellEntry.Text ?? string.Empty;
                bool isDisplayedErrorCode = currentText.StartsWith("#")
                    && !currentText.StartsWith("#N/A", StringComparison.OrdinalIgnoreCase);
                if (isDisplayedErrorCode)
                {
                    return;
                }

                _isUpdatingUI = true;
                targetCell.Expression = NormalizeExpression(currentText);
                txtFormula.Text = currentText;
                SafelyEvaluateAll();
                UpdateUI(cellName);
                _isUpdatingUI = false;
            }
        }

        private void RecalculateAndRefresh()
        {
            SafelyEvaluateAll();
            UpdateUI();
        }

        // Isolates cells with syntax or circular errors before calling EvaluateAll
        private void SafelyEvaluateAll()
        {
            Dictionary<string, Cell> allCells = GetInternalCells();
            Dictionary<string, string> isolated = new Dictionary<string, string>();

            foreach (KeyValuePair<string, Cell> pair in allCells)
            {
                pair.Value.HasError = false;
                pair.Value.ErrorMessage = string.Empty;
            }

            foreach (KeyValuePair<string, Cell> pair in allCells)
            {
                bool hasExpr = !string.IsNullOrWhiteSpace(pair.Value.Expression);
                if (hasExpr)
                {
                    string preError = ResolveCellError(pair.Key, new HashSet<string>());
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

            for (int i = 0; i < EvaluationPasses; ++i)
            {
                _manager.EvaluateAll();
            }

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

        // Determines the exact Excel error code (#REF!, #NAME?, #VALUE!, #DIV/0!, #N/A)
        private string ResolveCellError(string cellName, HashSet<string> visited)
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
            string err = cell.ErrorMessage != null ? cell.ErrorMessage.ToLowerInvariant() : string.Empty;

            bool isNa = expr.Equals("#N/A", StringComparison.OrdinalIgnoreCase)
                || expr.Equals("NA()", StringComparison.OrdinalIgnoreCase)
                || err.Contains("n/a");
            if (isNa)
            {
                return "#N/A";
            }

            bool isUnknownName = HasUnknownIdentifier(expr) || err.Contains("name");
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
                || Regex.IsMatch(expr, @"(\/|\bmod\b|\bdiv\b)\s*0(\.0+)?\b", RegexOptions.IgnoreCase);
            if (isRuntimeDivZero)
            {
                return "#DIV/0!";
            }

            return string.Empty;
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

        private string GetCellDisplayText(string cellName, Cell cell)
        {
            bool isEmpty = cell == null || string.IsNullOrWhiteSpace(cell.Expression);
            if (isEmpty)
            {
                return string.Empty;
            }

            if (_showFormulas)
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

        // Updates all cells in the UI grid
        private void UpdateUI(string excludeCellName = null)
        {
            bool prevState = _isUpdatingUI;
            _isUpdatingUI = true;

            foreach (KeyValuePair<string, Entry> pair in _entryMap)
            {
                if (pair.Key == excludeCellName)
                {
                    continue;
                }

                Cell cell = _manager.GetCell(pair.Key);
                if (cell != null)
                {
                    string errorCode = ResolveCellError(pair.Key, new HashSet<string>());
                    bool hasError = !string.IsNullOrEmpty(errorCode);
                    pair.Value.Text = GetCellDisplayText(pair.Key, cell);
                    pair.Value.TextColor = (hasError && !_showFormulas) ? Colors.OrangeRed : Colors.White;
                }
            }

            _isUpdatingUI = prevState;
        }

        private void ApplyFormulaFromInput()
        {
            bool cannotApply = _isUpdatingUI || string.IsNullOrEmpty(_selectedCellName);
            if (cannotApply)
            {
                return;
            }

            Cell selectedCell = _manager.GetCell(_selectedCellName);
            if (selectedCell != null)
            {
                selectedCell.Expression = NormalizeExpression(txtFormula.Text);
            }

            RecalculateAndRefresh();
        }

        private void FormulaEntry_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txtFormula.IsFocused)
            {
                ApplyFormulaFromInput();
            }
        }

        private void FormulaEntry_Completed(object sender, EventArgs e)
        {
            ApplyFormulaFromInput();
        }

        private void FormulaEntry_Unfocused(object sender, FocusEventArgs e)
        {
            ApplyFormulaFromInput();
        }

        private void ShowFormulasSwitch_Toggled(object sender, ToggledEventArgs e)
        {
            _showFormulas = e.Value;
            lblSwitchState.Text = _showFormulas ? "On" : "Off";
            UpdateUI();
        }

        private async void SaveButton_Clicked(object sender, EventArgs e)
        {
            string filePath = Path.Combine(FileSystem.AppDataDirectory, SaveFileName);
            _fileManager.Serialize(filePath, GetInternalCells());
            await DisplayAlert("Збереження", $"Таблицю збережено у файл:\n{filePath}", "OK");
        }

        private async void LoadButton_Clicked(object sender, EventArgs e)
        {
            string filePath = Path.Combine(FileSystem.AppDataDirectory, SaveFileName);
            Dictionary<string, Cell> restored = _fileManager.Deserialize(filePath);
            bool hasData = restored != null && restored.Count > 0;

            if (hasData)
            {
                foreach (KeyValuePair<string, Cell> item in restored)
                {
                    Cell cell = _manager.GetCell(item.Key);
                    if (cell != null)
                    {
                        cell.Expression = NormalizeExpression(item.Value.Expression);
                    }
                }

                RecalculateAndRefresh();
                await DisplayAlert("Завантаження", "Таблицю успішно завантажено!", "OK");
            }
            else
            {
                await DisplayAlert("Помилка", "Файл збереження ще не створено.", "OK");
            }
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

        private void ResizeGridWithPreservedCells(int newRows, int newCols)
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

            SafelyEvaluateAll();
            CreateGrid();
        }

        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            ResizeGridWithPreservedCells(_manager.CountRow + 1, _manager.CountColumn);
        }

        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            bool canDelete = _manager.CountRow > MinGridDimension;
            if (canDelete)
            {
                ResizeGridWithPreservedCells(_manager.CountRow - 1, _manager.CountColumn);
            }
        }

        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            ResizeGridWithPreservedCells(_manager.CountRow, _manager.CountColumn + 1);
        }

        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            bool canDelete = _manager.CountColumn > MinGridDimension;
            if (canDelete)
            {
                ResizeGridWithPreservedCells(_manager.CountRow, _manager.CountColumn - 1);
            }
        }

        private void AboutButton_Click(object sender, EventArgs e)
        {
            AboutButton_Clicked(sender, e);
        }

        private async void AboutButton_Clicked(object sender, EventArgs e)
        {
            string info = "Лабораторна робота № 1.Жикуліна Софія К-27. Варіант 1.\nЕлектронні таблиці (MAUI, ANTLR4, JSON).";
            await DisplayAlert("Про програму", info, "OK");
        }

        private async void HelpButton_Clicked(object sender, EventArgs e)
        {
            string help = "Операції: +, -, *, /, mod, div, ^, ( ), посилання (A1, B2).\n"
                + "Помилки: #DIV/0!, #REF!, #NAME?, #VALUE!, #N/A.";
            await DisplayAlert("Довідка", help, "OK");
        }

        private async void ExitButton_Clicked(object sender, EventArgs e)
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
            if (answer)
            {
                Application.Current?.Quit();
            }
        }
    }
}