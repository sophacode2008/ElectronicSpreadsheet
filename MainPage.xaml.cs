using System;
using System.Collections.Generic;
using Microsoft.Maui.Controls;
using Microsoft.Maui.Graphics;
using ElectronicSpreadsheet.Core;

namespace ElectronicSpreadsheet
{
    // View class responsible strictly for UI rendering and user event delegation
    public partial class MainPage : ContentPage
    {
        private const double RowHeaderWidth = 40.0;
        private const double ColumnCellWidth = 85.0;
        private const double GridRowHeight = 32.0;
        private const double CellFontSize = 13.0;
        private const string HeaderBackColorHex = "#121212";
        private const string CellBackColorHex = "#1C1C1C";

        private readonly SpreadsheetController _controller;
        private readonly Dictionary<string, Entry> _entryMap;
        private string _selectedCellName = string.Empty;
        private bool _showFormulas = false;
        private bool _isUpdatingUI = false;

        public MainPage()
        {
            InitializeComponent();

            _controller = new SpreadsheetController();
            _entryMap = new Dictionary<string, Entry>();

            CreateGrid();
        }

        // Builds the visual spreadsheet grid structure
        private void CreateGrid()
        {
            gridTable.RowDefinitions.Clear();
            gridTable.ColumnDefinitions.Clear();
            gridTable.Children.Clear();
            _entryMap.Clear();

            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }

        private void AddColumnsAndColumnLabels()
        {
            int colCount = _controller.ColumnCount;
            gridTable.ColumnDefinitions.Add(new ColumnDefinition { Width = RowHeaderWidth });
            gridTable.RowDefinitions.Add(new RowDefinition { Height = GridRowHeight });

            BoxView cornerBox = new BoxView();
            cornerBox.Color = Color.FromArgb(HeaderBackColorHex);
            AddViewToGrid(cornerBox, 0, 0);

            for (int c = 0; c < colCount; ++c)
            {
                gridTable.ColumnDefinitions.Add(new ColumnDefinition { Width = ColumnCellWidth });
                string colName = _controller.GetColumnHeader(c);
                Label headerLabel = CreateHeaderLabel(colName);
                AddViewToGrid(headerLabel, 0, c + 1);
            }
        }

        private void AddRowsAndCellEntries()
        {
            int rowCount = _controller.RowCount;
            int colCount = _controller.ColumnCount;

            for (int r = 0; r < rowCount; ++r)
            {
                int rowNum = r + 1;
                gridTable.RowDefinitions.Add(new RowDefinition { Height = GridRowHeight });

                Label rowHeader = CreateHeaderLabel(rowNum.ToString());
                AddViewToGrid(rowHeader, rowNum, 0);

                for (int c = 0; c < colCount; ++c)
                {
                    string cellName = _controller.GetColumnHeader(c) + rowNum;
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
            bool hasError = _controller.IsCellInError(cellName);

            Entry cellEntry = new Entry();
            cellEntry.Text = _controller.GetCellDisplayText(cellName, _showFormulas);
            cellEntry.TextColor = (hasError && !_showFormulas) ? Colors.OrangeRed : Colors.White;
            cellEntry.HorizontalTextAlignment = TextAlignment.Center;
            cellEntry.BackgroundColor = Color.FromArgb(CellBackColorHex);
            cellEntry.FontSize = CellFontSize;

            cellEntry.Focused += (s, e) => OnCellFocused(cellName, cellEntry);
            cellEntry.TextChanged += (s, e) => OnCellTextChanged(cellName, cellEntry);
            cellEntry.Completed += (s, e) => OnCellCompleted();
            cellEntry.Unfocused += (s, e) => OnCellUnfocused(cellName);

            return cellEntry;
        }

        private void OnCellFocused(string cellName, Entry cellEntry)
        {
            _isUpdatingUI = true;
            _selectedCellName = cellName;
            lblSelectedCell.Text = cellName;
            lblSelectedCell.IsVisible = true;

            UpdateUI(cellName);

            string expr = _controller.GetCellExpression(cellName);
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

            string currentText = cellEntry.Text ?? string.Empty;
            bool isDisplayedErrorCode = currentText.StartsWith("#")
                && !currentText.StartsWith("#N/A", StringComparison.OrdinalIgnoreCase);
            if (isDisplayedErrorCode)
            {
                return;
            }

            _isUpdatingUI = true;
            txtFormula.Text = currentText;
            _controller.UpdateCellExpression(cellName, currentText);
            UpdateUI(cellName);
            _isUpdatingUI = false;
        }

        private void OnCellCompleted()
        {
            _controller.RecalculateAll();
            UpdateUI();
        }

        private void OnCellUnfocused(string unfocusedCellName)
        {
            _controller.RecalculateAll();
            string activeCell = IsAnyCellFocused() ? _selectedCellName : null;
            UpdateUI(activeCell);
        }

        private bool IsAnyCellFocused()
        {
            bool hasSelection = !string.IsNullOrEmpty(_selectedCellName)
                && _entryMap.ContainsKey(_selectedCellName);
            if (!hasSelection)
            {
                return false;
            }

            return _entryMap[_selectedCellName].IsFocused;
        }

        // Refreshes displayed text and error highlighting across all cells in the UI
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

                bool hasError = _controller.IsCellInError(pair.Key);
                pair.Value.Text = _controller.GetCellDisplayText(pair.Key, _showFormulas);
                pair.Value.TextColor = (hasError && !_showFormulas) ? Colors.OrangeRed : Colors.White;
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

            _controller.UpdateCellExpression(_selectedCellName, txtFormula.Text);
            UpdateUI();
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
            string filePath = _controller.SaveToFile();
            await DisplayAlert("Збереження", $"Таблицю збережено у файл:\n{filePath}", "OK");
        }

        private async void LoadButton_Clicked(object sender, EventArgs e)
        {
            bool isLoaded = _controller.LoadFromFile();
            if (isLoaded)
            {
                UpdateUI();
                await DisplayAlert("Завантаження", "Таблицю успішно завантажено!", "OK");
            }
            else
            {
                await DisplayAlert("Помилка", "Файл збереження ще не створено.", "OK");
            }
        }

        private void AddRowButton_Clicked(object sender, EventArgs e)
        {
            _controller.AddRow();
            CreateGrid();
        }

        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            bool isDeleted = _controller.DeleteRow();
            if (isDeleted)
            {
                CreateGrid();
            }
        }

        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            _controller.AddColumn();
            CreateGrid();
        }

        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            bool isDeleted = _controller.DeleteColumn();
            if (isDeleted)
            {
                CreateGrid();
            }
        }

        private async void AboutButton_Clicked(object sender, EventArgs e)
        {
            string info = "Лабораторна робота № 1. Варіант 1.\nЕлектронні таблиці (MAUI, ANTLR4, JSON).";
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