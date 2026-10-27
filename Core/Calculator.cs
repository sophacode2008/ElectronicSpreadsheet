using Antlr4.Runtime;
using System;

namespace ElectronicSpreadsheet.Core
{
    public class Calculator
    {
        public double Evaluate(string expression, Func<string, double> getCellValue)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return 0.0;
            }

            // If its just a number without "=" or operations, we can directly parse it
            if (double.TryParse(expression, System.Globalization.CultureInfo.InvariantCulture, out double directNumber))
            {
                return directNumber;
            }

            // If the formula starts with "=" — we remove it
            if (expression.StartsWith("="))
            {
                expression = expression.Substring(1);
            }

            AntlrInputStream inputStream = new AntlrInputStream(expression);
            CalculatorLexer lexer = new CalculatorLexer(inputStream);
            CommonTokenStream tokens = new CommonTokenStream(lexer);
            CalculatorParser parser = new CalculatorParser(tokens);

            // Add custom error handling for ANTLR syntax errors
            parser.RemoveErrorListeners();

            CalculatorParser.CompileUnitContext tree = parser.compileUnit();
            LabCalculatorVisitor visitor = new LabCalculatorVisitor(getCellValue);

            return visitor.Visit(tree);
        }
    }
}