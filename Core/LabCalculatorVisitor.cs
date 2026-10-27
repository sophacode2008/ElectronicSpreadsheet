using System;
using System.Globalization;

namespace ElectronicSpreadsheet.Core
{
    public class LabCalculatorVisitor : CalculatorBaseVisitor<double>
    {
        private readonly Func<string, double> _getCellValue;

        public LabCalculatorVisitor(Func<string, double> getCellValue)
        {
            _getCellValue = getCellValue;
        }

        public override double VisitCompileUnit(CalculatorParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitNumberExpr(CalculatorParser.NumberExprContext context)
        {
            string text = context.GetText();
            return double.Parse(text, CultureInfo.InvariantCulture);
        }

        public override double VisitIdentifierExpr(CalculatorParser.IdentifierExprContext context)
        {
            string cellName = context.IDENTIFIER().GetText();
            return _getCellValue(cellName);
        }

        public override double VisitParenthesesExpr(CalculatorParser.ParenthesesExprContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitUnaryExpr(CalculatorParser.UnaryExprContext context)
        {
            double value = Visit(context.expression());
            if (context.@operator.Text == "-")
                return -value;
            return value;
        }

        public override double VisitPowerExpr(CalculatorParser.PowerExprContext context)
        {
            double left = Visit(context.left);
            double right = Visit(context.right);
            return Math.Pow(left, right);
        }

        public override double VisitMultiplicativeExpr(CalculatorParser.MultiplicativeExprContext context)
        {
            double left = Visit(context.left);
            double right = Visit(context.right);
            string op = context.@operator.Text;

            if (op == "*") return left * right;
            if (op == "/")
            {
                if (right == 0)
                {
                    throw new DivideByZeroException("Division by zero.");
                }
                return left / right;
            }
            if (op == "mod")
            {
                if (right == 0)
                {
                    throw new DivideByZeroException("Division by zero (mod).");
                }
                return left % right;
            }
            if (op == "div")
            {
                if (right == 0)
                {
                    throw new DivideByZeroException("Division by zero (div).");
                }
                return Math.Truncate(left / right);
            }
            return 0;
        }

        public override double VisitAdditiveExpr(CalculatorParser.AdditiveExprContext context)
        {
            double left = Visit(context.left);
            double right = Visit(context.right);
            string op = context.@operator.Text;

            if (op == "+") return left + right;
            if (op == "-") return left - right;
            return 0;
        }
    }
}