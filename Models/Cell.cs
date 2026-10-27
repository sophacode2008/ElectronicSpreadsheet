namespace ElectronicSpreadsheet.Models
{
    public class Cell
    {
        private string _expression;
        private double _value;
        private string _errorMessage;
        private bool _hasError;

        public string Name { get; set; }

        public string Expression
        {
            get { return _expression; }
            set { _expression = value; }
        }

        public double Value
        {
            get { return _value; }
            set { _value = value; }
        }

        public string ErrorMessage
        {
            get { return _errorMessage; }
            set { _errorMessage = value; }
        }

        public bool HasError
        {
            get { return _hasError; }
            set { _hasError = value; }
        }

        public Cell(string name)
        {
            Name = name;
            _expression = string.Empty;
            _errorMessage = string.Empty;
            _hasError = false;
            _value = 0.0;
        }
    }
}