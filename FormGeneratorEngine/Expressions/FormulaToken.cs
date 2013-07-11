
namespace FormGeneratorEngine.Expressions
{
    public class FormulaToken
    {
        private string value;
        private FormulaTokenType type;
        private FormulaTokenSubtype subtype;

        private FormulaToken() { }

        internal FormulaToken(string value, FormulaTokenType type) : this(value, type, FormulaTokenSubtype.Nothing) { }

        internal FormulaToken(string value, FormulaTokenType type, FormulaTokenSubtype subtype)
        {
            this.value = value;
            this.type = type;
            this.subtype = subtype;
        }

        public string Value
        {
            get { return value; }
            internal set { this.value = value; }
        }

        public FormulaTokenType Type
        {
            get { return type; }
            internal set { type = value; }
        }

        public FormulaTokenSubtype Subtype
        {
            get { return subtype; }
            internal set { subtype = value; }
        }

    }
}
