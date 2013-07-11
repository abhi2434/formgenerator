using System.Collections.Generic;

namespace FormGeneratorEngine.Expressions
{
    internal class FormulaTokens
    {
        private int index = -1;
        private List<FormulaToken> tokens;

        public FormulaTokens() : this(4) { }

        public FormulaTokens(int capacity)
        {
            tokens = new List<FormulaToken>(capacity);
        }

        public int Count
        {
            get { return tokens.Count; }
        }

        public bool BOF
        {
            get { return (index <= 0); }
        }

        public bool EOF
        {
            get { return (index >= (tokens.Count - 1)); }
        }

        public FormulaToken Current
        {
            get
            {
                if (index == -1) return null;
                return tokens[index];
            }
        }

        public FormulaToken Next
        {
            get
            {
                if (EOF) return null;
                return tokens[index + 1];
            }
        }

        public FormulaToken Previous
        {
            get
            {
                if (index < 1) return null;
                return tokens[index - 1];
            }
        }

        public FormulaToken Add(FormulaToken token)
        {
            tokens.Add(token);
            return token;
        }

        public bool MoveNext()
        {
            if (EOF) return false;
            index++;
            return true;
        }

        public void Reset()
        {
            index = -1;
        }

    }
}
