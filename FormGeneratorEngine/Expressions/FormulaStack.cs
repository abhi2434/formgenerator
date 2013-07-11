using System.Collections;
using System.Collections.Generic;

namespace FormGeneratorEngine.Expressions
{
    public class FormulaStack : IEnumerable<FormulaToken>, ICollection, IEnumerable
    {
        private Stack<FormulaToken> stack = new Stack<FormulaToken>();

        public FormulaStack() { }

        public void Push(FormulaToken token)
        {
            stack.Push(token);
        }

        public FormulaToken Pop()
        {
            if (stack.Count == 0) return null;
            return new FormulaToken("", stack.Pop().Type, FormulaTokenSubtype.Stop);
        }

        public FormulaToken Current
        {
            get { return (stack.Count > 0) ? stack.Peek() : null; }
        }


        public IEnumerator<FormulaToken> GetEnumerator()
        {
            return stack.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return stack.GetEnumerator();
        }

        public void CopyTo(FormulaToken[] array, int index)
        {
            this.stack.CopyTo(array, index);
        }

        public int Count
        {
            get { return this.stack.Count; }
        }

        public bool IsSynchronized
        {
            get { return true; }
        }

        public object SyncRoot
        {
            get { return null; }
        }

        public void CopyTo(System.Array array, int index)
        {
            this.CopyTo((FormulaToken[])array, index);
        }
    }
}
