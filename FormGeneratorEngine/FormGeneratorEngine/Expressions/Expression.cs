using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FormGeneratorEngine.Expressions
{
    public class Expression
    {
        public string Formula { get; set; }

        public FormulaParser Parser { get; set; }

        public Expression(string formula)
        {
            this.Formula = formula;

            this.Parser = new FormulaParser(formula);
        }

        public object Evaluate()
        {
            var postFixResult = this.Parser.GetInfixToPostfix();

            double operand1 = 0;
            double operand2 = 0;

            double result = 0;

            Stack<double> stacktokens = new Stack<double>();

            foreach (FormulaToken token in postFixResult)
            {
                if (token.Type == FormulaTokenType.OperatorInfix)
                {
                    switch (token.Value)
                    {
                        case "*":
                            operand1 = stacktokens.Pop();
                            operand2 = stacktokens.Pop();
                            stacktokens.Push(operand1 * operand2);
                            break;
                        case "-":
                            operand1 = stacktokens.Pop();
                            operand2 = stacktokens.Pop();
                            stacktokens.Push(operand1 - operand2);
                            break;
                        case "%":
                            operand1 = stacktokens.Pop();
                            operand2 = stacktokens.Pop();
                            stacktokens.Push(operand1 % operand2);
                            break;
                        case "+":
                            operand1 = stacktokens.Pop();
                            operand2 = stacktokens.Pop();
                            stacktokens.Push(operand1 + operand2);
                            break;
                        case "/":
                            operand1 = stacktokens.Pop();
                            operand2 = stacktokens.Pop();
                            stacktokens.Push(operand1 / operand2);
                            break;
                    }
                }
                else if (token.Type == FormulaTokenType.Operand)
                {
                    if (token.Value == "A1")
                        token.Value = "11";
                    if (token.Value == "B1")
                        token.Value = "22";
                    stacktokens.Push(Convert.ToDouble(token.Value));
                }
            }

            result = stacktokens.Pop();

            return result;
        }

    }
}
