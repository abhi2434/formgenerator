using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace FormGeneratorEngine.Expressions
{
    public class FormulaParser : IList<FormulaToken>
    {
        string formula;
        List<FormulaToken> tokens;

        private FormulaParser() { }

        public FormulaParser(string formula)
        {
            if (formula == null) throw new ArgumentNullException("formula");
            this.formula = formula.Trim();
            tokens = new List<FormulaToken>();
            ParseToTokens();
        }

        public string Formula
        {
            get { return formula; }
        }

        public FormulaToken this[int index]
        {
            get
            {
                return tokens[index];
            }
            set
            {
                throw new NotSupportedException();
            }
        }

        public int IndexOf(FormulaToken item)
        {
            return tokens.IndexOf(item);
        }

        public void Insert(int index, FormulaToken item)
        {
            throw new NotSupportedException();
        }

        public void RemoveAt(int index)
        {
            throw new NotSupportedException();
        }

        public void Add(FormulaToken item)
        {
            throw new NotSupportedException();
        }

        public void Clear()
        {
            throw new NotSupportedException();
        }

        public bool Contains(FormulaToken item)
        {
            return tokens.Contains(item);
        }

        public bool IsReadOnly
        {
            get { return true; }
        }

        public bool Remove(FormulaToken item)
        {
            throw new NotSupportedException();
        }

        public void CopyTo(FormulaToken[] array, int arrayIndex)
        {
            tokens.CopyTo(array, arrayIndex);
        }

        public int Count
        {
            get { return tokens.Count; }
        }

        public IEnumerator<FormulaToken> GetEnumerator()
        {
            return tokens.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public List<FormulaToken> GetInfixToPostfix()
        {
            List<FormulaToken> actualtokens = this.tokens;
            List<FormulaToken> postFix = new List<FormulaToken>();
            FormulaToken arrival;
            Stack<FormulaToken> oprerator = new Stack<FormulaToken>();
            foreach (FormulaToken ftoken in actualtokens)
            {
                if (ftoken.Type == FormulaTokenType.Operand)
                    postFix.Add(ftoken);
                else if (ftoken.Type == FormulaTokenType.Subexpression && ftoken.Value == "(")
                    oprerator.Push(ftoken);
                else if (ftoken.Type == FormulaTokenType.Subexpression && ftoken.Value == ")")
                {
                    arrival = oprerator.Pop();
                    while (arrival.Type != FormulaTokenType.Subexpression)
                    {
                        postFix.Add(arrival);
                        arrival = oprerator.Pop();
                    }
                }
                else
                {
                    if (oprerator.Count != 0 && this.Predecessor(oprerator.Peek(), ftoken))//If find an operator
                    {
                        arrival = oprerator.Pop();
                        while (this.Predecessor(arrival, ftoken))
                        {
                            postFix.Add(arrival);

                            if (oprerator.Count == 0)
                                break;

                            arrival = oprerator.Pop();
                        }
                        oprerator.Push(ftoken);
                    }
                    else
                        oprerator.Push(ftoken);//If Stack is empty or the operator has precedence 
                }
            }
            while (oprerator.Count > 0)
            {
                arrival = oprerator.Pop();
                postFix.Add(arrival);
            }

            return postFix;
        }
        private bool Predecessor(FormulaToken firstOperator, FormulaToken secondOperator)
        {
            string opString = "(+-*/%";
            
            int firstPoint, secondPoint;

            int[] precedence = { 0, 12, 12, 13, 13, 13 };// "(" has less prececence

            firstPoint = opString.IndexOf(firstOperator.Value);
            secondPoint = opString.IndexOf(secondOperator.Value);

            return (precedence[firstPoint] >= precedence[secondPoint]) ? true : false;
        }

        private void ParseToTokens()
        {

            // No attempt is made to verify formulas; assumes formulas are derived from Excel, where 
            // they can only exist if valid; stack overflows/underflows sunk as nulls without exceptions.

            if ((formula.Length < 2) || (formula[0] != '=')) return;

            FormulaTokens formulatokens = new FormulaTokens();
            FormulaStack formulastack = new FormulaStack();

            const char QUOTE_DOUBLE = '"';
            const char QUOTE_SINGLE = '\'';
            const char BRACKET_CLOSE = ']';
            const char BRACKET_OPEN = '[';
            const char BRACE_OPEN = '{';
            const char BRACE_CLOSE = '}';
            const char PAREN_OPEN = '(';
            const char PAREN_CLOSE = ')';
            const char SEMICOLON = ';';
            const char WHITESPACE = ' ';
            const char COMMA = ',';
            const char ERROR_START = '#';

            const string OPERATORS_SN = "+-";
            const string OPERATORS_INFIX = "+-*/^&=><";
            const string OPERATORS_POSTFIX = "%";

            string[] ERRORS = new string[] { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" };

            string[] COMPARATORS_MULTI = new string[] { ">=", "<=", "<>" };

            bool inString = false;
            bool inPath = false;
            bool inRange = false;
            bool inError = false;

            int index = 1;
            string value = "";

            while (index < formula.Length)
            {

                // state-dependent character evaluation (order is important)

                // double-quoted strings
                // embeds are doubled
                // end marks token

                if (inString)
                {
                    if (formula[index] == QUOTE_DOUBLE)
                    {
                        if (((index + 2) <= formula.Length) && (formula[index + 1] == QUOTE_DOUBLE))
                        {
                            value += QUOTE_DOUBLE;
                            index++;
                        }
                        else
                        {
                            inString = false;
                            formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand, FormulaTokenSubtype.Text));
                            value = "";
                        }
                    }
                    else
                    {
                        value += formula[index];
                    }
                    index++;
                    continue;
                }

                // single-quoted strings (links)
                // embeds are double
                // end does not mark a token

                if (inPath)
                {
                    if (formula[index] == QUOTE_SINGLE)
                    {
                        if (((index + 2) <= formula.Length) && (formula[index + 1] == QUOTE_SINGLE))
                        {
                            value += QUOTE_SINGLE;
                            index++;
                        }
                        else
                        {
                            inPath = false;
                        }
                    }
                    else
                    {
                        value += formula[index];
                    }
                    index++;
                    continue;
                }

                // bracked strings (R1C1 range index or linked workbook name)
                // no embeds (changed to "()" by Excel)
                // end does not mark a token

                if (inRange)
                {
                    if (formula[index] == BRACKET_CLOSE)
                    {
                        inRange = false;
                    }
                    value += formula[index];
                    index++;
                    continue;
                }

                // error values
                // end marks a token, determined from absolute list of values

                if (inError)
                {
                    value += formula[index];
                    index++;
                    if (Array.IndexOf(ERRORS, value) != -1)
                    {
                        inError = false;
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand, FormulaTokenSubtype.Error));
                        value = "";
                    }
                    continue;
                }

                // scientific notation check

                if ((OPERATORS_SN).IndexOf(formula[index]) != -1)
                {
                    if (value.Length > 1)
                    {
                        if (Regex.IsMatch(value, @"^[1-9]{1}(\.[0-9]+)?E{1}$"))
                        {
                            value += formula[index];
                            index++;
                            continue;
                        }
                    }
                }

                // independent character evaluation (order not important)

                // establish state-dependent character evaluations

                if (formula[index] == QUOTE_DOUBLE)
                {
                    if (value.Length > 0)
                    {  // unexpected
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Unknown));
                        value = "";
                    }
                    inString = true;
                    index++;
                    continue;
                }

                if (formula[index] == QUOTE_SINGLE)
                {
                    if (value.Length > 0)
                    { // unexpected
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Unknown));
                        value = "";
                    }
                    inPath = true;
                    index++;
                    continue;
                }

                if (formula[index] == BRACKET_OPEN)
                {
                    inRange = true;
                    value += BRACKET_OPEN;
                    index++;
                    continue;
                }

                if (formula[index] == ERROR_START)
                {
                    if (value.Length > 0)
                    { // unexpected
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Unknown));
                        value = "";
                    }
                    inError = true;
                    value += ERROR_START;
                    index++;
                    continue;
                }

                // mark start and end of arrays and array rows

                if (formula[index] == BRACE_OPEN)
                {
                    if (value.Length > 0)
                    { // unexpected
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Unknown));
                        value = "";
                    }
                    formulastack.Push(formulatokens.Add(new FormulaToken("ARRAY", FormulaTokenType.Function, FormulaTokenSubtype.Start)));
                    formulastack.Push(formulatokens.Add(new FormulaToken("ARRAYROW", FormulaTokenType.Function, FormulaTokenSubtype.Start)));
                    index++;
                    continue;
                }

                if (formula[index] == SEMICOLON)
                {
                    if (value.Length > 0)
                    {
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand));
                        value = "";
                    }
                    formulatokens.Add(formulastack.Pop());
                    formulatokens.Add(new FormulaToken(",", FormulaTokenType.Argument));
                    formulastack.Push(formulatokens.Add(new FormulaToken("ARRAYROW", FormulaTokenType.Function, FormulaTokenSubtype.Start)));
                    index++;
                    continue;
                }

                if (formula[index] == BRACE_CLOSE)
                {
                    if (value.Length > 0)
                    {
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand));
                        value = "";
                    }
                    formulatokens.Add(formulastack.Pop());
                    formulatokens.Add(formulastack.Pop());
                    index++;
                    continue;
                }

                // trim white-space

                if (formula[index] == WHITESPACE)
                {
                    if (value.Length > 0)
                    {
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand));
                        value = "";
                    }
                    formulatokens.Add(new FormulaToken("", FormulaTokenType.Whitespace));
                    index++;
                    while ((formula[index] == WHITESPACE) && (index < formula.Length))
                    {
                        index++;
                    }
                    continue;
                }

                // multi-character comparators

                if ((index + 2) <= formula.Length)
                {
                    if (Array.IndexOf(COMPARATORS_MULTI, formula.Substring(index, 2)) != -1)
                    {
                        if (value.Length > 0)
                        {
                            formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand));
                            value = "";
                        }
                        formulatokens.Add(new FormulaToken(formula.Substring(index, 2), FormulaTokenType.OperatorInfix, FormulaTokenSubtype.Logical));
                        index += 2;
                        continue;
                    }
                }

                // standard infix operators

                if ((OPERATORS_INFIX).IndexOf(formula[index]) != -1)
                {
                    if (value.Length > 0)
                    {
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand));
                        value = "";
                    }
                    formulatokens.Add(new FormulaToken(formula[index].ToString(), FormulaTokenType.OperatorInfix));
                    index++;
                    continue;
                }

                // standard postfix operators (only one)

                if ((OPERATORS_POSTFIX).IndexOf(formula[index]) != -1)
                {
                    if (value.Length > 0)
                    {
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand));
                        value = "";
                    }
                    formulatokens.Add(new FormulaToken(formula[index].ToString(), FormulaTokenType.OperatorPostfix));
                    index++;
                    continue;
                }

                // start subexpression or function

                if (formula[index] == PAREN_OPEN)
                {
                    if (value.Length > 0)
                    {
                        formulastack.Push(formulatokens.Add(new FormulaToken(value, FormulaTokenType.Function, FormulaTokenSubtype.Start)));
                        value = "";
                    }
                    else
                    {
                        formulastack.Push(formulatokens.Add(new FormulaToken("", FormulaTokenType.Subexpression, FormulaTokenSubtype.Start)));
                    }
                    index++;
                    continue;
                }

                // function, subexpression, or array parameters, or operand unions

                if (formula[index] == COMMA)
                {
                    if (value.Length > 0)
                    {
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand));
                        value = "";
                    }
                    if (formulastack.Current.Type != FormulaTokenType.Function)
                    {
                        formulatokens.Add(new FormulaToken(",", FormulaTokenType.OperatorInfix, FormulaTokenSubtype.Union));
                    }
                    else
                    {
                        formulatokens.Add(new FormulaToken(",", FormulaTokenType.Argument));
                    }
                    index++;
                    continue;
                }

                // stop subexpression

                if (formula[index] == PAREN_CLOSE)
                {
                    if (value.Length > 0)
                    {
                        formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand));
                        value = "";
                    }
                    formulatokens.Add(formulastack.Pop());
                    index++;
                    continue;
                }

                // token accumulation

                value += formula[index];
                index++;

            }

            // dump remaining accumulation

            if (value.Length > 0)
            {
                formulatokens.Add(new FormulaToken(value, FormulaTokenType.Operand));
            }

            // move tokenList to new set, excluding unnecessary white-space tokens and converting necessary ones to intersections

            FormulaTokens tokens2 = new FormulaTokens(formulatokens.Count);

            while (formulatokens.MoveNext())
            {

                FormulaToken token = formulatokens.Current;

                if (token == null) continue;

                if (token.Type != FormulaTokenType.Whitespace)
                {
                    tokens2.Add(token);
                    continue;
                }

                if ((formulatokens.BOF) || (formulatokens.EOF)) continue;

                FormulaToken previous = formulatokens.Previous;

                if (previous == null) continue;

                if (!(
                      ((previous.Type == FormulaTokenType.Function) && (previous.Subtype == FormulaTokenSubtype.Stop)) ||
                      ((previous.Type == FormulaTokenType.Subexpression) && (previous.Subtype == FormulaTokenSubtype.Stop)) ||
                      (previous.Type == FormulaTokenType.Operand)
                      )
                    ) continue;

                FormulaToken next = formulatokens.Next;

                if (next == null) continue;

                if (!(
                      ((next.Type == FormulaTokenType.Function) && (next.Subtype == FormulaTokenSubtype.Start)) ||
                      ((next.Type == FormulaTokenType.Subexpression) && (next.Subtype == FormulaTokenSubtype.Start)) ||
                      (next.Type == FormulaTokenType.Operand)
                      )
                    ) continue;

                tokens2.Add(new FormulaToken("", FormulaTokenType.OperatorInfix, FormulaTokenSubtype.Intersection));

            }

            // move tokens to final list, switching infix "-" operators to prefix when appropriate, switching infix "+" operators 
            // to noop when appropriate, identifying operand and infix-operator subtypes, and pulling "@" from function names

            tokens = new List<FormulaToken>(tokens2.Count);

            while (tokens2.MoveNext())
            {

                FormulaToken token = tokens2.Current;

                if (token == null) continue;

                FormulaToken previous = tokens2.Previous;
                FormulaToken next = tokens2.Next;

                if ((token.Type == FormulaTokenType.OperatorInfix) && (token.Value == "-"))
                {
                    if (tokens2.BOF)
                        token.Type = FormulaTokenType.OperatorPrefix;
                    else if (
                            ((previous.Type == FormulaTokenType.Function) && (previous.Subtype == FormulaTokenSubtype.Stop)) ||
                            ((previous.Type == FormulaTokenType.Subexpression) && (previous.Subtype == FormulaTokenSubtype.Stop)) ||
                            (previous.Type == FormulaTokenType.OperatorPostfix) ||
                            (previous.Type == FormulaTokenType.Operand)
                            )
                        token.Subtype = FormulaTokenSubtype.Math;
                    else
                        token.Type = FormulaTokenType.OperatorPrefix;

                    tokens.Add(token);
                    continue;
                }

                if ((token.Type == FormulaTokenType.OperatorInfix) && (token.Value == "+"))
                {
                    if (tokens2.BOF)
                        continue;
                    else if (
                            ((previous.Type == FormulaTokenType.Function) && (previous.Subtype == FormulaTokenSubtype.Stop)) ||
                            ((previous.Type == FormulaTokenType.Subexpression) && (previous.Subtype == FormulaTokenSubtype.Stop)) ||
                            (previous.Type == FormulaTokenType.OperatorPostfix) ||
                            (previous.Type == FormulaTokenType.Operand)
                            )
                        token.Subtype = FormulaTokenSubtype.Math;
                    else
                        continue;

                    tokens.Add(token);
                    continue;
                }

                if ((token.Type == FormulaTokenType.OperatorInfix) && (token.Subtype == FormulaTokenSubtype.Nothing))
                {
                    if (("<>=").IndexOf(token.Value.Substring(0, 1)) != -1)
                        token.Subtype = FormulaTokenSubtype.Logical;
                    else if (token.Value == "&")
                        token.Subtype = FormulaTokenSubtype.Concatenation;
                    else
                        token.Subtype = FormulaTokenSubtype.Math;

                    tokens.Add(token);
                    continue;
                }

                if ((token.Type == FormulaTokenType.Operand) && (token.Subtype == FormulaTokenSubtype.Nothing))
                {
                    double d;
                    bool isNumber = double.TryParse(token.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out d);
                    if (!isNumber)
                        if ((token.Value == "TRUE") || (token.Value == "FALSE"))
                            token.Subtype = FormulaTokenSubtype.Logical;
                        else
                            token.Subtype = FormulaTokenSubtype.Range;
                    else
                        token.Subtype = FormulaTokenSubtype.Number;

                    tokens.Add(token);
                    continue;
                }

                if (token.Type == FormulaTokenType.Function)
                {
                    if (token.Value.Length > 0)
                    {
                        if (token.Value.Substring(0, 1) == "@")
                        {
                            token.Value = token.Value.Substring(1);
                        }
                    }
                }

                tokens.Add(token);

            }

        }

    }
}