
namespace FormGeneratorEngine.Expressions
{
    public enum FormulaTokenType
    {
        Noop,
        Operand,
        Function,
        Subexpression,
        Argument,
        OperatorPrefix,
        OperatorInfix,
        OperatorPostfix,
        Whitespace,
        Unknown
    }

    public enum FormulaTokenSubtype
    {
        Nothing,
        Start,
        Stop,
        Text,
        Number,
        Logical,
        Error,
        Range,
        Math,
        Concatenation,
        Intersection,
        Union
    }
}
