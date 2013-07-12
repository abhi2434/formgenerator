using FormGeneratorEngine.Expressions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleTestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string formula = "=A1 + (20 * B1)/32";


            Expression expression = new Expression(formula);

            var result = expression.Evaluate();

            Console.WriteLine("Result is : {0}", result);

        }
    }
}
