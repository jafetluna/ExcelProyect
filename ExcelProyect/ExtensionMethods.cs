using ExcelProyect.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProyect
{
    public static class ExtensionMethods
    {
        public static string NameOfOperation(this Operation operation)
        {
            return operation switch
            {
                (byte)Operation.Sum => nameof(Operation.Sum),
                _ => throw new NotSupportedException()
            };
        }

        public static string SetOperation(this Operation operation)
        {
            return operation switch
            {
                Operation.Sum => "+",
                _ => throw new NotSupportedException()
            };
        }

    }
}
