using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProyect
{
    public class ExcelException(string error = "") : Exception
    {
        private readonly string message = error;

        public string GetException() => message;
    }
}
