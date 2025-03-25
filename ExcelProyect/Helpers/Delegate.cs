using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProyect.Helpers
{
    public class Delegate
    {
        public delegate (int, int) GetPosition(string rowText);
    }
}
