using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProyect
{
    public class ExcelModel(int row = 0, int column = 0)
    {
        public int Row { get; set; } = row;
        public int Column { get; set; } = column;

    }
}
