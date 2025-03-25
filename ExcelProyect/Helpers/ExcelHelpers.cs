using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelProyect.Helpers
{
    public static class ExcelHelpers
    {
        public static (int, int) GetPositionByString(string rowText)
        {
            if (!string.IsNullOrEmpty(rowText))
            {
                int row = CellReference.ConvertColStringToIndex(Regex.Match(rowText, @"[A-Za-z]+").Value);
                if (int.TryParse(Regex.Match(rowText, @"\d+").Value, out int col))
                {
                    return (row, --col);
                }

            }

            return (0, 0);
        }

    }
}
