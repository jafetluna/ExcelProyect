using ExcelProyect.Enums;
using ExcelProyect.Helpers;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static ExcelProyect.Helpers.Delegate;

namespace ExcelProyect
{
    public class ExcelFile(string path)
    {
        private readonly IWorkbook _workbook = new XSSFWorkbook();
        private ISheet _sheet;
        private ExcelModel _model = new();
        private readonly string _path = path;
        private readonly List<string> SheetNames = new();
        private readonly GetPosition getPosition = ExcelHelpers.GetPositionByString;
        IDrawing _drawing;
        IRow _row;
        ICell _cell;

        public void CreateSheet(string name = "")
        {
            if(string.IsNullOrEmpty(name)) name = $"Hoja {SheetNames.Count + 1}";
            _sheet = _workbook.CreateSheet(name);
            //_sheet.SetArrayFormula
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fp1">First position</param>
        /// <param name="fp2">Second Position</param>
        /// <param name="lp1">Last</param>
        /// <param name="lp2"></param>
        /// <param name="operation"></param>
        /// <exception cref="Exception"></exception>
        public void ArrayFormula(string fp1, string fp2, string lp1, string lp2, Operation operation, string positionFormulaStart = null)
        {
            try
            {
                CellRangeAddress range = DefineRangeByArrayPosition(fp1, fp2, lp1, lp2, positionFormulaStart);
                string formula = $"{fp1}:{lp1}{operation.SetOperation()}{fp2}:{lp2}";
                _sheet?.SetArrayFormula(formula, range);
            }
            catch (FormulaParseException ex)
            {
                Console.WriteLine(ex.ToString());
                throw new Exception();
            }
        }
        private CellRangeAddress DefineRangeByArrayPosition(string fp1, string fp2, string lp1, string lp2, string positionFormulaStart = null)
        {
            int positionCol = 0, positionRow = 0;
            var (_, firstRow) = getPosition(fp1);//A1 -> 0,0
            var (_, secondRow) = getPosition(fp2);//B1 -> 1, 0
            var (_, thirdRow) = getPosition(lp1);//A1 -> 0, 1
            var (fourthCol, fourthRow) = getPosition(lp2);//B1 -> 1, 1
            if (!string.IsNullOrEmpty(positionFormulaStart))
            {
                (positionCol, positionRow) = getPosition(positionFormulaStart);
            }

            bool isVertical = firstRow == secondRow && thirdRow== fourthRow;
            CellRangeAddress range;
            if (isVertical)
            {
                if (positionCol == 0) positionCol = fourthRow + 1;
                range = new((firstRow + positionRow), (fourthRow + positionRow), positionCol, positionCol);
            }
            else
            {
                if(positionRow == 0) positionRow = thirdRow + 1;
                range = new(positionRow, positionRow, (firstRow + positionRow), (fourthCol + positionCol));
            }
            return range;
        }

       
        public void CreateRow(int row = 0, int col = 0)
        {
            if(row > 0) _model.Row = row ;
             _row = _sheet.CreateRow(_model.Row++);
            _model.Column = col;
        }

        public string SetValue(int value, int column = 0, bool increment = true, ICellStyle style = null)
        {
            if(column  > 0) _model.Column = column ;
            _cell = _row.CreateCell(_model.Column);
            if(style != null)
            {
                _cell.CellStyle = style;
            }
            _cell.SetCellValue(value);
            if (increment) _model.Column++;
            return _cell.Address.ToString();
        }

        public void SetValue(double value, int column = 0, bool increment = true)
        {
            if (column > 0) _model.Column = column;
            _cell = _row.CreateCell(_model.Column);
            _cell.SetCellValue(value);
            if (increment) _model.Column++;
        }

        public void SetValue(string value, int column = 0, bool increment = true, ICellStyle style = null)
        {
            if (column > 0) _model.Column = column;
            _cell = _row.CreateCell(_model.Column);
            if (style != null)
            {
                _cell.CellStyle = style;
            }
            _cell.SetCellValue(value);
            if (increment) _model.Column++;
        }
        
        public void SetCommentary(string commentary, int col = 0, int row = 0)
        {
            _drawing ??= _sheet.CreateDrawingPatriarch();
            if (col == 0) col = _model.Column;
            if(row == 0 ) row = _model.Row;
            IRow iRow = _sheet.GetRow(row) is null ? _sheet.CreateRow(row) : _sheet.GetRow(row);
            ICell cell = iRow.GetCell(col) is null ? iRow.CreateCell(col) : iRow.GetCell(col);

            // Calculate the size of the comment box based on the length of the commentary
            int commentWidth = Math.Min(commentary.Length * 7 * 256, 10000); // Adjust the multiplier and max width as needed
            int commentHeight = 20 * (commentary.Length / 50 + 1); // Adjust the height calculation as needed

            IComment _comment = _drawing.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, col, row, col + 2, row + 2));
            _comment.String = new XSSFRichTextString(commentary);
            _comment.Author = "Author"; // Set the author if needed
            cell.CellComment = (_comment);

            // Adjust the comment box size
            ((XSSFClientAnchor)_comment.ClientAnchor).AnchorType = AnchorType.MoveDontResize;
            ((XSSFClientAnchor)_comment.ClientAnchor).Dx1 = 0;
            ((XSSFClientAnchor)_comment.ClientAnchor).Dy1 = 0;
            ((XSSFClientAnchor)_comment.ClientAnchor).Dx2 = commentWidth;
            ((XSSFClientAnchor)_comment.ClientAnchor).Dy2 = commentHeight;

            // Set the cell background color to yellow
            ICellStyle style = _sheet.Workbook.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.Yellow.Index;
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;
        }

        public void SetMerged(int lastRow, int lastCol, int firstRow = 0, int firstCol = 0)
        {
            if(firstRow == 0) firstRow = _model.Row;
            if(firstCol == 0) firstCol = _model.Column;

            CellRangeAddress region = new(firstRow, lastRow, firstCol, lastCol);
            _sheet.AddMergedRegion(region);
        }
        
        public void SetIncrementMerged(int lastRow, int lastCol)
        {
            CellRangeAddress region = new(_model.Row, (_model.Row + lastRow), _model.Column, (_model.Column + lastCol));
            _sheet.AddMergedRegion(region);
        }

        public void CreateFormula(string value1, string value2, Operation operation, int column = 0)
        {
            if(column > 0) _model.Row = column;
            _cell = _row.CreateCell(_model.Column);
            _cell.SetCellFormula($"{operation.NameOfOperation()}({value1}:{value2})");
            _model.Column++;
        }

        public void SetHeaders(List<string> headers, int column = 0, bool setFilter = false, IndexedColors color = null)
        {
            if(_row == null) CreateRow();
            ICellStyle headerStyle = null;
            if(color != null)
            {
                headerStyle = _workbook.CreateCellStyle();
                headerStyle.Alignment = HorizontalAlignment.Center;
                headerStyle.FillPattern = FillPattern.SolidForeground;
                headerStyle.FillForegroundColor = color.Index;
            }
            foreach(string header in headers)
            {
                SetValue(header, column, style: headerStyle);
            }
            int row = _model.Row - 1;
            if (setFilter) _sheet.SetAutoFilter(new CellRangeAddress(row, row, _model.Column - headers.Count, _model.Column - 1));
        }

        public void Finish(bool open = false)
        {
            FileStream file = File.Create(_path);
            _workbook.Write(file, false);
            file.Close();

            if (open)
            {
                var p = new Process
                {
                    StartInfo = new ProcessStartInfo(_path)
                    {
                        UseShellExecute = true
                    }
                };
                p.Start();
            }
        }

    }
}
