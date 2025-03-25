// create a new workbook
using ExcelProyect;
using ExcelProyect.Enums;
using NPOI.SS.UserModel;

string path = $"C://tmp//myWorkbook{DateTime.Now.Ticks}.xlsx";
try
{
    ExcelFile excelFile = new(path);
    excelFile.CreateSheet();
    List<string> headers = new() { "Header 1", "Header 2", "Header 3"};
    excelFile.SetHeaders(headers, setFilter: true, color: IndexedColors.Grey40Percent);
    excelFile.CreateRow();
    excelFile.SetValue("Hello world");
    //excelFile.SetComment("Test comment");
    excelFile.SetCommentary("Test comment");
    string val1 = excelFile.SetValue(1);
    string val2 = excelFile.SetValue(2);
    excelFile.CreateFormula(val1, val2, Operation.Sum);
    excelFile.SetIncrementMerged(1, 1);

    excelFile.CreateRow();
    string fp1 = excelFile.SetValue(1);
    string fp2 = excelFile.SetValue(2);
    excelFile.CreateRow();
    string lp1 = excelFile.SetValue(3);
    string lp2 = excelFile.SetValue(4);
    excelFile.ArrayFormula(fp1, fp2, lp1, lp2, Operation.Sum);
    excelFile.Finish(open: true);

}catch(Exception ex)
{ 
    Console.WriteLine(ex);
}
finally
{

}