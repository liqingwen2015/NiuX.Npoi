using NPOI.SS.UserModel;

namespace NiuX.Npoi.Extensions;

public static class NpoiExtensions
{
    /// <summary>
    /// get first sheet
    /// </summary>
    /// <param name="workbook"></param>
    /// <returns></returns>
    public static ISheet FirstSheet(this IWorkbook workbook) => workbook.GetSheetAt(0);
    
    /// <summary>
    /// create new row
    /// </summary>
    /// <param name="sheet"></param>
    /// <returns></returns>
    public static IRow CreateRow(this ISheet sheet) => sheet.CreateRow(0);
    
    /// <summary>
    /// create new cell
    /// </summary>
    /// <param name="row"></param>
    /// <returns></returns>
    public static ICell CreateCell(this IRow row) => row.CreateCell(0);
}