using NPOI.SS.UserModel;

namespace NiuX.Npoi.Extensions;

public static class NpoiExtensions
{
    public static ISheet FirstSheet(this IWorkbook workbook) => workbook.GetSheetAt(0);
    public static IRow CreateRow(this ISheet sheet) => sheet.CreateRow(0);
    public static ICell CreateCell(this IRow row) => row.CreateCell(0);
}