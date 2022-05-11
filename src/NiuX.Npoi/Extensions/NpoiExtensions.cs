using NPOI.SS.UserModel;

namespace NiuX.Npoi.Extensions;

/// <summary>
/// Extension methods for <see cref="ICell"/>.
/// </summary>
public static class NpoiExtensions
{
    /// <summary>
    /// Firsts the sheet.
    /// </summary>
    /// <param name="workbook">The workbook.</param>
    /// <returns></returns>
    public static ISheet FirstSheet(this IWorkbook workbook) => workbook.GetSheetAt(0);

    /// <summary>
    /// Creates the row.
    /// </summary>
    /// <param name="sheet">The sheet.</param>
    /// <returns></returns>
    public static IRow CreateRow(this ISheet sheet) => sheet.CreateRow(0);

    /// <summary>
    /// Creates the cell.
    /// </summary>
    /// <param name="row">The row.</param>
    /// <returns></returns>
    public static ICell CreateCell(this IRow row) => row.CreateCell(0);
}