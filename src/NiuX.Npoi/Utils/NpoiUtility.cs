using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using NiuX.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Threading.Tasks;
using NiuX.Npoi.Extensions;

namespace NiuX.Npoi.Utils;

public static class NpoiUtility
{
    /// <summary>
    /// 读取 Excel 并转换为 DataTable
    /// </summary>
    /// <returns></returns>
    public static DataTable ReadAsDataTable(string filePath)
    {
        var dataTable = new DataTable();
        var rowList = new List<string>();

        using var stream = new FileStream(filePath, FileMode.Open);
        stream.Position = 0;

        var xssWorkbook = new XSSFWorkbook(stream);
        var sheet = xssWorkbook.GetSheetAt(0);
        var headerRow = sheet.GetRow(0);
        int cellCount = headerRow.LastCellNum;

        for (int j = 0; j < cellCount; j++)
        {
            var cell = headerRow.GetCell(j);
            if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
            {
                dataTable.Columns.Add(cell.ToString());
            }
        }
        for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);

            if (row == null) continue;
            if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

            for (int j = row.FirstCellNum; j < cellCount; j++)
            {
                if (row.GetCell(j) != null)
                {
                    if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                    {
                        rowList.Add(row.GetCell(j).ToString());
                    }
                }
            }

            if (rowList.Count > 0)
                dataTable.Rows.Add(rowList.ToArray());
            rowList.Clear();
        }

        return dataTable;
    }

#if NETSTANDARD2_1_OR_GREATER
    
    /// <summary>
    /// 读取 Excel 并转换为 DataTable
    /// </summary>
    /// <returns></returns>
    public static async Task<DataTable> ReadAsDataTableAsync(string filePath)
    {
        var dataTable = new DataTable();
        var rowList = new List<string>();

        await using var stream = new FileStream(filePath, FileMode.Open);
        stream.Position = 0;

        var xssWorkbook = new XSSFWorkbook(stream);
        var sheet = xssWorkbook.GetSheetAt(0);
        var headerRow = sheet.GetRow(0);
        int cellCount = headerRow.LastCellNum;

        for (int j = 0; j < cellCount; j++)
        {
            var cell = headerRow.GetCell(j);
            if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
            {
                dataTable.Columns.Add(cell.ToString());
            }
        }
        for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);

            if (row == null) continue;
            if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

            for (int j = row.FirstCellNum; j < cellCount; j++)
            {
                if (row.GetCell(j) != null)
                {
                    if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                    {
                        rowList.Add(row.GetCell(j).ToString());
                    }
                }
            }

            if (rowList.Count > 0)
                dataTable.Rows.Add(rowList.ToArray());
            rowList.Clear();
        }

        return dataTable;
    }

    public static async Task<List<T>> ReadAsListAsync<T>(string filePath) => (await ReadAsDataTableAsync(filePath)).ToJson().FromJson<List<T>>();


#endif




    public static string ReadAsJson(string filePath) => ReadAsDataTable(filePath).ToJson();

    public static List<T> ReadAsList<T>(string filePath) => ReadAsDataTable(filePath).ToJson().FromJson<List<T>>();

    /// <summary>
    /// 写入 Excel
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="filePath"></param>
    /// <param name="items"></param>
    /// <param name="sheetName"></param>
    public static void Write<T>(string filePath, IEnumerable<T> items, string sheetName = "sheet1")
    {
        // Lets converts our object data to Datatable for a simplified logic.
        // Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

        var table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(items), typeof(DataTable))!;

        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);

        IWorkbook workbook = new XSSFWorkbook();
        var excelSheet = workbook.CreateSheet(sheetName);

        var columns = new List<string>();
        var row = excelSheet.CreateRow(0);
        var columnIndex = 0;

        foreach (DataColumn column in table.Columns)
        {
            columns.Add(column.ColumnName);
            row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
            columnIndex++;
        }

        var rowIndex = 1;
        foreach (DataRow dsrow in table.Rows)
        {
            row = excelSheet.CreateRow(rowIndex);
            var cellIndex = 0;

            foreach (var col in columns)
            {
                row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                cellIndex++;
            }

            rowIndex++;
        }

        workbook.Write(fs);
    }

    /// <summary>
    /// 写入 Excel
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="filePath"></param>
    /// <param name="items"></param>
    /// <param name="sheetName"></param>
    public static void Write<T>(string filePath, List<T> items, string sheetName = "sheet1")
    {

        // Lets converts our object data to Datatable for a simplified logic.
        // Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

        var table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(items), typeof(DataTable))!;

        using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);

        IWorkbook workbook = new XSSFWorkbook();
        var excelSheet = workbook.CreateSheet(sheetName);

        var columns = new List<string>();
        var row = excelSheet.CreateRow();
        var columnIndex = 0;

        foreach (DataColumn column in table.Columns)
        {
            columns.Add(column.ColumnName);
            row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
            columnIndex++;
        }

        var rowIndex = 1;
        foreach (DataRow dsrow in table.Rows)
        {
            row = excelSheet.CreateRow(rowIndex);
            var cellIndex = 0;

            foreach (var col in columns)
            {
                row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                cellIndex++;
            }

            rowIndex++;
        }

        workbook.Write(fs);
    }

#if NETSTANDARD2_1_OR_GREATER

    /// <summary>
    /// 写入 Excel
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="filePath"></param>
    /// <param name="items"></param>
    /// <param name="sheetName"></param>
    public static async Task WriteAsync<T>(string filePath, List<T> items, string sheetName = "sheet1")
    {

        // Lets converts our object data to Datatable for a simplified logic.
        // Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

        var table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(items), typeof(DataTable))!;

        await using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);

        IWorkbook workbook = new XSSFWorkbook();
        var excelSheet = workbook.CreateSheet(sheetName);

        var columns = new List<string>();
        var row = excelSheet.CreateRow(0);
        var columnIndex = 0;

        foreach (DataColumn column in table.Columns)
        {
            columns.Add(column.ColumnName);
            row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
            columnIndex++;
        }

        var rowIndex = 1;
        foreach (DataRow dsrow in table.Rows)
        {
            row = excelSheet.CreateRow(rowIndex);
            var cellIndex = 0;

            foreach (var col in columns)
            {
                row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                cellIndex++;
            }

            rowIndex++;
        }

        workbook.Write(fs);
    }

    /// <summary>
    /// 写入 Excel
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="filePath"></param>
    /// <param name="items"></param>
    /// <param name="sheetName"></param>
    public static async Task WriteAsync<T>(string filePath, IEnumerable<T> items, string sheetName = "sheet1")
    {

        // Lets converts our object data to Datatable for a simplified logic.
        // Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

        var table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(items), typeof(DataTable))!;

        await using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);

        IWorkbook workbook = new XSSFWorkbook();
        var excelSheet = workbook.CreateSheet(sheetName);

        var columns = new List<string>();
        var row = excelSheet.CreateRow(0);
        var columnIndex = 0;

        foreach (DataColumn column in table.Columns)
        {
            columns.Add(column.ColumnName);
            row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
            columnIndex++;
        }

        var rowIndex = 1;
        foreach (DataRow dsrow in table.Rows)
        {
            row = excelSheet.CreateRow(rowIndex);
            var cellIndex = 0;

            foreach (var col in columns)
            {
                row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                cellIndex++;
            }

            rowIndex++;
        }

        workbook.Write(fs);
    }

#endif


}