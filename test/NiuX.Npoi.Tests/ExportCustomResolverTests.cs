
using NiuX.Npoi.Extensions;
using NiuX.Npoi.Tests.Models;
using Xunit;

namespace NiuX.Npoi.Tests;

public class ExportCustomResolverTests : NpoiTestBase
{
    [Fact]
    public void ForHeader_ChangeHeaderStyle_ShouldChanged()
    {
        // Arrange
        const string str1 = "aBc";
        var workbook = GetBlankWorkbook();
        var row1 = workbook.GetSheetAt(0).CreateRow();
        row1.CreateCell(11).SetCellValue("StringProperty");

        var mapper = new Mapper(workbook);

        // Act
        mapper.ForHeader(cell =>
        {
            if (cell.ColumnIndex == 11)
            {
                var style = cell.Sheet.Workbook.CreateCellStyle();
                style.LeftBorderColor = 120;
                cell.CellStyle = style;
            }
        });
        mapper.Put(new[] { new SampleClass { StringProperty = str1 } });

        // Assert
        var row2 = workbook.GetSheetAt(0).GetRow(1);
        Assert.Equal(120, row1.GetCell(11).CellStyle.LeftBorderColor);
        Assert.Equal(str1, row2.GetCell(11).StringCellValue);

    }
}