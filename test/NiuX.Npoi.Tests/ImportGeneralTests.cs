using System;
using System.IO;
using System.Linq;
using NiuX;
using NiuX.Npoi;
using NiuX.Npoi.Extensions;
using NiuX.Npoi.Tests.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace NiuX.Npoi.Tests;

public class ImportGeneralTests : NpoiTestBase
{
    private class TestClass
    {
        public string String { get; set; }
        public DateTime DateTime { get; set; }
        public double Double { get; set; }
    }

    private class NullableClass
    {
        public DateTime? NullableDateTime { get; set; }
        public string NormalString { get; set; }
    }

    [Fact]
    public void ImporterWithoutAnyMapping()
    {
        // Arrange
        var stream = new FileStream("Book1.xlsx", FileMode.Open);

        // Act
        var importer = new Mapper(stream);
        var items = importer.Take<TestClass>("TestClass").ToList();

        // Assert
        Assert.NotNull(importer);
        Assert.NotNull(importer.Workbook);
        Assert.Equal(3, items.Count);
        Assert.True(items[1].Value.DateTime.Year == 2017);
        Assert.True(Math.Abs(items[1].Value.Double - 1.2345) < 0.00001);
    }

    [Fact]
    public void ImporterWithFormat()
    {
        // Arrange
        var stream = new FileStream("Book1.xlsx", FileMode.Open);

        // Act
        var importer = new Mapper(stream);
        importer.UseFormat(typeof(DateTime), "MM^dd^yyyy");
        var items = importer.Take<TestClass>("TestClass").ToList();

        // Assert
        Assert.NotNull(importer);
        Assert.NotNull(importer.Workbook);
        Assert.Equal(3, items.Count);
        Assert.True(items[1].Value.DateTime.Year == 2017);
        Assert.True(Math.Abs(items[1].Value.Double - 1.2345) < 0.00001);
    }

    [Fact]
    public void Import_ParseStringToNullableDateTime_Success()
    {
        // Arrange
        var importer = new Mapper("Book1.xlsx");

        // Act
        importer.UseFormat(typeof(DateTime), "MM^dd^yyyy");
        var items = importer.Take<NullableClass>("NullableClass").ToList();

        // Assert
        Assert.True(items[0].Value.NullableDateTime.Value.Year == 2017);
        Assert.True(items[1].Value.NullableDateTime.Value.Year == 2017);
        Assert.True(items[2].Value.NullableDateTime.Value.Year == 2017);
    }

    [Fact]
    public void Import_ErrorOnNullable_GetNullObject()
    {
        // Arrange
        var importer = new Mapper("Book1.xlsx");

        // Act
        var items = importer.Take<NullableClass>("NullableClass").ToList();

        // Assert
        Assert.Equal(0, items[3].ErrorColumnIndex);
        Assert.Null(items[3].Value.NullableDateTime);
    }

    [Fact]
    public void Import_IgnoreErrorOnNullable_GetNullProperty()
    {
        // Arrange
        var importer = new Mapper("Book1.xlsx");

        // Act
        importer.IgnoreErrorsFor<NullableClass>(o => o.NullableDateTime);
        var items = importer.Take<NullableClass>("NullableClass").ToList();

        // Assert
        Assert.Null(items[2].Value.NullableDateTime);
        Assert.NotNull(items[2].Value.NormalString);
        Assert.Null(items[3].Value.NullableDateTime);
        Assert.NotNull(items[3].Value.NormalString);
    }

    [Fact]
    public void ImporterConstructorWorkbookTest()
    {
        // Arrange
        var workbook = GetSimpleWorkbook(DateTime.MaxValue, "dummy");

        // Act
        var importer = new Mapper(workbook);

        // Assert
        Assert.NotNull(importer);
        Assert.NotNull(importer.Workbook);
    }

    [Fact]
    public void ImporterConstructorNullStreamTest()
    {
        // Arrange
        Stream nullStream = null;

        // Act
        var action = () => new Mapper(nullStream);

        // Assert
        Assert.Throws<ArgumentNullException>(action);
    }

    [Fact]
    public void ImporterConstructorNullWorkbookTest()
    {
        // Arrange
        IWorkbook nullWorkbook = null;

        // Act
        var action = () => new Mapper(nullWorkbook);

        // Assert
        Assert.Throws<ArgumentNullException>(action);
    }

    [Fact]
    public void ImporterConstructorFilePathTest()
    {
        // Arrange

        // Act
        var importer = new Mapper("Book1.xlsx");


        // Assert
        Assert.NotNull(importer);
        Assert.NotNull(importer.Workbook);
    }

    [Fact]
    public void ImporterConstructorFilePathNotExistTest()
    {
        // Arrange

        // Act
        var action = () => new Mapper("dummy.txt");

        // Assert
        Assert.Throws<FileNotFoundException>(action);
    }

    [Fact]
    public void ImporterNoElementTest()
    {
        // Arrange
        var workbook = new XSSFWorkbook();
        var header = workbook.CreateSheet("sheet1").CreateRow(0);
        header.CreateCell(0).SetCellValue("StringProperty");
        header.CreateCell(1).SetCellValue("Int32Property");
        var importer = new Mapper(workbook);

        // Act
        var objs = importer.Take<SampleClass>(0);

        // Assert
        Assert.NotNull(objs);
        Assert.Equal(0, objs.Count());
    }

    [Fact]
    public void ImporterEmptySheetTest()
    {
        // Arrange
        var workbook = new XSSFWorkbook();
        workbook.CreateSheet("sheet1");
        var importer = new Mapper(workbook);

        // Act
        var objs = importer.Take<SampleClass>(0);

        // Assert
        Assert.NotNull(objs);
        Assert.Equal(0, objs.Count());
    }

    [Fact]
    public void TakeByHeaderIndexTest()
    {
        // Arrange
        var date = DateTime.Now;
        const string str = "aBC";
        var workbook = GetSimpleWorkbook(date, str);
        var importer = new Mapper(workbook);

        // Act
        var objs = importer.Take<SampleClass>(1).ToList();

        // Assert
        Assert.NotNull(objs);
        Assert.Equal(1, objs.Count);

        var obj = objs[0];
        var objDate = obj.Value.DateProperty;

        Assert.Equal(date.ToLongDateString(), objDate.ToLongDateString());
        Assert.Equal(str, obj.Value.StringProperty);
    }

    [Fact]
    public void TakeByHeaderIndexOutOfRangeTest()
    {
        // Arrange
        var date = DateTime.Now;
        const string str = "aBC";
        var workbook = GetSimpleWorkbook(date, str);
        var importer = new Mapper(workbook);

        // Act
        var action = () => importer.Take<SampleClass>(10);

        // Assert
        Assert.Throws<ArgumentException>(action);
    }

    [Fact]
    public void TakeByHeaderNameTest()
    {
        // Arrange
        var date = DateTime.Now;
        const string str = "aBC";
        var workbook = GetSimpleWorkbook(date, str);
        var importer = new Mapper(workbook);

        // Act
        var objs = importer.Take<SampleClass>("sheet2").ToList();

        // Assert
        Assert.NotNull(objs);
        Assert.Single(objs);

        var obj = objs[0];
        var objDate = obj.Value.DateProperty;

        Assert.Equal(date.ToLongDateString(), objDate.ToLongDateString());
        Assert.Equal(str, obj.Value.StringProperty);
    }

    [Fact]
    public void TakeByHeaderNameNotExistTest()
    {
        // Arrange
        var date = DateTime.Now;
        const string str = "aBC";
        var workbook = GetSimpleWorkbook(date, str);
        var importer = new Mapper(workbook);

        // Act
        var objs = importer.Take<SampleClass>("notExistSheet").ToList();

        // Assert
        Assert.NotNull(objs);
        Assert.Empty(objs);
    }

    [Fact]
    public void Import_ConvertValueError_GotErrorColumnIndex()
    {
        // Arrange
        const double dou1 = 1.833;
        const string str1 = "aBC";
        const string str2 = "BCD";
        var workbook = GetBlankWorkbook();
        var sheet = workbook.GetSheetAt(0);
        sheet.CreateRow(0);
        sheet.CreateRow(1);
        sheet.CreateRow(2);

        // Header row
        sheet.GetRow(0).CreateCell(0).SetCellValue("DoubleProperty");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Int32Property");
        sheet.GetRow(0).CreateCell(2).SetCellValue("StringProperty");

        // Row #1
        sheet.GetRow(1).CreateCell(0).SetCellValue(dou1);
        sheet.GetRow(1).CreateCell(1).SetCellValue((string)null);
        sheet.GetRow(1).CreateCell(2).SetCellValue(str1);

        // Row #2
        sheet.GetRow(2).CreateCell(0).SetCellValue(dou1);
        sheet.GetRow(2).CreateCell(1).SetCellValue("dummy");
        sheet.GetRow(2).CreateCell(2).SetCellValue(str2);
        var mapper = new Mapper(workbook);

        // Act
        var items = mapper.Take<SampleClass>().ToList();

        // Assert
        Assert.Equal(default, items[0].Value.Int32Property);
        Assert.Equal(1, items[1].ErrorColumnIndex);
    }

    [Fact]
    public void Import_IgnoreValueTypeParseError_GetDefaultProperty()
    {
        // Arrange
        const double dou1 = 1.833;
        const int int1 = 22;
        const string str1 = "aBC";
        const string str2 = "BCD";
        var workbook = GetBlankWorkbook();
        var sheet = workbook.GetSheetAt(0);
        sheet.CreateRow(0);
        sheet.CreateRow(1);
        sheet.CreateRow(2);

        // Header row
        sheet.GetRow(0).CreateCell(0).SetCellValue("DoubleProperty");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Int32Property");
        sheet.GetRow(0).CreateCell(2).SetCellValue("StringProperty");

        // Row #1
        sheet.GetRow(1).CreateCell(0).SetCellValue(int1.ToString());
        sheet.GetRow(1).CreateCell(1).SetCellValue(dou1.ToString("f3"));
        sheet.GetRow(1).CreateCell(2).SetCellValue(str1);

        // Row #2
        sheet.GetRow(2).CreateCell(0).SetCellValue("dummy");
        sheet.GetRow(2).CreateCell(1).SetCellValue("dummy");
        sheet.GetRow(2).CreateCell(2).SetCellValue(str2);
        var mapper = new Mapper(workbook);

        // Act
        mapper.IgnoreErrorsFor<SampleClass>(o => o.DoubleProperty);
        mapper.IgnoreErrorsFor<SampleClass>(o => o.Int32Property);
        var items = mapper.Take<SampleClass>().ToList();

        // Assert
        Assert.Equal(int1, items[0].Value.DoubleProperty);
        Assert.Equal(Math.Round(dou1), items[0].Value.Int32Property);
        Assert.Equal(str1, items[0].Value.StringProperty);
        Assert.Equal(default, items[1].Value.DoubleProperty);
        Assert.Equal(default, items[1].Value.Int32Property);
        Assert.Equal(str2, items[1].Value.StringProperty);
    }

    [Fact]
    public void Import_ValidEnum_ShouldWork()
    {
        // Arrange
        var workbook = GetBlankWorkbook();
        var sheet = workbook.GetSheetAt(0);
        sheet.CreateRow(0);
        sheet.CreateRow(1);
        sheet.CreateRow(2);
        sheet.CreateRow(3);

        // Header row
        sheet.GetRow(0).CreateCell(0).SetCellValue("EnumProperty");

        // Row #1
        sheet.GetRow(1).CreateCell(0).SetCellValue(SampleEnum.Value1.ToString());

        // Row #2
        sheet.GetRow(2).CreateCell(0).SetCellValue(SampleEnum.Value2.ToString());

        // Row #3
        sheet.GetRow(3).CreateCell(0).SetCellValue("value3");

        var mapper = new Mapper(workbook);

        // Act
        var items = mapper.Take<SampleClass>().ToList();

        // Assert
        Assert.Equal(SampleEnum.Value1, items[0].Value.EnumProperty);
        Assert.Equal(SampleEnum.Value2, items[1].Value.EnumProperty);
        Assert.Equal(SampleEnum.Value3, items[2].Value.EnumProperty);
    }

    [Fact]
    public void Map_WithIndexAndName_ShouldImportByIndex()
    {
        // Arrange
        var workbook = GetEmptyWorkbook();
        const string nameString = "StringProperty";
        const string nameGeneral = "GeneralProperty";
        var sheet = workbook.CreateSheet();

        var headerRow = sheet.CreateRow(0);
        headerRow.CreateCell(0).SetCellValue(nameGeneral);
        headerRow.CreateCell(1).SetCellValue(nameString);

        var row1 = sheet.CreateRow(1);
        row1.CreateCell(0).SetCellValue("a");
        row1.CreateCell(1).SetCellValue("b");

        var mapper = new Mapper(workbook);

        // Act
        mapper.Map<SampleClass>(0, "StringProperty", nameString);
        mapper.Map<SampleClass>(1, "GeneralProperty", nameGeneral);
        var obj = mapper.Take<SampleClass>().Select(o => o.Value).ToArray()[0];

        // Assert
        Assert.Equal("a", obj.StringProperty);
        Assert.Equal("b", obj.GeneralProperty);
    }

    [Fact]
    public void Take_WithFirstRowIndex_ShouldImportExpectedRows()
    {
        // Arrange
        bool hasHeader = true;
        const int firstRowIndex = 100;
        const string sheetName = "sheet2";
        var workbook = GetSimpleWorkbook(DateTime.Now, "a");
        const string nameString = "StringProperty";
        const string nameGeneral = "GeneralProperty";
        var sheet = workbook.GetSheet(sheetName);

        if (hasHeader)
        {
            var headerRow = sheet.CreateRow(firstRowIndex);
            headerRow.CreateCell(0).SetCellValue(nameGeneral);
            headerRow.CreateCell(1).SetCellValue(nameString);
        }

        var firstDataRowIndex = hasHeader ? firstRowIndex + 1 : firstRowIndex;
        var row1 = sheet.CreateRow(firstDataRowIndex);
        row1.CreateCell(0).SetCellValue("a");
        row1.CreateCell(1).SetCellValue("b");
        var row2 = sheet.CreateRow(firstDataRowIndex + 1);
        row2.CreateCell(0).SetCellValue("c");
        row2.CreateCell(1).SetCellValue("d");

        var mapper = new Mapper(workbook) { HasHeader = hasHeader, FirstRowIndex = firstRowIndex };
        mapper.Map<SampleClass>(0, o => o.GeneralProperty);
        mapper.Map<SampleClass>(1, o => o.StringProperty);

        // Act
        var obj = mapper.Take<SampleClass>(sheetName).ToList();

        // Assert
        Assert.Equal(2, obj.Count);
        Assert.Equal("a", obj[0].Value.GeneralProperty);
        Assert.Equal("b", obj[0].Value.StringProperty);
        Assert.Equal("c", obj[1].Value.GeneralProperty);
        Assert.Equal("d", obj[1].Value.StringProperty);
    }

    private class TestTrimClass
    {
        public string StringProperty { get; set; }
    }

    [Fact]
    public void Take_MapByNameAndExtraSpaceInExcelColumnName_MapsAsTrimmed()
    {
        // Arrange
        const string str1 = "aBC";
        const string str2 = "BCD";
        var workbook = GetBlankWorkbook();
        var sheet = workbook.GetSheetAt(0);
        sheet.CreateRow(0);
        sheet.CreateRow(1);
        sheet.CreateRow(2);

        // Header row with extra spaces
        sheet.GetRow(0).CreateCell(0).SetCellValue(" Name  ");

        // Row #1
        sheet.GetRow(1).CreateCell(0).SetCellValue(str1);

        // Row #2
        sheet.GetRow(2).CreateCell(0).SetCellValue(str2);
        var mapper = new Mapper(workbook);
        mapper.Map<TestTrimClass>("Name", o => o.StringProperty);

        // Act
        var items = mapper.Take<TestTrimClass>().ToList();

        // Assert
        Assert.Equal(str1, items[0].Value.StringProperty);
        Assert.Equal(str2, items[1].Value.StringProperty);
    }

    private class TestGuidClass
    {
        public Guid ID { get; set; }
    }

    [Fact]
    public void Take_GuidColumn_ParseAndSetGuid()
    {
        // Arrange
        var id = Guid.NewGuid();
        var workbook = GetBlankWorkbook();
        var sheet = workbook.GetSheetAt(0);
        sheet.CreateRow(0);
        sheet.CreateRow(1);

        // Header row with extra spaces
        sheet.GetRow(0).CreateCell(0).SetCellValue("ID");

        // Row #1
        sheet.GetRow(1).CreateCell(0).SetCellValue(id.ToString());

        var mapper = new Mapper(workbook);

        // Act
        var items = mapper.Take<TestGuidClass>().ToList();

        // Assert
        Assert.Equal(id, items[0].Value.ID);
    }
}