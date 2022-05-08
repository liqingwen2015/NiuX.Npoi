using System;
using System.IO;
using System.Linq;
using NiuX;
using NiuX.Extensions;
using NiuX.Npoi;
using NiuX.Npoi.Extensions;
using NiuX.Npoi.Tests.Models;
using NiuX.Utils;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace NiuX.Npoi.Tests;

public class GeneralExportTests : NpoiTestBase
{
    SampleClass sampleObj = new SampleClass
    {
        ColumnIndexAttributeProperty = "Column Index",
        CustomFormatProperty = 0.87,
        DateProperty = DateTime.Now,
        DoubleProperty = 78,
        GeneralProperty = "general sting",
        StringProperty = "balabala",
        BoolProperty = true,
        EnumProperty = SampleEnum.Value3,
        IgnoredAttributeProperty = "Ignored column",
        Int32Property = 100,
        SingleColumnResolverProperty = "I'm here..."
    };

    private class DummyClass
    {
        public string String { get; set; }
        public DateTime DateTime { get; set; }
        public double Double { get; set; }
        public DateTime DateTime2 { get; set; }
    }

    private DummyClass dummyObj = new DummyClass
    {
        String = "My string",
        DateTime = DateTime.Now,
        Double = 0.4455,
        DateTime2 = DateTime.Now.AddDays(1)
    };

    const string FileName = "test.xlsx";


    [Fact]
    public void SaveSheetWithoutAnyMapping()
    {
        // Arrange
        var exporter = new Mapper();
        //var sheetName = "newSheet";
        FileUtility.Delete(FileName);

        // Act
        exporter.Save(FileName, new[] { dummyObj });
        var dateCell = exporter.Workbook.FirstSheet().GetRow(1).GetCell(1);

        // Assert
        Assert.NotNull(exporter.Workbook);
        Assert.Equal(2, exporter.Workbook.FirstSheet().PhysicalNumberOfRows);
        Assert.True(DateUtil.IsCellDateFormatted(dateCell));
        Assert.Equal(dummyObj.String, exporter.Take<DummyClass>().First().Value.String);
        Assert.Equal(dummyObj.Double, exporter.Take<DummyClass>().First().Value.Double);
    }

    [Fact]
    public void SaveSheetUseFormat()
    {
        // Arrange
        var exporter = new Mapper();

        var dateFormat = "yyyy.MM.dd hh.mm.ss";
        var doubleFormat = "0%";
        FileUtility.Delete(FileName);

        // Act
        exporter.UseFormat(typeof(DateTime), dateFormat);
        exporter.UseFormat(typeof(double), doubleFormat);
        exporter.Save(FileName, new[] { dummyObj });
        var items = exporter.Take<DummyClass>().ToList();
        var dateCell = exporter.Workbook.GetSheetAt(0).GetRow(1).GetCell(1);

        // Assert
        Assert.Equal(2, exporter.Workbook.FirstSheet().PhysicalNumberOfRows);
        Assert.True(DateUtil.IsCellDateFormatted(dateCell));
        Assert.Equal(dummyObj.DateTime.ToLongDateString(), items.First().Value.DateTime.ToLongDateString());
        Assert.Equal(dummyObj.Double, items.First().Value.Double);
        Assert.Equal(dummyObj.DateTime2.ToLongDateString(), items.First().Value.DateTime2.ToLongDateString());
    }

    [Fact]
    public void SaveSheet_UseFormat_ForNullable()
    {
        // Arrange
        var exporter = new Mapper();
        var dateFormat = "yyyy.MM.dd hh.mm.ss";
        var obj1 = new NullableClass { NullableDateTime = null, DummyString = "dummy" };
        var obj2 = new NullableClass { NullableDateTime = DateTime.Now };
        FileUtility.Delete(FileName);

        // Act
        exporter.UseFormat(typeof(DateTime?), dateFormat);

        // Issue #5, if the first data row has null value, then next rows will not be formated
        // So here we make the first date row has a null value for DateTime? property.
        exporter.Save(FileName, new[] { obj1, obj2 });

        var items = exporter.Take<NullableClass>().ToList();
        var dateCell = exporter.Workbook.FirstSheet().GetRow(2).GetCell(0);

        // Assert
        Assert.Equal(3, exporter.Workbook.FirstSheet().PhysicalNumberOfRows);
        Assert.Equal(obj1.DummyString, items.First().Value.DummyString);
        Assert.Equal(obj2.NullableDateTime.Value.ToLongDateString(), items.Skip(1).First().Value.NullableDateTime.Value.ToLongDateString());
        Assert.True(DateUtil.IsCellDateFormatted(dateCell));
        Assert.Equal(obj2.NullableDateTime.Value.ToLongDateString(), items.Skip(1).First().Value.NullableDateTime.Value.ToLongDateString());
        Assert.False(exporter.Take<NullableClass>().First().Value.NullableDateTime.HasValue);

        FileName.AsOpen();
    }

    [Fact]
    public void SaveSheet()
    {
        // Prepare
        var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
        var exporter = new Mapper(workbook);
        if (File.Exists(FileName)) File.Delete(FileName);
        var objs = exporter.Take<SampleClass>(1).ToList();

        // Act
        exporter.Save<SampleClass>(FileName, 1);

        // Assert
        Assert.NotNull(objs);
        Assert.NotNull(exporter);
        Assert.NotNull(exporter.Workbook);
    }


    [Fact]
    public void SaveObjects()
    {
        // Prepare
        var exporter = new Mapper();
        exporter.Map<SampleClass>("General Column", o => o.GeneralProperty);
        if (File.Exists(FileName)) File.Delete(FileName);

        // Act
        exporter.Save(FileName, new[] { sampleObj }, "newSheet");

        // Assert
        Assert.NotNull(exporter.Workbook);
        Assert.Equal(2, exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows);
    }

    [Fact]
    public void SaveTrackedObjectsTest()
    {
        // Prepare
        var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
        var exporter = new Mapper(workbook);
        FileUtility.Delete(FileName);
        var objs = exporter.Take<SampleClass>(1).ToList();

        // Act
        exporter.Save(FileName, objs.Select(info => info.Value), "newSheet");

        // Assert
        Assert.NotNull(exporter.Workbook);
        Assert.Equal(2, exporter.Workbook.GetSheet("newSheet").PhysicalNumberOfRows);
    }

    [Fact]
    public void FormatAttribute()
    {
        // Prepare
        var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
        var exporter = new Mapper(workbook);
        FileUtility.Delete(FileName);
        var objs = exporter.Take<SampleClass>(1).ToList();
        objs[0].Value.CustomFormatProperty = 100.234;

        // Act
        exporter.Map<SampleClass>(12, o => o.CustomFormatProperty);
        exporter.Save(FileName, objs.Select(info => info.Value), "newSheet");

        // Assert
        var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
        Assert.NotNull(exporter.Workbook);
        Assert.NotEqual(0, doubleStyle.DataFormat);
    }

    [Fact]
    public void FormatMethod()
    {
        // Prepare
        var workbook = GetSimpleWorkbook(DateTime.Now, "aBC");
        var exporter = new Mapper(workbook);
        FileUtility.Delete(FileName);
        var objs = exporter.Take<SampleClass>(1).ToList();
        objs[0].Value.DoubleProperty = 100.234;

        // Act
        exporter.Map<SampleClass>(11, o => o.DateProperty);
        exporter.Map<SampleClass>(12, o => o.DoubleProperty);
        exporter.Format<SampleClass>("0%", o => o.DoubleProperty);
        exporter.Save(FileName, objs.Select(info => info.Value), "newSheet");

        // Assert
        var doubleStyle = exporter.Workbook.GetSheet("newSheet").GetRow(1).GetCell(12).CellStyle;
        Assert.NotNull(exporter.Workbook);
        Assert.NotEqual(0, doubleStyle.DataFormat);
    }


    [Fact]
    public void NoHeader()
    {
        // Prepare
        var exporter = new Mapper { HasHeader = false };
        const string sheetName = "newSheet";
        FileUtility.Delete(FileName);

        // Act
        exporter.Save(FileName, new[] { sampleObj, }, sheetName);

        // Assert
        Assert.NotNull(exporter.Workbook);
        Assert.Equal(1, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);
    }


    [Fact]
    public void ExportXls()
    {
        // Prepare
        const string existingFile = "Book2.xlsx";
        const string sheetName = "newSheet";
        FileUtility.Delete(FileName);

        File.Copy("Book1.xlsx", existingFile);
        var exporter = new Mapper();

        // Act
        exporter.Save(existingFile, new[] { sampleObj, }, sheetName, true, false);

        // Assert
        Assert.NotNull(exporter.Workbook as HSSFWorkbook);
        Assert.Equal(2, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);
    }


    [Fact]
    public void OverwriteNewFile()
    {
        // Prepare
        const string existingFile = "Book2.xlsx";
        const string sheetName = "Allocations";
        FileUtility.Delete(FileName);
        File.Copy("Book1.xlsx", existingFile);
        var exporter = new Mapper();

        // Act
        exporter.Save(existingFile, new[] { sampleObj, }, sheetName, true);

        // Assert
        Assert.Equal(1, exporter.Workbook.NumberOfSheets);
    }

    [Fact]
    public void MergeToExistedRows()
    {
        // Prepare
        const string existingFile = "Book2.xlsx";
        const string sheetName = "Allocations";
        FileUtility.Delete(FileName);
        File.Copy("Book1.xlsx", existingFile);
        var exporter = new Mapper();
        exporter.Map<SampleClass>("Project Name", o => o.GeneralProperty);
        exporter.Map<SampleClass>("Allocation Month", o => o.DateProperty);

        // Act
        exporter.Save(existingFile, new[] { sampleObj, }, sheetName, false);

        // Assert
        var sheet = exporter.Workbook.GetSheet(sheetName);
        Assert.Equal(sampleObj.GeneralProperty, sheet.GetRow(4).GetCell(1).StringCellValue);
        Assert.Equal(sampleObj.DateProperty.Date, sheet.GetRow(4).GetCell(2).DateCellValue.Date);
    }

    [Fact]
    public void PutAppendRow()
    {
        // Prepare
        const string existingFile = "Book2.xlsx";
        const string sheetName = "Allocations";
        FileUtility.Delete(FileName);
        File.Copy("Book1.xlsx", existingFile);
        var exporter = new Mapper(existingFile);
        exporter.Map<SampleClass>("Project Name", o => o.GeneralProperty);
        exporter.Map<SampleClass>("Allocation Month", o => o.DateProperty);

        // Act
        exporter.Put(new[] { sampleObj, }, sheetName, false);
        var workbook = WriteAndReadBack(exporter.Workbook, existingFile);

        // Assert
        var sheet = workbook.GetSheet(sheetName);
        Assert.Equal(sampleObj.GeneralProperty, sheet.GetRow(4).GetCell(1).StringCellValue);
        Assert.Equal(sampleObj.DateProperty.Date, sheet.GetRow(4).GetCell(2).DateCellValue.Date);
    }

    [Fact]
    public void PutOverwriteRow()
    {
        // Prepare
        const string existingFile = "Book3.xlsx";
        const string sheetName = "Allocations";
        FileUtility.Delete(FileName);
        File.Copy("Book1.xlsx", existingFile);

        var exporter = new Mapper(existingFile);
        exporter.Map<SampleClass>("Project Name", o => o.GeneralProperty);
        exporter.Map<SampleClass>("Allocation Month", o => o.DateProperty);
        exporter.Map<SampleClass>("Name", o => o.StringProperty);
        exporter.Map<SampleClass>("email", o => o.BoolProperty);

        // Act
        exporter.Put(new[] { sampleObj, }, sheetName, true);
        exporter.Put(new[] { sampleObj }, "Resources");
        var workbook = WriteAndReadBack(exporter.Workbook, existingFile);

        // Assert
        var sheet = workbook.GetSheet(sheetName);
        Assert.Equal(sampleObj.GeneralProperty, sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal(sampleObj.DateProperty.Date, sheet.GetRow(1).GetCell(2).DateCellValue.Date);
    }

    [Fact]
    public void SaveWorkbookToFileTest()
    {
        // Prepare
        const string fileName = "temp4.xlsx";
        FileUtility.Delete(FileName);

        var exporter = new Mapper("Book1.xlsx");

        // Act
        exporter.Save(fileName);

        // Assert
        Assert.True(File.Exists(fileName));
        File.Delete(fileName);
    }

    [Fact]
    public void PutWithNotExistedSheetIndex_ShouldAutoPopulateSheets()
    {
        // Arrange
        var workbook = GetEmptyWorkbook();

        var mapper = new Mapper(workbook);

        // Act
        mapper.Put(new[] { new object(), }, 100);

        // Assert
        Assert.True(workbook.NumberOfSheets > 0);
    }

    [Fact]
    public void PutWithNotExistedSheetName_ShouldAutoPopulateSheets()
    {
        // Arrange
        var workbook = GetEmptyWorkbook();

        var mapper = new Mapper(workbook);

        // Act
        mapper.Put(new[] { new object(), }, "sheet100");

        // Assert
        Assert.True(workbook.NumberOfSheets > 0);
    }

    [Fact]
    public void Map_WithIndexAndName_ShouldExportCustomColumnName()
    {
        // Arrange
        var workbook = GetEmptyWorkbook();
        const string nameString = "string";
        const string nameInt = "int";
        const string nameBool = "bool";
        var sheet = workbook.CreateSheet();

        var mapper = new Mapper(workbook);

        // Act
        mapper.Map<SampleClass>(0, o => o.StringProperty, nameString);
        mapper.Map<SampleClass>(1, o => o.Int32Property, nameInt);
        mapper.Map<SampleClass>(2, o => o.BoolProperty, nameBool);
        mapper.Put(new[] { new SampleClass(), }, 0);

        // Assert
        var row = sheet.GetRow(0);
        Assert.Equal(nameString, row.GetCell(0).StringCellValue);
        Assert.Equal(nameInt, row.GetCell(1).StringCellValue);
        Assert.Equal(nameBool, row.GetCell(2).StringCellValue);
    }


    [Fact]
    public void Put_WithFirstRowIndex_ShouldExportExpectedRows()
    {
        // Arrange
        var hasHeader = true;
        const int firstRowIndex = 100;
        const string nameString = "StringProperty";
        var workbook = GetEmptyWorkbook();
        var sheet = workbook.CreateSheet();

        var item = new SampleClass { StringProperty = nameString };
        var mapper = new Mapper(workbook) { HasHeader = hasHeader, FirstRowIndex = firstRowIndex };
        mapper.Map<SampleClass>(0, o => o.StringProperty, "a");

        // Act
        mapper.Put(new[] { item }, 0);

        // Assert
        var firstDataRowIndex = hasHeader ? firstRowIndex + 1 : firstRowIndex;
        var row = sheet.GetRow(firstDataRowIndex);
        Assert.Equal(1 + (hasHeader ? 1 : 0), sheet.PhysicalNumberOfRows);
        Assert.Equal(nameString, row.GetCell(0).StringCellValue);
    }

    private class NullableClass
    {
        public DateTime? NullableDateTime { get; set; }
        public string DummyString { get; set; }
    }
}