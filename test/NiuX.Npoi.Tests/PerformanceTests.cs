﻿using System;
using System.Diagnostics;
using System.Linq;
using NiuX;
using NiuX.Npoi;
using Xunit;

namespace NiuX.Npoi.Tests;

public class PerformanceTests : NpoiTestBase
{
    [Fact]
    public void TakeDynamic_Performance_Tests()
    {
        // Arrange
        int count = 100;
        //int count=10_000;
        //int count=1_000_000;
        var watch = new Stopwatch();
        var now = DateTime.Now;
        var workbook = GetBlankWorkbook();
        var sheet = workbook.GetSheetAt(0);
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("string");
        header.CreateCell(1).SetCellValue("int");
        header.CreateCell(2).SetCellValue("date");

        for (var i = 1; i <= count; i++)
        {
            var row = sheet.CreateRow(i);
            row.CreateCell(0).SetCellValue("this is a dummy string!");
            row.CreateCell(1).SetCellValue(i);
            row.CreateCell(2).SetCellValue(now.AddSeconds(i));
        }

        var mapper = new Mapper(workbook);

        // Act
        watch.Start();
        var objs = mapper.Take<dynamic>().ToList();
        watch.Stop();

        // Assert
        Trace.WriteLine($"Total Row:{count:0000000} - {watch.ElapsedMilliseconds} ms");
        Assert.Equal(count, objs.Count);
    }
}