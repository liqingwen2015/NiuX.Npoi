﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using NiuX.Npoi.Attributes;

namespace NiuX.Npoi.Tests.Models;

/// <summary>
/// Sample class for testing purpose.
/// </summary>
public class SampleClass : BaseClass
{
    public SampleClass()
    {
        CollectionGenericProperty = new List<string>();
        GeneralCollectionProperty = new List<string>();
    }

    public SampleClass(ICollection<string> collectionGenericProperty)
    {
        CollectionGenericProperty = collectionGenericProperty;
    }

    public string StringProperty { get; set; }

    public int Int32Property { get; set; }

    public bool BoolProperty { get; set; }

    public DateTime DateProperty { get; set; }

    public double DoubleProperty { get; set; }

    public SampleEnum EnumProperty { get; set; }

    public object ObjectProperty { get; set; }

    public ICollection<string> CollectionGenericProperty { get; set; }

    public string SingleColumnResolverProperty { get; set; }

    [Column("By Name")]
    public string ColumnNameAttributeProperty { get; set; }

    [Column(11)]
    public string ColumnIndexAttributeProperty { get; set; }

    public string IndexOverNameAttributeProperty { get; set; }

    [UseLastNonBlankValue]
    public string UseLastNonBlankValueAttributeProperty { get; set; }

    [Ignore]
    public string IgnoredAttributeProperty { get; set; }

    [Display(Name = "Display Name")]
    public string DisplayNameProperty { get; set; }

    public string GeneralProperty { get; set; }

    public ICollection<string> GeneralCollectionProperty { get; set; }

    [Column(CustomFormat = "0%")]
    public double CustomFormatProperty { get; set; }
}

/// <summary>
/// The base class for sample classes.
/// </summary>
public class BaseClass
{
    public string BaseStringProperty { get; set; }

    [Ignore]
    public string BaseIgnoredProperty { get; set; }
}

/// <summary>
/// Sample enum for testing purpose.
/// </summary>
public enum SampleEnum
{
    Value1 = 0,
    Value2 = 1,
    Value3 = 2,
}