using System;

namespace NiuX.Npoi.Attributes;

/// <summary>
///    Defines the name of the column.
/// </summary>
/// <seealso cref="System.Attribute" />
[AttributeUsage(AttributeTargets.All, AllowMultiple = false)]
public class UseLastNonBlankValueAttribute : Attribute
{

}