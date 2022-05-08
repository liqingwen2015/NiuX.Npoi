using System;

namespace NiuX.Npoi.Attributes;

/// <summary>
/// Specifies to ignore a property for mapping.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class IgnoreAttribute : Attribute
{
}