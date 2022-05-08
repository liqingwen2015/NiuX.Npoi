using System;

namespace NiuX.Npoi.Extensions
{
    /// <summary>
    /// Extension methods for format.
    /// </summary>
    public static class FormatExtensions
    {
        /// <summary>
        /// Uses a custom format for all properties that have the same type.
        /// </summary>
        /// <param name="mapper">The <see cref="Mapper"/> object.</param>
        /// <param name="propertyType">The type of property to format.</param>
        /// <param name="customFormat">The custom format for the specified type.</param>
        /// <returns>The <see cref="Mapper"/> itself.</returns>
        public static Mapper UseFormat(this Mapper mapper, Type propertyType, string customFormat)
        {
            if (mapper == null) throw new ArgumentNullException(nameof(mapper));
            if (propertyType == null) throw new ArgumentNullException(nameof(propertyType));
            if (customFormat.IsNullOrWhiteSpace()) throw new ArgumentException($"Parameter '{nameof(customFormat)}' cannot be null or white space.");

            mapper.TypeFormats[propertyType] = customFormat;

            return mapper;
        }
    }
}
