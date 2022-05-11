using System;

namespace NiuX.Npoi.Utils;

/// <summary>
/// A utility class for manipulating byte arrays.
/// </summary>
public static class ExceptionUtils
{
    /// <summary>
    /// Throws the argument null exception.
    /// </summary>
    /// <param name="predicate">The predicate.</param>
    /// <param name="message">The message.</param>
    /// <exception cref="System.ArgumentNullException"></exception>
    public static void ThrowArgumentException(Func<bool> predicate, string message)
    {
        if (predicate())
        {
            throw new ArgumentException(message);
        }
    }

    /// <summary>
    /// Throws the argument null exception.
    /// </summary>
    /// <param name="predicate">The predicate.</param>
    /// <param name="message">The message.</param>
    /// <exception cref="System.ArgumentNullException"></exception>
    public static void ThrowArgumentNullException(Func<bool> predicate, string message)
    {
        if (predicate())
        {
            throw new ArgumentNullException(message);
        }
    }

    /// <summary>
    /// Throw Argument Null Exception If Null
    /// </summary>
    /// <param name="argument"></param>
    /// <param name="argumentName"></param>
    /// <exception cref="ArgumentNullException"></exception>
    public static void ThrowArgumentNullExceptionIfNull(object? argument, string argumentName)
    {
        if (argument == null)
        {
            throw new ArgumentNullException(argumentName);
        }
    }

    /// <summary>
    /// Throws the invalid operation exception.
    /// </summary>
    /// <param name="argument">The argument.</param>
    /// <param name="message">The message.</param>
    /// <exception cref="System.InvalidOperationException"></exception>
    public static void ThrowInvalidOperationExceptionIfNull(object? argument, string message)
    {
        if (argument == null)
        {
            throw new InvalidOperationException(message);
        }

    }

    /// <summary>
    /// Throws the invalid operation exception.
    /// </summary>
    /// <param name="predicate"></param>
    /// <param name="message">The message.</param>
    /// <exception cref="System.InvalidOperationException"></exception>
    public static void ThrowInvalidOperationException(Func<bool> predicate, string message)
    {
        if (predicate())
        {
            throw new InvalidOperationException(message);
        }
    }

    /// <summary>
    /// Throws the invalid operation exception.
    /// </summary>
    /// <param name="message">The message.</param>
    /// <exception cref="System.InvalidOperationException"></exception>
    public static void ThrowInvalidOperationException(string message)
    {
        throw new InvalidOperationException(message);
    }
}