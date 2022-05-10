using System;

namespace NiuX.Npoi.Utils;

public static class ExceptionUtils
{
    /// <summary>
    /// Throw Argument Null Exception If Null
    /// </summary>
    /// <param name="argument"></param>
    /// <param name="argumentName"></param>
    /// <exception cref="ArgumentNullException"></exception>
    public static void ThrowArgumentNullExceptionIfNull(object argument, string argumentName)
    {
        if (argument == null)
        {
            throw new ArgumentNullException(argumentName);
        }
    }

    public static void ThrowInvalidOperationException(string message)
    {
        throw new InvalidOperationException(message);
    }
}