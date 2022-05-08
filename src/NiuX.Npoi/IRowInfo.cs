namespace NiuX.Npoi;

/// <summary>
/// Information required for one row when mapping between object and file rows.
/// </summary>
public interface IRowInfo
{
    int RowNumber { get; set; }

    /// <summary>
    /// Column index of the first error.
    /// </summary>
    int ErrorColumnIndex { get; set; }

    /// <summary>
    /// Error message for the first error.
    /// </summary>
    string ErrorMessage { get; set; }
}