using UniverBlazored.Spreadsheets.Data.ConditionFormat;
using UniverBlazored.Spreadsheets.Data.Workbook;

namespace UniverBlazored.SpreadsheetConverter.Services.IO.Data;

/// <summary>
/// Data class to store conditional format information for a worksheet, including the ranges it applies to, the type of conditional format, and the style to be applied when the condition is met.
/// </summary>
public class UCondFormatData
{
    /// <summary>
    /// Ranges for the conditional format to apply
    /// </summary>
    /// <value></value>
    public List<URange> Ranges { get; set; } = new();

    /// <summary>
    /// Type of condition
    /// </summary>
    /// <value></value>
    public UConditionType Type { get; set; }

    /// <summary>
    /// Style to be applied when the condition is met
    /// </summary>
    /// <value></value>
    public UConditionFormatStyle Style { get; set; }

    /// <summary>
    /// Data class to store conditional format information for a worksheet, including the ranges it applies to, the type of conditional format, and the style to be applied when the condition is met.
    /// </summary>
    public UCondFormatData()
    {
        Style = new UConditionFormatStyle();
    }

    /// <summary>
    /// Data class to store conditional format information for a worksheet, including the ranges it applies to, the type of conditional format, and the style to be applied when the condition is met.
    /// </summary>
    /// <param name="ranges">Ranges for the conditional format to apply</param>
    /// <param name="type">Type of condition</param>
    /// <param name="style">Style to be applied when the condition is met</param>
    public UCondFormatData(List<URange> ranges, UConditionType type, UConditionFormatStyle style)
    {
        Ranges = ranges;
        Type = type;
        Style = style;
    }
}
