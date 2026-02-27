using UniverBlazored.Spreadsheets.Data.Styles;
using UniverBlazored.Spreadsheets.Data.Workbook;

namespace UniverBlazored.SpreadsheetConverter.Services.IO.Data;

/// <summary>
/// Store styling information for a given range in a worksheet.
/// </summary>
public class URangeStyles
{
    /// <summary>
    /// Ranges for the styles to apply
    /// </summary>
    /// <value></value>
    public List<URange> Ranges { get; set; } = new();

    /// <summary>
    /// Properties for the font style to apply to the ranges
    /// </summary>
    /// <value></value>
    public UFontProperties FontProperties { get; set; }

    /// <summary>
    /// Border information for each range
    /// </summary>
    public List<(EBorderType type, EBorderStyleType style, string color)> Borders { get; set; } = new();

    /// <summary>
    /// Store styling information for a given range in a worksheet.
    /// </summary>
    public URangeStyles()
    {
        FontProperties = new UFontProperties();
    }

    /// <summary>
    /// Store styling information for a given range in a worksheet.
    /// </summary>
    /// <param name="ranges">Ranges for the styles to apply</param>
    /// <param name="fontProperties">Properties for the font style to apply to the ranges</param>
    /// <param name="borders">Border information for each range</param>
    public URangeStyles(List<URange> ranges, UFontProperties fontProperties, List<(EBorderType type, EBorderStyleType style, string color)> borders)
    {
        Ranges = ranges;
        FontProperties = fontProperties;
        Borders = borders;
    }
}
