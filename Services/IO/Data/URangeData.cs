namespace UniverBlazored.SpreadsheetConverter.Services.IO.Data;

/// <summary>
/// Store values and formulas for a given range in a worksheet.
/// </summary>
public class URangeData
{
    /// <summary>
    /// Values for the range (If any)
    /// </summary>
    /// <value></value>
    public object[][] Values { get; set; }

    /// <summary>
    /// Formulas used in the range (If any)
    /// </summary>
    /// <value></value>
    public string[][] Formulas { get; set; }

    /// <summary>
    /// Store values and formulas for a given range in a worksheet.
    /// </summary>
    public URangeData()
    {
        Values = Array.Empty<object[]>();
        Formulas = Array.Empty<string[]>();
    }

    /// <summary>
    /// Store values and formulas for a given range in a worksheet.
    /// </summary>
    /// <param name="formulas">Values for the range (If any)</param>
    /// <param name="values">Formulas used in the range (If any)</param>
    public URangeData(object[][] values, string[][] formulas)
    {
        Values = values;
        Formulas = formulas;
    }
}
