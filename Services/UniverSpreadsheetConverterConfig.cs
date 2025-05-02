namespace UniverBlazored.SpreadsheetConverter.Services;

/// <summary>
/// Configuration object for the Converter Scoped
/// </summary>
public class UniverSpreadsheetConverterConfig
{
    /// <summary>
    /// Number of Cells to process per task in the workbook
    /// </summary>
    public int MaxCellsReaded { get; set; } = 1000;
}