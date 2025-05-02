namespace UniverBlazored.SpreadsheetConverter.Services;

/// <summary>
/// Option for get/set the information from Univer's Asgent
/// </summary>
public class SpreadsheetOptions
{
    /// <summary>
    /// True for get/set spreadsheet data
    /// </summary>
    public bool RecData { get; set; }

    /// <summary>
    /// True for get/set spreadsheet styles
    /// </summary>
    public bool RecStyles { get; set; }

    /// <summary>
    /// True for get/set spreadsheet merges
    /// </summary>
    public bool RecMerges { get; set; }

    /// <summary>
    /// True for get/set spreadsheet filters
    /// </summary>
    public bool RecFilters { get; set; }

    /// <summary>
    /// True for get/set spreadsheet freeze columns and rows
    /// </summary>
    public bool RecFreeze { get; set; }

    /// <summary>
    /// True for get/set spreadsheet comments
    /// </summary>
    public bool RecComments { get; set; }

    /// <summary>
    /// True for get/set spreadsheet columns and rows data
    /// </summary>
    public bool RecColumnsAndRows { get; set; }

    /// <summary>
    /// True for get/set spreadsheet images
    /// </summary>
    public bool RecImages { get; set; }

    /// <summary>
    /// True for get/set spreadsheet conditional formats
    /// </summary>
    public bool RecConditionalFormats { get; set; }

    /// <summary>
    /// True for get/set spreadsheet accesibility
    /// </summary>
    public bool RecAccesibility { get; set; }

    /// <summary>
    /// Option for get/set the information from Univer's Asgent
    /// </summary>
    public SpreadsheetOptions() { }

    /// <summary>
    /// Option for get/set the information from Univer's Asgent
    /// </summary>
    /// <param name="RecAll">True for enable all options. False to disable them</param>
    public SpreadsheetOptions(bool RecAll)
    {
        RecData = RecStyles = RecMerges = RecFilters = RecFreeze = RecComments = RecColumnsAndRows = RecImages = RecConditionalFormats = RecAccesibility = RecAll;
    }
}