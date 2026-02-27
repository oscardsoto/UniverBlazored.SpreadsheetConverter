using UniverBlazored.Spreadsheets.Data.Workbook;

namespace UniverBlazored.SpreadsheetConverter.Services.IO.Data;

/// <summary>
/// Comment reference for a given comment in a worksheet.
/// </summary>
public class UCommentReference
{
    /// <summary>
    /// Cell reference for the comment
    /// </summary>
    /// <value></value>
    public URange Reference { get; set; }

    /// <summary>
    /// Text value for the comment
    /// </summary>
    /// <value></value>
    public string TextValue { get; set; } = "";

    /// <summary>
    /// Comment reference for a given comment in a worksheet.
    /// </summary>
    public UCommentReference()
    {
        Reference = new URange();
    }

    /// <summary>
    /// Comment reference for a given comment in a worksheet.
    /// </summary>
    /// <param name="reference">Cell reference for the comment</param>
    /// <param name="textValue">Text value for the comment</param>
    public UCommentReference(URange reference, string textValue)
    {
        Reference = reference;
        TextValue = textValue;
    }   
}
