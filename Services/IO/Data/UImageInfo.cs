using UniverBlazored.Spreadsheets.Data.Workbook;

namespace UniverBlazored.SpreadsheetConverter.Services.IO.Data;

/// <summary>
/// Image information for a given image in a worksheet
/// </summary>
public class UImageInfo
{
    /// <summary>
    /// Information about the image, including the type and the extension of the image
    /// </summary>
    /// <value></value>
    public string DataUri { get; set; } = "";

    /// <summary>
    /// Cell reference for the image
    /// </summary>
    /// <value></value>
    public URange StartCell { get; set; }

    /// <summary>
    /// Amount of pixels to the left, starting from the left border of the start cell, for the image to be placed
    /// </summary>
    /// <value></value>
    public double PixelsLeft { get; set; }

    /// <summary>
    /// Amount of pixels to the top, starting from the top border of the start cell, for the image to be placed
    /// </summary>
    /// <value></value>
    public double PixelsTop { get; set; }

    /// <summary>
    /// Image information for a given image in a worksheet
    /// </summary>
    public UImageInfo()
    {
        StartCell = new URange();
    }

    /// <summary>
    /// Image information for a given image in a worksheet
    /// </summary>
    /// <param name="dataUri">Information about the image, including the type and the extension of the image</param>
    /// <param name="startCell">Cell reference for the image</param>
    /// <param name="pixelsLeft">Amount of pixels to the left, starting from the left border of the start cell, for the image to be placed</param>
    /// <param name="pixelsTop">Amount of pixels to the top, starting from the top border of the start cell, for the image to be placed</param>
    public UImageInfo(string dataUri, URange startCell, double pixelsLeft, double pixelsTop)
    {
        DataUri = dataUri;
        StartCell = startCell;
        PixelsLeft = pixelsLeft;
        PixelsTop = pixelsTop;
    }
}