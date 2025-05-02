using ClosedXML.Excel;

namespace UniverBlazored.SpreadsheetConverter;

/// <summary>
/// Default values for the spreadseet style
/// </summary>
internal readonly struct DefaultStyleValues
{
    /// <summary>
    /// XLStyle constant value for 'FontName'
    /// </summary>
    public static bool FontName(string font) => font.Equals("Calibri") || font.Equals("Arial");
    
    /// <summary>
    /// XLStyle constant value for 'Italic'
    /// </summary>
    public const bool Italic = false;
    
    /// <summary>
    /// XLStyle constant value for 'Bold'
    /// </summary>
    public const bool Bold = false;
    
    /// <summary>
    /// XLStyle constant value for 'FontSize' (not considered)
    /// </summary>
    public static bool FontSize(double fontSize) => fontSize == 11.0 || fontSize == 10.0;
    
    /// <summary>
    /// XLStyle constant value for 'FontColor'
    /// </summary>
    public const string FontColor = "FF000000";
    
    /// <summary>
    /// XLStyle constant value for 'Underline'
    /// </summary>
    public const XLFontUnderlineValues Underline = XLFontUnderlineValues.None;
    
    /// <summary>
    /// XLStyle constant value for 'WrapText'
    /// </summary>
    public const bool WrapText = false;
    
    /// <summary>
    /// XLStyle constant value for 'Horizontal'
    /// </summary>
    public const XLAlignmentHorizontalValues Horizontal = XLAlignmentHorizontalValues.General;
    
    /// <summary>
    /// XLStyle constant value for 'Vertical'
    /// </summary>
    public const XLAlignmentVerticalValues Vertical = XLAlignmentVerticalValues.Bottom;
    
    /// <summary>
    /// XLStyle constant value for 'NumberFormat'
    /// </summary>
    public const string NumberFormat = "";
    
    /// <summary>
    /// XLStyle constant value for 'DateFormat'
    /// </summary>
    public const string DateFormat = "";
    
    /// <summary>
    /// XLStyle constant value for 'BackgroundColor'
    /// </summary>
    public const string BackgroundColor = "00FFFFFF";
    
    /// <summary>
    /// XLStyle constant value for 'FontStrikethrough'
    /// </summary>
    public const bool FontStrikethrough = false;
    
    /// <summary>
    /// XLStyle constant value for 'TextRotation'
    /// </summary>
    public const int TextRotation = 0;
    
    /// <summary>
    /// XLStyle constant value for 'ReadingOrder'
    /// </summary>
    public const XLAlignmentReadingOrderValues ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent;

    // Borders

    /// <summary>
    /// XLStyle constant value for 'TopBorder'
    /// </summary>
    public const XLBorderStyleValues TopBorder = XLBorderStyleValues.None;

    /// <summary>
    /// XLStyle constant value for 'TopBorderColor'
    /// </summary>
    public const string TopBorderColor = "FF000000";
    
    /// <summary>
    /// XLStyle constant value for 'RightBorder'
    /// </summary>
    public const XLBorderStyleValues RightBorder = XLBorderStyleValues.None;

    /// <summary>
    /// XLStyle constant value for 'RightBorderColor'
    /// </summary>
    public const string RightBorderColor = "FF000000";
    
    /// <summary>
    /// XLStyle constant value for 'BottomBorder'
    /// </summary>
    public const XLBorderStyleValues BottomBorder = XLBorderStyleValues.None;

    /// <summary>
    /// XLStyle constant value for 'BottomBorderColor'
    /// </summary>
    public const string BottomBorderColor = "FF000000";
    
    /// <summary>
    /// XLStyle constant value for 'LeftBorder'
    /// </summary>
    public const XLBorderStyleValues LeftBorder = XLBorderStyleValues.None;

    /// <summary>
    /// XLStyle constant value for 'LeftBorderColor'
    /// </summary>
    public const string LeftBorderColor = "FF000000";
    
    /// <summary>
    /// XLStyle constant value for 'DiagonalBorder'
    /// </summary>
    public const XLBorderStyleValues DiagonalBorder = XLBorderStyleValues.None;

    /// <summary>
    /// XLStyle constant value for 'DiagonalBorderColor'
    /// </summary>
    public const string DiagonalBorderColor = "FF000000";
    
    /// <summary>
    /// XLStyle constant value for 'DiagonalUp'
    /// </summary>
    public const bool DiagonalUp = false;

    /// <summary>
    /// XLStyle constant value for 'DiagonalDown'
    /// </summary>
    public const bool DiagonalDown = false;
}