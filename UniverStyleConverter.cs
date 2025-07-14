using System.Text;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using UniverBlazored.Spreadsheets.Data.ConditionFormat;
using UniverBlazored.Spreadsheets.Data.Styles;

namespace UniverBlazored.SpreadsheetConverter;

internal static class UniverStyleConverter
{
    // Converts the stirng value of an enum into TittleCale
    static string ToEnumTitleCase(string input)
    {
        // To handle an empty string
        if (string.IsNullOrEmpty(input))
        {
            return input;
        }

        StringBuilder result = new StringBuilder();
        bool newWord = true;

        foreach (char c in input)
        {
            if (c == '_')
            {
                // Skip underscores, they separate words
                newWord = true;
            }
            else
            {
                if (newWord)
                {
                    // Capitalize the first letter of the word
                    result.Append(char.ToUpper(c));
                    newWord = false;
                }
                else
                {
                    // Make the rest of the letters lowercase
                    result.Append(char.ToLower(c));
                }
            }
        }

        return result.ToString();
    }

    /// <summary>
    /// Returns the type of border and its color to place in ClosedXML
    /// </summary>
    /// <param name="styleData"></param>
    /// <returns></returns>
    public static (XLBorderStyleValues border, XLColor color) GetBorderData(this UBorderStyleData styleData)
    {
        XLBorderStyleValues border = XLBorderStyleValues.None;
        if (styleData.s.HasValue)
        {
            string borderStyle = styleData.s.Value.ToString();
            border = Enum.Parse<XLBorderStyleValues>(ToEnumTitleCase(borderStyle)); // I'll take the risk :D
        }

        return (border, Toolbox.ConvertHexToARGB(styleData.cl.Value.rgb));
    }

    /// <summary>
    /// Returns the border style type for Univer
    /// </summary>
    /// <param name="border"></param>
    /// <returns></returns>
    public static EBorderStyleType ToBorderStyle(this XLBorderStyleValues border)
    {
        switch (border)
        {
            case XLBorderStyleValues.None: return EBorderStyleType.NONE;
            case XLBorderStyleValues.Thin: return EBorderStyleType.THIN;
            case XLBorderStyleValues.Hair: return EBorderStyleType.HAIR;
            case XLBorderStyleValues.Dotted: return EBorderStyleType.DOTTED;
            case XLBorderStyleValues.Dashed: return EBorderStyleType.DASHED;
            case XLBorderStyleValues.DashDot: return EBorderStyleType.DASH_DOT;
            case XLBorderStyleValues.DashDotDot: return EBorderStyleType.DASH_DOT_DOT;
            case XLBorderStyleValues.Double: return EBorderStyleType.DOUBLE;
            case XLBorderStyleValues.Medium: return EBorderStyleType.MEDIUM;
            case XLBorderStyleValues.MediumDashed: return EBorderStyleType.MEDIUM_DASHED;
            case XLBorderStyleValues.MediumDashDot: return EBorderStyleType.MEDIUM_DASH_DOT; // Agregado
            case XLBorderStyleValues.MediumDashDotDot: return EBorderStyleType.MEDIUM_DASH_DOT_DOT; // Agregado
            case XLBorderStyleValues.SlantDashDot: return EBorderStyleType.SLANT_DASH_DOT;
            case XLBorderStyleValues.Thick: return EBorderStyleType.THICK;
        }
        
        // Si no encuentra ningún caso de coincidencia, devuelve NONE
        return EBorderStyleType.NONE; 
    }

    /// <summary>
    /// Returns the vertical align type for ClosedXML
    /// </summary>
    /// <param name="align"></param>
    /// <returns></returns>
    public static XLAlignmentVerticalValues ToVerticalValue(this EVerticalAlign? align)
    {
        // I don't know why It must be like this... but it is
        if (align == null)
            return XLAlignmentVerticalValues.Bottom;

        switch (align)
        {
            case EVerticalAlign.TOP:
                return XLAlignmentVerticalValues.Top;

            case EVerticalAlign.BOTTOM:
                return XLAlignmentVerticalValues.Bottom;

            case EVerticalAlign.MIDDLE:
                return XLAlignmentVerticalValues.Center;

            case EVerticalAlign.UNSPECIFIED:
            default:
                return XLAlignmentVerticalValues.Top;
        }
    }

    /// <summary>
    /// Returns the vertical align type for Univer
    /// </summary>
    /// <param name="align"></param>
    /// <returns></returns>
    public static EVerticalAlign ToEVerticalValue(this XLAlignmentVerticalValues align)
    {
        switch (align)
        {
            case XLAlignmentVerticalValues.Top:
                return EVerticalAlign.TOP;
            case XLAlignmentVerticalValues.Bottom:
                return EVerticalAlign.BOTTOM;
            case XLAlignmentVerticalValues.Center:
                return EVerticalAlign.MIDDLE;
            default:
                return EVerticalAlign.UNSPECIFIED;
        }
    }

    /// <summary>
    /// Returns the vertical align type for ClosedXML
    /// </summary>
    /// <param name="align"></param>
    /// <returns></returns>
    public static XLAlignmentHorizontalValues ToHorizontalValue(this EHorizontalAlign align)
    {
        switch (align)
        {
            case EHorizontalAlign.LEFT:
                return XLAlignmentHorizontalValues.Left;

            case EHorizontalAlign.CENTER:
                return XLAlignmentHorizontalValues.Center;

            case EHorizontalAlign.RIGHT:
                return XLAlignmentHorizontalValues.Right;

            case EHorizontalAlign.JUSTIFIED:
                return XLAlignmentHorizontalValues.Justify;
            
            case EHorizontalAlign.BOTH:
                return XLAlignmentHorizontalValues.Fill;

            case EHorizontalAlign.DISTRIBUTED:
                return XLAlignmentHorizontalValues.Distributed;

            case EHorizontalAlign.UNSPECIFIED:
            default:
                return XLAlignmentHorizontalValues.General;
        }
    }

    /// <summary>
    /// Returns the horizontal align type for Univer
    /// </summary>
    /// <param name="align"></param>
    /// <returns></returns>
    public static EHorizontalAlign ToEHorizontalValue(this XLAlignmentHorizontalValues align)
    {
        switch (align)
        {
            case XLAlignmentHorizontalValues.Left:
                return EHorizontalAlign.LEFT;
            case XLAlignmentHorizontalValues.Center:
                return EHorizontalAlign.CENTER;
            case XLAlignmentHorizontalValues.Right:
                return EHorizontalAlign.RIGHT;
            case XLAlignmentHorizontalValues.Justify:
                return EHorizontalAlign.JUSTIFIED;
            case XLAlignmentHorizontalValues.Fill:
                return EHorizontalAlign.BOTH;
            case XLAlignmentHorizontalValues.Distributed:
                return EHorizontalAlign.DISTRIBUTED;
            default:
                return EHorizontalAlign.UNSPECIFIED;
        }
    }

    /// <summary>
    /// Returns the picture format, acording to the file type in the Data Uri
    /// </summary>
    /// <param name="imageType"></param>
    /// <returns></returns>
    public static XLPictureFormat ToPictureFormat(this string imageType)
    {
        switch (imageType)
        {
            case "image/tiff":
                return XLPictureFormat.Tiff;

            case "image/x-pcx":
                return XLPictureFormat.Pcx;

            case "image/x-icon":
                return XLPictureFormat.Icon;

            case "image/gif":
                return XLPictureFormat.Gif;

            case "image/bmp":
                return XLPictureFormat.Bmp;

            case "image/webp":
                return XLPictureFormat.Webp;

            case "image/png":
                return XLPictureFormat.Png;

            case "image/jpeg":
                return XLPictureFormat.Jpeg;

            default:
                return XLPictureFormat.Unknown;
        }
    }

    /// <summary>
    /// Returns the MIME type, depending of tge Picture format
    /// </summary>
    /// <param name="pictureFormat"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentException">If the Picture format is invalid for Univer</exception>
    public static string ToImageType(this XLPictureFormat pictureFormat)
    {
        switch (pictureFormat)
        {
            case XLPictureFormat.Tiff:
                return "image/tiff";
            case XLPictureFormat.Pcx:
                return "image/x-pcx";
            case XLPictureFormat.Icon:
                return "image/x-icon";
            case XLPictureFormat.Gif:
                return "image/gif";
            case XLPictureFormat.Bmp:
                return "image/bmp";
            case XLPictureFormat.Webp:
                return "image/webp";
            case XLPictureFormat.Png:
                return "image/png";
            case XLPictureFormat.Jpeg:
                return "image/jpeg";
            default:
                throw new ArgumentException("Invalid picture format");
        }
    }

    /// <summary>
    /// Returns the icon set, acording to the icon type, from Univer
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    public static XLIconSetStyle ToIconSet(this EIconType? type)
    {
        switch (type)
        {
            case EIconType.I_3ArrowsGray:
                return XLIconSetStyle.ThreeArrowsGray;
                
            case EIconType.I_4Arrows:
                return XLIconSetStyle.FourArrows;

            case EIconType.I_4ArrowsGray:
                return XLIconSetStyle.FourArrowsGray;

            case EIconType.I_5Arrows:
                return XLIconSetStyle.FiveArrows;

            case EIconType.I_5ArrowsGray:
                return XLIconSetStyle.FiveArrowsGray;

            case EIconType.I_3TrafficLights1:
                return XLIconSetStyle.ThreeTrafficLights1;

            case EIconType.I_3TrafficLights2:
                return XLIconSetStyle.ThreeTrafficLights2;

            case EIconType.I_3Signs:
                return XLIconSetStyle.ThreeSigns;

            case EIconType.I_3Symbols:
                return XLIconSetStyle.ThreeSymbols;

            case EIconType.I_3Symbols2:
                return XLIconSetStyle.ThreeSymbols2;

            case EIconType.I_3Flags:
                return XLIconSetStyle.ThreeFlags;

            case EIconType.I_4RedToBlack:
                return XLIconSetStyle.FourRedToBlack;

            case EIconType.I_4Rating:
                return XLIconSetStyle.FourRating;

            case EIconType.I_4TrafficLights:
                return XLIconSetStyle.FourTrafficLights;

            case EIconType.I_5Rating:
                return XLIconSetStyle.FiveRating;

            case EIconType.I_5Quarters:
                return XLIconSetStyle.FiveQuarters;

            // NOT SUPPORTED YET!! v

            case EIconType.I__5Felling:
            case EIconType.I_5Boxes:
                return XLIconSetStyle.FiveArrows;

            case EIconType.I_3Stars:
            case EIconType.I_3Triangles:
            case EIconType.I_3Arrows: // Supported...
            default:
                return XLIconSetStyle.ThreeArrows;
        }
    }

    /// <summary>
    /// Returns the Icon Type, acording to the icon set, from ClosedXML
    /// </summary>
    /// <param name="iconStyle"></param>
    /// <returns></returns>
    public static EIconType ToIconType(this XLIconSetStyle iconStyle)
    {
        switch (iconStyle)
        {
            case XLIconSetStyle.ThreeArrows: return EIconType.I_3Arrows;
            case XLIconSetStyle.ThreeArrowsGray: return EIconType.I_3ArrowsGray;
            case XLIconSetStyle.ThreeFlags: return EIconType.I_3Flags;
            case XLIconSetStyle.FourArrows: return EIconType.I_4Arrows;
            case XLIconSetStyle.FourArrowsGray: return EIconType.I_4ArrowsGray;
            case XLIconSetStyle.FiveArrows: return EIconType.I_5Arrows;
            case XLIconSetStyle.FiveArrowsGray: return EIconType.I_5ArrowsGray;
            case XLIconSetStyle.ThreeTrafficLights1: return EIconType.I_3TrafficLights1;
            case XLIconSetStyle.ThreeTrafficLights2: return EIconType.I_3TrafficLights2;
            case XLIconSetStyle.ThreeSigns: return EIconType.I_3Signs;
            case XLIconSetStyle.ThreeSymbols: return EIconType.I_3Symbols;
            case XLIconSetStyle.ThreeSymbols2: return EIconType.I_3Symbols2;
            case XLIconSetStyle.FourRedToBlack: return EIconType.I_4RedToBlack;
            case XLIconSetStyle.FourRating: return EIconType.I_4Rating;
            case XLIconSetStyle.FourTrafficLights: return EIconType.I_4TrafficLights;
            case XLIconSetStyle.FiveRating: return EIconType.I_5Rating;
            case XLIconSetStyle.FiveQuarters: return EIconType.I_5Quarters;
            default: 
                return EIconType.I_3Arrows; 
        }
    }

    /// <summary>
    /// Returns the Icon Set's operator, acording to the operator type, from Univer
    /// </summary>
    /// <param name="_operator"></param>
    /// <returns></returns>
    public static XLCFIconSetOperator ToIconSetOperator(this ECFOperators _operator)
    {
        switch (_operator)
        {
            case ECFOperators.greaterThanOrEqual: return XLCFIconSetOperator.EqualOrGreaterThan;
            case ECFOperators.greaterThan: default: return XLCFIconSetOperator.GreaterThan;
        }
    }

    /// <summary>
    /// Returns the conditional format operator for Univer, from ClosedXML
    /// </summary>
    /// <param name="_operator"></param>
    /// <returns></returns>
    public static ECFOperators ToOperator(this XLCFIconSetOperator _operator)
    {
        switch (_operator)
        {
            case XLCFIconSetOperator.EqualOrGreaterThan: return ECFOperators.greaterThanOrEqual;
            case XLCFIconSetOperator.GreaterThan: default: return ECFOperators.greaterThan;
        }
    }

    public static ECFValueType ToValueType(this XLCFContentType contentType)
    {
        switch (contentType)
        {
            case XLCFContentType.Percent: return ECFValueType.percent;
            case XLCFContentType.Formula: return ECFValueType.formula;
            case XLCFContentType.Percentile: return ECFValueType.percentile;
            case XLCFContentType.Minimum: return ECFValueType.min;
            case XLCFContentType.Maximum: return ECFValueType.max;
            case XLCFContentType.Number: return ECFValueType.num;
            default: return ECFValueType.num;
        }
    }

    /// <summary>
    /// Return true if the style is in default values
    /// </summary>
    /// <param name="style"></param>
    /// <param name="worksheetStyle">Default style used in the worksheet</param>
    /// <returns></returns>
    public static bool IsDefaultStyle(this IXLStyle style, IXLStyle worksheetStyle)
    {
        
        return
            IsDefaultFontName(style, worksheetStyle) &&
            IsDefaultItalic(style, worksheetStyle) &&
            IsDefaultBold(style, worksheetStyle) &&
            IsDefaultFontSize(style, worksheetStyle) &&
            IsDefaultUnderline(style, worksheetStyle) &&
            IsDefaultDateFormat(style, worksheetStyle) &&
            IsDefaultNumberFormat(style, worksheetStyle) &&
            IsDefaultTextRotation(style, worksheetStyle) &&
            IsDefaultReadingOrder(style, worksheetStyle) &&
            IsDefaultBackgroundColor(style, worksheetStyle) &&
            IsDefaultFontColor(style, worksheetStyle) &&
            IsDefaultWrapText(style, worksheetStyle) &&
            IsDefaultHorizontalAlignment(style, worksheetStyle) &&
            IsDefaultVerticalAlignment(style, worksheetStyle) &&
            IsDefaultFontStrikethrough(style, worksheetStyle) &&

            IsDefaultBottomBorder(style, worksheetStyle) &&
            IsDefaultLeftBorder(style, worksheetStyle) &&
            IsDefaultTopBorder(style, worksheetStyle) &&
            IsDefaultRightBorder(style, worksheetStyle) &&
            IsDefaultDiagonalBorder(style, worksheetStyle) &&
            IsDefaultDiagonalUpBorder(style, worksheetStyle) &&
            IsDefaultDiagonalDownBorder(style, worksheetStyle);
    }

    #region All Default Conditions

    static bool IsDefaultFontName(this IXLStyle style, IXLStyle worksheetStyle)
        => worksheetStyle.Font.FontName.Equals(style.Font.FontName);
    
    static bool IsDefaultItalic(this IXLStyle style, IXLStyle worksheetStyle)
        => style.Font.Italic == worksheetStyle.Font.Italic;

    static bool IsDefaultBold(this IXLStyle style, IXLStyle worksheetStyle)
        => style.Font.Bold == worksheetStyle.Font.Bold;
    
    static bool IsDefaultFontSize(this IXLStyle style, IXLStyle worksheetStyle)
        => worksheetStyle.Font.FontSize == style.Font.FontSize;

    static bool IsDefaultUnderline(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Font.Underline == worksheetStyle.Font.Underline;

    static bool IsDefaultFontStrikethrough(this IXLStyle style, IXLStyle worksheetStyle)
        => style.Font.Strikethrough == worksheetStyle.Font.Strikethrough;

    static bool IsDefaultFontColor(this IXLStyle style, IXLStyle worksheetStyle)
        => worksheetStyle.Font.FontColor.Equals(style.Font.FontColor);

    static bool IsDefaultWrapText(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Alignment.WrapText == worksheetStyle.Alignment.WrapText;

    static bool IsDefaultHorizontalAlignment(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Alignment.Horizontal == worksheetStyle.Alignment.Horizontal;

    static bool IsDefaultVerticalAlignment(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Alignment.Vertical == worksheetStyle.Alignment.Vertical;

    static bool IsDefaultNumberFormat(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.NumberFormat.Format == worksheetStyle.NumberFormat.Format;

    static bool IsDefaultDateFormat(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.DateFormat.Format == worksheetStyle.DateFormat.Format;

    static bool IsDefaultTextRotation(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Alignment.TextRotation == worksheetStyle.Alignment.TextRotation;

    static bool IsDefaultReadingOrder(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Alignment.ReadingOrder == worksheetStyle.Alignment.ReadingOrder;

    static bool IsDefaultBackgroundColor(this IXLStyle style, IXLStyle worksheetStyle)
        => worksheetStyle.Fill.BackgroundColor.Equals(style.Fill.BackgroundColor);

    // Borders
    static bool IsDefaultBottomBorder(this IXLStyle style, IXLStyle worksheetStyle)
        => style.Border.BottomBorder == worksheetStyle.Border.BottomBorder;

    static bool IsDefaultLeftBorder(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Border.LeftBorder == worksheetStyle.Border.LeftBorder;

    static bool IsDefaultTopBorder(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Border.TopBorder == worksheetStyle.Border.TopBorder;

    static bool IsDefaultRightBorder(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Border.RightBorder == worksheetStyle.Border.RightBorder;

    static bool IsDefaultDiagonalBorder(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Border.DiagonalBorder == worksheetStyle.Border.DiagonalBorder;

    static bool IsDefaultDiagonalUpBorder(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Border.DiagonalUp == worksheetStyle.Border.DiagonalUp;

    static bool IsDefaultDiagonalDownBorder(this IXLStyle style, IXLStyle worksheetStyle) 
        => style.Border.DiagonalDown == worksheetStyle.Border.DiagonalDown;

    #endregion

    /// <summary>
    /// Return the CLosedXML style as an Univer's cell style
    /// </summary>
    /// <param name="style"></param>
    /// <param name="theme">Spreadhseet Theme to extract the colors</param>
    /// <param name="worksheetStyle">Default style used in the worksheet</param>
    /// <returns></returns>
    public static UFontProperties ToFontProperties(this IXLStyle style, IXLTheme theme, IXLStyle worksheetStyle)
    {
        var fontProps = new UFontProperties();
        if (!style.IsDefaultFontName(worksheetStyle))
            fontProps.Family = style.Font.FontName;

        if (!style.IsDefaultItalic(worksheetStyle))
            fontProps.Italic = style.Font.Italic;

        if (!style.IsDefaultBold(worksheetStyle))
            fontProps.Bold = style.Font.Bold;

        if (!style.IsDefaultFontSize(worksheetStyle))
            fontProps.Size = style.Font.FontSize;

        if (!style.IsDefaultUnderline(worksheetStyle))
            fontProps.Underline = style.Font.Underline is not XLFontUnderlineValues.None;

        if (!style.IsDefaultDateFormat(worksheetStyle))
            fontProps.NumberFormat = style.DateFormat.Format;
            
        if (!style.IsDefaultNumberFormat(worksheetStyle))
            fontProps.NumberFormat = style.NumberFormat.Format;
        
        if (!style.IsDefaultTextRotation(worksheetStyle))
            fontProps.TextRotation = style.Alignment.TextRotation;

        if (!style.IsDefaultFontColor(worksheetStyle))
            fontProps.Color = style.Font.FontColor.ColorType is XLColorType.Theme ?
                                Toolbox.ColorToHexString(theme.ResolveThemeColor(style.Font.FontColor.ThemeColor).Color):
                                Toolbox.ColorToHexString(style.Font.FontColor.Color);
            

        if (!style.IsDefaultWrapText(worksheetStyle))
        {
            fontProps.IsWrap        = true;
            fontProps.WrapStrategy  = EWrapStrategy.WRAP;
        }
            

        if (!style.IsDefaultHorizontalAlignment(worksheetStyle))
        {
            switch (style.Alignment.Horizontal)
            {
                case XLAlignmentHorizontalValues.Center:
                    fontProps.HorizontalAlign = FHorizontalAligment.CENTER;
                    break;

                case XLAlignmentHorizontalValues.Left:
                    fontProps.HorizontalAlign = FHorizontalAligment.LEFT;
                    break;

                case XLAlignmentHorizontalValues.Right:
                default:
                    fontProps.HorizontalAlign = FHorizontalAligment.NORMAL;
                    break;
            }
        }

        if (!style.IsDefaultVerticalAlignment(worksheetStyle))
        {
            switch (style.Alignment.Vertical)
            {
                case XLAlignmentVerticalValues.Top:
                    fontProps.VerticalAlign = FVerticalAligment.TOP;
                    break;

                case XLAlignmentVerticalValues.Center:
                    fontProps.VerticalAlign = FVerticalAligment.MIDDLE;
                    break;

                case XLAlignmentVerticalValues.Bottom:
                default:
                    fontProps.VerticalAlign = FVerticalAligment.BOTTOM;
                    break;
            }
        }

        if (!style.IsDefaultBackgroundColor(worksheetStyle))
            fontProps.BackgroundColor = style.Fill.BackgroundColor.ColorType is XLColorType.Theme ?
                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(style.Fill.BackgroundColor.ThemeColor).Color, true, style.Fill.BackgroundColor.ThemeTint):
                                            Toolbox.ColorToHexString(style.Fill.BackgroundColor.Color, true);

        if (!style.IsDefaultFontStrikethrough(worksheetStyle))
            fontProps.Strikethrough = style.Font.Strikethrough;

        // Text Direction is ignored!
        return fontProps;
    }

    /// <summary>
    /// Returns each border property to put into Univer
    /// </summary>
    /// <param name="style"></param>
    /// <param name="theme">Spreadhseet Theme to extract the colors</param>
    /// <param name="worksheetStyle">Default style used in the worksheet</param>
    /// <returns></returns>
    public static List<(EBorderType type, EBorderStyleType style, string color)> ToBorderProperties(this IXLStyle style, IXLTheme theme, IXLStyle worksheetStyle)
    {
        var results = new List<(EBorderType type, EBorderStyleType style, string? color)>();
        
        if (!style.IsDefaultBottomBorder(worksheetStyle))
            results.Add(new (
                EBorderType.BOTTOM, 
                style.Border.BottomBorder.ToBorderStyle(), 
                style.Border.BottomBorderColor.ColorType is XLColorType.Theme ? 
                                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(style.Border.BottomBorderColor.ThemeColor).Color, true, style.Border.BottomBorderColor.ThemeTint):
                                                            Toolbox.ColorToHexString(style.Border.BottomBorderColor.Color, true)
            ));
        
        if (!style.IsDefaultTopBorder(worksheetStyle))
            results.Add(new (
                EBorderType.TOP, 
                style.Border.TopBorder.ToBorderStyle(), 
                style.Border.TopBorderColor.ColorType is XLColorType.Theme ?
                                                         Toolbox.ColorToHexString(theme.ResolveThemeColor(style.Border.TopBorderColor.ThemeColor).Color, true, style.Border.TopBorderColor.ThemeTint) :
                                                         Toolbox.ColorToHexString(style.Border.TopBorderColor.Color, true)
            ));

        if (!style.IsDefaultLeftBorder(worksheetStyle))
            results.Add(new(
                EBorderType.LEFT,
                style.Border.LeftBorder.ToBorderStyle(),
                style.Border.LeftBorderColor.ColorType is XLColorType.Theme ?
                                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(style.Border.LeftBorderColor.ThemeColor).Color, true, style.Border.LeftBorderColor.ThemeTint):
                                                            Toolbox.ColorToHexString(style.Border.LeftBorderColor.Color, true)
            ));
        
        if (!style.IsDefaultRightBorder(worksheetStyle))
            results.Add(new(
                EBorderType.RIGHT,
                style.Border.RightBorder.ToBorderStyle(),
                style.Border.RightBorderColor.ColorType is XLColorType.Theme ?
                                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(style.Border.RightBorderColor.ThemeColor).Color, true, style.Border.RightBorderColor.ThemeTint) :
                                                            Toolbox.ColorToHexString(style.Border.RightBorderColor.Color, true)
            ));

        if (!style.IsDefaultDiagonalBorder(worksheetStyle))
            results.Add(new(
                style.Border.DiagonalUp ? EBorderType.BLTR : EBorderType.TLBR,
                style.Border.DiagonalBorder.ToBorderStyle(),
                style.Border.DiagonalBorderColor.ColorType is XLColorType.Theme ?
                                                                Toolbox.ColorToHexString(theme.ResolveThemeColor(style.Border.DiagonalBorderColor.ThemeColor).Color, true, style.Border.DiagonalBorderColor.ThemeTint):
                                                                Toolbox.ColorToHexString(style.Border.DiagonalBorderColor.Color, true)
            ));
        
        return results;
    }

    /// <summary>
    /// Return the Time Period for Univer
    /// </summary>
    /// <param name="time"></param>
    /// <returns></returns>
    public static ECFOperators ToTimePeriod(this XLTimePeriod time)
    {
        switch (time)
        {
            case XLTimePeriod.Yesterday: return ECFOperators.yesterday;
            case XLTimePeriod.Today: return ECFOperators.today;
            case XLTimePeriod.Tomorrow: return ECFOperators.tomorrow;
            case XLTimePeriod.InTheLast7Days: return ECFOperators.last7Days;
            case XLTimePeriod.LastWeek: return ECFOperators.lastWeek;
            case XLTimePeriod.ThisWeek: return ECFOperators.thisWeek;
            case XLTimePeriod.NextWeek: return ECFOperators.nextWeek;
            case XLTimePeriod.LastMonth: return ECFOperators.lastMonth;
            case XLTimePeriod.ThisMonth: return ECFOperators.thisMonth;
            case XLTimePeriod.NextMonth: return ECFOperators.nextMonth;
            default: return ECFOperators.last7Days;
        }
    }
}