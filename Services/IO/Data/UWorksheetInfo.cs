using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using UniverBlazored.Spreadsheets.Data.ConditionFormat;
using UniverBlazored.Spreadsheets.Data.Workbook;

namespace UniverBlazored.SpreadsheetConverter.Services.IO.Data;

/// <inheritdoc/>
public class UWorksheetInfo : IWorksheetInfo<IXLWorksheet>
{
    /// <inheritdoc/>
    public USheetInfo SheetInfo { get; set; }

    /// <inheritdoc/>
    public Dictionary<URange, URangeData> RangesData { get; set; } = new();

    /// <inheritdoc/>
    public List<URangeStyles> RangeStyles { get; set; } = new();

    /// <inheritdoc/>
    public List<URange> MergedRanges { get; set; } = new();

    /// <inheritdoc/>
    public URange Filter { get; set; }

    /// <inheritdoc/>
    public URange FreezeReference { get; set; }

    /// <inheritdoc/>
    public List<UCommentReference> Comments { get; set; } = new();

    /// <inheritdoc/>
    public List<UCondFormatData> ConditionalFormats { get; set; } = new();

    /// <inheritdoc/>
    public List<(int column, double width)> ColumnWidths { get; set; } = new();

    /// <inheritdoc/>
    public List<(int row, double height)> RowHeights { get; set; } = new();

    /// <inheritdoc/>
    public List<UImageInfo> Images { get; set; } = new();

    /// <summary>
    /// All worksheet information to be stored in Univer after reading a spreadsheet
    /// </summary>
    /// <param name="worksheet">Worksheet to process</param>
    /// <param name="maxCellsReaded">Maximum number of cells to read</param>
    /// <param name="sheetInfo">Sheet info already extracted</param>
    public UWorksheetInfo(IXLWorksheet worksheet, int maxCellsReaded, USheetInfo sheetInfo)
    {
        SheetInfo           = sheetInfo;
        ColumnWidths        = GetColumnWidths(worksheet);
        RowHeights          = GetRowHeights(worksheet);
        Comments            = GetCommentsData(worksheet);
        ConditionalFormats  = GetConditionalFormatsData(worksheet);
        Filter              = GetFilter(worksheet);
        FreezeReference     = GetFreeze(worksheet);
        Images              = GetImagesData(worksheet);
        MergedRanges        = GetMergedRanges(worksheet);
        RangesData          = GetRangesData(worksheet, SheetInfo.maxUsed, maxCellsReaded);
        RangeStyles         = GetRangeStyles(worksheet, SheetInfo.maxUsed, maxCellsReaded);
    }

    /// <inheritdoc/>
    public List<(int column, double width)> GetColumnWidths(IXLWorksheet worksheet)
    {
        if (worksheet.ColumnsUsed().Count() == 0)
            return new List<(int column, double width)>();

        var cols = new List<(int column, double width)>();
        int[] colPositions = Toolbox.GetValuesInBetween(worksheet.ColumnsUsed().First().ColumnNumber(), worksheet.ColumnsUsed().Last().ColumnNumber());
        for (int i = 0; i < colPositions.Length; i++)
        {
            var colWidth = worksheet.Column(colPositions[i]).Width;
            cols.Add((colPositions[i] - 1, Toolbox.ConvertToColumnPixels(colWidth)[0]));
        }

        return cols;
    }

    /// <inheritdoc/>
    public List<UCommentReference> GetCommentsData(IXLWorksheet worksheet)
    {
        var comments = new List<UCommentReference>();
        var allComments = worksheet.CellsUsed().Select(c => c.HasComment ? new { range = c.Address.ToString(XLReferenceStyle.A1), comment = c.GetComment() } : null).ToList();

        foreach (var com in allComments)
        {
            if (com == null)
                continue;

            comments.Add(new(new URange(com.range), com.comment.Text));
        }
        return comments;
    }

    /// <inheritdoc/>
    public List<UCondFormatData> GetConditionalFormatsData(IXLWorksheet worksheet)
    {
        var condFormats = new List<UCondFormatData>();

        IXLTheme theme = worksheet.Workbook.Theme;
        IXLStyle styleFrmt;
        var separator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        foreach (var cdFt in worksheet.ConditionalFormats)
        {
            switch (cdFt.ConditionalFormatType)
            {
                case XLConditionalFormatType.IconSet:
                    // init the components
                    (bool isShowValue, UIconSetConfig[] config) iconConfig = new();
                    iconConfig.isShowValue = true;
                    List<UIconSetConfig> configTemp = new();

                    // extract the data for the conditional format
                    foreach (var iconOp in cdFt.IconSetOperators)
                    {
                        UIconSetConfig newConfig = new();
                        // Icon Id (position)
                        newConfig.iconId = (iconOp.Key - 1).ToString();
                        // Type
                        newConfig.SetIconType(cdFt.IconSetStyle.ToIconType());
                        // Operator (G or GE)
                        newConfig.SetOperator(iconOp.Key == 1 ? ECFOperators.lessThanOrEqual : ECFOperators.greaterThan);

                        // Depends of the value type, sets the value in Univer's config
                        var valueType = cdFt.ContentTypes[iconOp.Key].ToValueType();
                        if (valueType is ECFValueType.formula)
                            newConfig.value = new (
                                valueType, 
                                cdFt.Values[iconOp.Key] == null ? "" : cdFt.Values[iconOp.Key].Value
                            );
                        else
                            newConfig.value = new (
                                valueType, 
                                cdFt.Values[iconOp.Key] == null ? 0 : double.Parse(Regex.Replace(cdFt.Values[iconOp.Key].Value, "[.,]", separator))
                            );

                        // Add the config in the list
                        configTemp.Add(newConfig);
                    }
                    iconConfig.config = configTemp.ToArray();

                    var rangesIcon = new List<URange>();
                    foreach (var range in cdFt.Ranges)
                        rangesIcon.Add(new(range.RangeAddress.ToString(XLReferenceStyle.A1)));

                    condFormats.Add(new UCondFormatData(rangesIcon, new UConditionType() { IsIconSet = true }, new UConditionFormatStyle() { IconSet = iconConfig }));
                    break;

                case XLConditionalFormatType.DataBar:
                    var barConfig           = new UDataBarConfig();
                    barConfig.isShowValue   = true;
                    barConfig.isGradient    = true;
                    XLColor color           = cdFt.Colors.First().Value;
                    barConfig.positiveColor = color.ColorType is XLColorType.Theme ? Toolbox.ColorToHexString(theme.ResolveThemeColor(color.ThemeColor).Color, false, color.ThemeTint) : Toolbox.ColorToHexString(color.Color);
                    barConfig.nativeColor   = Toolbox.ColorToHexString(Color.Red); // Default value

                    // Min Value
                    var minValType = cdFt.ContentTypes[1].ToValueType();
                    if (minValType is ECFValueType.formula)
                        barConfig.min = new (
                            minValType,
                            cdFt.Values[1] == null ? "" : cdFt.Values[1].Value
                        );
                    else
                        barConfig.min = new (
                            minValType,
                            cdFt.Values[1] == null ? 0 : double.Parse(Regex.Replace(cdFt.Values[1].Value, "[.,]", separator))
                        );
                    
                    // Max Value
                    var maxValType = cdFt.ContentTypes[2].ToValueType();
                    if (maxValType is ECFValueType.formula)
                        barConfig.max = new (
                            maxValType,
                            cdFt.Values[2] == null ? "" : cdFt.Values[2].Value
                        );
                    else
                        barConfig.max = new (
                            maxValType,
                            cdFt.Values[2] == null ? 0 : double.Parse(Regex.Replace(cdFt.Values[2].Value, "[.,]", separator))
                        );

                    var rangesData = new List<URange>();
                    foreach (var range in cdFt.Ranges)
                        rangesData.Add(new(range.RangeAddress.ToString(XLReferenceStyle.A1)));

                    condFormats.Add(new UCondFormatData(rangesData, new UConditionType() { IsDataBar = true }, new UConditionFormatStyle() { DataBar = barConfig }));
                    break;

                case XLConditionalFormatType.ColorScale:
                    var colorConfig = new UColorScale();
                    var colorList   = new List<UColorScaleConfig>();

                    foreach (var kvp in cdFt.Values)
                    {
                        var config      = new UColorScaleConfig();
                        config.index    = kvp.Key;
                        config.color    = cdFt.Colors[kvp.Key].ColorType is XLColorType.Theme ? 
                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(cdFt.Colors[kvp.Key].ThemeColor).Color, false, cdFt.Colors[kvp.Key].ThemeTint):
                                            Toolbox.ColorToHexString(cdFt.Colors[kvp.Key].Color);

                        var valType = cdFt.ContentTypes[kvp.Key].ToValueType();
                        if (valType is ECFValueType.formula)
                            config.value = new (
                                valType,
                                cdFt.Values[kvp.Key] == null ? "" : cdFt.Values[kvp.Key].Value
                            );

                        else
                            config.value = new (
                                valType,
                                cdFt.Values[kvp.Key] == null ? 0 : double.Parse(Regex.Replace(cdFt.Values[kvp.Key].Value, "[.,]", separator))
                            );

                        colorList.Add(config);
                    }

                    var rangesColor = new List<URange>();
                    foreach (var range in cdFt.Ranges)
                        rangesColor.Add(new(range.RangeAddress.ToString(XLReferenceStyle.A1)));

                    colorConfig.config = colorList.ToArray();
                    condFormats.Add(new UCondFormatData(rangesColor, new UConditionType() { IsColorScale = true }, new UConditionFormatStyle() { ColorScale = colorConfig }));
                    break;

                default:
                    styleFrmt = cdFt.Style;
                    var condFormStyle = new UConditionFormatStyle()
                    {
                        Background      = styleFrmt.Fill.BackgroundColor.ColorType is XLColorType.Theme ?
                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(styleFrmt.Fill.BackgroundColor.ThemeColor).Color, false, styleFrmt.Fill.BackgroundColor.ThemeTint):
                                            Toolbox.ColorToHexString(styleFrmt.Fill.BackgroundColor.Color),
                        IsBold          = styleFrmt.Font.Bold,
                        FontColor       = styleFrmt.Font.FontColor.ColorType is XLColorType.Theme ?
                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(styleFrmt.Font.FontColor.ThemeColor).Color, false, styleFrmt.Font.FontColor.ThemeTint):
                                            Toolbox.ColorToHexString(styleFrmt.Font.FontColor.Color),
                        Italic          = styleFrmt.Font.Italic,
                        Strikethrough   = styleFrmt.Font.Strikethrough,
                        Underline       = styleFrmt.Font.Underline is not XLFontUnderlineValues.None
                    };
                    
                    var condType = new UConditionType();
                    switch (cdFt.ConditionalFormatType)
                    {
                        // Text
                        case XLConditionalFormatType.ContainsText:
                            condType.WhenTextContains = (cdFt.Values.Count == 0) ? string.Empty : cdFt.Values[1].Value;
                            break;

                        case XLConditionalFormatType.NotContainsText:
                            condType.WhenTextDoesNotContain = (cdFt.Values.Count == 0) ? string.Empty : cdFt.Values[1].Value;
                            break;

                        case XLConditionalFormatType.TimePeriod:
                            condType.WhenDate = cdFt.TimePeriod.ToTimePeriod();
                            break;

                        case XLConditionalFormatType.Expression:
                            condType.WhenFormulaSatisfied = (cdFt.Values.Count == 0) ? "" : $"={cdFt.Values[1].Value}";
                            break;

                        case XLConditionalFormatType.StartsWith:
                            condType.WhenTextStartsWith = (cdFt.Values.Count == 0) ? string.Empty : cdFt.Values[1].Value;
                            break;
                        
                        case XLConditionalFormatType.EndsWith:
                            condType.WhenTextEndsWith = (cdFt.Values.Count == 0) ? string.Empty : cdFt.Values[1].Value;
                            break;

                        // Values
                        case XLConditionalFormatType.CellIs:
                            var value = cdFt.Values[1].Value;
                            switch (cdFt.Operator)
                            {
                                case XLCFOperator.Equal:
                                    if (double.TryParse(value, out _))
                                        condType.WhenNumberEqualTo = double.Parse(Regex.Replace(value, "[.,]", separator));
                                    else
                                        condType.WhenTextEqualTo = value;
                                    break;

                                case XLCFOperator.LessThan:
                                    condType.WhenNumberLessThan = double.Parse(Regex.Replace(value, "[.,]", separator));
                                    break;

                                case XLCFOperator.GreaterThan:
                                    condType.WhenNumberGreaterThan = double.Parse(Regex.Replace(value, "[.,]", separator));
                                    break;

                                case XLCFOperator.EqualOrGreaterThan:
                                    condType.WhenNumberGreaterThanOrEqual = double.Parse(Regex.Replace(value, "[.,]", separator));
                                    break;

                                case XLCFOperator.EqualOrLessThan:
                                    condType.WhenNumberLessThanOrEqual = double.Parse(Regex.Replace(value, "[.,]", separator));
                                    break;

                                case XLCFOperator.Between:
                                    condType.WhenNumberInBetween = new(double.Parse(Regex.Replace(value, "[.,]", separator)), 
                                                                        double.Parse(Regex.Replace(cdFt.Values[2].Value, "[.,]", separator)));
                                    break;

                                case XLCFOperator.NotBetween:
                                    condType.WhenNumberNotBetween = new(double.Parse(Regex.Replace(value, "[.,]", separator)), 
                                                                        double.Parse(Regex.Replace(cdFt.Values[2].Value, "[.,]", separator)));
                                    break;

                                case XLCFOperator.NotEqual:
                                    if (double.TryParse(value, out _))
                                        condType.WhenNumberNotEqual = double.Parse(Regex.Replace(value, "[.,]", separator));
                                    // Text Not supported
                                    break;
                            }
                            break;

                        case XLConditionalFormatType.Top10:
                            condFormStyle.Rank = new(cdFt.Bottom, cdFt.Percent, double.Parse(Regex.Replace((cdFt.Values.Count == 0) ? "0.0" : cdFt.Values[1].Value, "[.,]", separator)));
                            break;

                        case XLConditionalFormatType.IsDuplicate:
                            condFormStyle.DuplicateValues = true;
                            break;

                        case XLConditionalFormatType.IsUnique:
                            condFormStyle.UniqueValues = true;
                            break;

                        case XLConditionalFormatType.IsBlank:
                            condType.WhenCellEmpty = true;
                            break;

                        case XLConditionalFormatType.NotBlank:
                            condType.WhenCellNotEmpty = true;
                            break;

                        /*
                            Not supported (yet):
                                AboveAverage
                                IsError
                                NotError
                        */
                        default:
                            continue;
                    }

                    var rangesCond = new List<URange>();
                    foreach (var range in cdFt.Ranges)
                        rangesCond.Add(new(range.RangeAddress.ToString(XLReferenceStyle.A1)));

                    condFormats.Add(new UCondFormatData(rangesCond, condType, condFormStyle));
                    break;
            }
        }

        return condFormats;
    }

    /// <inheritdoc/>
    public URange GetFilter(IXLWorksheet worksheet)
    {
        if (!worksheet.AutoFilter.IsEnabled)
            return new();

        string reference = "";
        var rangeFilter = worksheet.AutoFilter.Range;
        if (rangeFilter.RowCount() == 1)
        {
            // Coordenadas del encabezado
            int rowFirst        = rangeFilter.FirstCell().Address.RowNumber;
            int columnFirst     = rangeFilter.FirstCell().Address.ColumnNumber;
            int columnLast      = rangeFilter.LastCell().Address.ColumnNumber;

            int rowLast = rowFirst;
            for (int row = rowFirst + 1; row <= worksheet.LastRowUsed()?.RowNumber(); row++)
            {
                bool rowHasData = false;
                for (int col = columnFirst; col <= columnLast; col++)
                    if (!worksheet.Cell(row, col).IsEmpty())
                    {
                        rowHasData = true;
                        break;
                    }

                if (rowHasData)
                    rowLast = row;
                else
                    break; 
            }

            // Rango completo de datos filtrables
            reference = worksheet.Range(rowFirst, columnFirst, rowLast, columnLast).RangeAddress.ToString(XLReferenceStyle.A1);
        }
        else
            reference = rangeFilter.RangeAddress.ToString(XLReferenceStyle.A1);

        return new(reference);
    }

    /// <inheritdoc/>
    public URange GetFreeze(IXLWorksheet worksheet)
    {
        return new(worksheet.SheetView.SplitRow, worksheet.SheetView.SplitColumn);
    }

    /// <inheritdoc/>
    public List<UImageInfo> GetImagesData(IXLWorksheet worksheet)
    {
        var imagesList = new List<UImageInfo>();
        foreach (var img in worksheet.Pictures)
        {
            string dataURI = $"data:{img.Format.ToImageType()};base64,{Toolbox.ConvertToBase64(img.ImageStream)}";
            string cellTop = img.TopLeftCell.Address.ToString(XLReferenceStyle.A1);
            URange cellImg = new(cellTop);

            imagesList.Add(new UImageInfo(dataURI, cellImg, Toolbox.ConvertToPixels(img.Left), Toolbox.ConvertToPixels(img.Top)));
        }
        return imagesList;
    }

    /// <inheritdoc/>
    public List<URange> GetMergedRanges(IXLWorksheet worksheet)
    {
        var merges = new List<URange>();
        foreach (var merge in worksheet.MergedRanges)
        {
            string reference = merge.RangeAddress.ToString(XLReferenceStyle.A1);
            merges.Add(new(reference));
        }
        return merges;
    }

    /// <inheritdoc/>
    public Dictionary<URange, URangeData> GetRangesData(IXLWorksheet worksheet, URange maxUsed, int maxCellsReaded)
    {
        var data = new Dictionary<URange, URangeData>();

        // Selects the range where the data will be extracted
        int maxRow          = maxUsed.endRow,
            maxCol          = maxUsed.endColumn,
            rowCounter      = 0,
            rowsPerProcess  = Toolbox.MaxRowsPerProcess(maxCellsReaded, maxRow + 1, maxCol + 1);

        // Gets all results from each value
        do
        {
            var range = worksheet.Range(rowCounter + 1, 1, rowCounter + rowsPerProcess + 1, maxCol + 1);

            int firstRow    = 1,
                lastRow     = range.RowCount(),
                firstColumn = 1,
                lastColumn  = range.LastColumn().ColumnNumber(),
                rowPos      = firstRow,
                colPos      = firstColumn;

            var listValues      = new List<object[]>();
            var listFormulas    = new List<string[]>();
            var arrayValues     = new object[lastColumn];
            var arrayFormulas   = new string[lastColumn];
            do
            {
                var cell = range.Cell(rowPos, colPos);
                arrayFormulas[colPos - 1]   = cell.HasFormula ? $"={cell.FormulaA1}" : "";

                #pragma warning disable CS8601 // Possible null reference assignment. I need that null!! and those yellow lines are annoying
                arrayValues[colPos - 1]     = cell.HasFormula ? null : cell.Value.IsBlank ? null : 
                                                cell.Value.IsText ? (string) cell.Value :
                                                cell.Value.IsNumber ? (double) cell.Value :
                                                cell.Value.IsBoolean ? (bool) cell.Value : 
                                                cell.Value.IsDateTime ? cell.Value.GetDateTime().ToString("dd/MM/yyyy") : 
                                                cell.Value.IsError ? cell.Value.GetError().ToString() : 
                                                cell.Value.IsTimeSpan ? cell.Value.GetTimeSpan().ToString("HH:mm:ss") : null;
                #pragma warning restore CS8601 // Possible null reference assignment.

                if (colPos + 1 <= lastColumn)
                {
                    colPos++;
                    continue;
                }

                colPos = firstColumn;
                rowPos++;
                listValues.Add(arrayValues);
                listFormulas.Add(arrayFormulas);
                arrayValues = new object[lastColumn];
                arrayFormulas = new string[lastColumn];
                continue;
            }
            while (rowPos < lastRow);

            // Add a new KvP
            int endRow = rowCounter + rowsPerProcess - 1;
            if (endRow > maxRow)
                endRow = maxRow;
            data.Add(new URange(rowCounter, endRow, 0, maxCol), new URangeData(listValues.ToArray(), listFormulas.ToArray()));
            rowCounter += rowsPerProcess;
        }
        while (rowCounter < maxRow);

        return data;
    }

    (object[][] values, string[][] formulas) ExtractValuesAndFormulas(IXLRange range)
    {
        int firstRow    = 1,
            lastRow     = range.RowCount(),
            firstColumn = 1,
            lastColumn  = range.LastColumn().ColumnNumber(),
            rowPos      = firstRow,
            colPos      = firstColumn;

        var listValues      = new List<object[]>();
        var listFormulas    = new List<string[]>();
        var arrayValues     = new object[lastColumn];
        var arrayFormulas   = new string[lastColumn];
        do
        {
            var cell = range.Cell(rowPos, colPos);
            arrayFormulas[colPos - 1]   = cell.HasFormula ? $"={cell.FormulaA1}" : "";

            #pragma warning disable CS8601 // Possible null reference assignment. I need that null!! and those yellow lines are annoying
            arrayValues[colPos - 1]     = cell.HasFormula ? null : cell.Value.IsBlank ? null : 
                                            cell.Value.IsText ? (string) cell.Value :
                                            cell.Value.IsNumber ? (double) cell.Value :
                                            cell.Value.IsBoolean ? (bool) cell.Value : 
                                            cell.Value.IsDateTime ? cell.Value.GetDateTime().ToString("dd/MM/yyyy") : 
                                            cell.Value.IsError ? cell.Value.GetError().ToString() : 
                                            cell.Value.IsTimeSpan ? cell.Value.GetTimeSpan().ToString("HH:mm:ss") : null;
            #pragma warning restore CS8601 // Possible null reference assignment.

            if (colPos + 1 <= lastColumn)
            {
                colPos++;
                continue;
            }

            colPos = firstColumn;
            rowPos++;
            listValues.Add(arrayValues);
            listFormulas.Add(arrayFormulas);
            arrayValues = new object[lastColumn];
            arrayFormulas = new string[lastColumn];
            continue;
        }
        while (rowPos < lastRow);

        return (listValues.ToArray(), listFormulas.ToArray());
    }

    /// <inheritdoc/>
    public List<URangeStyles> GetRangeStyles(IXLWorksheet worksheet, URange maxUsed, int maxCellsReaded)
    {
        var stylesResult = new List<URangeStyles>(); 

        // Selects the range where the data will be extracted
        int maxRow          = maxUsed.endRow,
            maxCol          = maxUsed.endColumn,
            rowCounter      = 0,
            rowsPerProcess  = Toolbox.MaxRowsPerProcess(maxCellsReaded, maxRow + 1, maxCol + 1);

        IXLTheme theme          = worksheet.Workbook.Theme;
        IXLStyle defaultStyle   = worksheet.Style;        
        do
        {
            int firstRow    = rowCounter + 1,
                lastRow     = rowCounter + rowsPerProcess + 1;
            var range       = worksheet.Range(firstRow, 1, lastRow, maxCol + 1);

            IXLStyle[][] styles = range.Rows().Select(row => row.Cells().Select(c => c.Style).ToArray()).ToArray();

            var styleGroup      = new Dictionary<IXLStyle, List<URange>>();
            for (int row = 0; row < styles.Length; row++)
            {
                var rowStyles = styles[row];
                for (int col = 0; col < rowStyles.Length; col++)
                {
                    if (rowStyles[col].IsDefaultStyle(defaultStyle))
                        continue;

                    var cell = new URange(firstRow + row - 1, col);
                    styleGroup.TryAdd(rowStyles[col], new());
                    styleGroup[rowStyles[col]].Add(cell);
                }
            }

            foreach (var group in styleGroup)
            {
                var style = group.Key;
                var univerStyle = style.ToFontProperties(theme, defaultStyle);
                var univerBordr = style.ToBorderProperties(theme, defaultStyle);

                var ranges = RectangularizeRegion(group.Value);
                stylesResult.Add(new URangeStyles(ranges, univerStyle, univerBordr));
            }

            rowCounter += rowsPerProcess;
        }
        while (rowCounter < maxRow);

        return stylesResult;
    }

    List<URange> RectangularizeRegion(List<URange> regions)
    {
        var result = new List<URange>();
        var visited = new HashSet<URange>();
        var grid = new HashSet<URange>(regions);

        foreach (var region in regions)
        {
            if (visited.Contains(region))
                continue;

            int endCol = region.startColumn;
            // Expand to the right
            while (grid.Contains(new(region.startRow, endCol + 1)) && !visited.Contains(new(region.startRow, endCol + 1)))
                endCol++;

            int endRow = region.startRow;
            bool fullRowMatch;

            // Expand downward while full row match
            do
            {
                endRow++;
                fullRowMatch = true;
                for (int col = region.startColumn; col <= endCol; col++)
                {
                    if (!grid.Contains(new(endRow, col)) || visited.Contains(new(endRow, col)))
                    {
                        fullRowMatch = false;
                        break;
                    }
                }
            } while (fullRowMatch);

            // Final endRow is the last matching one
            endRow--;

            // Mark all block as visited
            for (int row = region.startRow; row <= endRow; row++)
                for (int col = region.startColumn; col <= endCol; col++)
                    visited.Add(new(row, col));

            result.Add(new(region.startRow, endRow, region.startColumn, endCol));
        }

        return result;
    }

    /// <inheritdoc/>
    public List<(int row, double height)> GetRowHeights(IXLWorksheet worksheet)
    {
        if (worksheet.RowsUsed().Count() == 0)
            return new List<(int row, double height)>();

        var rowsVal = new List<(int row, double height)>();
        int[] rowPositions = Toolbox.GetValuesInBetween(worksheet.RowsUsed().First().RowNumber(), worksheet.RowsUsed().Last().RowNumber());
        for (int i = 0; i < rowPositions.Length; i++)
        {
            var rowHeight = worksheet.Row(rowPositions[i]).Height;
            rowsVal.Add((rowPositions[i] - 1, Toolbox.ConvertToRowPixels(rowHeight)[0]));
        }
        return rowsVal;
    }
}
