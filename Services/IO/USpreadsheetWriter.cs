using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using UniverBlazored.Generic.Services;
using UniverBlazored.Spreadsheets.Data.ConditionFormat;
using UniverBlazored.Spreadsheets.Data.Styles;
using UniverBlazored.Spreadsheets.Data.Workbook;
using UniverBlazored.Spreadsheets.Services;

namespace UniverBlazored.SpreadsheetConverter.Services.IO;

/// <summary>
/// Writing operations for the worksheet
/// </summary>
public class USpreadsheetWriter : ISpreadsheetWriter<IXLWorksheet>
{
    /// <inheritdoc/>
    public async Task SetAccesibilityAsync(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        var permissions = await agent.Accessibility().OnSheet().GetWorksheetPermissions();
        if (permissions == null || permissions.Count == 0)
            return;

        bool canEdit = permissions.TryGetValue("WorksheetEdit", out var edit) && edit;
        if (canEdit)
            return;

        XLSheetProtectionElements allowedElements = XLSheetProtectionElements.None;

        if (permissions.TryGetValue("WorksheetSort", out var canSort) && canSort)
            allowedElements |= XLSheetProtectionElements.Sort;

        if (permissions.TryGetValue("WorksheetFilter", out var canFilter) && canFilter)
            allowedElements |= XLSheetProtectionElements.AutoFilter;

        if (permissions.TryGetValue("WorksheetPivotTable", out var canPivot) && canPivot)
            allowedElements |= XLSheetProtectionElements.PivotTables;

        if (permissions.TryGetValue("WorksheetInsertColumn", out var canInsertColumn) && canInsertColumn)
            allowedElements |= XLSheetProtectionElements.InsertColumns;

        if (permissions.TryGetValue("WorksheetInsertRow", out var canInsertRow) && canInsertRow)
            allowedElements |= XLSheetProtectionElements.InsertRows;

        if (permissions.TryGetValue("WorksheetInsertHyperlink", out var canInsertHyperlink) && canInsertHyperlink)
            allowedElements |= XLSheetProtectionElements.InsertHyperlinks;

        if (permissions.TryGetValue("WorksheetDeleteColumn", out var canDeleteColumn) && canDeleteColumn)
            allowedElements |= XLSheetProtectionElements.DeleteColumns;

        if (permissions.TryGetValue("WorksheetDeleteRow", out var canDeleteRow) && canDeleteRow)
            allowedElements |= XLSheetProtectionElements.DeleteRows;

        if (permissions.TryGetValue("WorksheetSetCellStyle", out var canSetCellStyle) && canSetCellStyle)
            allowedElements |= XLSheetProtectionElements.FormatCells;

        if (permissions.TryGetValue("WorksheetSetColumnStyle", out var canSetColumnStyle) && canSetColumnStyle)
            allowedElements |= XLSheetProtectionElements.FormatColumns;

        if (permissions.TryGetValue("WorksheetSetRowStyle", out var canSetRowStyle) && canSetRowStyle)
            allowedElements |= XLSheetProtectionElements.FormatRows;

        if (permissions.TryGetValue("WorksheetEditExtraObject", out var canEditObjects) && canEditObjects)
            allowedElements |= XLSheetProtectionElements.EditObjects;

        if (permissions.TryGetValue("WorksheetSelectProtectedCells", out var canSelectProtectedCells) && canSelectProtectedCells)
            allowedElements |= XLSheetProtectionElements.SelectLockedCells;

        if (permissions.TryGetValue("WorksheetSelectUnProtectedCells", out var canSelectUnprotectedCells) && canSelectUnprotectedCells)
            allowedElements |= XLSheetProtectionElements.SelectUnlockedCells;

        worksheet.Protect(allowedElements);
    }

    /// <inheritdoc/>
    public async Task SetColumnsAndRowsAsync(UniverSpreadsheetAgent agent, IXLWorksheet worksheet, URange maxUsed)
    {
        // Columns
        int[] colPositions  = Toolbox.GetValuesInBetween(maxUsed.startColumn, maxUsed.endColumn);
        double[] colWidths  = await agent.RowColumns().OnSheet().GetColumnWidth(colPositions);
        colWidths           = Toolbox.ConvertToColumnPoints(colWidths);
        for (int i = 0; i < colPositions.Length; i++)
            worksheet.Column(colPositions[i] + 1).Width = colWidths[i];

        // Rows
        int[] rowPositions  = Toolbox.GetValuesInBetween(maxUsed.startRow, maxUsed.endRow);
        double[] rowHeights = await agent.RowColumns().OnSheet().GetRowsHeights(rowPositions);
        rowHeights          = Toolbox.ConvertToRowPoints(rowHeights);
        for (int i = 0; i < rowPositions.Length; i++)
            worksheet.Row(rowPositions[i] + 1).Height = rowHeights[i];
    }

    /// <inheritdoc/>
    public async Task SetCommentsAsync(UniverSpreadsheetAgent agent, UniverUserManager userManager, IXLWorksheet worksheet)
    {
        var comments = await agent.Comments().OnSheet().GetComments();
        if (comments == null)
            return;

        foreach (var comment in comments)
        {
            var user = await userManager.GetUser(comment.personId);
            worksheet.Cell(comment.reference).CreateComment()
                .SetAuthor(user.name)
                .AddSignature()
                .AddText(comment.text.dataStream);
        }
    }

    /// <inheritdoc/>
    public async Task SetConditionalFormatsAsync(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        var conditionals = await agent.ConditionalFormats().OnSheet().GetAllConditionalFormats();
        if (conditionals == null)
            return;

        foreach (var cond in conditionals)
        {
            var wsCond = worksheet.AddConditionalFormat();
            wsCond.SetStopIfTrue(cond.stopIfTrue);
            var ruleType = cond.GetTypeConditionalFormat();
            switch (ruleType)
            {
                case ECFRuleType.dataBar:
                    var configDB  = cond.GetDataBarConfig();
                    var dataBar = wsCond.DataBar(Toolbox.ConvertHexToARGB(configDB?.positiveColor), Toolbox.ConvertHexToARGB(configDB?.nativeColor), (bool)configDB?.isShowValue);
                    var typeDataMin = Toolbox.ConvertToContentType(configDB?.min.GetValueType());
                    IXLCFDataBarMax max = default;
                    JsonElement valMin;
                    switch (typeDataMin)
                    {
                        case XLCFContentType.Minimum:
                            dataBar.LowestValue().HighestValue();
                            break;

                        case XLCFContentType.Formula:
                            valMin = (JsonElement) configDB.Value.min.value;
                            max = dataBar.Minimum(typeDataMin, valMin.GetString());
                            break;

                        default:
                            valMin = (JsonElement) configDB.Value.min.value;
                            max = dataBar.Minimum(typeDataMin, valMin.GetDouble());
                            break;

                    }

                    var typeDataMax = Toolbox.ConvertToContentType(configDB?.max.GetValueType());
                    JsonElement valMax;
                    switch (typeDataMax)
                    {
                        case XLCFContentType.Maximum: // Ignores it
                            break;

                        case XLCFContentType.Formula:
                            valMax = (JsonElement) configDB.Value.max.value;
                            max.Maximum(typeDataMax, valMax.GetString());
                            break;

                        default:
                            valMax = (JsonElement) configDB.Value.max.value;
                            max.Maximum(typeDataMax, valMax.GetDouble());
                            break;
                    }
                    break;

                case ECFRuleType.colorScale:
                    var configCS = cond.GetColorScaleConfigs();
                    var colorScale = wsCond.ColorScale();

                    (IXLCFColorScaleMid mid, IXLCFColorScaleMax max) selectPos = new();
                    for (int i = 0; i < configCS.Length; i++)
                    {
                        var posValType    = Toolbox.ConvertToContentType(configCS[i].value.GetValueType());
                        var posColor      = Toolbox.ConvertHexToARGB(configCS[i].color);
                        switch (posValType)
                        {
                            case XLCFContentType.Minimum:
                                if (i == 0)
                                {
                                    selectPos.mid = colorScale.LowestValue(posColor);
                                    continue;
                                }
                                break;
                            
                            case XLCFContentType.Maximum:
                                if (i == 1 && configCS.Length > 2)
                                {
                                    selectPos.mid.HighestValue(Toolbox.ConvertHexToARGB(configCS[configCS.Length - 1].color));
                                    continue;
                                }
                                break;

                            case XLCFContentType.Formula:
                                JsonElement configFormula = (JsonElement) configCS[i].value.value;
                                if (i == 0)
                                {
                                    selectPos.mid = colorScale.Minimum(posValType, configFormula.GetString(), posColor);
                                    continue;
                                }

                                if (i == 1)
                                {
                                    if (configCS.Length == 2)
                                    {
                                        selectPos.mid.Maximum(posValType, configFormula.GetString(), posColor);
                                        continue;
                                    }

                                    selectPos.max = selectPos.mid.Midpoint(posValType, configFormula.GetString(), posColor);
                                    continue;
                                }

                                selectPos.max.Maximum(posValType, configFormula.GetString(), posColor);
                                break;

                            default:
                                JsonElement configValue = (JsonElement) configCS[i].value.value;
                                if (i == 0)
                                {
                                    selectPos.mid = colorScale.Minimum(posValType, configValue.GetDouble(), posColor);
                                    continue;
                                }

                                if (i == 1)
                                {
                                    if (configCS.Length == 2)
                                    {
                                        selectPos.mid.Maximum(posValType, configValue.GetDouble(), posColor);
                                        continue;
                                    }

                                    selectPos.max = selectPos.mid.Midpoint(posValType, configValue.GetDouble(), posColor);
                                    continue;
                                }

                                selectPos.max.Maximum(posValType, configValue.GetDouble(), posColor);
                                break;
                        }
                    }
                    
                    break;

                case ECFRuleType.iconSet:
                    var configIS    = cond.GetIconSetConfigs();
                    var iconSet     = wsCond.IconSet(configIS[0].GetIconType().ToIconSet());
                    for (int i = 0; i < configIS.Length; i++)
                    {
                        var typeDataIcon = Toolbox.ConvertToContentType(configIS[i].value.GetValueType());
                        JsonElement valIcon;
                        if (typeDataIcon is XLCFContentType.Formula)
                        {
                            valIcon = (JsonElement) configIS[i].value.value;
                            iconSet.AddValue(configIS[i].GetOperator().ToIconSetOperator(), valIcon.GetString(), typeDataIcon);
                        }
                        else
                        {
                            valIcon = (JsonElement) configIS[i].value.value;
                            iconSet.AddValue(configIS[i].GetOperator().ToIconSetOperator(), valIcon.GetDouble(), typeDataIcon);
                        }
                    }
                    break;

                case ECFRuleType.highlightCell:
                    var subType = cond.GetSubType();

                    // Average is not supported!!
                    if (subType is ECFSubRuleType.average)
                        break;

                    var styleBase = cond.GetStyleUsed();
                    if (styleBase == null)
                        throw new NullReferenceException("The style in the highlihgtCell Conditional Format must exist!");

                    ECFOperators? _opertr = null;
                    IXLStyle style;
                    var jsonRule = cond.rule;

                    // Gets the style for the condition
                    switch (subType)
                    {
                        case ECFSubRuleType.uniqueValues:
                            style = wsCond.WhenIsUnique();
                            SetStyleSettingsInExcel(styleBase, style);
                            break;

                        case ECFSubRuleType.duplicateValues:
                            style = wsCond.WhenIsDuplicate();
                            SetStyleSettingsInExcel(styleBase, style);
                            break;

                        case ECFSubRuleType.rank:
                            int rValue      = (int)jsonRule["value"];
                            bool isPercent  = (bool)jsonRule["isPercent"];
                            style           = wsCond.WhenIsTop(rValue, isPercent ? XLTopBottomType.Percent : XLTopBottomType.Items);
                            SetStyleSettingsInExcel(styleBase, style);
                            break;

                        case ECFSubRuleType.formula:
                            string sValue   = jsonRule["value"].ToString();
                            style           = wsCond.WhenIsTrue(sValue);
                            SetStyleSettingsInExcel(styleBase, style);
                            break;

                        case ECFSubRuleType.text:
                            _opertr = cond.GetOperator();
                            string tValue = jsonRule["value"]?.ToString();
                            switch (_opertr)
                            {
                                case ECFOperators.beginsWith:
                                    style = wsCond.WhenStartsWith(tValue);
                                    break;

                                case ECFOperators.endsWith:
                                    style = wsCond.WhenEndsWith(tValue);
                                    break;
                                
                                case ECFOperators.containsText:
                                    style = wsCond.WhenContains(tValue);
                                    break;

                                case ECFOperators.notContainsText:
                                    style = wsCond.WhenNotContains(tValue);
                                    break;

                                case ECFOperators.notEqual:
                                    style = wsCond.WhenNotEquals(tValue);
                                    break;

                                case ECFOperators.containsBlanks:
                                    style = wsCond.WhenIsBlank();
                                    break;

                                case ECFOperators.notContainsBlanks:
                                    style = wsCond.WhenNotBlank();
                                    break;

                                case ECFOperators.containsErrors:
                                    style = wsCond.WhenIsError();
                                    break;

                                case ECFOperators.notContainsErrors:
                                    style = wsCond.WhenNotError();
                                    break;

                                case ECFOperators.equal:
                                default:
                                    style = wsCond.WhenEquals(tValue);
                                    break;
                            }
                            SetStyleSettingsInExcel(styleBase, style);
                            break;

                        case ECFSubRuleType.timePeriod:
                            _opertr = cond.GetOperator();
                            switch (_opertr)
                            {
                                case ECFOperators.yesterday:
                                    style = wsCond.WhenDateIs(XLTimePeriod.Yesterday);
                                    break;

                                case ECFOperators.last7Days:
                                    style = wsCond.WhenDateIs(XLTimePeriod.InTheLast7Days);
                                    break;

                                case ECFOperators.tomorrow:
                                    style = wsCond.WhenDateIs(XLTimePeriod.Tomorrow);
                                    break;

                                case ECFOperators.thisMonth:
                                    style = wsCond.WhenDateIs(XLTimePeriod.ThisMonth);
                                    break;

                                case ECFOperators.lastMonth:
                                    style = wsCond.WhenDateIs(XLTimePeriod.LastMonth);
                                    break;

                                case ECFOperators.nextMonth:
                                    style = wsCond.WhenDateIs(XLTimePeriod.NextMonth);
                                    break;

                                case ECFOperators.thisWeek:
                                    style = wsCond.WhenDateIs(XLTimePeriod.ThisWeek);
                                    break;

                                case ECFOperators.lastWeek:
                                    style = wsCond.WhenDateIs(XLTimePeriod.LastWeek);
                                    break;

                                case ECFOperators.nextWeek:
                                    style = wsCond.WhenDateIs(XLTimePeriod.NextWeek);
                                    break;
                                    
                                case ECFOperators.today:
                                default:
                                    style = wsCond.WhenDateIs(XLTimePeriod.Today);
                                    break;
                            }
                            SetStyleSettingsInExcel(styleBase, style);
                            break;

                        case ECFSubRuleType.number:
                            _opertr         = cond.GetOperator();
                            double nValue   = 0;
                            double[] values = [];
                            switch (_opertr)
                            {
                                case ECFOperators.greaterThan:
                                    nValue  = (double)jsonRule["value"];
                                    style   = wsCond.WhenGreaterThan(nValue);
                                    break;

                                case ECFOperators.greaterThanOrEqual:
                                    nValue  = (double)jsonRule["value"];
                                    style   = wsCond.WhenEqualOrGreaterThan(nValue);
                                    break;

                                case ECFOperators.lessThan:
                                    nValue  = (double)jsonRule["value"];
                                    style   = wsCond.WhenLessThan(nValue);
                                    break;

                                case ECFOperators.lessThanOrEqual:
                                    nValue  = (double)jsonRule["value"];
                                    style   = wsCond.WhenEqualOrLessThan(nValue);
                                    break;

                                case ECFOperators.between:
                                    values  = jsonRule["value"].Deserialize<double[]>();
                                    style   = wsCond.WhenBetween(values[0], values[1]);
                                    break;

                                case ECFOperators.notBetween:
                                    values  = jsonRule["value"].Deserialize<double[]>();
                                    style   = wsCond.WhenNotBetween(values[0], values[1]);
                                    break;

                                case ECFOperators.notEqual:
                                    nValue = (double)jsonRule["value"];
                                    style = wsCond.WhenNotEquals(nValue);
                                    break;

                                case ECFOperators.equal:
                                default:
                                    nValue = (double)jsonRule["value"];
                                    style = wsCond.WhenEquals(nValue);
                                    break;
                            }
                            SetStyleSettingsInExcel(styleBase, style);
                            break;
                    }
                    break;
            }

            cond.ranges.ForEach(r => wsCond.Ranges.Add(worksheet.Range(r.ToA1Notation())));
        }
    }

    void SetStyleSettingsInExcel(IStyleBase settings, IXLStyle style)
    {
        style.Font.SetBold(settings.bl != 0);
        style.Font.SetItalic(settings.it != 0);
        if (settings.ul.s != 0)                             // Univer doesnt support different types of underline
            style.Font.SetUnderline();
        style.Font.SetStrikethrough(settings.st.s != 0);    // Doesnt support color!!
        if (settings.cl.HasValue)
           style.Font.SetFontColor(Toolbox.ConvertHexToARGB(settings.cl.Value.rgb));
        if (settings.bg.HasValue)
            style.Fill.SetBackgroundColor(Toolbox.ConvertHexToARGB(settings.bg.Value.rgb));
    }

    /// <inheritdoc/>
    public async Task SetDataAsync(UniverSpreadsheetAgent agent, IXLWorksheet worksheet, URange maxUsed, int maxCells)
    {
        // Selects the range where the data will be extracted
        int maxRow = maxUsed.endRow,
            maxCol = maxUsed.endColumn,
            rowCounter = 0,
            rowsPerProcess = Toolbox.MaxRowsPerProcess(maxCells, maxRow + 1, maxCol + 1);
        
        // Gets all result from both values and formulas for each chunk
        var listResults = new List<(URange chunk, object[][] values, string[][] formules)>();
        do
        {
            URange chunk = new(rowCounter, rowCounter + rowsPerProcess, 0, maxCol);
            if (Toolbox.IsOutsideMaxRange(chunk, maxUsed))
            {
                chunk.endRow = maxUsed.endRow;
                chunk.endColumn = maxUsed.endColumn;
            }

            var _getValues = await agent.Data().OnSheet().OnRange(chunk).GetValues();
            var _getFormulas = await agent.Data().OnSheet().OnRange(chunk).GetFormulas();
            listResults.Add(new (chunk, _getValues, _getFormulas));
            rowCounter += rowsPerProcess;
        }
        while (rowCounter < maxRow);

        // Check and sets every value in the chunk. If there's a formula in the position, the value is ignored. Otherwise insert the value
        foreach (var result in listResults)
        {
            var formules    = result.formules;
            var values      = result.values;

            if (Toolbox.IsEmpty(values))
                continue;

            // Check wich position of each array has values
            int x = 0, 
                y = 0;
            do
            {
                string posFormula   = formules[x][y];
                if (!string.IsNullOrEmpty(posFormula))
                    worksheet.Cell(result.chunk.startRow + 1 + x, result.chunk.startColumn + 1 + y).FormulaA1 = posFormula;
                else if (values[x][y] != null)
                {
                    JsonElement posVal = (JsonElement) values[x][y];
                    var cell = worksheet.Cell(result.chunk.startRow + 1 + x, result.chunk.startColumn + 1 + y);
                    switch (posVal.ValueKind)
                    {
                        case JsonValueKind.True:
                        case JsonValueKind.False:
                            cell.SetValue(posVal.GetBoolean());
                            break;

                        case JsonValueKind.Number:
                            cell.SetValue(posVal.GetDouble());
                            break;

                        case JsonValueKind.String:
                            var val = posVal.GetString();
                            if (DateTime.TryParse(val, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
                            {
                                if (date.Year >= 1000 && date.Year <= 9999)
                                {
                                    cell.SetValue(date);
                                    break;
                                }
                            }

                            if (TimeSpan.TryParse(val, CultureInfo.InvariantCulture, out var tiempo))
                            {
                                cell.SetValue(tiempo);
                                break;
                            }
                            cell.SetValue(val);
                            break;

                        default:
                            break;
                    }
                }
                
                if (y + 1 < formules[x].Length)
                {
                    y++;
                    continue;
                }

                y = 0;
                x++;
            }
            while (x < formules.Length);
        }
    }

    /// <inheritdoc/>
    public async Task SetFiltersAsync(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        if (!await agent.Ranges().OnSheet().HasFilter())
            return;

        var filter = await agent.Ranges().OnSheet().GetFilter();
        worksheet.Range(filter.Value.ToA1Notation()).SetAutoFilter();
    }

    /// <inheritdoc/>
    public async Task SetFreezeAsync(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        var freeze = await agent.RowColumns().OnSheet().GetFreeze();
        if (freeze.startRow != -1)
            worksheet.SheetView.FreezeRows(freeze.startRow);

        if (freeze.startColumn != -1)
            worksheet.SheetView.FreezeColumns(freeze.startColumn);
    }

    /// <inheritdoc/>
    public async Task SetImagesAsync(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        var imagesId = await agent.Images().OnSheet().GetImagesId();
        foreach (var imgId in imagesId)
        {
            var imageInfo = await agent.Images().OnSheet().GetImage(imgId, false);
            imageInfo.source = await agent.Images().OnSheet().GetImageSource(imgId);
            byte[] imageBytes = Convert.FromBase64String(imageInfo.GetBase64());
            using (MemoryStream stream = new(imageBytes))
            {
                var img     = worksheet.AddPicture(stream);
                img.MoveTo(
                    worksheet.Cell(
                        imageInfo.sheetTransform.from.row + 1,
                        imageInfo.sheetTransform.from.column + 1
                    ),
                    (int) imageInfo.sheetTransform.from.columnOffset,
                    (int) imageInfo.sheetTransform.from.rowOffset)
                    .WithSize(
                        Convert.ToInt32(imageInfo.transform.width), 
                        Convert.ToInt32(imageInfo.transform.height)
                    );
                // Doesnt accept crop images!!!
            }
        }
    }

    /// <inheritdoc/>
    public async Task SetMergesAsync(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        var merges = await agent.Ranges().OnSheet().GetAllMerges();
        if (merges.Length == 0)
            return;

        foreach (var merge in merges)
        {
            var fstCell = merge.GetFirstCellOfRange().ToA1Notation();
            worksheet.Cell(fstCell).Style.Alignment.Horizontal  = XLAlignmentHorizontalValues.Center;
            worksheet.Cell(fstCell).Style.Alignment.Vertical    = XLAlignmentVerticalValues.Center;
            worksheet.Range(merge.ToA1Notation()).Merge();
        }
    }

    /// <inheritdoc/>
    public async Task SetStylesAsync(UniverSpreadsheetAgent agent, IXLWorksheet worksheet, URange maxUsed, int maxCells)
    {
         // Selects the range where the data will be extracted
        int maxRow          = maxUsed.endRow,
            maxCol          = maxUsed.endColumn,
            rowCounter      = 0,
            rowsPerProcess  = Toolbox.MaxRowsPerProcess(maxCells, maxRow + 1, maxCol + 1);

        // Gets all results from styles for each chunk
        var listResults = new List<Dictionary<UStyleData, URange[]>>();
        do
        {
            URange chunk = new(rowCounter, rowCounter + rowsPerProcess, 0, maxCol);
            if (Toolbox.IsOutsideMaxRange(chunk, maxUsed))
            {
                chunk.endRow = maxUsed.endRow;
                chunk.endColumn = maxUsed.endColumn;
            }
            
            var styles = await agent.Styles().OnSheet().OnRange(chunk).GetStyles();
            listResults.Add(styles);
            rowCounter += rowsPerProcess;
        }
        while (rowCounter < maxRow);

        foreach (var result in listResults)
        {
            foreach (var styleData in result)
            {
                var posData = styleData.Key;
                foreach (var range in styleData.Value)
                {
                    var rangeStyle = worksheet.Range(range.ToA1Notation()).Style;

                    var border = rangeStyle.Border;
                    #region Borders
                    if (posData.bd.HasValue)
                    {
                        UBorderData borderData = posData.bd.Value;
                        if (borderData.t.HasValue)
                        {
                            var data = borderData.t.Value.GetBorderData();
                            border.SetTopBorder(data.border);
                            border.SetTopBorderColor(data.color);
                        }

                        if (borderData.r.HasValue)
                        {
                            var data = borderData.r.Value.GetBorderData();
                            border.SetRightBorder(data.border);
                            border.SetRightBorderColor(data.color);
                        }

                        if (borderData.b.HasValue)
                        {
                            var data = borderData.b.Value.GetBorderData();
                            border.SetBottomBorder(data.border);
                            border.SetBottomBorderColor(data.color);
                        }

                        if (borderData.l.HasValue)
                        {
                            var data = borderData.l.Value.GetBorderData();
                            border.SetLeftBorder(data.border);
                            border.SetLeftBorderColor(data.color);
                        }

                        if (borderData.tl_br.HasValue)
                        {
                            var data = borderData.tl_br.Value.GetBorderData();
                            border.SetDiagonalBorder(data.border);
                            border.SetDiagonalBorderColor(data.color);
                            border.SetDiagonalDown();
                        }

                        if (borderData.bl_tr.HasValue)
                        {
                            var data = borderData.bl_tr.Value.GetBorderData();
                            border.SetDiagonalBorder(data.border);
                            border.SetDiagonalBorderColor(data.color);
                            border.SetDiagonalUp();
                        }
                    }
                    // Other borders are not supported!!
                    #endregion

                    var font = rangeStyle.Font;
                    #region Font
                    font.SetBold(posData.bl != 0);
                    font.SetFontSize(posData.fs);
                    font.SetFontName(posData.ff);
                    if (posData.cl.HasValue && !string.IsNullOrEmpty(posData.cl.Value.rgb))
                        font.SetFontColor(Toolbox.ConvertHexToARGB(posData.cl.Value.rgb));

                    font.SetItalic(posData.it != 0);
                    font.SetStrikethrough(posData.st.s != 0);   // Doesnt support color!!
                    if (posData.ul.s != 0)                      // Univer doesnt support different types of underline
                        font.SetUnderline();
                    
                    rangeStyle.Alignment.SetVertical(posData.vt.ToVerticalValue());
                    rangeStyle.Alignment.SetHorizontal(posData.ht.ToHorizontalValue());
                    if (posData.bg.HasValue && !string.IsNullOrEmpty(posData.bg.Value.rgb))
                        rangeStyle.Fill.SetBackgroundColor(Toolbox.ConvertHexToARGB(posData.bg.Value.rgb));

                    rangeStyle.Alignment.SetWrapText(!posData.tb.Equals(EWrapStrategy.UNSPECIFIED)); // Doesnt support wrap strategies!!
                    switch (posData.td)
                    {
                        case ETextDirection.LEFT_TO_RIGHT:
                            rangeStyle.Alignment.ReadingOrder = XLAlignmentReadingOrderValues.LeftToRight;
                            break;

                        case ETextDirection.RIGHT_TO_LEFT:
                            rangeStyle.Alignment.ReadingOrder = XLAlignmentReadingOrderValues.RightToLeft;
                            break;

                        case ETextDirection.UNSPECIFIED:
                            rangeStyle.Alignment.ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent;
                            break;
                    }
                    if (posData.tr.HasValue)
                        rangeStyle.Alignment.SetTextRotation(posData.tr.Value.a);

                    if (posData.n.HasValue)
                        if (posData.n.Value.IsForNumber())
                            rangeStyle.NumberFormat.SetFormat(posData.n.Value.pattern);
                        else if (posData.n.Value.IsForDate())
                            rangeStyle.DateFormat.SetFormat(posData.n.Value.pattern);
                    
                    /*
                        Ignored:
                        - Overline
                        - Padding
                        - Bottom Border Line
                        - Subscript (for chinese)
                    */
                    #endregion
                }
            }
        }
    }
}
