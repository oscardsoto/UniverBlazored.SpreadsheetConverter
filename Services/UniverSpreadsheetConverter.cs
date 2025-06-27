namespace UniverBlazored.SpreadsheetConverter.Services;

using System.Drawing;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Microsoft.Extensions.Options;
using UniverBlazored.Generic.Data;
using UniverBlazored.Generic.Services;
using UniverBlazored.Spreadsheets.Data.ConditionFormat;
using UniverBlazored.Spreadsheets.Data.Styles;
using UniverBlazored.Spreadsheets.Data.Workbook;
using UniverBlazored.Spreadsheets.Services;

/// <summary>
/// Interface for managing spreadsheet data.
/// Provides methods to set/get data into/from an XLWorkbook using agent and userManager details, and optionally just populate data.
/// </summary>
public class UniverSpreadsheetConverter : IUniverSpreadsheetConverter
{
    private readonly UniverSpreadsheetConverterConfig config;

    /// <summary>
    /// Interface for managing spreadsheet data.
    /// Provides methods to set/get data into/from an XLWorkbook using agent and userManager details, and optionally just populate data.
    /// </summary>
    /// <param name="options">Configuration object</param>
    public UniverSpreadsheetConverter(IOptions<UniverSpreadsheetConverterConfig> options)
    {
        config = options.Value;
    }

    /// <summary>
    /// Sets spreadsheet data based on provided agent and userManager.
    /// </summary>
    /// <param name="agent">Agent for setting up the workbook.</param>
    /// <param name="userManager">User service for setting up the workbook.</param>
    /// <param name="options">Flags to select wich data will be setted</param>
    public async Task<XLWorkbook> SetInformationAsync(UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options)
    {
        var workbook = new XLWorkbook();
        var sheetInf = await agent.GetSheetsInfo();
        foreach (var sheet in sheetInf)
        {
            // Set worksheet info in workbook (Excel)
            var worksheet = workbook.Worksheets.Add(sheet.name);
            if (!string.IsNullOrEmpty(sheet.tabColor))
                worksheet.SetTabColor(Toolbox.ConvertHexToARGB(sheet.tabColor));

            // Selects the sheet in univer
            await agent.SetActiveSheet(sheet.id);

            // All operations to sheet (to bad that the sheet cant work asyncronous... it will be more faster)
            if (options.RecData)
                await SetData(agent, worksheet, sheet.maxUsed);
            if (options.RecStyles)
                await SetStyles(agent, worksheet, sheet.maxUsed);
            if (options.RecMerges)
                await SetMerges(agent, worksheet);
            if (options.RecFilters)
                await SetFilters(agent, worksheet);
            if (options.RecFreeze)
                await SetFreeze(agent, worksheet);
            if (options.RecComments)
                await SetComments(agent, userManager, worksheet);
            if (options.RecColumnsAndRows)
                await SetColumnsAndRows(agent, worksheet, sheet.maxUsed);
            if (options.RecImages)
                await SetImages(agent, worksheet);
            if (options.RecConditionalFormats)
                await SetConditionalFormats(agent, worksheet);
            // Pending...
            if (options.RecAccesibility)
                await SetAccesibility(agent, worksheet);
        }

        return workbook;
    }

    /// <summary>
    /// Sets spreadsheet data based on provided agent and userManager.
    /// </summary>
    /// <param name="workbook">Workbook to put the sheet data</param>
    /// <param name="unvrSheet">Data of the sheet (Univer) that will be readed</param>
    /// <param name="agent">Univer's agent</param>
    /// <param name="userManager">Univer's user manager</param>
    /// <param name="options">Flags to select wich data will be setted</param>
    /// <returns></returns>
    public async Task SetInformationInSheetAsync(XLWorkbook workbook, USheetInfo unvrSheet, UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options)
    {
        var worksheet = workbook.Worksheets.Add(unvrSheet.name);
        if (!string.IsNullOrEmpty(unvrSheet.tabColor))
            worksheet.SetTabColor(Toolbox.ConvertHexToARGB(unvrSheet.tabColor));

        // Selects the sheet in univer
        await agent.SetActiveSheet(unvrSheet.id);

        // All operations to sheet (to bad that the sheet cant work asyncronous... it will be more faster)
        if (options.RecData)
            await SetData(agent, worksheet, unvrSheet.maxUsed);
        if (options.RecStyles)
            await SetStyles(agent, worksheet, unvrSheet.maxUsed);
        if (options.RecMerges)
            await SetMerges(agent, worksheet);
        if (options.RecFilters)
            await SetFilters(agent, worksheet);
        if (options.RecFreeze)
            await SetFreeze(agent, worksheet);
        if (options.RecComments)
            await SetComments(agent, userManager, worksheet);
        if (options.RecColumnsAndRows)
            await SetColumnsAndRows(agent, worksheet, unvrSheet.maxUsed);
        if (options.RecImages)
            await SetImages(agent, worksheet);
        if (options.RecConditionalFormats)
            await SetConditionalFormats(agent, worksheet);
        // Pending...
        if (options.RecAccesibility)
            await SetAccesibility(agent, worksheet);
    }

    async Task SetData(UniverSpreadsheetAgent agent, IXLWorksheet worksheet, URange maxUsed)
    {
        // Selects the range where the data will be extracted
        int maxRow = maxUsed.endRow,
            maxCol = maxUsed.endColumn,
            rowCounter = 0,
            rowsPerProcess = Toolbox.MaxRowsPerProcess(config.MaxCellsReaded, maxRow + 1, maxCol + 1);
        
        // Gets all result from both values and formulas for each chunk
        var listResults = new List<(URange chunk, object[][] values, string[][] formules)>();
        do
        {
            URange chunk = new(rowCounter, rowCounter + rowsPerProcess, 0, maxCol);
            await agent.SetActiveRange(chunk);
            var _getValues = await agent.GetValues();
            var _getFormulas = await agent.GetFormulas();
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
                            if (DateTime.TryParse(val, out _))
                            {
                                cell.SetValue(DateTime.Parse(val));
                                break;
                            }

                            if (TimeSpan.TryParse(val, out _))
                            {
                                cell.SetValue(TimeSpan.Parse(val));
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

    async Task SetMerges(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        var merges = await agent.GetAllMerges();
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

    async Task SetFreeze(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        var freeze = await agent.GetFreeze();
        if (freeze.startColumn == -1 && freeze.startRow == -1)
            return;
        worksheet.SheetView.FreezeRows(freeze.startRow + freeze.xSplit);
        worksheet.SheetView.FreezeColumns(freeze.startColumn + freeze.ySplit);
    }

    async Task SetStyles(UniverSpreadsheetAgent agent, IXLWorksheet worksheet, URange maxUsed)
    {        
        // Selects the range where the data will be extracted
        int maxRow          = maxUsed.endRow,
            maxCol          = maxUsed.endColumn,
            rowCounter      = 0,
            rowsPerProcess  = Toolbox.MaxRowsPerProcess(config.MaxCellsReaded, maxRow + 1, maxCol + 1);

        // Gets all results from styles for each chunk
        var listResults = new List<(URange chunk, UStyleData[][] styles)>();
        do
        {
            URange chunk = new(rowCounter, rowCounter + rowsPerProcess, 0, maxCol);
            await agent.SetActiveRange(chunk);
            listResults.Add(new (chunk, await agent.GetStyles()));
            rowCounter += rowsPerProcess;
        }
        while (rowCounter < maxRow);

        // Check and sets every value in the chunk. If there's a formula in the position, the value is ignored. Otherwise insert the value
        foreach (var result in listResults)
        {
            var styles = result.styles;

            // Check wich position of each array has styles. 
            int x = 0,
                y = 0;
            do
            {
                UStyleData posData = styles[x][y];

                // If the style is null or default, it will be ignored - UStyleData.IsDefault(posData) is ignored for testing
                if (posData != null)
                {
                    var cellStyle = worksheet.Cell(result.chunk.startRow + 1 + x, result.chunk.startColumn + 1 + y).Style;
                    
                    var border = cellStyle.Border;
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

                    var font = cellStyle.Font;
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
                    
                    cellStyle.Alignment.SetVertical(posData.vt.ToVerticalValue());
                    cellStyle.Alignment.SetHorizontal(posData.ht.ToHorizontalValue());
                    if (posData.bg.HasValue && !string.IsNullOrEmpty(posData.bg.Value.rgb))
                        cellStyle.Fill.SetBackgroundColor(Toolbox.ConvertHexToARGB(posData.bg.Value.rgb));

                    cellStyle.Alignment.SetWrapText(!posData.tb.Equals(EWrapStrategy.UNSPECIFIED)); // Doesnt support wrap strategies!!
                    switch (posData.td)
                    {
                        case ETextDirection.LEFT_TO_RIGHT:
                            cellStyle.Alignment.ReadingOrder = XLAlignmentReadingOrderValues.LeftToRight;
                            break;

                        case ETextDirection.RIGHT_TO_LEFT:
                            cellStyle.Alignment.ReadingOrder = XLAlignmentReadingOrderValues.RightToLeft;
                            break;

                        case ETextDirection.UNSPECIFIED:
                            cellStyle.Alignment.ReadingOrder = XLAlignmentReadingOrderValues.ContextDependent;
                            break;
                    }
                    if (posData.tr.HasValue)
                        cellStyle.Alignment.SetTextRotation(posData.tr.Value.a);

                    if (posData.n.HasValue)
                        if (posData.n.Value.IsForNumber())
                            cellStyle.NumberFormat.SetFormat(posData.n.Value.pattern);
                        else if (posData.n.Value.IsForDate())
                            cellStyle.DateFormat.SetFormat(posData.n.Value.pattern);
                    
                    /*
                        Ignored:
                        - Overline
                        - Padding
                        - Bottom Border Line
                        - Subscript (for chinese)
                    */
                    #endregion
                }
                
                if (y + 1 < styles[x].Length)
                {
                    y++;
                    continue;
                }

                y = 0;
                x++;
            }
            while (x < styles.Length);
        }

    }

    async Task SetFilters(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        if (!await agent.HasFilter())
            return;

        var filter = await agent.GetFilter();
        worksheet.Range(filter.Value.ToA1Notation()).SetAutoFilter();
    }

    async Task SetConditionalFormats(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        var conditionals = await agent.GetAllConditionalFormats();
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

    async Task SetImages(UniverSpreadsheetAgent agent, IXLWorksheet worksheet)
    {
        var imagesId = await agent.GetImagesId();
        foreach (var imgId in imagesId)
        {
            var imageInfo = await agent.GetImage(imgId, false);
            imageInfo.source = await agent.GetImageSource(imgId);
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

    async Task SetComments(UniverSpreadsheetAgent agent, UniverUserManager userManager, IXLWorksheet worksheet)
    {
        var comments = await agent.GetComments();
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

    async Task SetColumnsAndRows(UniverSpreadsheetAgent agent, IXLWorksheet worksheet, URange maxUsed)
    {
        // Columns
        int[] colPositions  = Toolbox.GetValuesInBetween(maxUsed.startColumn, maxUsed.endColumn);
        double[] colWidths  = await agent.GetColumnWidth(colPositions);
        colWidths           = Toolbox.ConvertToColumnPoints(colWidths);
        for (int i = 0; i < colPositions.Length; i++)
            worksheet.Column(colPositions[i] + 1).Width = colWidths[i];

        // Rows
        int[] rowPositions  = Toolbox.GetValuesInBetween(maxUsed.startRow, maxUsed.endRow);
        double[] rowHeights = await agent.GetRowsHeights(rowPositions);
        rowHeights          = Toolbox.ConvertToRowPoints(rowHeights);
        for (int i = 0; i < rowPositions.Length; i++)
            worksheet.Row(rowPositions[i] + 1).Height = rowHeights[i];
    }

    // Pending...
    async Task SetAccesibility(UniverSpreadsheetAgent agent, IXLWorksheet worksheet) { }

    /// <summary>
    /// Get data into a workbook using provided agent and worksheet.
    /// </summary>
    /// <param name="workbook">Worksheet containing source data for getting into an agent.</param>
    /// <param name="userManager">User service for setting up the workbook.</param>
    /// <param name="agent">Agent for storing fetched data.</param>
    /// <param name="options">Flags to select wich data will be setted</param>
    public async Task GetInformationInAgentAsync(XLWorkbook workbook, UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options)
    {
        foreach (var worksheet in workbook.Worksheets)
            await GetInformationInAgentAsync(worksheet, agent, userManager, options);
    }

    /// <summary>
    /// Gets information into an agent from a worksheet (IXLWorksheet). 
    /// </summary>
    /// <param name="worksheet">Worksheet containing source data for getting into an agent.</param>
    /// <param name="userManager">User service for setting up the workbook.</param>
    /// <param name="agent">Agent for storing fetched data.</param>
    /// <param name="options">Flags to select wich data will be setted</param>
    public async Task GetInformationInAgentAsync(IXLWorksheet worksheet, UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options)
    {
        // Adding the sheet, sets it as active
        if (worksheet.LastRowUsed() == null || worksheet.LastColumnUsed() == null)
            return;

        string colorHex = worksheet.TabColor.ColorType is XLColorType.Theme ?
                            Toolbox.ColorToHexString(worksheet.Workbook.Theme.ResolveThemeColor(worksheet.TabColor.ThemeColor).Color):
                            Toolbox.ColorToHexString(worksheet.TabColor.Color);

        var sheetAdded = await agent.AddNewSheet(
                worksheet.Name, 
                worksheet.LastRowUsed().RowNumber(), 
                worksheet.LastColumnUsed().ColumnNumber(), 
                colorHex.Equals(Toolbox.ColorToHexString(Color.FromArgb(0,0,0,0))) ? null : colorHex);

        if (options.RecData)
            await GetData(worksheet, agent, sheetAdded.maxUsed);
        if (options.RecStyles)
            await GetStyles(worksheet, agent, sheetAdded.maxUsed);
        if (options.RecMerges)
            await GetMerges(worksheet, agent);
        if (options.RecFilters)
            await GetFilter(worksheet, agent);
        if (options.RecFreeze)
            await GetFreeze(worksheet, agent);
        if (options.RecComments)
            await GetComments(worksheet, agent, userManager);
        if (options.RecColumnsAndRows)
            await GetColumnsAndRows(worksheet, agent);
        if (options.RecImages)
            await GetImages(worksheet, agent);
        if (options.RecConditionalFormats)
            await GetConditionalFormats(worksheet, agent);
        // Pending...
        if (options.RecAccesibility)
            await GetAccesibility(worksheet, agent);
    }

    async Task GetData(IXLWorksheet worksheet, UniverSpreadsheetAgent agent, URange maxUsed)
    {
        // Selects the range where the data will be extracted
        int maxRow          = maxUsed.endRow,
            maxCol          = maxUsed.endColumn,
            rowCounter      = 0,
            rowsPerProcess  = Toolbox.MaxRowsPerProcess(config.MaxCellsReaded, maxRow + 1, maxCol + 1);

        // Gets all results from each value
        do
        {
            int firstRow    = rowCounter + 1,
                lastRow     = rowCounter + rowsPerProcess + 1;
            var range       = worksheet.Range(firstRow, 1, lastRow, maxCol + 1);
            
            var results = ExtractValuesAndFormulas(range);

            // Sets here, synchronously, each chunk of data in the Agent
            await agent.SetActiveRange(new(firstRow - 1, lastRow - 2, 0, maxCol));
            await agent.SetValue(results.values);
            await agent.SetFormula(results.formulas);
            rowCounter += rowsPerProcess;
        }
        while (rowCounter < maxRow);
    }

    (object[][] values, string[][] formulas) ExtractValuesAndFormulas(IXLRange range)
    {
        int firstRow    = range.FirstRow().RowNumber(),
            lastRow     = range.LastRow().RowNumber(),
            firstColumn = range.FirstColumn().ColumnNumber(),
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

    async Task GetStyles(IXLWorksheet worksheet, UniverSpreadsheetAgent agent, URange maxUsed)
    {
        // Selects the range where the data will be extracted
        int maxRow          = maxUsed.endRow,
            maxCol          = maxUsed.endColumn,
            rowCounter      = 0,
            rowsPerProcess  = Toolbox.MaxRowsPerProcess(config.MaxCellsReaded, maxRow + 1, maxCol + 1);

        IXLTheme theme = worksheet.Workbook.Theme;
        do
        {
            int firstRow    = rowCounter + 1,
                lastRow     = rowCounter + rowsPerProcess + 1;
            var range       = worksheet.Range(firstRow, 1, lastRow, maxCol + 1);

            IXLStyle[][] styles = range.Rows().Select(row => row.Cells().Select(c => c.Style).ToArray()).ToArray();
            
            int x = 0,
                y = 0;
            do
            {
                var style = styles[x][y];
                if (!style.IsDefaultStyle(theme))
                {
                    var univerStyle = style.ToFontProperties(theme);
                    var univerBordr = style.ToBorderProperties(theme);

                    await agent.SetActiveRange(new(x, y));
                    await agent.SetFontProperties(univerStyle);
                    foreach (var data in univerBordr)
                        await agent.SetBorderStyle(data.type, data.style, data.color);
                }

                if (y + 1 < styles[x].Length)
                {
                    y++;
                    continue;
                }

                y = 0;
                x++;
            }
            while(x < styles.Length);

            rowCounter += rowsPerProcess;
        }
        while (rowCounter < maxRow);
    }

    async Task GetMerges(IXLWorksheet worksheet, UniverSpreadsheetAgent agent)
    {
        var merges = worksheet.MergedRanges;
        foreach (var merge in merges)
        {
            string reference    = merge.RangeAddress.ToString(XLReferenceStyle.A1);
            URange mergeRange   = new(reference);
            await agent.SetActiveRange(mergeRange);
            await agent.Merge(MergeStrategy.ALL, true);
        }
    }

    async Task GetFilter(IXLWorksheet worksheet, UniverSpreadsheetAgent agent)
    {
        if (worksheet.AutoFilter.IsEnabled)
        {
            var reference       = worksheet.AutoFilter.Range.RangeAddress.ToString(XLReferenceStyle.A1);
            URange filterRange  = new(reference);
            await agent.SetActiveRange(filterRange);
            await agent.CreateFilter();
        }
    }

    async Task GetFreeze(IXLWorksheet worksheet, UniverSpreadsheetAgent agent)
    {
        await agent.SetFreeze(worksheet.SheetView.SplitRow, worksheet.SheetView.SplitColumn);
    }

    async Task GetComments(IXLWorksheet worksheet, UniverSpreadsheetAgent agent, UniverUserManager userManager)
    {
        UniverUser current = await userManager.GetCurrentUser();

        var allComments = worksheet.CellsUsed().Select(c => c.HasComment ? new { range = c.Address.ToString(XLReferenceStyle.A1), comment = c.GetComment() } : null).ToList();

        foreach (var com in allComments)
        {
            if (com == null)
                continue;

            await agent.SetActiveRange(new URange(com.range));

            UniverComment comment = new();
            comment.SetDateTime();
            comment.SetText(com.comment.Text);
            comment.id = Toolbox.GenerateRandomId();
            comment.personId = current.userID;
            await agent.InsertComment(comment);
        }
    }

    async Task GetImages(IXLWorksheet worksheet, UniverSpreadsheetAgent agent)
    {
        foreach (var img in worksheet.Pictures)
        {
            string dataURI = $"data:{img.Format.ToImageType()};base64,{Toolbox.ConvertToBase64(img.ImageStream)}";
            string cellTop = img.TopLeftCell.Address.ToString(XLReferenceStyle.A1);
            URange cellImg = new(cellTop);
            await agent.AddImage(dataURI, cellImg.startRow, cellImg.startColumn, Toolbox.ConvertToPixels(img.Left), Toolbox.ConvertToPixels(img.Top));
        }
    }

    async Task GetConditionalFormats(IXLWorksheet worksheet, UniverSpreadsheetAgent agent)
    {
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

                    await agent.AddConditionalFormat(
                        new() { IsIconSet = true }, 
                        new() { IconSet = iconConfig },
                        rangesIcon.ToArray()
                    );
                    break;

                case XLConditionalFormatType.DataBar:
                    var barConfig           = new UDataBarConfig();
                    barConfig.isShowValue   = true;
                    barConfig.isGradient    = true;
                    XLColor color           = cdFt.Colors.First().Value;
                    barConfig.positiveColor = color.ColorType is XLColorType.Theme ? Toolbox.ColorToHexString(theme.ResolveThemeColor(color.ThemeColor).Color) : Toolbox.ColorToHexString(color.Color);
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

                    await agent.AddConditionalFormat(
                        new() { IsDataBar = true },
                        new() { DataBar = barConfig },
                        rangesData.ToArray()
                    );
                    break;

                case XLConditionalFormatType.ColorScale:
                    var colorConfig = new UColorScale();
                    var colorList   = new List<UColorScaleConfig>();

                    foreach (var kvp in cdFt.Values)
                    {
                        var config      = new UColorScaleConfig();
                        config.index    = kvp.Key;
                        config.color    = cdFt.Colors[kvp.Key].ColorType is XLColorType.Theme ? 
                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(cdFt.Colors[kvp.Key].ThemeColor).Color):
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
                    await agent.AddConditionalFormat(
                        new() { IsColorScale = true },
                        new() { ColorScale = colorConfig },
                        rangesColor.ToArray()
                    );
                    break;

                default:
                    styleFrmt = cdFt.Style;
                    var condFormStyle = new UConditionFormatStyle()
                    {
                        Background      = styleFrmt.Fill.BackgroundColor.ColorType is XLColorType.Theme ?
                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(styleFrmt.Fill.BackgroundColor.ThemeColor).Color):
                                            Toolbox.ColorToHexString(styleFrmt.Fill.BackgroundColor.Color),
                        IsBold          = styleFrmt.Font.Bold,
                        FontColor       = styleFrmt.Font.FontColor.ColorType is XLColorType.Theme ?
                                            Toolbox.ColorToHexString(theme.ResolveThemeColor(styleFrmt.Font.FontColor.ThemeColor).Color):
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

                    await agent.AddConditionalFormat(condType, condFormStyle, rangesCond.ToArray());
                    break;
            }
        }
    }

    async Task GetColumnsAndRows(IXLWorksheet worksheet, UniverSpreadsheetAgent agent)
    {
        if (worksheet.ColumnsUsed().Count() == 0 || worksheet.RowsUsed().Count() == 0)
            return;

        // Columns
        int[] colPositions = Toolbox.GetValuesInBetween(worksheet.ColumnsUsed().First().ColumnNumber(), worksheet.ColumnsUsed().Last().ColumnNumber());
        for (int i = 0; i < colPositions.Length; i++)
        {
            var colWidth = worksheet.Column(colPositions[i]).Width;
            await agent.SetColumnWidth(colPositions[i] - 1, Toolbox.ConvertToColumnPixels(colWidth)[0]);
        }

        // Rows
        int[] rowPositions = Toolbox.GetValuesInBetween(worksheet.RowsUsed().First().RowNumber(), worksheet.RowsUsed().Last().RowNumber());
        for (int i = 0; i < rowPositions.Length; i++)
        {
            var rowHeight = worksheet.Row(rowPositions[i]).Height;
            await agent.SetRowHeight(rowPositions[i] - 1, Toolbox.ConvertToRowPixels(rowHeight)[0]);
        }
    }

    // Pending...
    async Task GetAccesibility(IXLWorksheet worksheet, UniverSpreadsheetAgent agent) { } 
}