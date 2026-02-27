using ClosedXML.Excel;
using UniverBlazored.Generic.Data;
using UniverBlazored.Generic.Services;
using UniverBlazored.SpreadsheetConverter.Services.IO.Data;
using UniverBlazored.Spreadsheets.Data.Styles;
using UniverBlazored.Spreadsheets.Data.Workbook;
using UniverBlazored.Spreadsheets.Services;

namespace UniverBlazored.SpreadsheetConverter.Services.IO;

/// <summary>
/// Reading operations for the worksheet
/// </summary>
public class USpreadsheetReader : ISpreadsheetReader<UWorksheetInfo>
{
    /// <inheritdoc/>
    public async Task GetAccesibilityAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        // Pending...
    }

    /// <inheritdoc/>
    public async Task GetColumnsAndRowsAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        foreach (var col in worksheet.ColumnWidths)
            await agent.RowColumns.OnSheet(worksheet.SheetInfo).SetColumnWidth(col.column, col.width);
    }

    /// <inheritdoc/>
    public async Task GetCommentsAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent, UniverUserManager userManager)
    {
        UniverUser current = await userManager.GetCurrentUser();
        foreach (var com in worksheet.Comments)
        {
            UniverComment comment = new();
            comment.SetDateTime();
            comment.SetText(com.TextValue);
            comment.id = Toolbox.GenerateRandomId();
            comment.personId = current.userID;
            await agent.Comments.OnSheet(worksheet.SheetInfo).OnRange(com.Reference).InsertComment(comment);
        }
    }

    /// <inheritdoc/>
    public async Task GetConditionalFormatsAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        foreach (var condFormat in worksheet.ConditionalFormats)
            await agent.ConditionalFormats.OnSheet(worksheet.SheetInfo).AddConditionalFormat(condFormat.Type, condFormat.Style, condFormat.Ranges.ToArray());
    }

    /// <inheritdoc/>
    public async Task GetDataAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        foreach (var dataKvP in worksheet.RangesData)
        {
            await agent.Data.OnSheet(worksheet.SheetInfo).OnRange(dataKvP.Key).SetValue(dataKvP.Value.Values);
            await agent.Data.OnSheet(worksheet.SheetInfo).OnRange(dataKvP.Key).SetFormula(dataKvP.Value.Formulas);
        }
    }

    /// <inheritdoc/>
    public async Task GetFilterAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        await agent.Ranges.OnSheet(worksheet.SheetInfo).OnRange(worksheet.Filter).CreateFilter();
    }

    /// <inheritdoc/>
    public async Task GetFreezeAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        await agent.RowColumns.OnSheet(worksheet.SheetInfo).SetFreeze(worksheet.FreezeReference.endRow, worksheet.FreezeReference.endColumn);
    }

    /// <inheritdoc/>
    public async Task GetImagesAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        foreach (var image in worksheet.Images)
            await agent.Images.OnSheet(worksheet.SheetInfo).AddImage(image.DataUri, image.StartCell.startRow, image.StartCell.startColumn, image.PixelsLeft, image.PixelsTop);
    }

    /// <inheritdoc/>
    public async Task GetMergesAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        foreach (var merge in worksheet.MergedRanges)
        {
            // Any merge that is outside of the limits, will be ignored
            if (IsOutsideMaxRange(merge, worksheet.SheetInfo.maxUsed))
                continue;
            await agent.Ranges.OnSheet(worksheet.SheetInfo).OnRange(merge).Merge(MergeStrategy.ALL, true);
        }
    }

    /// <inheritdoc/>
    public async Task GetStylesAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        foreach (var style in worksheet.RangeStyles)
        {
            var rangesStyle = style.Ranges.Where(r => !IsOutsideMaxRange(r, worksheet.SheetInfo.maxUsed));
            if (rangesStyle.Count() == 0)
                continue;
            await agent.Styles.OnSheet(worksheet.SheetInfo).SetStylesAsync(style.FontProperties, rangesStyle.ToArray());
            await agent.Styles.OnSheet(worksheet.SheetInfo).SetBordersAsync(style.Borders, rangesStyle.ToArray());
        }
    }

    bool IsOutsideMaxRange(URange range, URange maxRange) => range.startRow > maxRange.endRow || range.endRow > maxRange.endRow || range.startColumn > maxRange.endColumn || range.endColumn > maxRange.endColumn;
}