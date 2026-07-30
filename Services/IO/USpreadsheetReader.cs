using ClosedXML.Excel;
using UniverBlazored.Generic.Data;
using UniverBlazored.Generic.Services;
using UniverBlazored.SpreadsheetConverter.Services.IO.Data;
using UniverBlazored.Spreadsheets.Data.Accessibility;
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
        if (worksheet.PermissionConfig.Points == null)
            return;

        var canEdit = worksheet.PermissionConfig.Points.TryGetValue("WorksheetEdit", out bool editAllowed) && editAllowed;
        if (canEdit)
        {
            await agent.Accessibility().OnSheet(worksheet.SheetInfo).SetWorksheetPermission(EWorksheetPermissionPoint.Edit, true);
            return;
        }

        await agent.Accessibility().OnSheet(worksheet.SheetInfo).ApplyWorksheetConfig(worksheet.PermissionConfig);
    }

    /// <inheritdoc/>
    public async Task GetColumnsAndRowsAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        var tasks = new List<Task>();
        foreach (var col in worksheet.ColumnWidths)
            tasks.Add(agent.RowColumns().OnSheet(worksheet.SheetInfo).SetColumnWidth(col.column, col.width));
        await Task.WhenAll(tasks);
    }

    /// <inheritdoc/>
    public async Task GetCommentsAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent, UniverUserManager userManager)
    {
        UniverUser current = await userManager.GetCurrentUser();
        var tasks = new List<Task>();
        foreach (var com in worksheet.Comments)
        {
            UniverComment comment = new();
            comment.SetDateTime();
            comment.SetText(com.TextValue);
            comment.id = Toolbox.GenerateRandomId();
            comment.personId = current.userID;
            tasks.Add(agent.Comments().OnSheet(worksheet.SheetInfo).OnRange(com.Reference).InsertComment(comment));
        }
        await Task.WhenAll(tasks);
    }

    /// <inheritdoc/>
    public async Task GetConditionalFormatsAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        var tasks = new List<Task>();
        foreach (var condFormat in worksheet.ConditionalFormats)
            tasks.Add(agent.ConditionalFormats().OnSheet(worksheet.SheetInfo).AddConditionalFormat(condFormat.Type, condFormat.Style, condFormat.Ranges.ToArray()));
        await Task.WhenAll(tasks);
    }

    /// <inheritdoc/>
    public async Task GetDataAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        var tasks = new List<Task>();
        foreach (var dataKvP in worksheet.RangesData)
        {
            tasks.Add(agent.Data().OnSheet(worksheet.SheetInfo).OnRange(dataKvP.Key).SetValue(dataKvP.Value.Values));
            tasks.Add(agent.Data().OnSheet(worksheet.SheetInfo).OnRange(dataKvP.Key).SetFormula(dataKvP.Value.Formulas));
        }
        await Task.WhenAll(tasks);
    }

    /// <inheritdoc/>
    public async Task GetFilterAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        await agent.Ranges().OnSheet(worksheet.SheetInfo).OnRange(worksheet.Filter).CreateFilter();
    }

    /// <inheritdoc/>
    public async Task GetFreezeAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        await agent.RowColumns().OnSheet(worksheet.SheetInfo).SetFreeze(worksheet.FreezeReference.endRow, worksheet.FreezeReference.endColumn);
    }

    /// <inheritdoc/>
    public async Task GetImagesAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        var tasks = new List<Task>();
        foreach (var image in worksheet.Images)
            tasks.Add(agent.Images().OnSheet(worksheet.SheetInfo).AddImage(image.DataUri, image.StartCell.startRow, image.StartCell.startColumn, image.PixelsLeft, image.PixelsTop));
        await Task.WhenAll(tasks);
    }

    /// <inheritdoc/>
    public async Task GetMergesAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        var tasks = new List<Task>();
        foreach (var merge in worksheet.MergedRanges)
        {
            // Any merge that is outside of the limits, will be ignored
            if (Toolbox.IsOutsideMaxRange(merge, worksheet.SheetInfo.maxUsed))
                continue;
            tasks.Add(agent.Ranges().OnSheet(worksheet.SheetInfo).OnRange(merge).Merge(MergeStrategy.ALL, true));
        }
        await Task.WhenAll(tasks);
    }

    /// <inheritdoc/>
    public async Task GetStylesAsync(UWorksheetInfo worksheet, UniverSpreadsheetAgent agent)
    {
        var tasks = new List<Task>();
        foreach (var style in worksheet.RangeStyles)
        {
            var rangesStyle = style.Ranges.Where(r => !Toolbox.IsOutsideMaxRange(r, worksheet.SheetInfo.maxUsed));
            if (rangesStyle.Count() == 0)
                continue;
            tasks.Add(agent.Styles().OnSheet(worksheet.SheetInfo).SetStylesAsync(style.FontProperties, rangesStyle.ToArray()));
            tasks.Add(agent.Styles().OnSheet(worksheet.SheetInfo).SetBordersAsync(style.Borders, rangesStyle.ToArray()));
        }
        await Task.WhenAll(tasks);
    }
}
