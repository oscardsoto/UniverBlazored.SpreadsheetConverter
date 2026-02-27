namespace UniverBlazored.SpreadsheetConverter.Services;

using System.Drawing;
using ClosedXML.Excel;

using Microsoft.Extensions.Options;
using UniverBlazored.Generic.Services;
using UniverBlazored.SpreadsheetConverter.Services.IO;
using UniverBlazored.SpreadsheetConverter.Services.IO.Data;
using UniverBlazored.Spreadsheets.Data.Workbook;
using UniverBlazored.Spreadsheets.Services;

/// <summary>
/// Interface for managing spreadsheet data.
/// Provides methods to set/get data into/from an XLWorkbook using agent and userManager details, and optionally just populate data.
/// </summary>
public class UniverSpreadsheetConverter : IUniverSpreadsheetConverter<XLWorkbook, IXLWorksheet>
{
    private readonly UniverSpreadsheetConverterConfig _config;
    private readonly ISpreadsheetReader<UWorksheetInfo> _reader;
    private readonly ISpreadsheetWriter<IXLWorksheet> _writer;

    /// <summary>
    /// Interface for managing spreadsheet data.
    /// Provides methods to set/get data into/from an XLWorkbook using agent and userManager details, and optionally just populate data.
    /// </summary>
    /// <param name="reader">Reader operations</param>
    /// <param name="writer">Writer operations</param>
    /// <param name="options">Configuration object</param>
    public UniverSpreadsheetConverter(ISpreadsheetReader<UWorksheetInfo> reader, ISpreadsheetWriter<IXLWorksheet> writer, IOptions<UniverSpreadsheetConverterConfig> options)
    {
        _reader = reader;
        _writer = writer;
        _config = options.Value;
    }

    /// <inheritdoc/>
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
                await _writer.SetDataAsync(agent, worksheet, sheet.maxUsed, _config.MaxCellsReaded);
            if (options.RecStyles)
                await _writer.SetStylesAsync(agent, worksheet, sheet.maxUsed, _config.MaxCellsReaded);
            if (options.RecMerges)
                await _writer.SetMergesAsync(agent, worksheet);
            if (options.RecFilters)
                await _writer.SetFiltersAsync(agent, worksheet);
            if (options.RecFreeze)
                await _writer.SetFreezeAsync(agent, worksheet);
            if (options.RecComments)
                await _writer.SetCommentsAsync(agent, userManager, worksheet);
            if (options.RecColumnsAndRows)
                await _writer.SetColumnsAndRowsAsync(agent, worksheet, sheet.maxUsed);
            if (options.RecImages)
                await _writer.SetImagesAsync(agent, worksheet);
            if (options.RecConditionalFormats)
                await _writer.SetConditionalFormatsAsync(agent, worksheet);
            // Pending...
            if (options.RecAccesibility)
                await _writer.SetAccesibilityAsync(agent, worksheet);

            // Last thing is hidden parts
            foreach (var rowsHidden in sheet.rowsHidden)
                for (int r = rowsHidden.startRow + 1; r < rowsHidden.endRow + 1; r++)
                    worksheet.Row(r).Hide();

            foreach (var colsHidden in sheet.columnsHidden)
                for (int c = colsHidden.startColumn + 1; c < colsHidden.endColumn + 1; c++)
                    worksheet.Column(c).Hide();

            if (sheet.isHidden)
                worksheet.Hide();
        }

        return workbook;
    }

    /// <inheritdoc/>
    public async Task SetInformationInSheetAsync(XLWorkbook workbook, USheetInfo unvrSheet, UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options)
    {
        var worksheet = workbook.Worksheets.Add(unvrSheet.name);
        if (!string.IsNullOrEmpty(unvrSheet.tabColor))
            worksheet.SetTabColor(Toolbox.ConvertHexToARGB(unvrSheet.tabColor));

        // Selects the sheet in univer
        await agent.SetActiveSheet(unvrSheet.id);

        // All operations to sheet (to bad that the sheet cant work asyncronous... it will be more faster)
        if (options.RecData)
            await _writer.SetDataAsync(agent, worksheet, unvrSheet.maxUsed, _config.MaxCellsReaded);
        if (options.RecStyles)
            await _writer.SetStylesAsync(agent, worksheet, unvrSheet.maxUsed, _config.MaxCellsReaded);
        if (options.RecMerges)
            await _writer.SetMergesAsync(agent, worksheet);
        if (options.RecFilters)
            await _writer.SetFiltersAsync(agent, worksheet);
        if (options.RecFreeze)
            await _writer.SetFreezeAsync(agent, worksheet);
        if (options.RecComments)
            await _writer.SetCommentsAsync(agent, userManager, worksheet);
        if (options.RecColumnsAndRows)
            await _writer.SetColumnsAndRowsAsync(agent, worksheet, unvrSheet.maxUsed);
        if (options.RecImages)
            await _writer.SetImagesAsync(agent, worksheet);
        if (options.RecConditionalFormats)
            await _writer.SetConditionalFormatsAsync(agent, worksheet);
        // Pending...
        if (options.RecAccesibility)
            await _writer.SetAccesibilityAsync(agent, worksheet);

        // Last thing is hidden parts
        foreach (var rowsHidden in unvrSheet.rowsHidden)
            for (int r = rowsHidden.startRow + 1; r < rowsHidden.endRow + 1; r++)
                worksheet.Row(r).Hide();

        foreach (var colsHidden in unvrSheet.columnsHidden)
            for (int c = colsHidden.startColumn + 1; c < colsHidden.endColumn + 1; c++)
                worksheet.Column(c).Hide();

        if (unvrSheet.isHidden)
            worksheet.Hide();
    }

    /// <inheritdoc/>
    public async Task GetInformationInAgentAsync(XLWorkbook workbook, UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options)
    {
        var sheetTasks = new List<Task<USheetInfo>>();
        foreach (var worksheet in workbook.Worksheets)
        {
            // Adding the sheet, sets it as active
            if (worksheet.LastRowUsed() == null || worksheet.LastColumnUsed() == null)
                continue;

            string colorHex = worksheet.TabColor.ColorType is XLColorType.Theme ?
                            Toolbox.ColorToHexString(worksheet.Workbook.Theme.ResolveThemeColor(worksheet.TabColor.ThemeColor).Color, false, worksheet.TabColor.ThemeTint) :
                            Toolbox.ColorToHexString(worksheet.TabColor.Color);

            string sheetName = worksheet.Name;
            int lastRow      = worksheet.RowsUsed().Last().RowNumber(),
                lastColumn   = worksheet.ColumnsUsed().Last().ColumnNumber();
            sheetTasks.Add(agent.AddNewSheet(sheetName, lastRow, lastColumn, colorHex.Equals(Toolbox.ColorToHexString(Color.FromArgb(0, 0, 0, 0))) ? null : colorHex));
        }

        foreach (var t in sheetTasks)
        {
            USheetInfo sheetInfo = await t;
            var closedSheet = workbook.Worksheets.First(w => w.Name == sheetInfo.name);

            // Wait for it to charge
            var univerSheet = new UWorksheetInfo(closedSheet, _config.MaxCellsReaded, sheetInfo);
            var tasks = new List<Task>();

            // Get all data in Univer (Asynchronously (｡˃ ᵕ ˂ )👌 ❤️)
            if (options.RecData)
                tasks.Add(_reader.GetDataAsync(univerSheet, agent));
            if (options.RecStyles)
                tasks.Add(_reader.GetStylesAsync(univerSheet, agent));
            if (options.RecMerges)
                tasks.Add(_reader.GetMergesAsync(univerSheet, agent));
            if (options.RecFilters)
                tasks.Add(_reader.GetFilterAsync(univerSheet, agent));
            if (options.RecFreeze)
                tasks.Add(_reader.GetFreezeAsync(univerSheet, agent));
            if (options.RecComments)
                tasks.Add(_reader.GetCommentsAsync(univerSheet, agent, userManager));
            if (options.RecColumnsAndRows)
                tasks.Add(_reader.GetColumnsAndRowsAsync(univerSheet, agent));
            if (options.RecImages)
                tasks.Add(_reader.GetImagesAsync(univerSheet, agent));
            if (options.RecConditionalFormats)
                tasks.Add(_reader.GetConditionalFormatsAsync(univerSheet, agent));
            // Pending...
            if (options.RecAccesibility)
                tasks.Add(_reader.GetAccesibilityAsync(univerSheet, agent));

            // Last thing is hidden parts :D
            foreach (var col in closedSheet.ColumnsUsed())
                if (col.IsHidden)
                {
                    int colNumber = col.ColumnNumber();
                    tasks.Add(agent.RowColumns.OnSheet(univerSheet.SheetInfo).HideColumns(colNumber - 1, 1));
                }

            foreach (var row in closedSheet.RowsUsed())
                if (row.IsHidden)
                {
                    int rowNumber = row.RowNumber();
                    tasks.Add(agent.RowColumns.OnSheet(univerSheet.SheetInfo).HideRows(rowNumber - 1, 1));
                }

            if (closedSheet.Visibility is XLWorksheetVisibility.Hidden)
                tasks.Add(agent.HideSheet(univerSheet.SheetInfo.id));

            await Task.WhenAll(tasks);
        }
    }

    /// <inheritdoc/>
    public async Task GetInformationInAgentAsync(IXLWorksheet worksheet, UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options)
    {
        // Adding the sheet, sets it as active
        if (worksheet.LastRowUsed() == null || worksheet.LastColumnUsed() == null)
            return;

        string colorHex = worksheet.TabColor.ColorType is XLColorType.Theme ?
                            Toolbox.ColorToHexString(worksheet.Workbook.Theme.ResolveThemeColor(worksheet.TabColor.ThemeColor).Color, false, worksheet.TabColor.ThemeTint) :
                            Toolbox.ColorToHexString(worksheet.TabColor.Color);

        string sheetName = worksheet.Name;
        int lastRow      = worksheet.RowsUsed().Last().RowNumber(),
            lastColumn   = worksheet.ColumnsUsed().Last().ColumnNumber();
        var agentSheet   = await agent.AddNewSheet(sheetName, lastRow, lastColumn, colorHex.Equals(Toolbox.ColorToHexString(Color.FromArgb(0, 0, 0, 0))) ? null : colorHex);

        // Wait for it to charge
        var univerSheet = new UWorksheetInfo(worksheet, _config.MaxCellsReaded, agentSheet);
        var tasks = new List<Task>();

        // Get all data in Univer (Asynchronously (｡˃ ᵕ ˂ )👌 ❤️)
        if (options.RecData)
            tasks.Add(_reader.GetDataAsync(univerSheet, agent));
        if (options.RecStyles)
            tasks.Add(_reader.GetStylesAsync(univerSheet, agent));
        if (options.RecMerges)
            tasks.Add(_reader.GetMergesAsync(univerSheet, agent));
        if (options.RecFilters)
            tasks.Add(_reader.GetFilterAsync(univerSheet, agent));
        if (options.RecFreeze)
            tasks.Add(_reader.GetFreezeAsync(univerSheet, agent));
        if (options.RecComments)
            tasks.Add(_reader.GetCommentsAsync(univerSheet, agent, userManager));
        if (options.RecColumnsAndRows)
            tasks.Add(_reader.GetColumnsAndRowsAsync(univerSheet, agent));
        if (options.RecImages)
            tasks.Add(_reader.GetImagesAsync(univerSheet, agent));
        if (options.RecConditionalFormats)
            tasks.Add(_reader.GetConditionalFormatsAsync(univerSheet, agent));
        // Pending...
        if (options.RecAccesibility)
            tasks.Add(_reader.GetAccesibilityAsync(univerSheet, agent));

        // Last thing is hidden parts :D
        foreach (var col in worksheet.ColumnsUsed())
            if (col.IsHidden)
            {
                int colNumber = col.ColumnNumber();
                tasks.Add(agent.RowColumns.OnSheet(univerSheet.SheetInfo).HideColumns(colNumber - 1, 1));
            }

        foreach (var row in worksheet.RowsUsed())
            if (row.IsHidden)
            {
                int rowNumber = row.RowNumber();
                tasks.Add(agent.RowColumns.OnSheet(univerSheet.SheetInfo).HideRows(rowNumber - 1, 1));
            }

        if (worksheet.Visibility is XLWorksheetVisibility.Hidden)
            tasks.Add(agent.HideSheet(univerSheet.SheetInfo.id));

        await Task.WhenAll(tasks);
    }
}