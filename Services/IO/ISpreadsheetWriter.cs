using UniverBlazored.Generic.Services;
using UniverBlazored.Spreadsheets.Data.Workbook;
using UniverBlazored.Spreadsheets.Services;

namespace UniverBlazored.SpreadsheetConverter.Services.IO;

/// <summary>
/// Spreadhseet service for writing operations
/// </summary>
/// <typeparam name="TWorksheet">Worksheet object to apply all writing operations</typeparam>
public interface ISpreadsheetWriter<TWorksheet>
{
    /// <summary>
    /// Sets all data from Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="maxUsed">Maximum number of rows and columns used</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="maxCells">Maximum number of cells to process per task</param>
    /// <returns></returns>
    Task SetDataAsync(UniverSpreadsheetAgent agent, TWorksheet worksheet, URange maxUsed, int maxCells);

    /// <summary>
    /// Sets all merges from Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <returns></returns>
    Task SetMergesAsync(UniverSpreadsheetAgent agent, TWorksheet worksheet);

    /// <summary>
    /// Sets all freeze rows/columns from Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <returns></returns>
    Task SetFreezeAsync(UniverSpreadsheetAgent agent, TWorksheet worksheet);

    /// <summary>
    /// Sets all styles from Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="maxUsed">Maximum number of rows and columns used</param>
    /// <param name="maxCells">Maximum number of cells to process per task</param>
    /// <returns></returns>
    Task SetStylesAsync(UniverSpreadsheetAgent agent, TWorksheet worksheet, URange maxUsed, int maxCells);

    /// <summary>
    /// Sets all filters from Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <returns></returns>
    Task SetFiltersAsync(UniverSpreadsheetAgent agent, TWorksheet worksheet);

    /// <summary>
    /// Sets all conditional formats from Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <returns></returns>
    Task SetConditionalFormatsAsync(UniverSpreadsheetAgent agent, TWorksheet worksheet);

    /// <summary>
    /// Sets all images form Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <returns></returns>
    Task SetImagesAsync(UniverSpreadsheetAgent agent, TWorksheet worksheet);

    /// <summary>
    /// Sets all comment threads from Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="userManager">User Manager (for user context)</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <returns></returns>
    Task SetCommentsAsync(UniverSpreadsheetAgent agent, UniverUserManager userManager, TWorksheet worksheet);

    /// <summary>
    /// Sets all columns and rows settings from Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="maxUsed">Maximum number of rows and columns used</param>
    /// <returns></returns>
    Task SetColumnsAndRowsAsync(UniverSpreadsheetAgent agent, TWorksheet worksheet, URange maxUsed);

    /// <summary>
    /// Sets all accesibility settings and blocked cells from Univer into the worksheet
    /// </summary>
    /// <param name="agent">Univer Agent</param>
    /// <param name="worksheet">Worksheet object</param>
    /// <returns></returns>
    Task SetAccesibilityAsync(UniverSpreadsheetAgent agent, TWorksheet worksheet);
}