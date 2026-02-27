using UniverBlazored.Generic.Services;
using UniverBlazored.Spreadsheets.Data.Workbook;
using UniverBlazored.Spreadsheets.Services;

namespace UniverBlazored.SpreadsheetConverter.Services.IO;

/// <summary>
/// Spreadsheet service for reading operations
/// </summary>
/// <typeparam name="TWorksheet">Worksheet object to apply all reading operations</typeparam>
public interface ISpreadsheetReader<TWorksheet>
{
    /// <summary>
    /// Gets all worksheet data in Univer
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <returns></returns>
    Task GetDataAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent);

    /// <summary>
    /// Gets all worksheet styles in Univer
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <returns></returns>
    Task GetStylesAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent);

    /// <summary>
    /// Gets all worksheet merges in Univer
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <returns></returns>
    Task GetMergesAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent);

    /// <summary>
    /// Gets all worksheet filters in Univer 
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <returns></returns>
    Task GetFilterAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent);

    /// <summary>
    /// Gets all frozen rows and columns from the worksheet in Univer
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <returns></returns>
    Task GetFreezeAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent);

    /// <summary>
    /// Gets all worksheet comments in Univer
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <param name="userManager">User Manager (for user context)</param>
    /// <returns></returns>
    Task GetCommentsAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent, UniverUserManager userManager);

    /// <summary>
    /// Gets all worksheet images in Univer
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <returns></returns>
    Task GetImagesAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent);

    /// <summary>
    /// Gets all conditional formats from the worksheet in Univer
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <returns></returns>
    Task GetConditionalFormatsAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent);

    /// <summary>
    /// Gets all rows and columns configuration from the worksheet in Univer
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <returns></returns>
    Task GetColumnsAndRowsAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent);

    /// <summary>
    /// Gets all accesibility configurations and blocked cells from the worksheet in Univer
    /// </summary>
    /// <param name="worksheet">Worksheet object</param>
    /// <param name="agent">Univer Agent</param>
    /// <returns></returns>
    Task GetAccesibilityAsync(TWorksheet worksheet, UniverSpreadsheetAgent agent);
}