using ClosedXML.Excel;
using UniverBlazored.Generic.Services;
using UniverBlazored.Spreadsheets.Data.Workbook;
using UniverBlazored.Spreadsheets.Services;

namespace UniverBlazored.SpreadsheetConverter.Services;

/// <summary>
/// Interface for managing spreadsheet data.
/// Provides methods to set/get data into/from an XLWorkbook using agent and userManager details, and optionally just populate data.
/// </summary>
public interface IUniverSpreadsheetConverter
{
    /// <summary>
    /// Setup spreadsheet data based on provided agent and userManager.
    /// </summary>
    /// <param name="agent">Agent for setting up the workbook.</param>
    /// <param name="userManager">User service for setting up the workbook.</param>
    /// <param name="options">Flags to select wich data will be setted</param>
    Task<XLWorkbook> SetInformationAsync(UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options);

    /// <summary>
    /// Get data into a workbook using provided agent and worksheet.
    /// </summary>
    /// <param name="excelWorkbook">Workbook containing source data for getting into an agent.</param>
    /// <param name="agent">Agent for storing fetched data.</param>
    /// <param name="userManager">User service for setting up the workbook.</param>
    /// <param name="options">Flags to select wich data will be setted</param>
    Task GetInformationInAgentAsync(XLWorkbook excelWorkbook, UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options);

    /// <summary>
    /// Gets information into an agent from a worksheet (IXLWorksheet). 
    /// </summary>
    /// <param name="worksheet">Worksheet containing source data for getting into an agent.</param>
    /// <param name="agent">Agent for storing fetched data.</param>
    /// <param name="userManager">User service for setting up the workbook.</param>
    /// <param name="options">Flags to select wich data will be setted</param>
    Task GetInformationInAgentAsync(IXLWorksheet worksheet, UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options);

    /// <summary>
    /// Sets spreadsheet data based on provided agent and userManager.
    /// </summary>
    /// <param name="workbook">Workbook to put the sheet data</param>
    /// <param name="unvrSheet">Data of the sheet (Univer) that will be readed</param>
    /// <param name="agent">Univer's agent</param>
    /// <param name="userManager">Univer's user manager</param>
    /// <param name="options">Flags to select wich data will be setted</param>
    Task SetInformationInSheetAsync(XLWorkbook workbook, USheetInfo unvrSheet, UniverSpreadsheetAgent agent, UniverUserManager userManager, SpreadsheetOptions options);
}