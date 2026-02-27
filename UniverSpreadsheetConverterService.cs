using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using UniverBlazored.SpreadsheetConverter.Services;
using UniverBlazored.SpreadsheetConverter.Services.IO;
using UniverBlazored.SpreadsheetConverter.Services.IO.Data;

namespace UniverBlazored.SpreadsheetConverter;

/// <summary>
/// Converter service to get/set the data on Univer
/// </summary>
public static class UniverSpreadsheetConverterService
{
    /// <summary>
    /// Adds the Univer's service for spreadsheets
    /// </summary>
    /// <param name="services"></param>
    /// <param name="configuration">Configuration object</param>
    public static void AddUniverSpreadsheetsConverter(this IServiceCollection services, Action<UniverSpreadsheetConverterConfig>? configuration = null)
    {
        services.Configure(configuration == null ? config => {} : configuration);
        services.AddScoped<ISpreadsheetReader<UWorksheetInfo>, USpreadsheetReader>();
        services.AddScoped<ISpreadsheetWriter<IXLWorksheet>, USpreadsheetWriter>();
        services.AddScoped<IUniverSpreadsheetConverter<XLWorkbook, IXLWorksheet>, UniverSpreadsheetConverter>();
    }    
}