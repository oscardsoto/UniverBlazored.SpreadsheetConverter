using Microsoft.Extensions.DependencyInjection;
using UniverBlazored.SpreadsheetConverter.Services;

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
        services.AddScoped<IUniverSpreadsheetConverter, UniverSpreadsheetConverter>();
    }    
}