using UniverBlazored.Spreadsheets.Data.Workbook;

namespace UniverBlazored.SpreadsheetConverter.Services.IO.Data;

/// <summary>
/// All worksheet information to be stored in Univer after reading a spreadsheet
/// </summary>
public interface IWorksheetInfo<TWorksheet>
{
    /// <summary>
    /// Sheet's main information
    /// </summary>
    /// <value></value>
    public USheetInfo SheetInfo { get; set; }

    /// <summary>
    /// Data for the ranges in the worksheet, including values and formulas (If any)
    /// </summary>
    /// <returns></returns>
    public Dictionary<URange, URangeData> RangesData { get; set; }

    /// <summary>
    /// Styles to apply to the ranges in the worksheet
    /// </summary>
    /// <returns></returns>
    public List<URangeStyles> RangeStyles { get; set; }

    /// <summary>
    /// All the merged ranges in the worksheet
    /// </summary>
    public List<URange> MergedRanges { get; set; }

    /// <summary>
    /// Range for the filter in the worksheet (If any)
    /// </summary>
    /// <value></value>
    public URange Filter { get; set; }

    /// <summary>
    /// Range for the freeze in the worksheet (If any)
    /// </summary>
    /// <value></value>
    public URange FreezeReference { get; set; }

    /// <summary>
    /// Comments information in the worksheet
    /// </summary>
    /// <returns></returns>
    public List<UCommentReference> Comments { get; set; }

    /// <summary>
    /// Images information in the worksheet
    /// </summary>
    /// <returns></returns>
    public List<UImageInfo> Images { get; set; }

    /// <summary>
    /// Conditional formats in the worksheet, including the ranges they apply to, the type of condition, and the style to be applied when the condition is met.
    /// </summary>
    /// <returns></returns>
    public List<UCondFormatData> ConditionalFormats { get; set; }

    /// <summary>
    /// Column widths in the worksheet
    /// </summary>
    /// <returns></returns>
    public List<(int column, double width)> ColumnWidths { get; set; }

    /// <summary>
    /// Rows heights in the worksheet
    /// </summary>
    /// <returns></returns>
    public List<(int row, double height)> RowHeights { get; set; }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="maxUsed"></param>
    /// <param name="maxCellsReaded"></param>    
    /// <returns></returns>
    public Dictionary<URange, URangeData> GetRangesData(TWorksheet worksheet, URange maxUsed, int maxCellsReaded);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="maxUsed"></param>
    /// <param name="maxCellsReaded"></param>
    /// <returns></returns>
    public List<URangeStyles> GetRangeStyles(TWorksheet worksheet, URange maxUsed, int maxCellsReaded);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <returns></returns>
    public List<URange> GetMergedRanges(TWorksheet worksheet);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <returns></returns>
    public URange GetFilter(TWorksheet worksheet);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <returns></returns>
    public URange GetFreeze(TWorksheet worksheet);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <returns></returns>
    public List<UCommentReference> GetCommentsData(TWorksheet worksheet);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <returns></returns>
    public List<UImageInfo> GetImagesData(TWorksheet worksheet);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="worksheet"></param>
    /// <returns></returns>
    public List<UCondFormatData> GetConditionalFormatsData(TWorksheet worksheet);

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
    public List<(int column, double width)> GetColumnWidths(TWorksheet worksheet);

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
    public List<(int row, double height)> GetRowHeights(TWorksheet worksheet);
}