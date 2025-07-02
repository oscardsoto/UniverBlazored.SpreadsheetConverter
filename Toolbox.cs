using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using UniverBlazored.Spreadsheets.Data.ConditionFormat;

namespace UniverBlazored.SpreadsheetConverter;

internal static class Toolbox
{
    /// <summary>
    /// Returns the values of ARGB of the Hexadecimal value in XLColor, from ClosedXML
    /// </summary>
    /// <param name="hexColor"></param>
    public static XLColor ConvertHexToARGB(string hexColor)
    {
        // Sometimes... Univer decide to set a color to "null", which means the default color. This should fix it
        if (hexColor.Equals("null"))
            return XLColor.FromArgb(255, 0, 0, 0);

        // Sometimes, Univer decide to, instead of using hexadecimal, use "rgb(R, G, B)" pattern. This fragment solve that issue
        if (hexColor.StartsWith("rgb"))
        {
            var match_rgb = Regex.Match(hexColor, @"rgb\((\d+),(\d+),(\d+)\)");
            if (match_rgb.Success)
            {
                int r = int.Parse(match_rgb.Groups[1].Value);
                int g = int.Parse(match_rgb.Groups[2].Value);
                int b = int.Parse(match_rgb.Groups[3].Value);

                return XLColor.FromArgb(r, g, b);
            }

            var match_rgba = Regex.Match(hexColor, @"rgba\((\d+),(\d+),(\d+),(\d+)\)");
            if (match_rgba.Success)
            {
                int r = int.Parse(match_rgba.Groups[1].Value);
                int g = int.Parse(match_rgba.Groups[2].Value);
                int b = int.Parse(match_rgba.Groups[3].Value);
                int a = int.Parse(match_rgba.Groups[4].Value);

                return XLColor.FromArgb(a, r, g, b);
            }
        }

        if (hexColor.StartsWith("#"))
            hexColor = hexColor.Substring(1);  // Eliminar el carácter '#'

        int argb;
        if (hexColor.Length == 6)
        {
            // Si es un código de 6 dígitos, asumimos Alpha = 255
            hexColor = hexColor + "FF";             // Agregar Alpha al final
            argb = Convert.ToInt32(hexColor, 16);   // Convertir a entero
        }
        else if (hexColor.Length == 8)
            // Si es un código de 8 dígitos, convertir directamente
            argb = Convert.ToInt32(hexColor, 16);
        
        else
            throw new ArgumentException("El código hexadecimal debe ser de 6 o 8 dígitos.");

        int red     = (argb >> 24) & 0xFF;  // Los 8 primeros bits son para Alpha
        int green   = (argb >> 16) & 0xFF;  // Los siguientes 8 bits son para Red
        int blue    = (argb >> 8) & 0xFF;   // Los siguientes 8 bits son para Green
        int alpha   = argb & 0xFF;          // Los últimos 8 bits son para Blue

        // Devolver los valores en un array [Alpha, Red, Green, Blue]
        return XLColor.FromArgb(alpha, red, green, blue);
    }

    /// <summary>
    /// Return the maximum amount of rows that will be processed on each information thread
    /// </summary>
    /// <param name="maxCells">Maximum amount of cells that will be readed</param>
    /// <param name="maxColumns">Maximum amount of columns used in Univer</param>
    /// <param name="maxRows">Maximum amount of rows used in Univer</param>
    /// <returns></returns>
    public static int MaxRowsPerProcess(int maxCells, int maxRows, int maxColumns)
    {
        double rows = maxCells / maxColumns;
        return (rows > maxRows) ? maxRows : Convert.ToInt32(Math.Round(rows, 0, MidpointRounding.AwayFromZero));
    }

    /// <summary>
    /// Return the maximum amount of threads that will process the information in Univer
    /// </summary>
    /// <param name="maxCells">Maximum amount of cells that will be readed</param>
    /// <param name="maxColumns">Maximum amount of columns used in Univer</param>
    /// <param name="maxRows">Maximum amount of rows used in Univer</param>
    /// <returns></returns>
    public static int MaxProcessCount(int maxCells, int maxRows, int maxColumns)
    {
        if (maxRows * maxColumns < maxCells)
            maxCells = maxRows * maxColumns;

        if (maxCells < maxColumns)
            maxCells = maxColumns;

        double max = maxRows / MaxRowsPerProcess(maxCells, maxRows, maxColumns);
        if (max % 1 == 0)
            return Convert.ToInt32(max);
        else
            return Convert.ToInt32(Math.Round(max, 0, MidpointRounding.AwayFromZero)) + 1;
    }

    /// <summary>
    /// Return the Content Type for ClosedXML, depending of the content type for Univer
    /// </summary>
    /// <param name="type">Content type enum from Univer</param>
    /// <returns></returns>
    public static XLCFContentType ConvertToContentType(ECFValueType? type)
    {
        if (type == null)
            return XLCFContentType.Number;
            
        switch (type)
        {
            case ECFValueType.min:
                return XLCFContentType.Minimum;

            case ECFValueType.max:
                return XLCFContentType.Maximum;

            case ECFValueType.percent:
                return XLCFContentType.Percent;

            case ECFValueType.percentile:
                return XLCFContentType.Percentile;

            case ECFValueType.formula:
                return XLCFContentType.Formula;

            case ECFValueType.num:
            default:
                return XLCFContentType.Number;
        }
    }

    /// <summary>
    /// Returns the color as a Hexadecimal value
    /// </summary>
    /// <param name="color">Color to convert</param>
    /// <param name="onlyRGB">True to only return RGB values</param>
    /// <returns></returns>
    public static string ColorToHexString(Color color, bool onlyRGB = false)
    {
        return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2") + (onlyRGB ? "" : color.A.ToString("X2"));
    }

    /// <summary>
    /// Return a new random id
    /// </summary>
    /// <returns></returns>
    public static string GenerateRandomId() => Guid.NewGuid().ToString().Split('-')[0];

    /// <summary>
    /// Returns the base64 of the data alocated in the Memory stream
    /// </summary>
    /// <param name="ms">Memory Stream to get the base64</param>
    /// <returns></returns>
    public static string ConvertToBase64(MemoryStream ms)
    {
        string base64 = "";
        using (ms)
        {
            byte[] bytes = new byte[ms.Length];
            ms.Position = 0;
            ms.Read(bytes, 0, (int)ms.Length);
            base64 = Convert.ToBase64String(bytes);
        }
        return base64;
    }

    /// <summary>
    /// Return an array of values that are in between two natural numbers
    /// </summary>
    /// <param name="a"></param>
    /// <param name="b"></param>
    /// <returns></returns>
    public static int[] GetValuesInBetween(int a, int b)
    {
        if (a > b)
            throw new ArgumentException("a cant be higher than b");

        int[] values = new int[b - a + 1];
        for(int i = 0; i < values.Length; i++)
            values[i] = a + i;
        return values;
    }

    /// <summary>
    /// Return an array of values converted from (px) to (pt) for ClosedXML
    /// </summary>
    /// <param name="values">Values (in px)</param>
    /// <returns></returns>
    public static double[] ConvertToRowPoints(params double[] values)
    {
        //      RowHeight   -> Univer/24 (px), ClosedXML/15 (pt)
        double[] results = new double[values.Length];
        for (int i = 0; i < results.Length; i++)
            results[i] = values[i] == 24.0 ? 15.0 : values[i] * 15.0 / 24.0;
        return results;
    }

    /// <summary>
    /// Return an array of values converted from (pt) to (px) for Univer
    /// </summary>
    /// <param name="values">Values (in pt)</param>
    /// <returns></returns>
    public static double[] ConvertToRowPixels(params double[] values)
    {
        //      RowHeight   -> Univer/24 (px), ClosedXML/15 (pt)
        double[] results = new double[values.Length];
        for (int i = 0; i < results.Length; i++)
            results[i] = values[i] == 15.0 ? 24.0 : values[i] * 24.0 / 15.0;
        return results;
    }

    /// <summary>
    /// Return an array of values converted from (px) to (NoC) for ClosedXML
    /// </summary>
    /// <param name="values">Values (in px)</param>
    /// <returns></returns>
    public static double[] ConvertToColumnPoints(params double[] values)
    {
        //      ColumnWidth -> Univer/88 (px), ClosedXML/8.43 (NoC)
        double[] results = new double[values.Length];
        for (int i = 0; i < results.Length; i++)
            results[i] = values[i] == 88.0 ? 8.43 : values[i] * 8.43 / 88.0;
        return results;
    }

    /// <summary>
    /// Return an array of values converted from (NoC) to (px) for Univer
    /// </summary>
    /// <param name="values">Values (in NoC)</param>
    /// <returns></returns>
    public static double[] ConvertToColumnPixels(params double[] values)
    {
        //      ColumnWidth -> Univer/88 (px), ClosedXML/8.43 (NoC)
        double[] results = new double[values.Length];
        for (int i = 0; i < results.Length; i++)
            results[i] = values[i] == 8.43 ? 88.0 : values[i] * 88.0 / 8.43;
        return results;
    }

    /// <summary>
    /// Return the points in pixels value
    /// </summary>
    /// <param name="points"></param>
    /// <returns></returns>
    public static double ConvertToPixels(double points) => points * 1.33333333333;

    /// <summary>
    /// Return the pixels in points value
    /// </summary>
    /// <param name="pixels"></param>
    /// <returns></returns>
    public static double ConvertToPoints(double pixels) => pixels * 0.75;

    /// <summary>
    /// Returns true if the array is empty.
    /// </summary>
    /// <param name="values"></param>
    /// <returns></returns>
    public static bool IsEmpty(object[][] values)
    {
        int x = 0,
            y = 0;

        do 
        {
            var val = values[x][y];
            if (val != null)
                return false;

            if (y + 1 < values[x].Length)
            {
                y++;
                continue;
            }

            y = 0;
            x++;
        }
        while (x < values.Length);
        return true;
    }
}