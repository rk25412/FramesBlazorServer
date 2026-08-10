using System.Globalization;
using FramesServer.Models;
using Microsoft.AspNetCore.Components.Forms;
using OfficeOpenXml;

namespace FramesServer;

public static class Helper
{
    public static IEnumerable<DropdownDto> GetMonths()
    {
        var months = CultureInfo.CurrentCulture.DateTimeFormat.MonthNames.ToList();
        List<DropdownDto> monthsDdl = [];
        months.ForEach(x => { monthsDdl.Add(new DropdownDto { Value = months.IndexOf(x) + 1, Text = x }); });
        return monthsDdl;
    }

    public static IEnumerable<DropdownDto> GetYears()
    {
        var yr = DateTime.Now.Year;
        var years = Enumerable.Range(yr - 5, 10);
        return years.Select(x => new DropdownDto { Value = x, Text = x.ToString() });
    }

    public static async Task<string?> ReadFileAsString(IBrowserFile file)
    {
        if (!Path.GetExtension(file.Name).Equals(".txt", StringComparison.OrdinalIgnoreCase))
            return null;

        await using var stream = file.OpenReadStream(maxAllowedSize: 5 * 1024 * 1024);
        using var reader = new StreamReader(stream);

        var fileContents = await reader.ReadToEndAsync();
        return fileContents;
    }

    public static int GetNoOfDays(int year, int month)
    {
        var noOfDays = DateTime.DaysInMonth(year, month);
        return noOfDays;
    }

    public static string GetExcelTemplateName(int days) => days switch
    {
        28 => "28.xlsx",
        29 => "29.xlsx",
        30 => "30.xlsx",
        31 => "31.xlsx",
        _ => throw new InvalidDataException("Invalid no of days")
    };
}