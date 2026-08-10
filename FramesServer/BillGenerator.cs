using System.Globalization;
using FramesServer.Components.Pages.Home;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FramesServer;

public static class BillGenerator
{
    public static async Task GenerateBill(
        int year,
        int month,
        ExcelWorksheet ws,
        List<GridItem> items,
        Dictionary<DateOnly, decimal> paymentInfo,
        decimal lastMonthAmt)
    {
        Dictionary<string, (int, int)> allFinalCoords = [];
        SetFileHeader(ws, year, month);

        // Adding First Item
        var ringFinalCoords = SetItemData(ws, year, month, items[0], 6, 2, true);
        allFinalCoords.Add(items[0].Name, ringFinalCoords);
    }

    private static (int, int) SetItemData(
        ExcelWorksheet ws, 
        int year, int month,
        GridItem item,
        int row, int col, bool all = false)
    {
        (int row, int col) s = (row, col);
        SetItemHeader(ws, s.row, s.col, item.Name, all);

        if (all) s.row++;

        var listWithDate = item.Lines.Select((x, i) => (x, i))
            .ToDictionary(x => new DateTime(year, month, x.i + 1), x => x.x);

        foreach (var (key, val) in listWithDate)
        {
            if (!all && val is 0)
                continue;

            ws.Cells[s.row, s.col].Value = $"{key:dd-MMM-yyyy}";
            ws.Cells[s.row++, s.col + 1].Value = val is 0 ? "-" : val;
        }

        SetItemSummary(ws, item.Rate ?? 0, row + 1, s.row, col);
        return (s.row, s.col);
    }

    private static void SetItemSummary(ExcelWorksheet ws, decimal rate, int dataStartRow, int dataEndRow, int summaryCol)
    {
        //var totalCell = ws.Cells[dataEndRow + 1, summaryCol];
        //totalCell.Value = "Total";

        //var totalCellVal = ws.Cells[dataEndRow + 1, summaryCol + 1];
        //totalCellVal.Value = $"SUM({col3Char}{startRow}:{col3Char}{exactRow - 1})"; ;

        //var rateCell = ws.Cells[dataEndRow + 2, summaryCol];
        //rateCell.Value = "Rate";

        //var rateCellVal = ws.Cells[dataEndRow + 2, summaryCol + 1];
        //rateCellVal.Value = 
    }

    private static char GetChar(int col) => (char)(64 + col);

    private static void SetItemHeader(ExcelWorksheet ws, int row, int col, string itemName, bool all = false)
    {
        var cells = ws.Cells[row, col, row, col + 1];
        cells.Value = itemName;
        cells.Merge = true;
        cells.Style.Font.Bold = true;
        cells.Style.Font.Size = 11;
        cells.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        cells.Style.Border.Left.Style = ExcelBorderStyle.Medium;
        cells.Style.Border.Right.Style = ExcelBorderStyle.Medium;
        cells.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
    }

    private static void SetFileHeader(ExcelWorksheet ws, int year, int month)
    {
        var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month);
        ws.Cells[1, 1].Value = $"{monthName} {year}";
    }
}