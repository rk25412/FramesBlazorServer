using FramesServer.Components.Pages.Home;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Radzen;
using System.Globalization;
using System.Reflection.Metadata.Ecma335;

namespace FramesServer.Services;

public class BillGeneratorService
{
    private ExcelWorksheet _ws;
    private readonly IReadOnlyList<GridItem> _itemData;
    private int _year, _month;
    private IReadOnlyDictionary<DateOnly, decimal> _paidData;
    private decimal _lastMonthAmt;

    public BillGeneratorService(
            ExcelWorksheet ws,
            int year, int month,
            List<GridItem> itemData,
            IReadOnlyDictionary<DateOnly, decimal> paid,
            decimal lastMonthAmt)
    {
        _ws = ws;
        _itemData = itemData;
        (_month, _year) = (month, year);
        _paidData = paid;
        _lastMonthAmt = lastMonthAmt;
    }

    public void GenerateBill()
    {
        // Set Month and Year at the top cell
        SetFileHeader();

        // Adding 1st Item
        SetItemData(_itemData[0], 6, 2, true);

        // Adding 2nd Item
        SetItemData(_itemData[1], 6, 6, true);

        Dictionary<string, (int, int)> leftCoords = [];

        // Adding all other data

    }

    private void SetItemData(GridItem item, int row, int col, bool all = false)
    {
        (int row, int col) s = (row, col);
        SetItemHeader(item.Name, s.row, s.col);
        s.row++;

        var listWithDate = item.Lines.Select((x, i) => (x, i))
            .ToDictionary(x => new DateTime(_year, _month, x.i + 1), x => x.x);

        foreach (var (key, val) in listWithDate)
        {
            if (!all && val is 0)
                continue;

            _ws.Cells[s.row, s.col].Value = $"{key:dd-MMM-yyyy}";
            _ws.Cells[s.row, s.col + 1].Value = val is 0 ? "-" : val;
            s.row++;
        }

        SetItemSummary(item.Rate ?? 0, row + 1, s.row - 1, col);
    }

    private void SetItemSummary(decimal rate, int dataStartRow, int dataEndRow, int summaryStartCol)
    {
        var totalCell = GetCell(dataEndRow + 1, summaryStartCol);
        totalCell.Value = "Total";

        var totalCellVal = GetCell(dataEndRow + 1, summaryStartCol + 1);
        var dataColChar = GetChar(summaryStartCol + 1);
        totalCellVal.Formula = $"SUM({dataColChar}{dataStartRow}:{dataColChar}{dataEndRow})";

        var rateCell = GetCell(dataEndRow + 2, summaryStartCol);
        rateCell.Value = "Rate";

        var rateCellVal = GetCell(dataEndRow + 2, summaryStartCol + 1);
        rateCellVal.Value = rate;

        var grandTotalCell = GetCells(dataEndRow + 3, summaryStartCol, dataEndRow + 3, summaryStartCol + 1);
        grandTotalCell.Merge = true;
        grandTotalCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
        grandTotalCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        grandTotalCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        grandTotalCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        grandTotalCell.Style.Font.Bold = true;
        grandTotalCell.Formula = $"{totalCellVal.Address}*{rateCellVal.Address}";
    }

    private ExcelRange GetCell(int row, int col) 
        => _ws.Cells[row, col];
    private ExcelRange GetCells(int fromRow, int fromCol, int toRow, int toCol)
        => _ws.Cells[fromRow, fromCol, toRow, toCol];

    private static char GetChar(int col) => (char)(64 + col);

    private void SetItemHeader(string itemName, int headerRow, int headerCol)
    {
        var cells = _ws.Cells[headerRow, headerCol, headerRow, headerCol + 1];
        cells.Value = itemName;
        cells.Merge = true;
        cells.Style.Font.Bold = true;
        cells.Style.Font.Size = 11;
        cells.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        cells.Style.Border.Left.Style = ExcelBorderStyle.Medium;
        cells.Style.Border.Right.Style = ExcelBorderStyle.Medium;
        cells.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
    }

    private void SetFileHeader()
    {
        var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(_month);
        _ws.Cells[1, 1].Value = $"{monthName} {_year}";
    }
}
