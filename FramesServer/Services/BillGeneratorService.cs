using FramesServer.Components.Pages.Home;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;

namespace FramesServer.Services;

public class BillGeneratorService
{
    private readonly ExcelWorksheet _ws;
    private readonly List<GridItem> _itemData;
    private readonly int _year, _month;
    private readonly IReadOnlyDictionary<DateOnly, decimal> _paidData;
    private readonly decimal _lastMonthAmt;
    private readonly Dictionary<string, (int row, int col)> _allCoords = [];

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

        bool left = true;
        int leftRow = 51, rightRow = 51;

        // Adding all other data
        foreach (var item in _itemData.Skip(2))
        {
            if (left)
                (leftRow, _) = SetItemData(item, leftRow + 3, 2);
            else
                (rightRow, _) = SetItemData(item, rightRow + 3, 6);

            left = !left;
        }

        var lastRow = leftRow >= rightRow ? (leftRow is 51 ? 54 : leftRow) : rightRow;

        // Adding all items summary
        var summaryEndRow = SetAllItemSummary(lastRow);

        // Adding paid data
        var paidDataEndRow = SetPaidData(lastRow);

        lastRow = summaryEndRow > paidDataEndRow ? summaryEndRow : paidDataEndRow;

        // Adding last month summary
        var leftDataRow = SetLastMonthSummary(lastRow, paidDataEndRow);

        // Adding Grand Total
        SetGrandTotal(lastRow, summaryEndRow, leftDataRow);
    }

    private int SetLastMonthSummary(int lastRow, int paidDataEndRow)
    {
        var lastMonth = new DateTime(DateTime.Now.Year, _month, 1).AddMonths(-1).ToString("MMM");
        int startRow = lastRow + 2;

        var monthNameCell = GetCell(startRow, 2);
        monthNameCell.Value = lastMonth;

        var monthNameCellVal = GetCell(startRow, 3);
        monthNameCellVal.Value = _lastMonthAmt;

        var paidAmtCell = GetCell(startRow + 1, 2);
        paidAmtCell.Value = "Paid";

        var paidAmtCellVal = GetCell(startRow + 1, 3);
        paidAmtCellVal.Formula = GetCell(paidDataEndRow, 7).Address;

        var totalCell = GetCell(startRow + 2, 2);
        totalCell.Value = "Left";
        totalCell.Style.Font.Bold = true;
        totalCell.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        totalCell.Style.Border.Left.Style = ExcelBorderStyle.Medium;
        totalCell.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

        var totalCellVal = GetCell(startRow + 2, 3);
        totalCellVal.Formula = $"{monthNameCellVal.Address}-{paidAmtCellVal.Address}";
        totalCellVal.Style.Font.Bold = true;
        totalCellVal.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        totalCellVal.Style.Border.Right.Style = ExcelBorderStyle.Medium;
        totalCellVal.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

        return startRow + 2;
    }

    private void SetGrandTotal(int lastRow, int itemSummaryDataEndRow, int leftDataLastRow)
    {
        var curMonthName = new DateTime(DateTime.Now.Year, _month, 1).ToString("MMM");
        int startRow = lastRow + 2;

        var curMonthCell = GetCell(startRow, 6);
        curMonthCell.Value = curMonthName;

        var curMonthAmtCell = GetCell(startRow, 7);
        curMonthAmtCell.Formula = GetCell(itemSummaryDataEndRow, 3).Address;

        var leftCell = GetCell(startRow + 1, 6);
        leftCell.Value = "Left";

        var leftAmtCell = GetCell(startRow + 1, 7);
        leftAmtCell.Formula = GetCell(leftDataLastRow, 3).Address;

        var grandTotalCell = GetCell(startRow + 2, 6);
        grandTotalCell.Value = $"{curMonthName} Total";
        grandTotalCell.Style.Font.Bold = true;
        grandTotalCell.Style.Border.Top.Style = ExcelBorderStyle.Thick;
        grandTotalCell.Style.Border.Left.Style = ExcelBorderStyle.Thick;
        grandTotalCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;

        var grandTotalAmtCell = GetCell(startRow + 2, 7);
        grandTotalAmtCell.Formula = $"{curMonthAmtCell.Address}+{leftAmtCell.Address}";
        grandTotalAmtCell.Style.Font.Bold = true;
        grandTotalAmtCell.Style.Border.Top.Style = ExcelBorderStyle.Thick;
        grandTotalAmtCell.Style.Border.Right.Style = ExcelBorderStyle.Thick;
        grandTotalAmtCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
    }

    private int SetPaidData(int lastRow)
    {
        SetItemHeader("Paid", lastRow + 2, 6);
        int startRow = lastRow + 3;
        var curRow = startRow;

        foreach (var (key, amt) in _paidData)
        {
            var dateCell = GetCell(curRow, 6);
            dateCell.Value = $"{key:dd-MMM-yyyy}";

            var amtCell = GetCell(curRow, 7);
            amtCell.Value = amt;

            curRow++;
        }

        var totalCell = GetCell(curRow, 6);
        totalCell.Value = "Total";
        totalCell.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        totalCell.Style.Border.Left.Style = ExcelBorderStyle.Medium;
        totalCell.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        totalCell.Style.Font.Bold = true;

        var totalCellVal = GetCell(curRow, 7);
        var valColChar = GetChar(7);
        totalCellVal.Formula = $"SUM({valColChar}{startRow}:{valColChar}{curRow - 1})";
        totalCellVal.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        totalCellVal.Style.Border.Right.Style = ExcelBorderStyle.Medium;
        totalCellVal.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        totalCellVal.Style.Font.Bold = true;

        return curRow;
    }

    private int SetAllItemSummary(int lastRow)
    {
        var startRow = lastRow + 3;
        var curRow = startRow;

        foreach (var (key, (row, col)) in _allCoords)
        {
            var nameCell = GetCell(curRow, 2);
            nameCell.Value = key;

            var valCell = GetCell(curRow, 3);
            valCell.Formula = GetCell(row, col).Address;

            curRow++;
        }

        var totalCell = GetCell(curRow, 2);
        totalCell.Value = "Total";
        totalCell.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        totalCell.Style.Border.Left.Style = ExcelBorderStyle.Medium;
        totalCell.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        totalCell.Style.Font.Bold = true;

        var totalCellVal = GetCell(curRow, 3);
        var valColChar = GetChar(3);
        totalCellVal.Formula = $"SUM({valColChar}{startRow}:{valColChar}{curRow - 1})";
        totalCellVal.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        totalCellVal.Style.Border.Right.Style = ExcelBorderStyle.Medium;
        totalCellVal.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
        totalCellVal.Style.Font.Bold = true;

        return curRow;
    }
    private (int, int) SetItemData(GridItem item, int row, int col, bool all = false)
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

            GetCell(s.row, s.col).Value = $"{key:dd-MMM-yyyy}";
            GetCell(s.row, s.col + 1).Value = val is 0 ? "-" : val;
            s.row++;
        }

        SetItemSummary(item.Rate ?? 0, row + 1, s.row - 1, col);
        _allCoords.Add(item.Name, (s.row + 2, s.col));

        return (s.row + 2, s.col);
    }
    private void SetItemSummary(decimal rate, int dataStartRow, int dataEndRow, int summaryStartCol)
    {
        var totalCell = GetCell(dataEndRow + 1, summaryStartCol);
        totalCell.Value = "Total";
        totalCell.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        totalCell.Style.Border.Left.Style = ExcelBorderStyle.Medium;

        var totalCellVal = GetCell(dataEndRow + 1, summaryStartCol + 1);
        var dataColChar = GetChar(summaryStartCol + 1);
        totalCellVal.Formula = $"SUM({dataColChar}{dataStartRow}:{dataColChar}{dataEndRow})";
        totalCellVal.Style.Border.Top.Style = ExcelBorderStyle.Medium;
        totalCellVal.Style.Border.Right.Style = ExcelBorderStyle.Medium;

        var rateCell = GetCell(dataEndRow + 2, summaryStartCol);
        rateCell.Value = "Rate";
        rateCell.Style.Border.Left.Style = ExcelBorderStyle.Medium;

        var rateCellVal = GetCell(dataEndRow + 2, summaryStartCol + 1);
        rateCellVal.Value = rate;
        rateCellVal.Style.Border.Right.Style = ExcelBorderStyle.Medium;

        var grandTotalCell = GetCells(dataEndRow + 3, summaryStartCol, dataEndRow + 3, summaryStartCol + 1);
        SetBorderToExcelRange(grandTotalCell, ExcelBorderStyle.Medium);
        grandTotalCell.Style.Font.Bold = true;
        grandTotalCell.Formula = $"{totalCellVal.Address}*{rateCellVal.Address}";
    }
    private void SetItemHeader(string itemName, int headerRow, int headerCol)
    {
        var cells = GetCells(headerRow, headerCol, headerRow, headerCol + 1);
        cells.Value = itemName;
        cells.Style.Font.Bold = true;
        cells.Style.Font.Size = 11;
        SetBorderToExcelRange(cells, ExcelBorderStyle.Medium);
    }
    private void SetFileHeader()
    {
        var monthName = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(_month);
        GetCell(1, 1).Value = $"{monthName} {_year}";
    }
    private ExcelRange GetCell(int row, int col)
        => _ws.Cells[row, col];
    private ExcelRange GetCells(int fromRow, int fromCol, int toRow, int toCol)
        => _ws.Cells[fromRow, fromCol, toRow, toCol];
    private static void SetBorderToExcelRange(ExcelRange excelRange, ExcelBorderStyle borderStyle)
    {
        excelRange.Merge = true;
        excelRange.Style.Border.Top.Style = borderStyle;
        excelRange.Style.Border.Left.Style = borderStyle;
        excelRange.Style.Border.Right.Style = borderStyle;
        excelRange.Style.Border.Bottom.Style = borderStyle;
    }
    private static char GetChar(int col) => (char)(64 + col);
}
