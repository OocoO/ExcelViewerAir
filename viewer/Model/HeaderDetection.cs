namespace ExcelViewer.Model;

/// <summary>
/// 表头行识别。
///
/// 项目里的配表带 GDE 标记行，结构通常是：
///   row0  GDE_IGNORE      —— 标记行（不参与导出）
///   row1  GDE_FIELD_NAMES —— 字段名
///   row2  中文描述
///   row3  GDE_FIELD_TYPES —— 数据类型
/// 纯 xlsx/csv 查看器不应该假定结构，所以这里只做"探测"：
/// 在前若干行里找标记，找到就在界面上把那一行标成表头；
/// 找不到就回退成"第一行当表头"（普通表格的直觉行为）。
/// </summary>
public static class HeaderDetection
{
    public const int NoHeader = -1;

    private const string FieldNamesMarker = "GDE_FIELD_NAMES";
    private const int ScanLimit = 8;

    public sealed record Result(int HeaderRowIndex, int NotesRowIndex, bool FromGdeMarker);

    public static Result Detect(Sheet sheet)
    {
        if (sheet.RowCount == 0)
        {
            return new Result(NoHeader, NoHeader, false);
        }

        var rows = Math.Min(ScanLimit, sheet.RowCount);

        // 1) 找 GDE_FIELD_NAMES：它所在行就是字段名行
        for (var r = 0; r < rows; r++)
        {
            if (RowContains(sheet, r, FieldNamesMarker))
            {
                // 紧跟在字段名下面的那一行通常是中文说明
                var notes = r + 1 < sheet.RowCount ? r + 1 : NoHeader;
                return new Result(r, notes, true);
            }
        }

        // 2) 没有 GDE 标记：把第一行当表头（只要它不是空的）
        return RowHasAnyText(sheet, 0)
            ? new Result(0, 1 < sheet.RowCount ? 1 : NoHeader, false)
            : new Result(NoHeader, NoHeader, false);
    }

    private static bool RowContains(Sheet sheet, int row, string marker)
    {
        for (var c = 0; c < sheet.ColCount; c++)
        {
            if (sheet.Get(row, c).Contains(marker, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private static bool RowHasAnyText(Sheet sheet, int row)
    {
        for (var c = 0; c < sheet.ColCount; c++)
        {
            if (sheet.Get(row, c).Length != 0)
            {
                return true;
            }
        }

        return false;
    }
}
