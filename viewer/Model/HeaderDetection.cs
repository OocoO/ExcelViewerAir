namespace ExcelViewer.Model;

/// <summary>
/// 表头识别。
///
/// 首选信号是 <b>Excel 自己的冻结窗格配置</b>（<see cref="Sheet.FreezeRows"/>）：
/// 编辑者把哪几行冻住，哪几行就是表头块 —— 这是文件里的显式配置，
/// 任何配表结构都通用，不依赖项目内的标记约定。
///
/// 没有冻结配置时（csv / 没冻过的表）才退回到"探测"：
/// 前若干行里找 GDE_FIELD_NAMES 标记，找不到就把第一行当表头。
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

        // 1) 冻结窗格：冻结块整体就是表头，字段名行取标记行；没有标记就用冻结块最后一行
        if (sheet.FreezeRows > 0)
        {
            var frozen = Math.Min(sheet.FreezeRows, sheet.RowCount);
            var fieldRow = FindMarkerRow(sheet, frozen);
            var headerRow = fieldRow >= 0 ? fieldRow : frozen - 1;
            var notes = headerRow + 1 < sheet.RowCount ? headerRow + 1 : NoHeader;
            return new Result(headerRow, notes, fieldRow >= 0);
        }

        // 2) 没有冻结配置：找 GDE_FIELD_NAMES，它所在行就是字段名行
        var marker = FindMarkerRow(sheet, Math.Min(ScanLimit, sheet.RowCount));
        if (marker >= 0)
        {
            // 紧跟在字段名下面的那一行通常是中文说明
            var notes = marker + 1 < sheet.RowCount ? marker + 1 : NoHeader;
            return new Result(marker, notes, true);
        }

        // 3) 什么都没有：把第一行当表头（只要它不是空的）
        return RowHasAnyText(sheet, 0)
            ? new Result(0, 1 < sheet.RowCount ? 1 : NoHeader, false)
            : new Result(NoHeader, NoHeader, false);
    }

    /// <summary>在前 <paramref name="rows"/> 行里找标记行，找不到返回 -1。</summary>
    private static int FindMarkerRow(Sheet sheet, int rows)
    {
        for (var r = 0; r < rows && r < sheet.RowCount; r++)
        {
            if (RowContains(sheet, r, FieldNamesMarker))
            {
                return r;
            }
        }

        return -1;
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
