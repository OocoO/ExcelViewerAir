namespace ExcelViewer.Model;

/// <summary>一处单元格改动。</summary>
public sealed record CellChange(int Row, int Col, string OldValue, string NewValue);

/// <summary>一个工作表的比对结果。</summary>
public sealed class SheetDiff
{
    public required string SheetName { get; init; }

    public int OldRowCount { get; init; }

    public int NewRowCount { get; init; }

    public required List<CellChange> Changes { get; init; }

    /// <summary>只在旧文件里有的列（索引, 旧表头）。</summary>
    public required List<int> RemovedColumns { get; init; }

    /// <summary>只在新文件里有的列。</summary>
    public required List<int> AddedColumns { get; init; }

    public bool HasChanges => Changes.Count > 0 || RemovedColumns.Count > 0 || AddedColumns.Count > 0;

    public int ChangedCellCount => Changes.Count;
}

public sealed class DiffReport
{
    public required string OldPath { get; init; }

    public required string NewPath { get; init; }

    public required List<SheetDiff> Sheets { get; init; }

    /// <summary>旧文件里有、新文件里没有的工作表。</summary>
    public required List<string> RemovedSheets { get; init; }

    public required List<string> AddedSheets { get; init; }

    public long ElapsedMs { get; init; }

    public bool HasChanges =>
        Sheets.Any(s => s.HasChanges) || RemovedSheets.Count > 0 || AddedSheets.Count > 0;

    public int TotalChangedCells => Sheets.Sum(s => s.ChangedCellCount);

    /// <summary>可显示在窗口标题/横幅上的一句话摘要。</summary>
    public string Summary
    {
        get
        {
            if (!HasChanges)
            {
                return "两个文件内容一致，没有发现改动。";
            }

            var parts = new List<string>();
            var changedSheets = Sheets.Count(s => s.HasChanges);
            if (changedSheets > 0)
            {
                parts.Add($"{changedSheets} 个工作表有改动");
            }

            parts.Add($"共 {TotalChangedCells:N0} 处单元格");
            if (AddedSheets.Count > 0)
            {
                parts.Add($"新增工作表 {AddedSheets.Count} 个");
            }

            if (RemovedSheets.Count > 0)
            {
                parts.Add($"缺少工作表 {RemovedSheets.Count} 个");
            }

            return string.Join(" · ", parts);
        }
    }
}

/// <summary>
/// 两个工作簿的内容比对。
///
/// 用途：编辑走 temp 副本，保存后拿副本和原文件比一比，确认改了哪些格子再决定是否写回，
/// 避免"直接覆盖原文件"这种不可回退的操作。
/// 比对粒度是**单元格文本**（和表格里看到的一致），不做 OOXML 结构级 diff。
/// </summary>
public static class DiffService
{
    /// <summary>两个文件的工作表数量差太多时不做逐格比对，避免错位比较。</summary>
    private const int MaxReportedChangesPerSheet = 200_000;

    public static DiffReport Compare(
        string oldPath,
        string newPath,
        IProgress<string>? progress = null,
        CancellationToken ct = default)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();

        var oldNames = WorkbookLoader.ReadSheetNames(oldPath);
        var newNames = WorkbookLoader.ReadSheetNames(newPath);

        var removedSheets = oldNames.Where(n => !newNames.Contains(n, StringComparer.Ordinal)).ToList();
        var addedSheets = newNames.Where(n => !oldNames.Contains(n, StringComparer.Ordinal)).ToList();

        var pool = new StringPool();
        var results = new List<SheetDiff>();

        // 按位置配对（GDE 配表里改 sheet 名的情况很少，位置比名字更可靠）
        var pairCount = Math.Min(oldNames.Count, newNames.Count);
        for (var i = 0; i < pairCount; i++)
        {
            ct.ThrowIfCancellationRequested();

            var oldName = oldNames[i];
            var newName = newNames[i];
            if (removedSheets.Contains(oldName) || addedSheets.Contains(newName))
            {
                continue;
            }

            progress?.Report($"正在比对工作表「{oldName}」…");
            var oldSheet = WorkbookLoader.LoadSheet(oldPath, i, pool, ct);
            var newSheet = WorkbookLoader.LoadSheet(newPath, i, pool, ct);
            results.Add(CompareSheets(oldSheet, newSheet, oldName));
        }

        sw.Stop();
        return new DiffReport
        {
            OldPath = oldPath,
            NewPath = newPath,
            Sheets = results,
            RemovedSheets = removedSheets,
            AddedSheets = addedSheets,
            ElapsedMs = sw.ElapsedMilliseconds,
        };
    }

    private static SheetDiff CompareSheets(Sheet oldSheet, Sheet newSheet, string name)
    {
        var changes = new List<CellChange>();
        var removedCols = new List<int>();
        var addedCols = new List<int>();

        // 列数不同的部分先记下来：多出来的列整列都算"新增/缺失"
        var commonCols = Math.Min(oldSheet.ColCount, newSheet.ColCount);
        for (var c = commonCols; c < oldSheet.ColCount; c++)
        {
            removedCols.Add(c);
        }

        for (var c = commonCols; c < newSheet.ColCount; c++)
        {
            addedCols.Add(c);
        }

        var rows = Math.Max(oldSheet.RowCount, newSheet.RowCount);
        for (var r = 0; r < rows; r++)
        {
            if (changes.Count >= MaxReportedChangesPerSheet)
            {
                break;
            }

            for (var c = 0; c < commonCols; c++)
            {
                var oldValue = oldSheet.Get(r, c);
                var newValue = newSheet.Get(r, c);
                if (!string.Equals(oldValue, newValue, StringComparison.Ordinal))
                {
                    changes.Add(new CellChange(r, c, oldValue, newValue));
                }
            }
        }

        return new SheetDiff
        {
            SheetName = name,
            OldRowCount = oldSheet.RowCount,
            NewRowCount = newSheet.RowCount,
            Changes = changes,
            RemovedColumns = removedCols,
            AddedColumns = addedCols,
        };
    }

    /// <summary>
    /// 把某个表的改动拼成一张"看得见"的清单表：只列出有改动的格子，
    /// 每行 = [行号, 列标, 原值, 新值]，可以直接在查看器里浏览、复制、搜索。
    /// </summary>
    public static Sheet BuildChangeSheet(
        string sheetName,
        IReadOnlyList<CellChange> changes,
        StringPool pool,
        Func<int, string> columnNameAt)
    {
        var sheet = new Sheet(sheetName, pool);

        foreach (var ch in changes)
        {
            sheet.BeginRow();
            sheet.AddCell((ch.Row + 1).ToString());   // 行号与 Excel 一致，从 1 开始
            sheet.AddCell(columnNameAt(ch.Col));      // 列标
            sheet.AddCell(ch.OldValue);
            sheet.AddCell(ch.NewValue);
            sheet.EndRow();
        }

        sheet.Finish();
        return sheet;
    }

    /// <summary>改动清单的列标题。</summary>
    public static readonly string[] ChangeSheetHeaders = { "行号", "列", "原值", "新值" };
}
