using System.Text;

namespace ExcelViewer.Model;

/// <summary>一处单元格改动。<paramref name="Row"/> 是新文件里的行号（0-based）。</summary>
public sealed record CellChange(int Row, int Col, string OldValue, string NewValue);

/// <summary>一行在比对里的角色。</summary>
public enum RowChangeKind
{
    /// <summary>两边都有这一行，但格子里有内容变了。</summary>
    Modified,

    /// <summary>只有新文件里有这一行。</summary>
    Added,

    /// <summary>只有旧文件里有这一行。</summary>
    Removed,
}

/// <summary>
/// 行级改动。
/// <see cref="Modified"/> 的 <see cref="Cells"/> 是"改过的格子"；
/// <see cref="Added"/> / <see cref="Removed"/> 的 <see cref="RowText"/> 是整行内容（底部预览条里能看全文）。
/// </summary>
public sealed record RowChange(
    RowChangeKind Kind,
    int OldRow,
    int NewRow,
    string Key,
    string RowText,
    IReadOnlyList<CellChange> Cells)
{
    public static RowChange Modified(int oldRow, int newRow, string key, List<CellChange> cells) =>
        new(RowChangeKind.Modified, oldRow, newRow, key, string.Empty, cells);

    public static RowChange Added(int newRow, string key, string rowText) =>
        new(RowChangeKind.Added, -1, newRow, key, rowText, Array.Empty<CellChange>());

    public static RowChange Removed(int oldRow, string key, string rowText) =>
        new(RowChangeKind.Removed, oldRow, -1, key, rowText, Array.Empty<CellChange>());

    /// <summary>显示用的行号（0-based）：新增行没有旧行号，删除行没有新行号。</summary>
    public int DisplayRow => NewRow >= 0 ? NewRow : OldRow;

    public string KindText => Kind switch
    {
        RowChangeKind.Added => "新增行",
        RowChangeKind.Removed => "删除行",
        _ => "修改",
    };
}

/// <summary>一个工作表的比对结果。</summary>
public sealed class SheetDiff
{
    public required string SheetName { get; init; }

    public int OldRowCount { get; init; }

    public int NewRowCount { get; init; }

    /// <summary>按表格自上而下顺序排列的行级改动。</summary>
    public required List<RowChange> Rows { get; init; }

    /// <summary>所有"改了格子"的明细（不含整行新增/删除）。</summary>
    public required List<CellChange> Changes { get; init; }

    /// <summary>只在旧文件里有的列（索引）。</summary>
    public required List<int> RemovedColumns { get; init; }

    /// <summary>只在新文件里有的列。</summary>
    public required List<int> AddedColumns { get; init; }

    public int ChangedCellCount => Changes.Count;

    public int AddedRowCount => Rows.Count(r => r.Kind == RowChangeKind.Added);

    public int RemovedRowCount => Rows.Count(r => r.Kind == RowChangeKind.Removed);

    public bool HasChanges => Rows.Count > 0 || RemovedColumns.Count > 0 || AddedColumns.Count > 0;

    /// <summary>改动条数：改了的格子数 + 新增行数 + 删除行数。</summary>
    public int ChangeCount => ChangedCellCount + AddedRowCount + RemovedRowCount;

    /// <summary>标签页上用的极短概括。</summary>
    public string ShortSummary => ChangeCount == 0 ? "列有变化" : $"{ChangeCount:N0} 处";

    /// <summary>一句话概括，例如「12 处单元格 · 新增 3 行」。</summary>
    public string ChangeSummary
    {
        get
        {
            var parts = new List<string>();
            if (ChangedCellCount > 0)
            {
                parts.Add($"{ChangedCellCount:N0} 处单元格");
            }

            if (AddedRowCount > 0)
            {
                parts.Add($"新增 {AddedRowCount:N0} 行");
            }

            if (RemovedRowCount > 0)
            {
                parts.Add($"删除 {RemovedRowCount:N0} 行");
            }

            if (AddedColumns.Count > 0)
            {
                parts.Add($"新增 {AddedColumns.Count} 列");
            }

            if (RemovedColumns.Count > 0)
            {
                parts.Add($"缺少 {RemovedColumns.Count} 列");
            }

            return parts.Count == 0 ? "没有改动" : string.Join(" · ", parts);
        }
    }
}

public sealed class DiffReport
{
    public required string OldPath { get; init; }

    public required string NewPath { get; init; }

    public required List<SheetDiff> Sheets { get; init; }

    /// <summary>旧文件里有、新文件里没有的工作表。</summary>
    public required List<string> RemovedSheets { get; init; }

    /// <summary>新文件里有、旧文件里没有的工作表。</summary>
    public required List<string> AddedSheets { get; init; }

    public long ElapsedMs { get; init; }

    public bool HasChanges =>
        Sheets.Any(s => s.HasChanges) || RemovedSheets.Count > 0 || AddedSheets.Count > 0;

    public int TotalChangedCells => Sheets.Sum(s => s.ChangedCellCount);

    public int TotalAddedRows => Sheets.Sum(s => s.AddedRowCount);

    public int TotalRemovedRows => Sheets.Sum(s => s.RemovedRowCount);

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
            if (TotalAddedRows > 0)
            {
                parts.Add($"新增 {TotalAddedRows:N0} 行");
            }

            if (TotalRemovedRows > 0)
            {
                parts.Add($"删除 {TotalRemovedRows:N0} 行");
            }

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
/// 用途有两类：
///   * 编辑走 temp 副本，保存后拿副本和原文件比一比，确认改了哪些格子再决定是否写回；
///   * 当作 SVN 的外部 diff 工具（`--diff-svn`），看某个版本到工作副本改了些什么。
///
/// 比对粒度是**单元格文本**（和表格里看到的一致），不做 OOXML 结构级 diff。
///
/// 关键点是**先做行对齐再比格子**：早期版本按行列下标一一对应，只要中间插入或删除一行，
/// 后面所有行都会错位，200 行的表插一行就能报出 300 多处"改动"，完全没法看。
/// 现在分两步：
///   1. 挑一个"非空且互不重复"的列当行主键（配表的 ID 列），用主键把两边的行对上；
///      没有这样的列时退化成"整行文本"做锚，只认两边都唯一出现的那些行；
///   2. 对上的一对行逐格比较，对不上的区间才判成"整行新增 / 整行删除"。
/// 行序用 LIS 保证单调，所以插入一行只会报一行新增。
/// </summary>
public static class DiffService
{
    /// <summary>单个工作表最多记录多少处改动，防止两张完全不同的表把内存吃光。</summary>
    private const int MaxReportedChangesPerSheet = 200_000;

    /// <summary>探测主键列时采样的行数。</summary>
    private const int KeySampleRows = 400;

    /// <summary>主键列要求"非空值里至少 98% 互不重复"。</summary>
    private const double KeyUniquenessThreshold = 0.98;

    /// <summary>没有主键时，区间内"两边行数乘积"超过这个数就不做逐对配对（避免 O(n²)）。</summary>
    private const int MaxPairingWork = 4096;

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

    /// <summary>比对两张已经加载好的表。internal 是为了让 --diff-selftest 能直接构造用例。</summary>
    internal static SheetDiff CompareSheets(Sheet oldSheet, Sheet newSheet, string name)
    {
        var commonCols = Math.Min(oldSheet.ColCount, newSheet.ColCount);

        // 列数不同的部分先记下来：多出来的列整列都算"新增/缺失"
        var removedCols = new List<int>();
        for (var c = commonCols; c < oldSheet.ColCount; c++)
        {
            removedCols.Add(c);
        }

        var addedCols = new List<int>();
        for (var c = commonCols; c < newSheet.ColCount; c++)
        {
            addedCols.Add(c);
        }

        var keyCol = DetectKeyColumn(oldSheet, newSheet, commonCols);
        var rows = AlignRows(oldSheet, newSheet, commonCols, keyCol);

        return new SheetDiff
        {
            SheetName = name,
            OldRowCount = oldSheet.RowCount,
            NewRowCount = newSheet.RowCount,
            Rows = rows,
            Changes = rows
                .Where(r => r.Kind == RowChangeKind.Modified)
                .SelectMany(r => r.Cells)
                .Take(MaxReportedChangesPerSheet)
                .ToList(),
            RemovedColumns = removedCols,
            AddedColumns = addedCols,
        };
    }

    // ================= 行对齐 =================

    private static List<RowChange> AlignRows(Sheet oldSheet, Sheet newSheet, int commonCols, int keyCol)
    {
        // 整行全空的行不参与比对：它们在两张表里都没内容，硬比只会报出成百上千条
        // "删一行 / 加一行"的噪声（配表中间常常夹着大片空行）。
        var oldRows = NonEmptyRows(oldSheet);
        var newRows = NonEmptyRows(newSheet);

        var oldAnchor = new string[oldRows.Count];
        var newAnchor = new string[newRows.Count];

        for (var i = 0; i < oldRows.Count; i++)
        {
            oldAnchor[i] = Anchor(oldSheet, oldRows[i], commonCols, keyCol);
        }

        for (var j = 0; j < newRows.Count; j++)
        {
            newAnchor[j] = Anchor(newSheet, newRows[j], commonCols, keyCol);
        }

        var result = new List<RowChange>();
        var anchors = MatchUniqueAnchors(oldAnchor, newAnchor);

        var oldPos = 0;
        var newPos = 0;
        foreach (var (anchorOld, anchorNew) in anchors)
        {
            EmitGap(
                oldSheet, newSheet, oldRows, newRows,
                oldPos, anchorOld, newPos, anchorNew,
                commonCols, keyCol, result);

            var oldRow = oldRows[anchorOld];
            var newRow = newRows[anchorNew];
            var key = keyCol >= 0 ? oldSheet.Get(oldRow, keyCol) : string.Empty;
            var cells = CompareRowCells(oldSheet, newSheet, oldRow, newRow, commonCols);
            if (cells.Count > 0)
            {
                result.Add(RowChange.Modified(oldRow, newRow, key, cells));
            }

            oldPos = anchorOld + 1;
            newPos = anchorNew + 1;
        }

        EmitGap(
            oldSheet, newSheet, oldRows, newRows,
            oldPos, oldRows.Count, newPos, newRows.Count,
            commonCols, keyCol, result);

        return result;
    }

    /// <summary>有内容的行号（0-based），保持原顺序；空行直接跳过。</summary>
    private static List<int> NonEmptyRows(Sheet sheet)
    {
        var rows = new List<int>(sheet.RowCount);
        for (var r = 0; r < sheet.RowCount; r++)
        {
            if (!sheet.IsRowEmpty(r))
            {
                rows.Add(r);
            }
        }

        return rows;
    }

    /// <summary>两个锚之间对不上的区间：能配的配成"修改"，剩下的分别算新增 / 删除。</summary>
    private static void EmitGap(
        Sheet oldSheet,
        Sheet newSheet,
        IReadOnlyList<int> oldRows,
        IReadOnlyList<int> newRows,
        int oldStart,
        int oldEnd,
        int newStart,
        int newEnd,
        int commonCols,
        int keyCol,
        List<RowChange> result)
    {
        var oldLen = Math.Max(0, oldEnd - oldStart);
        var newLen = Math.Max(0, newEnd - newStart);
        if (oldLen == 0 && newLen == 0)
        {
            return;
        }

        // 一段"内容被改过"的行会掉进同一个区间（主键变了，或者本来就没有主键）。
        // 按"相等的格子数"贪心配对，就能把"删一行 + 加一行"还原成"这一行哪几格改了"。
        // 有主键时阈值更严：主键不同说明本来就不是同一行，不能因为碰巧有一格相同就硬配。
        if (oldLen > 0 && newLen > 0)
        {
            var threshold = keyCol < 0 ? 1 : Math.Max(1, commonCols / 2);
            var pairs = (long)oldLen * newLen <= MaxPairingWork
                ? GreedyPairRows(oldSheet, newSheet, oldRows, oldStart, oldLen, newRows, newStart, newLen, commonCols, threshold)
                : PositionalPairs(oldSheet, newSheet, oldRows, oldStart, oldLen, newRows, newStart, newLen, commonCols, threshold);

            if (pairs.Count > 0)
            {
                var matchedOld = new HashSet<int>();
                var oldByNew = new Dictionary<int, int>();
                foreach (var (o, n) in pairs)
                {
                    matchedOld.Add(o);
                    oldByNew[n] = o;
                }

                for (var k = 0; k < newLen; k++)
                {
                    var newRow = newRows[newStart + k];
                    if (!oldByNew.TryGetValue(newRow, out var oldRow))
                    {
                        result.Add(AddedRow(newSheet, newRow, commonCols, keyCol));
                        continue;
                    }

                    var cells = CompareRowCells(oldSheet, newSheet, oldRow, newRow, commonCols);
                    if (cells.Count > 0)
                    {
                        var key = keyCol >= 0 ? oldSheet.Get(oldRow, keyCol) : string.Empty;
                        result.Add(RowChange.Modified(oldRow, newRow, key, cells));
                    }
                }

                for (var k = 0; k < oldLen; k++)
                {
                    var oldRow = oldRows[oldStart + k];
                    if (!matchedOld.Contains(oldRow))
                    {
                        result.Add(RemovedRow(oldSheet, oldRow, commonCols, keyCol));
                    }
                }

                return;
            }
        }

        for (var k = 0; k < oldLen; k++)
        {
            result.Add(RemovedRow(oldSheet, oldRows[oldStart + k], commonCols, keyCol));
        }

        for (var k = 0; k < newLen; k++)
        {
            result.Add(AddedRow(newSheet, newRows[newStart + k], commonCols, keyCol));
        }
    }

    /// <summary>
    /// 区间太大时不做 O(n²) 的贪心配对，改成"第 k 行对第 k 行"，但每一对仍然要过相似度门槛。
    /// 配不上的一对会退化成"删除 + 新增"，所以"删 N 行 + 加 N 行"不会被硬说成 N 处修改。
    /// 注意：两个文件一模一样时，重复行（不是唯一锚）会落到这里——按位置配上、逐格比较无差异，
    /// 结果就是 0 改动，不会报出一堆假增删。
    /// </summary>
    private static List<(int Old, int New)> PositionalPairs(
        Sheet oldSheet,
        Sheet newSheet,
        IReadOnlyList<int> oldRows,
        int oldStart,
        int oldLen,
        IReadOnlyList<int> newRows,
        int newStart,
        int newLen,
        int commonCols,
        int threshold)
    {
        var pairs = new List<(int Old, int New)>();
        var count = Math.Min(oldLen, newLen);
        for (var k = 0; k < count; k++)
        {
            var oldRow = oldRows[oldStart + k];
            var newRow = newRows[newStart + k];
            if (RowSimilarity(oldSheet, newSheet, oldRow, newRow, commonCols) >= threshold)
            {
                pairs.Add((oldRow, newRow));
            }
        }

        return pairs;
    }

    /// <summary>两行有多像：相同的非空格子数。</summary>
    private static int RowSimilarity(Sheet oldSheet, Sheet newSheet, int oldRow, int newRow, int commonCols)
    {
        var score = 0;
        for (var c = 0; c < commonCols; c++)
        {
            var value = oldSheet.Get(oldRow, c);
            if (value.Length != 0 && string.Equals(value, newSheet.Get(newRow, c), StringComparison.Ordinal))
            {
                score++;
            }
        }

        return score;
    }

    /// <summary>
    /// 在一个区间里把两边最像的行配起来：相似度 = 相同的非空格子数，从高到低贪心取，
    /// 分数相同时优先"位置差得少"的，保证结果稳定、可复现。返回绝对行号对，按旧行号升序。
    /// </summary>
    private static List<(int Old, int New)> GreedyPairRows(
        Sheet oldSheet,
        Sheet newSheet,
        IReadOnlyList<int> oldRows,
        int oldStart,
        int oldLen,
        IReadOnlyList<int> newRows,
        int newStart,
        int newLen,
        int commonCols,
        int threshold)
    {
        var candidates = new List<(int Score, int Old, int New)>();
        for (var o = 0; o < oldLen; o++)
        {
            var oldRow = oldRows[oldStart + o];
            for (var n = 0; n < newLen; n++)
            {
                var newRow = newRows[newStart + n];
                var score = RowSimilarity(oldSheet, newSheet, oldRow, newRow, commonCols);
                if (score >= threshold)
                {
                    candidates.Add((score, oldRow, newRow));
                }
            }
        }

        if (candidates.Count == 0)
        {
            return new List<(int Old, int New)>();
        }

        candidates.Sort((x, y) =>
        {
            var byScore = y.Score.CompareTo(x.Score);
            return byScore != 0 ? byScore : Math.Abs(x.Old - x.New).CompareTo(Math.Abs(y.Old - y.New));
        });

        var usedOld = new HashSet<int>();
        var usedNew = new HashSet<int>();
        var pairs = new List<(int Old, int New)>();
        foreach (var (_, o, n) in candidates)
        {
            if (usedOld.Contains(o) || usedNew.Contains(n))
            {
                continue;
            }

            usedOld.Add(o);
            usedNew.Add(n);
            pairs.Add((o, n));
        }

        pairs.Sort((x, y) => x.Old.CompareTo(y.Old));
        return pairs;
    }

    /// <summary>行锚：有主键用主键，没有（或是空值）就用整行文本，前缀区分两者防止撞车。</summary>
    private static string Anchor(Sheet sheet, int row, int commonCols, int keyCol)
    {
        if (keyCol >= 0)
        {
            var key = sheet.Get(row, keyCol);
            if (key.Length > 0)
            {
                return "K\u0001" + key;
            }
        }

        return "C\u0001" + RowText(sheet, row, commonCols);
    }

    /// <summary>
    /// 找行主键列：从左往右挑第一个"基本不空、而且基本不重复"的列。
    /// 配表的 ID 列满足这个条件；说明列/标记列（GDE_IGNORE 这种）会因为大量重复被跳过。
    /// </summary>
    private static int DetectKeyColumn(Sheet oldSheet, Sheet newSheet, int commonCols)
    {
        for (var c = 0; c < commonCols; c++)
        {
            if (ColumnUniqueness(oldSheet, c) < KeyUniquenessThreshold)
            {
                continue;
            }

            if (ColumnUniqueness(newSheet, c) < KeyUniquenessThreshold)
            {
                continue;
            }

            return c;
        }

        return -1;
    }

    private static double ColumnUniqueness(Sheet sheet, int col)
    {
        var rows = Math.Min(sheet.RowCount, KeySampleRows);
        if (rows < 3)
        {
            return 0;
        }

        var seen = new HashSet<string>(StringComparer.Ordinal);
        var nonEmpty = 0;
        for (var r = 0; r < rows; r++)
        {
            var v = sheet.Get(r, col);
            if (v.Length == 0)
            {
                continue;
            }

            nonEmpty++;
            seen.Add(v);
        }

        // 大量空值的列不适合当主键
        if (nonEmpty < rows * 0.8)
        {
            return 0;
        }

        return (double)seen.Count / nonEmpty;
    }

    /// <summary>
    /// 取"两边都恰好出现一次"的行当骨架，再对新文件侧的位置求最长上升子序列。
    /// 这样既不会把插入的行错当成"后面的行都改了"，也不会把换过位置的行乱配对。
    /// </summary>
    private static List<(int Old, int New)> MatchUniqueAnchors(string[] oldAnchor, string[] newAnchor)
    {
        var newPos = new Dictionary<string, int>(StringComparer.Ordinal);
        var newCount = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var j = 0; j < newAnchor.Length; j++)
        {
            var k = newAnchor[j];
            if (k.Length == 0)
            {
                continue;
            }

            newCount[k] = newCount.TryGetValue(k, out var c) ? c + 1 : 1;
            newPos[k] = j;
        }

        var oldCount = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < oldAnchor.Length; i++)
        {
            var k = oldAnchor[i];
            if (k.Length == 0)
            {
                continue;
            }

            oldCount[k] = oldCount.TryGetValue(k, out var c) ? c + 1 : 1;
        }

        var candidates = new List<(int Old, int New)>();
        for (var i = 0; i < oldAnchor.Length; i++)
        {
            var k = oldAnchor[i];
            if (k.Length == 0 || oldCount[k] != 1)
            {
                continue;
            }

            if (!newCount.TryGetValue(k, out var nc) || nc != 1)
            {
                continue;
            }

            candidates.Add((i, newPos[k]));
        }

        return LongestIncreasingSubsequence(candidates);
    }

    /// <summary>对 (Old 升序的) 候选对按 New 求最长上升子序列，返回保序的配对表。</summary>
    private static List<(int Old, int New)> LongestIncreasingSubsequence(List<(int Old, int New)> candidates)
    {
        if (candidates.Count == 0)
        {
            return candidates;
        }

        var tails = new int[candidates.Count];
        var prev = new int[candidates.Count];
        var len = 0;

        for (var i = 0; i < candidates.Count; i++)
        {
            var v = candidates[i].New;
            var lo = 0;
            var hi = len;
            while (lo < hi)
            {
                var mid = (lo + hi) / 2;
                if (candidates[tails[mid]].New < v)
                {
                    lo = mid + 1;
                }
                else
                {
                    hi = mid;
                }
            }

            prev[i] = lo > 0 ? tails[lo - 1] : -1;
            tails[lo] = i;
            if (lo == len)
            {
                len++;
            }
        }

        var result = new List<(int Old, int New)>(len);
        var idx = tails[len - 1];
        while (idx >= 0)
        {
            result.Add(candidates[idx]);
            idx = prev[idx];
        }

        result.Reverse();
        return result;
    }

    // ================= 行内逐格比较 / 行文本 =================

    private static List<CellChange> CompareRowCells(
        Sheet oldSheet,
        Sheet newSheet,
        int oldRow,
        int newRow,
        int commonCols)
    {
        var cells = new List<CellChange>();
        for (var c = 0; c < commonCols; c++)
        {
            var oldValue = oldSheet.Get(oldRow, c);
            var newValue = newSheet.Get(newRow, c);
            if (!string.Equals(oldValue, newValue, StringComparison.Ordinal))
            {
                cells.Add(new CellChange(newRow, c, oldValue, newValue));
            }
        }

        return cells;
    }

    private static RowChange AddedRow(Sheet sheet, int row, int commonCols, int keyCol) =>
        RowChange.Added(row, keyCol >= 0 ? sheet.Get(row, keyCol) : string.Empty, RowText(sheet, row, commonCols));

    private static RowChange RemovedRow(Sheet sheet, int row, int commonCols, int keyCol) =>
        RowChange.Removed(row, keyCol >= 0 ? sheet.Get(row, keyCol) : string.Empty, RowText(sheet, row, commonCols));

    /// <summary>整行拼成一段文本（只取两表共有的列），用于整行新增/删除的展示。</summary>
    private static string RowText(Sheet sheet, int row, int cols)
    {
        var sb = new StringBuilder();
        for (var c = 0; c < cols; c++)
        {
            var v = sheet.Get(row, c);
            if (v.Length == 0)
            {
                continue;
            }

            if (sb.Length > 0)
            {
                sb.Append("  |  ");
            }

            sb.Append(v);
        }

        return sb.ToString();
    }

    // ================= 改动清单 =================

    /// <summary>
    /// 把某个表的改动拼成一张"看得见"的清单表：
    /// 改格子是 [修改, 行号, 列, 原值, 新值]，整行新增/删除则是 [新增行/删除行, 行号, 整行, 原值, 新值]。
    /// 清单本身也是一张普通的表，可以直接在查看器里浏览、复制、搜索。
    /// </summary>
    public static Sheet BuildChangeSheet(
        SheetDiff diff,
        StringPool pool,
        Func<int, string> columnNameAt)
    {
        var sheet = new Sheet(diff.SheetName, pool);

        foreach (var row in diff.Rows)
        {
            if (row.Kind == RowChangeKind.Modified)
            {
                foreach (var ch in row.Cells)
                {
                    sheet.BeginRow();
                    sheet.AddCell(row.KindText);
                    sheet.AddCell((ch.Row + 1).ToString());   // 行号与 Excel 一致，从 1 开始
                    sheet.AddCell(columnNameAt(ch.Col));      // 列标
                    sheet.AddCell(ch.OldValue);
                    sheet.AddCell(ch.NewValue);
                    sheet.EndRow();
                }

                continue;
            }

            var isAdded = row.Kind == RowChangeKind.Added;
            sheet.BeginRow();
            sheet.AddCell(row.KindText);
            sheet.AddCell((row.DisplayRow + 1).ToString());
            sheet.AddCell("整行");
            sheet.AddCell(isAdded ? string.Empty : row.RowText);
            sheet.AddCell(isAdded ? row.RowText : string.Empty);
            sheet.EndRow();
        }

        sheet.Finish();
        return sheet;
    }

    /// <summary>改动清单的列标题。</summary>
    public static readonly string[] ChangeSheetHeaders = { "类型", "行号", "列", "原值", "新值" };
}
