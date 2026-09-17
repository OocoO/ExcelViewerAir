using System.Diagnostics;
using System.Text.RegularExpressions;
using ExcelViewer.Model;

namespace ExcelViewer.Search;

public enum SearchMode
{
    /// <summary>包含匹配（默认），对中文和配表 id 都最直观。</summary>
    Contains,
    /// <summary>正则匹配，给高级用法留口子。</summary>
    Regex,
}

public sealed class SearchOptions
{
    public string Query { get; init; } = string.Empty;

    public SearchMode Mode { get; init; } = SearchMode.Contains;

    public bool CaseSensitive { get; init; }

    /// <summary>限定列范围（含）。null 表示不限。</summary>
    public int? ColumnFrom { get; init; }

    public int? ColumnTo { get; init; }

    public bool IsEmpty => Query.Length == 0;
}

/// <summary>
/// 全表搜索。
///
/// 结果只保留「每行是否有命中」的一个 byte 数组 + 命中总数：
///   * 命中的列在渲染时按可视行现场重扫，代价极低（可视区最多几十行）；
///   * 避免为一次搜索建一个几十万条目的命中列表（那才是真正吃内存的地方）。
/// 同一查询重复执行直接走缓存，上下条跳转、窗口缩放都不会重扫。
/// </summary>
public sealed class SearchResult
{
    public static readonly SearchResult Empty = new();

    public string Query { get; init; } = string.Empty;

    public SearchOptions? Options { get; init; }

    /// <summary>索引 = 行号，非 0 表示该行有命中。</summary>
    public byte[] RowHits { get; init; } = Array.Empty<byte>();

    /// <summary>命中单元格总数（仅在未被上限截断时精确）。</summary>
    public long CellHitCount { get; init; }

    public int RowHitCount { get; init; }

    public bool Truncated { get; init; }

    public bool Cancelled { get; init; }

    public long ElapsedMs { get; init; }

    public Sheet? Sheet { get; init; }

    public bool IsEmpty => Query.Length == 0;

    public bool HasHits => RowHitCount > 0;

    public bool RowIsHit(int row) => (uint)row < (uint)RowHits.Length && RowHits[row] != 0;

    /// <summary>找到 row 之后（不含）的第一个命中行，找不到返回 -1。</summary>
    public int NextHitRow(int row)
    {
        for (var r = row + 1; r < RowHits.Length; r++)
        {
            if (RowHits[r] != 0)
            {
                return r;
            }
        }

        return -1;
    }

    /// <summary>找到 row 之前（不含）的最后一个命中行，找不到返回 -1。</summary>
    public int PrevHitRow(int row)
    {
        for (var r = Math.Min(row - 1, RowHits.Length - 1); r >= 0; r--)
        {
            if (RowHits[r] != 0)
            {
                return r;
            }
        }

        return -1;
    }

    /// <summary>第 n 个命中行（1-based 序号），找不到返回 -1。</summary>
    public int HitRowAt(int ordinal)
    {
        if (ordinal <= 0)
        {
            return -1;
        }

        var seen = 0;
        for (var r = 0; r < RowHits.Length; r++)
        {
            if (RowHits[r] != 0 && ++seen == ordinal)
            {
                return r;
            }
        }

        return -1;
    }

    /// <summary>某行是第几个命中行（1-based），非命中行返回 0。</summary>
    public int OrdinalOfRow(int row)
    {
        if (!RowIsHit(row))
        {
            return 0;
        }

        var n = 0;
        for (var r = 0; r <= row && r < RowHits.Length; r++)
        {
            if (RowHits[r] != 0)
            {
                n++;
            }
        }

        return n;
    }
}

public static class SearchService
{
    /// <summary>命中总数上限：到顶后停止累计（防止 "*" 这类查询把 UI 拖住），只保留行标记。</summary>
    private const int MaxCellHits = 2_000_000;

    private static readonly Dictionary<string, SearchResult> Cache = new(StringComparer.Ordinal);
    private const int MaxCacheEntries = 8;

    public static SearchResult Search(
        Sheet sheet,
        SearchOptions options,
        IProgress<int>? progress = null,
        CancellationToken ct = default)
    {
        if (options.IsEmpty)
        {
            return SearchResult.Empty;
        }

        var key = CacheKey(sheet, options);
        lock (Cache)
        {
            if (Cache.TryGetValue(key, out var cached))
            {
                return cached;
            }
        }

        var result = Scan(sheet, options, progress, ct);

        if (!result.Cancelled)
        {
            lock (Cache)
            {
                if (Cache.Count >= MaxCacheEntries)
                {
                    Cache.Clear();
                }

                Cache[key] = result;
            }
        }

        return result;
    }

    public static void ClearCache()
    {
        lock (Cache)
        {
            Cache.Clear();
        }
    }

    private static string CacheKey(Sheet sheet, SearchOptions o) =>
        $"{sheet.Name}\u0001{o.Query}\u0001{o.Mode}\u0001{o.CaseSensitive}\u0001{o.ColumnFrom}\u0001{o.ColumnTo}\u0001{sheet.RowCount}x{sheet.ColCount}";

    private static SearchResult Scan(
        Sheet sheet,
        SearchOptions options,
        IProgress<int>? progress,
        CancellationToken ct)
    {
        var sw = Stopwatch.StartNew();
        var rows = sheet.RowCount;
        var cols = sheet.ColCount;
        var marks = new byte[rows];
        long cellHits = 0;
        var rowHits = 0;
        var truncated = false;

        var from = Math.Max(0, options.ColumnFrom ?? 0);
        var to = Math.Min(cols - 1, options.ColumnTo ?? (cols - 1));

        var comparison = options.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        Regex? regex = null;
        if (options.Mode == SearchMode.Regex)
        {
            regex = new Regex(
                options.Query,
                RegexOptions.CultureInvariant | (options.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase));
        }

        var lastPercent = -1;
        for (var c = from; c <= to && !ct.IsCancellationRequested; c++)
        {
            var col = sheet.ColumnBuffer(c);
            var limit = Math.Min(rows, col.Length);
            for (var r = 0; r < limit; r++)
            {
                // 每 256 格检查一次取消：兼顾响应速度与循环开销
                if ((r & 0xFF) == 0 && ct.IsCancellationRequested)
                {
                    break;
                }

                var v = col[r];
                if (v.Length == 0)
                {
                    continue;
                }

                var hit = regex is null
                    ? v.Contains(options.Query, comparison)
                    : regex.IsMatch(v);

                if (!hit)
                {
                    continue;
                }

                if (marks[r] == 0)
                {
                    marks[r] = 1;
                    rowHits++;
                }

                if (cellHits < MaxCellHits)
                {
                    cellHits++;
                }
                else
                {
                    truncated = true;
                }
            }

            if (progress is not null && to > from)
            {
                var percent = (int)(((c - from + 1) * 100.0) / (to - from + 1));
                if (percent != lastPercent)
                {
                    lastPercent = percent;
                    progress.Report(percent);
                }
            }
        }

        sw.Stop();
        return new SearchResult
        {
            Query = options.Query,
            Options = options,
            RowHits = marks,
            CellHitCount = cellHits,
            RowHitCount = rowHits,
            Truncated = truncated,
            Cancelled = ct.IsCancellationRequested,
            ElapsedMs = sw.ElapsedMilliseconds,
            Sheet = sheet,
        };
    }

    /// <summary>取出某一行里所有命中的列号，供高亮使用。</summary>
    public static List<int> HitColumnsInRow(Sheet sheet, SearchResult result, int row, int maxCols = 4096)
    {
        var list = new List<int>();
        if (result.Options is null)
        {
            return list;
        }

        var o = result.Options;
        var comparison = o.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        Regex? regex = o.Mode == SearchMode.Regex
            ? new Regex(o.Query, RegexOptions.CultureInvariant | (o.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase))
            : null;

        var from = Math.Max(0, o.ColumnFrom ?? 0);
        var to = Math.Min(sheet.ColCount - 1, o.ColumnTo ?? (sheet.ColCount - 1));
        for (var c = from; c <= to && list.Count < maxCols; c++)
        {
            var v = sheet.Get(row, c);
            if (v.Length == 0)
            {
                continue;
            }

            var hit = regex is null ? v.Contains(o.Query, comparison) : regex.IsMatch(v);
            if (hit)
            {
                list.Add(c);
            }
        }

        return list;
    }

    /// <summary>某行第一个命中的列，找不到返回 -1（用于横向滚动定位）。</summary>
    public static int FirstHitColumn(Sheet sheet, SearchResult result, int row)
    {
        var cols = HitColumnsInRow(sheet, result, row, maxCols: 1);
        return cols.Count > 0 ? cols[0] : -1;
    }
}
