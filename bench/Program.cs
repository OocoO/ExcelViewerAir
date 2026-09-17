using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;

namespace Bench;

/// <summary>
/// 基准脚本：比较各 xlsx 解析方案在全量读取 + 全表搜索上的耗时与内存。
/// 用法: dotnet run -c Release -- [要测的 xlsx 路径...]
/// </summary>
internal static class Program
{
    private const int SearchRepeats = 5;
    private const string Needle = "1";

    private static readonly string[] DefaultFiles =
    {
        @"C:\Users\18223\Programs\excel2csv\test\Languages.xlsx",
        @"C:\Users\18223\Programs\excel2csv\test\Level.xlsx",
        @"C:\Users\18223\Programs\excel2csv\test\Item.xlsx",
    };

    private static async Task<int> Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        if (args.Length > 0 && args[0] == "--diag")
        {
            foreach (var p in args.Skip(1))
            {
                Diag.Run(p);
            }

            return 0;
        }

        if (args.Length > 0 && args[0] == "--cols")
        {
            foreach (var p in args.Skip(1))
            {
                Columns.Run(p);
            }

            return 0;
        }

        if (args.Length > 0 && args[0] == "--perf")
        {
            var iters = 4;
            var rest = args.Skip(1).ToList();
            if (rest.Count > 0 && rest[0].StartsWith("--iter", StringComparison.Ordinal))
            {
                iters = int.Parse(rest[0].Split('=')[1]);
                rest.RemoveAt(0);
            }

            Perf.Run(rest.Count > 0 ? rest.ToArray() : DefaultFiles, iters);
            return 0;
        }

        var files = args.Length > 0
            ? args
            : DefaultFiles;

        foreach (var f in files)
        {
            if (!File.Exists(f))
            {
                Console.WriteLine($"!! missing: {f}");
                continue;
            }

            var size = new FileInfo(f).Length / 1048576.0;
            Console.WriteLine();
            Console.WriteLine($"=== {Path.GetFileName(f)} ({size:F2} MB) ===");

            await RunAsync("ExcelDataReader", f, ReadWithExcelDataReaderAsync);
            
            await RunAsync("OpenXml", f, ReadWithOpenXmlAsync);
        }

        return 0;
    }

    private static async Task RunAsync(string name, string file, Func<string, Task<SheetTable>> reader)
    {
        try
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            var sw = Stopwatch.StartNew();
            var table = await reader(file);
            sw.Stop();
            var loadMs = sw.Elapsed.TotalMilliseconds;

            var beforeSearch = GC.GetTotalAllocatedBytes(false);
            sw.Restart();
            var hits = Search(table, Needle);
            sw.Stop();
            var firstSearchMs = sw.Elapsed.TotalMilliseconds;

            var totalMs = 0.0;
            for (var i = 0; i < SearchRepeats; i++)
            {
                sw.Restart();
                Search(table, Needle);
                sw.Stop();
                totalMs += sw.Elapsed.TotalMilliseconds;
            }

            var avgSearchMs = totalMs / SearchRepeats;
            var allocMb = (GC.GetTotalAllocatedBytes(false) - beforeSearch) / 1048576.0;

            Console.WriteLine(
                $"  {name,-16} load={loadMs,8:F0}ms  rows={table.RowCount,7}  cols={table.ColCount,3}  " +
                $"cells={table.CellCount,9}  search1={firstSearchMs,7:F0}ms  searchAvg={avgSearchMs,7:F0}ms  " +
                $"hits={hits,7}  searchAlloc={allocMb,7:F1}MB  managed={GC.GetTotalMemory(false) / 1048576.0,7:F1}MB");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  {name,-16} FAILED: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static int Search(SheetTable table, string needle)
    {
        var n = 0;
        for (var r = 0; r < table.RowCount; r++)
        {
            for (var c = 0; c < table.ColCount; c++)
            {
                var s = table.Get(r, c);
                if (s is not null && s.Contains(needle, StringComparison.OrdinalIgnoreCase))
                {
                    n++;
                }
            }
        }

        return n;
    }

    private static string? Normalize(object? value) => value switch
    {
        null => null,
        string s => s,
        double d => d.ToString("R"),
        DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss"),
        bool b => b ? "TRUE" : "FALSE",
        _ => Convert.ToString(value),
    };

    // ---------------- ExcelDataReader ----------------

    private static Task<SheetTable> ReadWithExcelDataReaderAsync(string path) => Task.Run(() =>
    {
        using var stream = File.OpenRead(path);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var table = new SheetTable();
        do
        {
            table.BeginSheet(reader.Name ?? "Sheet");
            while (reader.Read())
            {
                for (var c = 0; c < reader.FieldCount; c++)
                {
                    table.Set(reader.IsDBNull(c) ? null : Normalize(reader.GetValue(c)));
                }

                table.EndRow(reader.FieldCount);
            }
        } while (reader.NextResult());

        return table;
    });

    // ---------------- DocumentFormat.OpenXml (流式逐行) ----------------

    private static Task<SheetTable> ReadWithOpenXmlAsync(string path) => Task.Run(() =>
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var wbPart = doc.WorkbookPart!;
        var sst = wbPart.SharedStringTablePart?.SharedStringTable;
        var table = new SheetTable();

        foreach (var sheet in wbPart.Workbook.Sheets!.Elements<Sheet>())
        {
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
            table.BeginSheet(sheet.Name?.Value ?? "Sheet");
            foreach (var row in wsPart.Worksheet.Descendants<Row>())
            {
                var cells = row.Elements<Cell>().ToList();
                if (cells.Count == 0)
                {
                    table.EndRow(0);
                    continue;
                }

                var last = ColumnIndex(cells[^1].CellReference?.Value);
                var buf = new string?[last + 1];
                foreach (var cell in cells)
                {
                    var col = ColumnIndex(cell.CellReference?.Value);
                    if (col < buf.Length)
                    {
                        buf[col] = CellText(cell, sst);
                    }
                }

                for (var c = 0; c < buf.Length; c++)
                {
                    table.Set(buf[c]);
                }

                table.EndRow(buf.Length);
            }
        }

        return table;
    });

    private static string? CellText(Cell cell, SharedStringTable? sst)
    {
        var v = cell.CellValue?.InnerText;
        if (v is null)
        {
            return cell.InlineString?.Text?.Text;
        }

        if (cell.DataType?.Value == CellValues.SharedString && sst is not null)
        {
            return int.TryParse(v, out var idx) && idx >= 0 && idx < sst.ChildElements.Count
                ? sst.ChildElements[idx].InnerText
                : v;
        }

        return v;
    }

    private static int ColumnIndex(string? reference)
    {
        if (string.IsNullOrEmpty(reference))
        {
            return 0;
        }

        var n = 0;
        foreach (var ch in reference)
        {
            if (ch is >= 'A' and <= 'Z')
            {
                n = (n * 26) + (ch - 'A' + 1);
            }
            else
            {
                break;
            }
        }

        return Math.Max(0, n - 1);
    }
}

/// <summary>基准用的极简列式表：按 sheet 顺序追加，列数组按需增长。</summary>
internal sealed class SheetTable
{
    private readonly List<string?[]> _cols = new();
    private string?[] _row = new string?[32];
    private int _rowLen;
    private int _rowIndex;

    public int RowCount { get; private set; }

    public int ColCount { get; private set; }

    public long CellCount { get; private set; }

    public void BeginSheet(string name)
    {
        _cols.Clear();
        RowCount = 0;
        ColCount = 0;
        CellCount = 0;
        _rowIndex = 0;
        _rowLen = 0;
    }

    public void Set(string? value)
    {
        if (_rowLen >= _row.Length)
        {
            Array.Resize(ref _row, _row.Length * 2);
        }

        _row[_rowLen++] = value;
    }

    public void EndRow(int width)
    {
        Flush(_rowLen);
        _rowLen = 0;
    }

    private void Flush(int count)
    {
        while (_cols.Count < count)
        {
            _cols.Add(new string?[256]);
        }

        for (var c = 0; c < count; c++)
        {
            var col = _cols[c];
            if (_rowIndex >= col.Length)
            {
                Array.Resize(ref col, col.Length * 2);
                _cols[c] = col;
            }

            var v = _row[c];
            col[_rowIndex] = v;
            if (v is not null)
            {
                CellCount++;
            }
        }

        _rowIndex++;
        RowCount = _rowIndex;
        ColCount = Math.Max(ColCount, count);
    }

    public string? Get(int r, int c)
    {
        if (c >= _cols.Count)
        {
            return null;
        }

        var col = _cols[c];
        return r < col.Length ? col[r] : null;
    }
}

