using System.Diagnostics;
using System.Text;
using ExcelDataReader;

namespace Bench;

/// <summary>
/// 精准基准：每种方案跑 Warmup+Iter 次，分别统计"完整读取"和"按需读一个 sheet"的耗时。
/// 用法: dotnet run -c Release -- [--diag|--iter N] [xlsx...]
/// </summary>
internal static class Perf
{
    public static void Run(string[] files, int iters)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Console.WriteLine($"iters={iters}   (每项先跑 1 次热身, 不计入)");

        foreach (var f in files)
        {
            if (!File.Exists(f))
            {
                Console.WriteLine($"!! missing {f}");
                continue;
            }

            Console.WriteLine();
            Console.WriteLine($"=== {Path.GetFileName(f)} ({new FileInfo(f).Length / 1048576.0:F2} MB) ===");

            MeasureAllSheets(f, iters);
            MeasureFirstSheetOnly(f, iters);
        }
    }

    /// <summary>一次性读完整个工作簿（用户抱怨的"打开慢"就是这种模式）。</summary>
    private static void MeasureAllSheets(string path, int iters)
    {
        var (rows, cols, cells) = (0, 0, 0L);
        RunReader(path, r =>
        {
            rows = 0;
            cols = 0;
            cells = 0;
            do
            {
                var c = r.FieldCount;
                cols = Math.Max(cols, c);
                while (r.Read())
                {
                    rows++;
                    for (var i = 0; i < c; i++)
                    {
                        if (!r.IsDBNull(i))
                        {
                            _ = r.GetValue(i);
                            cells++;
                        }
                    }
                }
            } while (r.NextResult());
        }, iters, "全工作簿");

        Console.WriteLine($"         -> rows={rows} cols={cols} cells={cells}");
    }

    /// <summary>只读第一个 sheet 后立即停下（首屏所需）。</summary>
    private static void MeasureFirstSheetOnly(string path, int iters)
    {
        var rows = 0;
        RunReader(path, r =>
        {
            rows = 0;
            while (r.Read())
            {
                rows++;
                for (var i = 0; i < r.FieldCount; i++)
                {
                    if (!r.IsDBNull(i))
                    {
                        _ = r.GetValue(i);
                    }
                }
            }
        }, iters, "只读首 sheet");

        Console.WriteLine($"         -> rows={rows}");
    }

    private static void RunReader(string path, Action<IExcelDataReader> body, int iters, string label)
    {
        var times = new List<double>();
        for (var i = 0; i <= iters; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            var sw = Stopwatch.StartNew();
            using (var fs = File.OpenRead(path))
            using (var reader = ExcelReaderFactory.CreateReader(fs))
            {
                body(reader);
            }

            sw.Stop();
            if (i > 0)
            {
                times.Add(sw.Elapsed.TotalMilliseconds);
            }
        }

        var managed = GC.GetTotalMemory(false) / 1048576.0;
        Console.WriteLine(
            $"  {label,-10} min={times.Min(),7:F0}ms  med={Median(times),7:F0}ms  max={times.Max(),7:F0}ms  " +
            $"managed={managed,7:F1}MB");
    }

    private static double Median(List<double> v)
    {
        var s = v.OrderBy(x => x).ToList();
        return s.Count % 2 == 1 ? s[s.Count / 2] : (s[(s.Count / 2) - 1] + s[s.Count / 2]) / 2;
    }
}
