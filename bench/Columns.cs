using System.Text;
using ExcelDataReader;
using ExcelViewer.Model;

namespace Bench;

/// <summary>
/// 列数核对：把 ExcelDataReader 逐行给出的 FieldCount 分布，
/// 与查看器加载后的 ColCount 对照，确认有没有丢列。
/// </summary>
internal static class Columns
{
    public static void Run(string path, int sheetIndex = 0)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Console.WriteLine($"### {Path.GetFileName(path)} sheet#{sheetIndex}");

        // 1) 原始 reader：看 FieldCount 分布与前几行的真实列数
        using (var fs = File.OpenRead(path))
        using (var reader = ExcelReaderFactory.CreateReader(fs))
        {
            for (var i = 0; i < sheetIndex && reader.NextResult(); i++)
            {
            }

            Console.WriteLine($"  reader.Name={reader.Name} reader.FieldCount(初始)={reader.FieldCount}");
            var histogram = new SortedDictionary<int, int>();
            var rowNo = 0;
            var samples = new List<string>();
            while (reader.Read())
            {
                var n = reader.FieldCount;
                histogram[n] = histogram.GetValueOrDefault(n) + 1;
                if (rowNo < 4)
                {
                    var cells = new List<string>();
                    for (var c = 0; c < n; c++)
                    {
                        cells.Add(reader.IsDBNull(c) ? "∅" : Trunc(reader.GetValue(c)?.ToString() ?? string.Empty, 12));
                    }

                    samples.Add($"    row{rowNo} n={n}: {string.Join(" | ", cells)}");
                }

                rowNo++;
            }

            Console.WriteLine($"  总行数={rowNo}");
            Console.WriteLine("  FieldCount 分布: " + string.Join(", ", histogram.Select(kv => $"{kv.Key}列×{kv.Value}行")));
            foreach (var s in samples)
            {
                Console.WriteLine(s);
            }
        }

        // 2) 查看器加载结果
        var pool = new StringPool();
        var sheet = WorkbookLoader.LoadSheet(path, sheetIndex, pool);
        Console.WriteLine($"  查看器: RowCount={sheet.RowCount} ColCount={sheet.ColCount} "
            + $"SourceCol={sheet.SourceColCount} 非空={sheet.NonEmptyCellCount}");

        var sb = new StringBuilder("  每列非空计数: ");
        for (var c = 0; c < Math.Min(sheet.ColCount, 50); c++)
        {
            var n = 0;
            for (var r = 0; r < sheet.RowCount; r++)
            {
                if (sheet.Get(r, c).Length != 0)
                {
                    n++;
                }
            }

            sb.Append($"{c}:{n} ");
        }

        Console.WriteLine(sb.ToString());
    }

    private static string Trunc(string s, int n) => s.Length <= n ? s : s[..n] + "…";
}

