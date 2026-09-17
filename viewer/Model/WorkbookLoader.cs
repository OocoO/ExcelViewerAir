using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using ExcelDataReader;

namespace ExcelViewer.Model;

/// <summary>工作簿加载器：xlsx / xls / xlsm / csv 只读解析，不修改源文件。</summary>
public static class WorkbookLoader
{
    static WorkbookLoader()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public static bool IsSupported(string path)
    {
        var ext = Path.GetExtension(path);
        return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".xlsb", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".xls", StringComparison.OrdinalIgnoreCase)
            || ext.Equals(".csv", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>先只拿 sheet 名单，用于建标签页（避免为了列名解析一遍数据）。</summary>
    public static List<string> ReadSheetNames(string path)
    {
        if (Path.GetExtension(path).Equals(".csv", StringComparison.OrdinalIgnoreCase))
        {
            return new List<string> { Path.GetFileNameWithoutExtension(path) };
        }

        using var stream = OpenRead(path);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var names = new List<string>();
        do
        {
            names.Add(reader.Name ?? $"Sheet{names.Count + 1}");
        } while (reader.NextResult());

        return names.Count == 0 ? new List<string> { "Sheet1" } : names;
    }

    /// <summary>
    /// 读取指定序号的工作表。<paramref name="sheetIndex"/> 从 0 开始。
    /// 会在后台线程调用，不触碰 UI。
    /// </summary>
    public static Sheet LoadSheet(string path, int sheetIndex, StringPool pool, CancellationToken ct = default)
    {
        var sw = Stopwatch.StartNew();
        Sheet sheet;
        if (Path.GetExtension(path).Equals(".csv", StringComparison.OrdinalIgnoreCase))
        {
            sheet = LoadCsv(path, pool, ct);
        }
        else
        {
            sheet = LoadExcelSheet(path, sheetIndex, pool, ct);
        }

        sheet.Finish();
        sheet.LoadMs = sw.ElapsedMilliseconds;
        return sheet;
    }

    private static Sheet LoadExcelSheet(string path, int sheetIndex, StringPool pool, CancellationToken ct)
    {
        using var stream = OpenRead(path);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        var index = 0;
        while (index < sheetIndex && reader.NextResult())
        {
            index++;
        }

        if (index != sheetIndex)
        {
            throw new InvalidOperationException($"工作表序号 {sheetIndex} 不存在。");
        }

        var name = reader.Name ?? $"Sheet{sheetIndex + 1}";
        var sheet = new Sheet(name, pool);

        while (reader.Read())
        {
            ct.ThrowIfCancellationRequested();
            sheet.BeginRow();
            var n = reader.FieldCount;
            for (var c = 0; c < n; c++)
            {
                if (reader.IsDBNull(c))
                {
                    sheet.AddCell(null);
                    continue;
                }

                // 日期列整体是 DateTime，混排列要逐格判类型，这里只做一次 GetFieldType
                var type = reader.GetFieldType(c);
                var text = type == typeof(DateTime)
                    ? FormatDate(reader.GetDateTime(c))
                    : FormatValue(reader.GetValue(c), out _);
                sheet.AddCell(text);
            }

            sheet.EndRow();
        }

        return sheet;
    }

    private static Sheet LoadCsv(string path, StringPool pool, CancellationToken ct)
    {
        var sheet = new Sheet(Path.GetFileNameWithoutExtension(path), pool);

        // 编码探测：有 BOM 按 BOM，否则按 UTF-8 严格解码，失败再退回 GBK。
        using var stream = OpenRead(path);
        using var reader = new StreamReader(stream, DetectEncoding(stream), detectEncodingFromByteOrderMarks: true);

        var parser = new CsvParser();
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            ct.ThrowIfCancellationRequested();
            sheet.BeginRow();
            parser.Parse(line, sheet);
            sheet.EndRow();
        }

        return sheet;
    }

    private static Encoding DetectEncoding(Stream stream)
    {
        Span<byte> head = stackalloc byte[3];
        var read = stream.Read(head);
        stream.Seek(-read, SeekOrigin.Current);
        if (read >= 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF)
        {
            return new UTF8Encoding(true);
        }

        if (read >= 2 && head[0] == 0xFF && head[1] == 0xFE)
        {
            return Encoding.Unicode;
        }

        if (read >= 2 && head[0] == 0xFE && head[1] == 0xFF)
        {
            return Encoding.BigEndianUnicode;
        }

        return new UTF8Encoding(false);
    }

    private static FileStream OpenRead(string path) =>
        new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1 << 16, FileOptions.SequentialScan);

    private static string FormatDate(DateTime dt) => dt.TimeOfDay == TimeSpan.Zero
        ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        : dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

    /// <summary>把单元格值转成显示文本。数字用最短往返表示，避免 1.0000000000000002 之类噪声。</summary>
    public static string FormatValue(object? value, out bool isNumeric)
    {
        switch (value)
        {
            case null:
                isNumeric = false;
                return string.Empty;
            case string s:
                isNumeric = false;
                return DisplayWidth.Flatten(s.Trim());
            case double d:
                isNumeric = true;
                return FormatDouble(d);
            case float f:
                isNumeric = true;
                return FormatDouble(f);
            case decimal m:
                isNumeric = true;
                return m.ToString(CultureInfo.InvariantCulture);
            case int or long or short or byte or sbyte or uint or ulong or ushort:
                isNumeric = true;
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            case bool b:
                isNumeric = false;
                return b ? "TRUE" : "FALSE";
            case DateTime dt:
                isNumeric = false;
                return FormatDate(dt);
            case TimeSpan ts:
                isNumeric = false;
                return ts.ToString();
            default:
                isNumeric = false;
                return DisplayWidth.Flatten(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
        }
    }

    private static string FormatDouble(double d)
    {
        if (double.IsNaN(d) || double.IsInfinity(d))
        {
            return d.ToString(CultureInfo.InvariantCulture);
        }

        // 整数用定点输出，避免出现 1E+06 这种表格里很别扭的写法
        if (d == Math.Floor(d) && Math.Abs(d) < 1e15)
        {
            return ((long)d).ToString(CultureInfo.InvariantCulture);
        }

        return d.ToString("R", CultureInfo.InvariantCulture);
    }
}

/// <summary>最小 RFC4180 CSV 解析器（支持引号包裹、双引号转义、字段内换行由调用方保证单行）。</summary>
internal sealed class CsvParser
{
    private readonly StringBuilder _field = new(64);

    public void Parse(string line, Sheet sheet)
    {
        _field.Clear();
        var inQuotes = false;
        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];
            if (inQuotes)
            {
                if (ch == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        _field.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    _field.Append(ch);
                }
            }
            else if (ch == '"')
            {
                inQuotes = true;
            }
            else if (ch == ',')
            {
                sheet.AddCell(_field.ToString());
                _field.Clear();
            }
            else
            {
                _field.Append(ch);
            }
        }

        sheet.AddCell(_field.ToString());
    }
}
