namespace ExcelViewer.Model;

/// <summary>
/// 单元格文本池：同一个值在工作簿里只保留一份字符串实例。
/// 配表里重复值非常多（语言表同一句文本出现几十次，空值更是几十万次），
/// 去重后内存与字符串比较开销都显著下降。
/// </summary>
public sealed class StringPool
{
    private readonly Dictionary<string, string> _map = new(StringComparer.Ordinal);

    public int UniqueCount => _map.Count;

    public long HitCount { get; private set; }

    public long MissCount { get; private set; }

    public string Intern(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        if (_map.TryGetValue(value, out var existing))
        {
            HitCount++;
            return existing;
        }

        _map[value] = value;
        MissCount++;
        return value;
    }

    /// <summary>不做池化的快速路径：调用方已确认该字符串来自池。</summary>
    public void Clear() => _map.Clear();
}
