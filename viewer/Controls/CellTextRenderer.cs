using System.Globalization;
using System.Windows;
using System.Windows.Media;

namespace ExcelViewer.Controls;

/// <summary>
/// 单元格文本排版/绘制器。
///
/// 这里不用 TextFormatter：实测它的 TextRunCache 会按 TextSource 实例复用上一次的排版结果，
/// 导致同一实例连续绘制不同文本时字形/宽度对不上（表现为文字被画到别的列）。
/// 改用 FormattedText —— 结果确定，配合下面两层缓存把开销压到可接受：
///   1. 宽度缓存：列宽自适应只需量一次，按字符串引用缓存；
///   2. 排版缓存：可视单元格按 (字符串, 对齐, 可用宽度) 缓存 FormattedText，滚动时直接复用。
/// 配表里重复值极多（且字符串池已去重），命中率很高。
/// </summary>
internal sealed class CellTextRenderer
{
    /// <summary>0=左对齐（数据行），1=居中（表头）。</summary>
    public const int AlignLeft = 0;

    public const int AlignCenter = 1;

    /// <summary>
    /// 排版缓存条目上限。可视区一次大约 30 行 x 40 列 = 1200 格，
    /// 取 4 倍余量，保证来回滚动时不至于把刚排版好的再丢一次。
    /// 超限时按插入顺序淘汰一半（而不是整体清空），避免出现"整帧重建"的性能尖刺。
    /// </summary>
    private const int MaxFormattedCache = 16384;

    private readonly Typeface _typeface;
    private readonly double _fontSize;
    private readonly double _pixelsPerDip;
    private readonly Brush _foreground;
    private readonly Dictionary<string, double> _widthCache = new(ReferenceEqualityComparer.Instance);
    private readonly Dictionary<LayoutKey, FormattedText> _layoutCache = new();
    private readonly Queue<LayoutKey> _insertOrder = new();

    public CellTextRenderer(Typeface typeface, double fontSize, double pixelsPerDip, Brush foreground)
    {
        _typeface = typeface;
        _fontSize = fontSize;
        _pixelsPerDip = pixelsPerDip;
        _foreground = foreground;
    }

    /// <summary>测量单行文本宽度（不绘制），用于列宽自适应与数字右对齐。</summary>
    public double Measure(string text)
    {
        if (text.Length == 0)
        {
            return 0;
        }

        if (_widthCache.TryGetValue(text, out var cached))
        {
            return cached;
        }

        var ft = Create(text, double.PositiveInfinity, AlignLeft);
        var w = ft.WidthIncludingTrailingWhitespace;
        _widthCache[text] = w;
        return w;
    }

    /// <summary>排版并绘制。超出可用宽度时按字符截断并加省略号。</summary>
    public void Draw(
        DrawingContext dc,
        string text,
        int alignment,
        double x,
        double y,
        double maxWidth,
        double lineHeight)
    {
        if (text.Length == 0 || maxWidth <= 1)
        {
            return;
        }

        var ft = GetLayout(text, alignment, maxWidth);
        var dy = y + ((lineHeight - ft.Height) / 2);
        dc.DrawText(ft, new Point(x, dy));
    }

    /// <summary>仅供诊断：当前排版缓存条目数。</summary>
    public int LayoutCacheCount => _layoutCache.Count;

    private FormattedText GetLayout(string text, int alignment, double maxWidth)
    {
        // 宽度取整到 0.5 像素做键，避免因浮点差异让缓存失效
        var rounded = Math.Round(maxWidth * 2, MidpointRounding.AwayFromZero) / 2;
        var key = new LayoutKey(text, (byte)alignment, rounded);
        if (_layoutCache.TryGetValue(key, out var cached))
        {
            return cached;
        }

        if (_layoutCache.Count >= MaxFormattedCache)
        {
            EvictOldest(MaxFormattedCache / 2);
        }

        var ft = Create(text, rounded, alignment);
        _layoutCache[key] = ft;
        _insertOrder.Enqueue(key);
        return ft;
    }

    /// <summary>按插入顺序淘汰，每次丢一半，摊还成本 O(1)。</summary>
    private void EvictOldest(int count)
    {
        for (var i = 0; i < count && _insertOrder.Count > 0; i++)
        {
            _layoutCache.Remove(_insertOrder.Dequeue());
        }
    }

    private FormattedText Create(string text, double maxWidth, int alignment)
    {
        var ft = new FormattedText(
            text,
            CultureInfo.CurrentUICulture,
            FlowDirection.LeftToRight,
            _typeface,
            _fontSize,
            _foreground,
            _pixelsPerDip);

        if (alignment == AlignCenter)
        {
            ft.TextAlignment = TextAlignment.Center;
        }

        if (!double.IsPositiveInfinity(maxWidth))
        {
            ft.MaxTextWidth = Math.Max(1, maxWidth);

            // 行高是固定的（默认 22px），只画一行：放不下就用 … 截断。
            // 不设 MaxLineCount 的话 FormattedText 会按列宽自动折行，整个文本块再被 Draw 垂直居中，
            // 于是多出来的行会画到上下相邻的行上去 —— 表现就是"格子里文字叠成一团"。
            ft.MaxLineCount = 1;
            ft.Trimming = TextTrimming.CharacterEllipsis;
        }

        return ft;
    }

    /// <summary>DPI/字号变化时重建（缓存里的对象绑定旧 DPI，必须一起丢弃）。</summary>
    public CellTextRenderer WithTypeface(Typeface typeface, double fontSize, double pixelsPerDip) =>
        new(typeface, fontSize, pixelsPerDip, _foreground);

    public void ClearCache()
    {
        _widthCache.Clear();
        _layoutCache.Clear();
        _insertOrder.Clear();
    }

    private readonly record struct LayoutKey(string Text, byte Align, double MaxWidth);

    /// <summary>按引用比较字符串：池化后的相同文本就是同一个实例，比逐字符比较快得多。</summary>
    private sealed class ReferenceEqualityComparer : IEqualityComparer<string>
    {
        public static readonly ReferenceEqualityComparer Instance = new();

        public bool Equals(string? x, string? y) => ReferenceEquals(x, y);

        public int GetHashCode(string obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
    }
}
