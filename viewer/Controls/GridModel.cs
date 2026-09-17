using ExcelViewer.Model;

namespace ExcelViewer.Controls;

/// <summary>
/// 表格布局模型：负责列宽、行高、视口坐标换算。
///
/// 设计要点：行高固定、列宽存累积数组，任何 (行,列) 到像素的换算都是 O(log n) 或 O(1)，
/// 因此滚动到第 3 万行和第 3 行的开销完全一致 —— 这是自绘虚拟化表格相对 DataGrid/WPS 的核心优势。
/// </summary>
public sealed class GridModel
{
    public const double DefaultRowHeight = 22.0;
    public const double MinColumnWidth = 36.0;
    public const double MaxColumnWidth = 520.0;
    public const double CellPaddingX = 6.0;
    public const double RowNumberGutterWidth = 56.0;

    private double[] _colWidths = Array.Empty<double>();
    private double[] _colOffsets = Array.Empty<double>(); // 长度 = 列数 + 1，最后一项是总宽
    private Sheet? _sheet;
    private bool _showRowNumbers = true;

    public Sheet? Sheet => _sheet;

    /// <summary>识别出的表头行（-1 表示没有），列宽自适应时保证它被采样到。</summary>
    public int HeaderRowIndex { get; set; } = -1;

    public int RowCount => _sheet?.RowCount ?? 0;

    public int ColCount => _sheet?.ColCount ?? 0;

    public double RowHeight { get; set; } = DefaultRowHeight;

    public double HeaderHeight { get; set; } = 26.0;

    public bool ShowRowNumbers
    {
        get => _showRowNumbers;
        set
        {
            _showRowNumbers = value;
            RebuildOffsets();
        }
    }

    /// <summary>行号栏宽度（不显示时为 0）。</summary>
    public double GutterWidth => ShowRowNumbers ? RowNumberGutterWidth : 0;

    public double TotalWidth => _colOffsets.Length == 0 ? 0 : _colOffsets[^1];

    public double TotalHeight => HeaderHeight + (RowCount * RowHeight);

    public void SetSheet(Sheet? sheet, bool resetWidths = true)
    {
        _sheet = sheet;
        AutoFitColumns();
    }

    /// <summary>
    /// 真实文本测量回调：由 ExcelGrid 在初始化时注入（用当前字体/DPI 排版一次）。
    /// 为 null 时退化为按半角单位估算，保证不依赖 UI 也能算出可用的列宽。
    /// </summary>
    public static Func<string, double>? TextMeasurer { get; set; }

    /// <summary>
    /// 计算初始列宽。
    /// 采样范围：前若干行 + 识别出的表头行（保证字段名完整显示），
    /// 太长或太宽的列一律截断到上限 —— 配表里某些列是整段说明文本，全量测量既慢又没意义。
    /// </summary>
    public void AutoFitColumns()
    {
        if (_sheet is null || _sheet.ColCount == 0)
        {
            _colWidths = Array.Empty<double>();
            RebuildOffsets();
            return;
        }

        const int sampleRows = 80;
        const int maxUnits = 110; // 约 40 个汉字，列宽上限的判定基准

        var sheet = _sheet;
        var n = sheet.ColCount;
        var rows = Math.Min(sheet.RowCount, sampleRows);
        var widths = new double[n];
        var measurer = TextMeasurer;
        var headerRow = HeaderRowIndex;

        for (var c = 0; c < n; c++)
        {
            var maxPx = 0.0;
            for (var r = 0; r < rows; r++)
            {
                maxPx = Accumulate(sheet.Get(r, c), maxPx, measurer, maxUnits);
                if (maxPx >= MaxColumnWidth)
                {
                    break;
                }
            }

            // 表头行可能在采样窗口之外，单独补一次，避免字段名被裁成 "GDE_FIELD_NA…"
            if (headerRow >= rows && headerRow < sheet.RowCount && maxPx < MaxColumnWidth)
            {
                maxPx = Accumulate(sheet.Get(headerRow, c), maxPx, measurer, maxUnits);
            }

            widths[c] = maxPx == double.MaxValue
                ? MaxColumnWidth
                : Math.Clamp(maxPx + (CellPaddingX * 2) + 1, MinColumnWidth, MaxColumnWidth);
        }

        _colWidths = widths;
        RebuildOffsets();
    }

    /// <summary>累计某列已见到的最大文本宽度；返回 double.MaxValue 表示"超出上限"。</summary>
    private static double Accumulate(string value, double current, Func<string, double>? measurer, int maxUnits)
    {
        if (value.Length == 0)
        {
            return current;
        }

        double px;
        if (measurer is not null)
        {
            px = measurer(value);
        }
        else
        {
            // 没有字体测量能力时按半角单位估算
            if (DisplayWidth.MeasureCapped(value, maxUnits) > maxUnits)
            {
                return double.MaxValue;
            }

            px = DisplayWidth.Measure(value) * 7.2;
        }

        return px > current ? px : current;
    }

    /// <summary>半角单位 -> 像素宽度。用经验系数即可，避免逐格测量字体。</summary>
    public static double WidthFromUnits(int units)
    {
        var px = (units * 7.2) + (CellPaddingX * 2) + 1;
        return Math.Clamp(px, MinColumnWidth, MaxColumnWidth);
    }

    public void SetColumnWidth(int col, double width)
    {
        if (col < 0 || col >= _colWidths.Length)
        {
            return;
        }

        _colWidths[col] = Math.Clamp(width, MinColumnWidth, MaxColumnWidth);
        RebuildOffsets();
    }

    public double GetColumnWidth(int col) =>
        (uint)col < (uint)_colWidths.Length ? _colWidths[col] : MinColumnWidth;

    private void RebuildOffsets()
    {
        if (_colWidths.Length == 0)
        {
            _colOffsets = Array.Empty<double>();
            return;
        }

        var offsets = new double[_colWidths.Length + 1];
        var acc = GutterWidth;
        offsets[0] = acc;
        for (var c = 0; c < _colWidths.Length; c++)
        {
            acc += _colWidths[c];
            offsets[c + 1] = acc;
        }

        _colOffsets = offsets;
    }

    /// <summary>列左边界（含行号栏偏移）。</summary>
    public double ColumnLeft(int col) => (uint)col < (uint)_colWidths.Length ? _colOffsets[col] : 0;

    public double ColumnRight(int col) => (uint)col < (uint)_colWidths.Length ? _colOffsets[col + 1] : 0;

    /// <summary>内容区总宽（不含行号栏）。</summary>
    public double ContentWidth => Math.Max(0, TotalWidth - GutterWidth);

    /// <summary>由 x（grid 内部坐标，已含横向滚动偏移）反查列号，越界返回 -1。</summary>
    public int ColumnAt(double x)
    {
        if (_colWidths.Length == 0 || x < GutterWidth)
        {
            return -1;
        }

        var lo = 0;
        var hi = _colWidths.Length - 1;
        while (lo < hi)
        {
            var mid = (lo + hi + 1) / 2;
            if (_colOffsets[mid] <= x)
            {
                lo = mid;
            }
            else
            {
                hi = mid - 1;
            }
        }

        return lo < _colWidths.Length && x < _colOffsets[lo + 1] ? lo : -1;
    }

    /// <summary>行号（0-based）。y 是内容坐标，调用方需先减去表头高度。</summary>
    public int RowAt(double y)
    {
        if (y < 0)
        {
            return -1;
        }

        var r = (int)(y / RowHeight);
        return r < RowCount ? r : -1;
    }

    /// <summary>指定列区间内实际有内容的最大宽度（用于双击列边界自适应）。</summary>
    public double MeasureColumn(int col, int maxRows, Func<int, int, string> getCell)
    {
        if (_sheet is null)
        {
            return MinColumnWidth;
        }

        var rows = Math.Min(RowCount, maxRows);
        var maxUnits = 0;
        for (var r = 0; r < rows; r++)
        {
            var v = getCell(r, col);
            if (v.Length == 0)
            {
                continue;
            }

            var w = DisplayWidth.MeasureCapped(v, 120);
            if (w > maxUnits)
            {
                maxUnits = w;
                if (maxUnits >= 120)
                {
                    break;
                }
            }
        }

        return WidthFromUnits(maxUnits);
    }
}
