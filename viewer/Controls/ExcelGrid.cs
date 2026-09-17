using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using ExcelViewer.Model;
using ExcelViewer.Search;

namespace ExcelViewer.Controls;

/// <summary>
/// 自绘只读表格。
///
/// 性能策略：
///   1. 只绘制可视区的行/列（虚拟化），滚动到第 3 万行的代价与第 3 行相同；
///   2. 绘制发生在 <see cref="OnRender"/>，只在滚动/选中/搜索/尺寸变化时触发，不做定时重绘；
///   3. 单元格文本走 <see cref="CellTextRenderer"/>，绘制阶段几乎零分配；
///   4. 列宽取自加载时的估算，渲染路径上不做字体测量。
///
/// 直接继承 FrameworkElement 而不是 Panel/DataGrid：完全掌控视觉树与绘制，
/// 滚动条作为两个子元素手工布局，避免 DataGrid 那套容器+绑定带来的开销。
/// </summary>
public sealed class ExcelGrid : FrameworkElement
{
    private const int HeaderRowIndex = 0;

    private readonly GridModel _model = new();
    private readonly ScrollBar _vBar;
    private readonly ScrollBar _hBar;
    private readonly DrawingVisual _visual = new();

    private CellTextRenderer? _text;
    private CellTextRenderer? _boldText;
    private double _vOffset;
    private double _hOffset;
    private double _fontSizePt = 13;
    private int _activeRow = HeaderRowIndex;
    private int _activeCol;
    private int _anchorRow = HeaderRowIndex;
    private int _anchorCol;
    private bool _dragging;
    private bool _resizing;
    private int _resizeCol = -1;
    private double _resizeStartX;
    private double _resizeStartWidth;
    private SearchResult _search = SearchResult.Empty;
    private List<int> _activeRowHitCols = new();
    private int _badgeCol = -1;
    private string _badgeText = string.Empty;
    private bool _suspendScrollSync;

    /// <summary>半角数字在当前字号下的估算宽度（13pt 微软雅黑约 7.2px）；随字号线性缩放。</summary>
    private double _estimatedDigitWidth = 7.2;

    public ExcelGrid()
    {
        ClipToBounds = true;
        SnapsToDevicePixels = true;
        Focusable = true;
        FocusVisualStyle = null;
        UseLayoutRounding = true;

        _vBar = new ScrollBar
        {
            Orientation = Orientation.Vertical,
            Width = 15,
            Minimum = 0,
            SmallChange = 40,
            LargeChange = 400,
            Focusable = false,
        };
        _hBar = new ScrollBar
        {
            Orientation = Orientation.Horizontal,
            Height = 15,
            Minimum = 0,
            SmallChange = 80,
            LargeChange = 400,
            Focusable = false,
        };
        _vBar.Scroll += (_, _) => OnScrollChanged();
        _hBar.Scroll += (_, _) => OnScrollChanged();

        AddVisualChild(_visual);
        AddLogicalChild(_visual);
        AddVisualChild(_vBar);
        AddVisualChild(_hBar);
        AddLogicalChild(_vBar);
        AddLogicalChild(_hBar);

        RebuildTextRenderer();
    }

    // ---------------- 对外状态 ----------------

    public GridModel Model => _model;

    public Sheet? Sheet => _model.Sheet;

    public int ActiveRow => _activeRow;

    public int ActiveCol => _activeCol;

    public SearchResult Search
    {
        get => _search;
        set
        {
            _search = value ?? SearchResult.Empty;
            _badgeCol = -1;
            RefreshActiveRowHits();
            InvalidateVisual();
            RaiseViewportChanged();
        }
    }

    public double FontSizePt
    {
        get => _fontSizePt;
        set
        {
            var v = Math.Clamp(value, 8, 28);
            if (Math.Abs(_fontSizePt - v) < 0.01)
            {
                return;
            }

            _fontSizePt = v;
            RebuildTextRenderer();
            _model.AutoFitColumns();
            UpdateScrollBars();
            InvalidateVisual();
        }
    }

    public bool ShowRowNumbers
    {
        get => _model.ShowRowNumbers;
        set
        {
            _model.ShowRowNumbers = value;
            UpdateScrollBars();
            InvalidateVisual();
        }
    }

    /// <summary>识别出的表头行（-1 表示没有）。该行用加粗+淡蓝底显示，方便一眼定位字段名。</summary>
    public int DetectedHeaderRow
    {
        get => _headerRowIndex;
        set
        {
            if (_headerRowIndex == value)
            {
                return;
            }

            _headerRowIndex = value;
            InvalidateVisual();
        }
    }

    private int _headerRowIndex = -1;

    /// <summary>是否冻结首行（把第一行当表头显示，居中加粗）。</summary>
    public bool HideHeaderRow
    {
        get => _hideHeaderRow;
        set
        {
            if (_hideHeaderRow == value)
            {
                return;
            }

            _hideHeaderRow = value;
            InvalidateVisual();
        }
    }

    private bool _hideHeaderRow;

    public int VisibleRowCount => Math.Max(1, (int)((ActualHeight - _model.HeaderHeight - _hBar.Height) / _model.RowHeight));

    /// <summary>诊断用：当前排版缓存条目数。</summary>
    public int TextLayoutCacheCount => _text?.LayoutCacheCount ?? 0;

    /// <summary>当前视口顶部对应的行号（0 为表头行）。</summary>
    public int TopVisibleRow => Math.Max(0, (int)(_vOffset / _model.RowHeight));

    public event EventHandler? ViewportChanged;

    public event EventHandler? ActiveCellChanged;

    // ---------------- 视觉树 ----------------

    protected override int VisualChildrenCount => 3;

    protected override Visual GetVisualChild(int index) => index switch
    {
        0 => _visual,
        1 => _vBar,
        2 => _hBar,
        _ => throw new ArgumentOutOfRangeException(nameof(index)),
    };

    protected override Size MeasureOverride(Size availableSize)
    {
        _vBar.Measure(new Size(_vBar.Width, availableSize.Height));
        _hBar.Measure(new Size(availableSize.Width, _hBar.Height));
        return new Size(0, 0);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        var vBarW = _vBar.Visibility == Visibility.Visible ? _vBar.Width : 0;
        var hBarH = _hBar.Visibility == Visibility.Visible ? _hBar.Height : 0;
        _vBar.Arrange(new Rect(Math.Max(0, finalSize.Width - vBarW), 0, vBarW, Math.Max(0, finalSize.Height - hBarH)));
        _hBar.Arrange(new Rect(0, Math.Max(0, finalSize.Height - hBarH), Math.Max(0, finalSize.Width - vBarW), hBarH));

        UpdateScrollBars();
        InvalidateVisual();
        return finalSize;
    }

    protected override void OnRenderSizeChanged(SizeChangedInfo info)
    {
        base.OnRenderSizeChanged(info);
        UpdateScrollBars();
        InvalidateVisual();
        RaiseViewportChanged();
    }

    protected override void OnDpiChanged(DpiScale oldDpi, DpiScale newDpi)
    {
        base.OnDpiChanged(oldDpi, newDpi);
        RebuildTextRenderer();
        InvalidateVisual();
    }

    private void UpdateScrollBars()
    {
        var vBarW = 15.0;
        var hBarH = 15.0;

        var totalH = _model.TotalHeight;
        var totalW = _model.TotalWidth;

        var needV = totalH > ActualHeight;
        var needH = totalW > ActualWidth;

        // 一个方向出现滚动条会压缩另一个方向的可视区，最多迭代两次即可稳定
        for (var i = 0; i < 3; i++)
        {
            var viewportH = Math.Max(0, ActualHeight - (needH ? hBarH : 0));
            var viewportW = Math.Max(0, ActualWidth - (needV ? vBarW : 0));
            var newNeedV = totalH > viewportH;
            var newNeedH = totalW > viewportW;
            if (newNeedV == needV && newNeedH == needH)
            {
                break;
            }

            needV = newNeedV;
            needH = newNeedH;
        }

        _vBar.Visibility = needV ? Visibility.Visible : Visibility.Collapsed;
        _hBar.Visibility = needH ? Visibility.Visible : Visibility.Collapsed;

        if (!needV && !needH)
        {
            return;
        }

        var contentH = Math.Max(0, ActualHeight - (needH ? hBarH : 0));
        var contentW = Math.Max(0, ActualWidth - (needV ? vBarW : 0));

        _suspendScrollSync = true;
        _vBar.ViewportSize = Math.Max(1, contentH - _model.HeaderHeight);
        _vBar.Maximum = Math.Max(0, totalH - _model.HeaderHeight - _vBar.ViewportSize);
        _hBar.ViewportSize = Math.Max(1, contentW - _model.GutterWidth);
        _hBar.Maximum = Math.Max(0, totalW - _model.GutterWidth - _hBar.ViewportSize);

        _vOffset = Math.Clamp(_vOffset, 0, _vBar.Maximum);
        _hOffset = Math.Clamp(_hOffset, 0, _hBar.Maximum);
        _vBar.Value = _vOffset;
        _hBar.Value = _hOffset;
        _suspendScrollSync = false;
    }

    private void OnScrollChanged()
    {
        if (_suspendScrollSync)
        {
            return;
        }

        _vOffset = _vBar.Value;
        _hOffset = _hBar.Value;
        InvalidateVisual();
        RaiseViewportChanged();
    }

    // ---------------- 数据 ----------------

    public void SetSheet(Sheet? sheet, bool keepScroll)
    {
        _model.HeaderRowIndex = _headerRowIndex;
        _model.SetSheet(sheet);
        if (!keepScroll)
        {
            _vOffset = 0;
            _hOffset = 0;
            _activeRow = HeaderRowIndex;
            _activeCol = 0;
            _anchorRow = HeaderRowIndex;
            _anchorCol = 0;
        }

        _activeRowHitCols.Clear();
        _badgeCol = -1;
        _badgeText = string.Empty;
        UpdateScrollBars();
        InvalidateVisual();
        RaiseViewportChanged();
        RaiseActiveCellChanged();
    }

    public void ClearSearch() => Search = SearchResult.Empty;

    // ---------------- 绘制 ----------------

    protected override void OnRender(DrawingContext dc)
    {
        var sheet = _model.Sheet;
        var width = ActualWidth;
        var height = ActualHeight;

        dc.DrawRectangle(GridTheme.EmptyAreaBackground, null, new Rect(0, 0, width, height));
        if (sheet is null || sheet.RowCount == 0)
        {
            DrawEmptyMessage(dc, sheet is null ? "把 Excel / CSV 拖到这里，或按 Ctrl+O 打开" : "这个工作表没有内容");
            return;
        }

        _text ??= CreateTextRendererFallback();

        var contentTop = _model.HeaderHeight;
        var contentBottom = Math.Max(contentTop, height - (float)_hBar.Height);
        var contentLeft = _model.GutterWidth;
        var contentRight = Math.Max(contentLeft, width - (float)_vBar.Width);

        var rowH = _model.RowHeight;
        var firstRow = Math.Max(0, (int)(_vOffset / rowH));
        var lastRow = Math.Min(sheet.RowCount - 1, (int)((_vOffset + (contentBottom - contentTop)) / rowH) + 1);

        var firstCol = 0;
        for (var c = 0; c < sheet.ColCount; c++)
        {
            if (_model.ColumnRight(c) > _hOffset + contentLeft)
            {
                firstCol = c;
                break;
            }
        }

        var lastCol = sheet.ColCount - 1;
        for (var c = firstCol; c < sheet.ColCount; c++)
        {
            if (_model.ColumnLeft(c) > _hOffset + contentRight)
            {
                lastCol = Math.Max(firstCol, c - 1);
                break;
            }
        }

        DrawCells(dc, sheet, firstRow, lastRow, firstCol, lastCol, contentTop, contentBottom, contentLeft, contentRight);
        DrawColumnHeaders(dc, firstCol, lastCol, contentLeft, contentRight);
        DrawGutter(dc, firstRow, lastRow, contentTop, contentBottom);
        DrawCorner(dc);
        DrawBadge(dc);
    }

    private void DrawCells(
        DrawingContext dc,
        Sheet sheet,
        int firstRow,
        int lastRow,
        int firstCol,
        int lastCol,
        double contentTop,
        double contentBottom,
        double contentLeft,
        double contentRight)
    {
        var rowH = _model.RowHeight;
        var selTop = Math.Min(_anchorRow, _activeRow);
        var selBottom = Math.Max(_anchorRow, _activeRow);
        var selLeft = Math.Min(_anchorCol, _activeCol);
        var selRight = Math.Max(_anchorCol, _activeCol);

        dc.PushClip(new RectangleGeometry(new Rect(contentLeft, contentTop, Math.Max(0, contentRight - contentLeft), Math.Max(0, contentBottom - contentTop))));

        for (var r = firstRow; r <= lastRow; r++)
        {
            var y = contentTop + (r * rowH) - _vOffset;
            var isPinnedHeader = r == HeaderRowIndex && !_hideHeaderRow;
            var isDetectedHeader = r == _headerRowIndex;
            var isSearchRow = _search.RowIsHit(r);
            var inSelection = r >= selTop && r <= selBottom;

            Brush background;
            if (isPinnedHeader)
            {
                background = GridTheme.HeaderBackground;
            }
            else if (isDetectedHeader)
            {
                background = GridTheme.DetectedHeaderBackground;
            }
            else if (isSearchRow)
            {
                background = GridTheme.SearchRowBackground;
            }
            else if (inSelection)
            {
                background = GridTheme.SelectionBackground;
            }
            else
            {
                background = (r & 1) == 1 ? GridTheme.RowAltBackground : GridTheme.RowBackground;
            }

            dc.DrawRectangle(background, null, new Rect(contentLeft, y, Math.Max(0, contentRight - contentLeft), rowH));

            if (isSearchRow && !isPinnedHeader && _search.Options is not null)
            {
                var hits = r == _activeRow
                    ? _activeRowHitCols
                    : SearchService.HitColumnsInRow(sheet, _search, r, maxCols: 128);
                foreach (var hc in hits)
                {
                    if (hc < firstCol || hc > lastCol)
                    {
                        continue;
                    }

                    var hx = _model.ColumnLeft(hc) - _hOffset;
                    var hw = _model.GetColumnWidth(hc);
                    dc.DrawRectangle(GridTheme.MatchHighlight, null, new Rect(hx + 1, y + 1, Math.Max(1, hw - 1), rowH - 1));
                }
            }

            for (var c = firstCol; c <= lastCol; c++)
            {
                var value = sheet.Get(r, c);
                if (value.Length == 0)
                {
                    continue;
                }

                var x = _model.ColumnLeft(c) - _hOffset;
                var w = _model.GetColumnWidth(c);

                if (isPinnedHeader)
                {
                    _text!.Draw(dc, value, CellTextRenderer.AlignCenter, x + 2, y, w - 4, rowH);
                }
                else if (isDetectedHeader)
                {
                    // 字段名行：加粗显示，长字段名仍然靠 TextFormatter 省略
                    _boldText!.Draw(dc, value, CellTextRenderer.AlignLeft, x + GridModel.CellPaddingX, y, w - (GridModel.CellPaddingX * 2), rowH);
                }
                else if (LooksNumeric(value))
                {
                    // 数字右对齐：配表里数值列很多，右对齐方便快速比对位数。
                    // 这里用半角单位估算文本宽度（数字等宽，误差很小），避免为对齐多做一次排版。
                    var units = DisplayWidth.Measure(value);
                    var tw = units * _estimatedDigitWidth;
                    var tx = Math.Max(x + 1, x + w - GridModel.CellPaddingX - tw);
                    _text!.Draw(dc, value, CellTextRenderer.AlignLeft, tx, y, w - GridModel.CellPaddingX, rowH);
                }
                else
                {
                    _text!.Draw(dc, value, CellTextRenderer.AlignLeft, x + GridModel.CellPaddingX, y, w - (GridModel.CellPaddingX * 2), rowH);
                }
            }
        }

        var linePen = new Pen(GridTheme.GridLine, 1);
        linePen.Freeze();
        for (var r = firstRow; r <= lastRow + 1; r++)
        {
            var y = Math.Round(contentTop + (r * rowH) - _vOffset) + 0.5;
            if (y < contentTop - 1 || y > contentBottom + 1)
            {
                continue;
            }

            dc.DrawLine(linePen, new Point(contentLeft, y), new Point(contentRight, y));
        }

        var lastLineCol = Math.Min(lastCol + 1, sheet.ColCount);
        for (var c = firstCol; c <= lastLineCol; c++)
        {
            var x = Math.Round(_model.ColumnLeft(c) - _hOffset) + 0.5;
            if (x < contentLeft - 1 || x > contentRight + 1)
            {
                continue;
            }

            dc.DrawLine(linePen, new Point(x, contentTop), new Point(x, contentBottom));
        }

        if (selRight >= selLeft && selBottom >= selTop)
        {
            var x0 = _model.ColumnLeft(selLeft) - _hOffset;
            var x1 = _model.ColumnRight(selRight) - _hOffset;
            var y0 = contentTop + (selTop * rowH) - _vOffset;
            var y1 = y0 + ((selBottom - selTop + 1) * rowH);
            var selPen = new Pen(GridTheme.SelectionBorder, 1.4);
            selPen.Freeze();
            dc.DrawRectangle(null, selPen, new Rect(x0, y0, Math.Max(1, x1 - x0), Math.Max(1, y1 - y0)));
        }

        // 当前搜索命中行：整行描边，配合状态栏“第 n / m 条”定位
        if (_search.HasHits && _activeRow != HeaderRowIndex && _search.RowIsHit(_activeRow))
        {
            var y0 = contentTop + (_activeRow * rowH) - _vOffset;
            var pen = new Pen(GridTheme.ActiveMatchBorder, 2);
            pen.Freeze();
            dc.DrawRectangle(null, pen, new Rect(contentLeft + 1, y0 + 1, Math.Max(1, contentRight - contentLeft - 2), rowH - 2));
        }

        dc.Pop();
    }

    private void DrawColumnHeaders(DrawingContext dc, int firstCol, int lastCol, double contentLeft, double contentRight)
    {
        var h = _model.HeaderHeight;
        dc.DrawRectangle(GridTheme.HeaderBackground, null, new Rect(0, 0, ActualWidth, h));

        dc.PushClip(new RectangleGeometry(new Rect(contentLeft, 0, Math.Max(0, contentRight - contentLeft), h)));
        var pen = new Pen(GridTheme.HeaderLine, 1);
        pen.Freeze();
        for (var c = firstCol; c <= lastCol; c++)
        {
            var x = _model.ColumnLeft(c) - _hOffset;
            var w = _model.GetColumnWidth(c);
            _text!.Draw(dc, ColumnName(c), CellTextRenderer.AlignCenter, x + 2, 0, w - 4, h);
            var lineX = Math.Round(x) + 0.5;
            if (lineX > contentLeft)
            {
                dc.DrawLine(pen, new Point(lineX, 4), new Point(lineX, h - 2));
            }
        }

        dc.Pop();

        var bottomPen = new Pen(GridTheme.HeaderLine, 1);
        bottomPen.Freeze();
        dc.DrawLine(bottomPen, new Point(0, h - 0.5), new Point(ActualWidth, h - 0.5));
    }

    private void DrawGutter(DrawingContext dc, int firstRow, int lastRow, double contentTop, double contentBottom)
    {
        if (!_model.ShowRowNumbers)
        {
            return;
        }

        var w = _model.GutterWidth;
        dc.DrawRectangle(GridTheme.GutterBackground, null, new Rect(0, contentTop, w, Math.Max(0, contentBottom - contentTop)));
        dc.PushClip(new RectangleGeometry(new Rect(0, contentTop, w, Math.Max(0, contentBottom - contentTop))));

        var rowH = _model.RowHeight;
        var selTop = Math.Min(_anchorRow, _activeRow);
        var selBottom = Math.Max(_anchorRow, _activeRow);
        for (var r = firstRow; r <= lastRow; r++)
        {
            var y = contentTop + (r * rowH) - _vOffset;
            if (r >= selTop && r <= selBottom)
            {
                dc.DrawRectangle(GridTheme.SelectionBackground, null, new Rect(0, y, w, rowH));
            }

            _text!.Draw(dc, r == HeaderRowIndex ? "H" : (r + 1).ToString(), CellTextRenderer.AlignCenter, 2, y, w - 6, rowH);
        }

        dc.Pop();

        var pen = new Pen(GridTheme.HeaderLine, 1);
        pen.Freeze();
        dc.DrawLine(pen, new Point(w - 0.5, 0), new Point(w - 0.5, ActualHeight));
    }

    private void DrawCorner(DrawingContext dc)
    {
        if (_model.ShowRowNumbers)
        {
            dc.DrawRectangle(GridTheme.HeaderBackground, null, new Rect(0, 0, _model.GutterWidth, _model.HeaderHeight));
        }
    }

    private void DrawBadge(DrawingContext dc)
    {
        if (_badgeCol < 0 || _badgeText.Length == 0)
        {
            return;
        }

        var x = _model.ColumnLeft(_badgeCol) - _hOffset;
        var w = _model.GetColumnWidth(_badgeCol);
        if (x + w < _model.GutterWidth || x > ActualWidth)
        {
            return;
        }

        var typeface = new Typeface(new FontFamily("Microsoft YaHei UI, Segoe UI"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
        var ft = new FormattedText(
            _badgeText,
            System.Globalization.CultureInfo.CurrentUICulture,
            FlowDirection.LeftToRight,
            typeface,
            12,
            Brushes.White,
            VisualTreeHelper.GetDpi(this).PixelsPerDip);

        var bw = ft.Width + 12;
        const double bh = 18;
        var bx = Math.Max(2, Math.Min(x + w, ActualWidth - bw - 2));
        var by = ActualHeight - _hBar.Height - bh - 8;
        var bg = new SolidColorBrush(Color.FromArgb(0xEE, 0x33, 0x33, 0x33));
        bg.Freeze();
        dc.DrawRoundedRectangle(bg, null, new Rect(bx, by, bw, bh), 4, 4);
        dc.DrawText(ft, new Point(bx + 6, by + ((bh - ft.Height) / 2)));
    }

    private void DrawEmptyMessage(DrawingContext dc, string message)
    {
        var typeface = new Typeface(new FontFamily("Microsoft YaHei UI, Segoe UI"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
        var ft = new FormattedText(
            message,
            System.Globalization.CultureInfo.CurrentUICulture,
            FlowDirection.LeftToRight,
            typeface,
            14,
            GridTheme.MutedText,
            VisualTreeHelper.GetDpi(this).PixelsPerDip);
        dc.DrawText(ft, new Point(Math.Max(0, (ActualWidth - ft.Width) / 2), Math.Max(0, (ActualHeight - ft.Height) / 2)));
    }

    /// <summary>字号/DPI 变化后若尚未重建，保证绘制器一定存在。</summary>
    private CellTextRenderer CreateTextRendererFallback()
    {
        RebuildTextRenderer();
        return _text!;
    }

    private void RebuildTextRenderer()
    {
        var dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        if (dpi <= 0)
        {
            dpi = 1.0;
        }

        var family = new FontFamily("Microsoft YaHei UI, Segoe UI");
        var regular = new Typeface(family, FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
        var bold = new Typeface(family, FontStyles.Normal, FontWeights.SemiBold, FontStretches.Normal);

        _text = new CellTextRenderer(regular, _fontSizePt, dpi, GridTheme.CellText);
        _boldText = new CellTextRenderer(bold, _fontSizePt, dpi, GridTheme.CellText);
        _estimatedDigitWidth = _fontSizePt * (7.2 / 13.0);

        // 列宽自适应复用同一个排版器，量过的宽度会被缓存
        GridModel.TextMeasurer = _text.Measure;
    }

    // ---------------- 鼠标交互 ----------------

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonDown(e);
        Focus();
        var p = e.GetPosition(this);

        if (p.Y <= _model.HeaderHeight)
        {
            var edge = HitColumnEdge(p.X, out _);
            if (edge >= 0)
            {
                _resizing = true;
                _resizeCol = edge;
                _resizeStartX = p.X;
                _resizeStartWidth = _model.GetColumnWidth(edge);
                CaptureMouse();
                e.Handled = true;
                return;
            }

            var headerCol = _model.ColumnAt(p.X + _hOffset);
            if (headerCol >= 0)
            {
                _activeRow = HeaderRowIndex;
                _anchorRow = HeaderRowIndex;
                _activeCol = headerCol;
                _anchorCol = headerCol;
                InvalidateVisual();
                RaiseActiveCellChanged();
            }

            return;
        }

        var row = _model.RowAt(p.Y - _model.HeaderHeight + _vOffset);
        var col = _model.ColumnAt(p.X + _hOffset);
        if (row < 0 || col < 0)
        {
            return;
        }

        if (e.ClickCount == 2)
        {
            AutoFitColumn(col);
            return;
        }

        _dragging = true;
        _anchorRow = row;
        _anchorCol = col;
        SetActiveCell(row, col, ensureVisible: false);
        CaptureMouse();
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        var p = e.GetPosition(this);

        if (_resizing && _resizeCol >= 0)
        {
            _model.SetColumnWidth(_resizeCol, _resizeStartWidth + (p.X - _resizeStartX));
            UpdateScrollBars();
            InvalidateVisual();
            return;
        }

        if (_dragging && e.LeftButton == MouseButtonState.Pressed)
        {
            var row = _model.RowAt(p.Y - _model.HeaderHeight + _vOffset);
            var col = _model.ColumnAt(p.X + _hOffset);
            if (row >= 0 && col >= 0 && (row != _activeRow || col != _activeCol))
            {
                _activeRow = row;
                _activeCol = col;
                RefreshActiveRowHits();
                InvalidateVisual();
                RaiseActiveCellChanged();
            }

            return;
        }

        if (p.Y <= _model.HeaderHeight)
        {
            Cursor = HitColumnEdge(p.X, out _) >= 0 ? Cursors.SizeWE : Cursors.Arrow;
        }
        else if (Cursor != Cursors.Arrow)
        {
            Cursor = Cursors.Arrow;
        }
    }

    protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonUp(e);
        if (_resizing)
        {
            _resizing = false;
            _resizeCol = -1;
            ReleaseMouseCapture();
        }

        if (_dragging)
        {
            _dragging = false;
            ReleaseMouseCapture();
        }
    }

    protected override void OnMouseWheel(MouseWheelEventArgs e)
    {
        base.OnMouseWheel(e);
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
        {
            ScrollHorizontal(-e.Delta * 0.5);
        }
        else
        {
            ScrollVertical(-e.Delta * 0.7);
        }

        e.Handled = true;
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        var shift = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift);
        var ctrl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
        var handled = true;

        switch (e.Key)
        {
            case Key.Up:
                MoveActive(-1, 0, shift);
                break;
            case Key.Down:
                MoveActive(1, 0, shift);
                break;
            case Key.Left:
                MoveActive(0, -1, shift);
                break;
            case Key.Right:
                MoveActive(0, 1, shift);
                break;
            case Key.PageUp:
                MoveActive(-VisibleRowCount, 0, shift);
                break;
            case Key.PageDown:
                MoveActive(VisibleRowCount, 0, shift);
                break;
            case Key.Home:
                if (ctrl)
                {
                    GoToCell(HeaderRowIndex, 0);
                }
                else
                {
                    SetActiveCell(_activeRow, 0, ensureVisible: true);
                }

                break;
            case Key.End:
                if (Sheet is not null && ctrl)
                {
                    GoToCell(Sheet.RowCount - 1, Sheet.ColCount - 1);
                }
                else if (Sheet is not null)
                {
                    SetActiveCell(_activeRow, Sheet.ColCount - 1, ensureVisible: true);
                }

                break;
            case Key.C when ctrl:
                CopySelectionToClipboard();
                break;
            default:
                handled = false;
                break;
        }

        if (handled)
        {
            e.Handled = true;
        }
    }

    private void MoveActive(int rowDelta, int colDelta, bool extendSelection)
    {
        var sheet = Sheet;
        if (sheet is null)
        {
            return;
        }

        var row = Math.Clamp(_activeRow + rowDelta, HeaderRowIndex, Math.Max(HeaderRowIndex, sheet.RowCount - 1));
        var col = Math.Clamp(_activeCol + colDelta, 0, Math.Max(0, sheet.ColCount - 1));
        if (!extendSelection)
        {
            _anchorRow = row;
            _anchorCol = col;
        }

        SetActiveCell(row, col, ensureVisible: true);
    }

    public void SetActiveCell(int row, int col, bool ensureVisible)
    {
        var changed = row != _activeRow || col != _activeCol;
        _activeRow = row;
        _activeCol = col;
        if (ensureVisible)
        {
            EnsureVisible(row, col);
        }

        if (changed)
        {
            RefreshActiveRowHits();
            InvalidateVisual();
            RaiseActiveCellChanged();
        }
    }

    /// <summary>把纵向滚动位置设到指定像素（打开文件时定位表头行用）。</summary>
    public void ScrollToRowOffset(double offset)
    {
        _vOffset = Math.Clamp(offset, 0, _vBar.Maximum);
        _vBar.Value = _vOffset;
        InvalidateVisual();
        RaiseViewportChanged();
    }

    /// <summary>显示一张"仅用于展示"的表（如改动清单）。</summary>
    public void ShowVirtualSheet(Sheet sheet)
    {
        _model.HeaderRowIndex = -1;
        SetSheet(sheet, keepScroll: false);
        DetectedHeaderRow = 0;
    }

    /// <summary>跳到指定单元格：滚动使其可见并选中（搜索结果导航用）。</summary>
    public void GoToCell(int row, int col)
    {
        var sheet = Sheet;
        if (sheet is null || sheet.RowCount == 0)
        {
            return;
        }

        row = Math.Clamp(row, HeaderRowIndex, sheet.RowCount - 1);
        col = Math.Clamp(col, 0, Math.Max(0, sheet.ColCount - 1));
        _anchorRow = row;
        _anchorCol = col;
        _activeRow = row;
        _activeCol = col;

        // 目标行落在视口上方 1/3 处，上下都留出上下文
        var viewportRows = VisibleRowCount;
        var desiredTop = Math.Max(0, row - HeaderRowIndex - (viewportRows / 3));
        _vOffset = Math.Min(desiredTop * _model.RowHeight, _vBar.Maximum);
        _vBar.Value = _vOffset;

        EnsureVisible(row, col);
        RefreshActiveRowHits();
        UpdateScrollBars();
        InvalidateVisual();
        RaiseViewportChanged();
        RaiseActiveCellChanged();
    }

    private void EnsureVisible(int row, int col)
    {
        var rowH = _model.RowHeight;
        var top = row * rowH;
        var bottom = top + rowH;
        var viewH = Math.Max(rowH, ActualHeight - _model.HeaderHeight - _hBar.Height);
        if (top < _vOffset)
        {
            _vOffset = top;
        }
        else if (bottom > _vOffset + viewH)
        {
            _vOffset = bottom - viewH;
        }

        _vOffset = Math.Clamp(_vOffset, 0, _vBar.Maximum);
        _vBar.Value = _vOffset;

        var left = _model.ColumnLeft(col);
        var right = _model.ColumnRight(col);
        var viewW = Math.Max(40, ActualWidth - _model.GutterWidth - _vBar.Width);
        var viewLeft = _hOffset + _model.GutterWidth;
        if (left < viewLeft)
        {
            _hOffset = Math.Max(0, left - _model.GutterWidth);
        }
        else if (right > viewLeft + viewW)
        {
            _hOffset = Math.Max(0, right - viewW - _model.GutterWidth + GridModel.CellPaddingX);
        }

        _hOffset = Math.Clamp(_hOffset, 0, _hBar.Maximum);
        _hBar.Value = _hOffset;
    }

    private void ScrollVertical(double delta)
    {
        _vOffset = Math.Clamp(_vOffset + delta, 0, _vBar.Maximum);
        _vBar.Value = _vOffset;
        InvalidateVisual();
        RaiseViewportChanged();
    }

    private void ScrollHorizontal(double delta)
    {
        _hOffset = Math.Clamp(_hOffset + delta, 0, _hBar.Maximum);
        _hBar.Value = _hOffset;
        InvalidateVisual();
    }

    private int HitColumnEdge(double x, out double edgeX)
    {
        edgeX = 0;
        var gridX = x + _hOffset;
        var count = Math.Max(1, _model.ColCount);
        for (var c = 0; c < count; c++)
        {
            var right = _model.ColumnRight(c);
            if (Math.Abs(gridX - right) <= 3)
            {
                edgeX = right - _hOffset;
                return c;
            }
        }

        return -1;
    }

    private void AutoFitColumn(int col)
    {
        var sheet = Sheet;
        if (sheet is null)
        {
            return;
        }

        var w = _model.MeasureColumn(col, 400, (r, c) => sheet.Get(r, c));
        _model.SetColumnWidth(col, w + 8);
        UpdateScrollBars();
        InvalidateVisual();
    }

    public void AutoFitAllColumns()
    {
        _model.AutoFitColumns();
        UpdateScrollBars();
        InvalidateVisual();
    }

    private void RefreshActiveRowHits()
    {
        if (_search.HasHits && _search.RowIsHit(_activeRow) && Sheet is not null)
        {
            _activeRowHitCols = SearchService.HitColumnsInRow(Sheet, _search, _activeRow, maxCols: 256);
        }
        else
        {
            _activeRowHitCols.Clear();
        }
    }

    public void ShowBadge(int col, string text)
    {
        _badgeCol = col;
        _badgeText = text;
        InvalidateVisual();
    }

    public void ClearBadge()
    {
        if (_badgeCol < 0 && _badgeText.Length == 0)
        {
            return;
        }

        _badgeCol = -1;
        _badgeText = string.Empty;
        InvalidateVisual();
    }

    public string GetActiveCellText()
    {
        var sheet = Sheet;
        return sheet is null ? string.Empty : sheet.Get(_activeRow, _activeCol);
    }

    private void CopySelectionToClipboard()
    {
        var sheet = Sheet;
        if (sheet is null)
        {
            return;
        }

        var r0 = Math.Min(_anchorRow, _activeRow);
        var r1 = Math.Max(_anchorRow, _activeRow);
        var c0 = Math.Min(_anchorCol, _activeCol);
        var c1 = Math.Max(_anchorCol, _activeCol);

        var sb = new StringBuilder();
        for (var r = r0; r <= r1; r++)
        {
            for (var c = c0; c <= c1; c++)
            {
                if (c > c0)
                {
                    sb.Append('\t');
                }

                sb.Append(sheet.Get(r, c));
            }

            if (r < r1)
            {
                sb.AppendLine();
            }
        }

        try
        {
            Clipboard.SetText(sb.ToString());
        }
        catch (Exception)
        {
            // 剪贴板偶发被其他进程占用；只读查看器静默跳过
        }
    }

    private static bool LooksNumeric(string s)
    {
        var seenDigit = false;
        foreach (var ch in s)
        {
            if (ch is >= '0' and <= '9')
            {
                seenDigit = true;
            }
            else if (ch is not ('.' or '-' or '+' or ',' or 'e' or 'E' or ' ' or '%'))
            {
                return false;
            }
        }

        return seenDigit;
    }

    /// <summary>0 -> A, 25 -> Z, 26 -> AA</summary>
    public static string ColumnName(int index)
    {
        if (index < 0)
        {
            return string.Empty;
        }

        Span<char> buf = stackalloc char[8];
        var i = buf.Length;
        index++;
        while (index > 0 && i > 0)
        {
            index--;
            buf[--i] = (char)('A' + (index % 26));
            index /= 26;
        }

        return new string(buf[i..]);
    }

    private void RaiseViewportChanged() => ViewportChanged?.Invoke(this, EventArgs.Empty);

    private void RaiseActiveCellChanged() => ActiveCellChanged?.Invoke(this, EventArgs.Empty);
}
