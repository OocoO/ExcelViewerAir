using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using ExcelViewer.Controls;
using ExcelViewer.Model;
using ExcelViewer.Search;
using Microsoft.Win32;

namespace ExcelViewer;

public partial class MainWindow
{
    private readonly DispatcherTimer _searchDebounce;
    private readonly StringPool _pool = new();

    private string? _filePath;
    private List<string> _sheetNames = new();
    private Sheet? _sheet;
    private int _sheetIndex = -1;
    private CancellationTokenSource? _loadCts;
    private CancellationTokenSource? _searchCts;
    private SearchResult _searchResult = SearchResult.Empty;
    private bool _hideHeaderRow;
    private long _openStopwatchTicks;
    private DiffView? _diffView;

    /// <summary>diff 视图里"写回"的目标（原始文件路径）。</summary>
    private string? _diffTargetPath;

    public MainWindow()
    {
        InitializeComponent();

        _diffView = new DiffView(
            this,
            GridHost,
            DiffBanner,
            DiffBannerText,
            DiffDetailText,
            SheetTabs,
            BtnDiffWriteBack,
            BtnDiffClose,
            OnWriteBackRequested);
        DiffBanner.Visibility = Visibility.Collapsed;

        _searchDebounce = new DispatcherTimer(DispatcherPriority.Background)
        {
            Interval = TimeSpan.FromMilliseconds(220),
        };
        _searchDebounce.Tick += (_, _) =>
        {
            _searchDebounce.Stop();
            RunSearch();
        };

        Grid.ViewportChanged += (_, _) => UpdateStatus();
        Grid.ActiveCellChanged += (_, _) =>
        {
            UpdateStatus();
            UpdatePreview();
        };

        // 拖分隔线调过高度后记下来：收起再展开不会跳回默认值
        PreviewSplitter.DragCompleted += OnPreviewSplitterDragCompleted;

        AppPreviewPane(true);
        UpdatePreview();

        SearchPlaceholder.Visibility = Visibility.Visible;
        BtnPrev.IsEnabled = false;
        BtnNext.IsEnabled = false;
        BtnClearSearch.IsEnabled = false;

        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        Grid.Focus();
        UpdateStatus();
    }

    // ================= 打开 / 加载 =================

    public async void OpenFileAsync(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            ShowBusyError($"找不到文件：{path}");
            return;
        }

        if (!WorkbookLoader.IsSupported(path))
        {
            ShowBusyError("这个文件类型暂时看不了，目前支持 xlsx / xlsm / xls / xlsb / csv。");
            return;
        }

        _loadCts?.Cancel();
        var cts = new CancellationTokenSource();
        _loadCts = cts;

        _openStopwatchTicks = Stopwatch.GetTimestamp();
        ShowBusy("正在打开…", Path.GetFileName(path));

        try
        {
            var names = await Task.Run(() => WorkbookLoader.ReadSheetNames(path), cts.Token);
            if (cts.IsCancellationRequested)
            {
                return;
            }

            _filePath = path;
            _sheetNames = names;
            _pool.Clear();
            SearchService.ClearCache();
            _sheet = null;
            _sheetIndex = -1;
            _searchResult = SearchResult.Empty;
            Grid.Search = SearchResult.Empty;
            SearchBox.Text = string.Empty;
            ClearSearchUi();

            BuildSheetTabs();
            Title = $"Excel 查看器 — {Path.GetFileName(path)}";

            await LoadSheetAsync(0, keepScroll: false);
        }
        catch (OperationCanceledException)
        {
            // 用户切了另一个文件，忽略
        }
        catch (Exception ex)
        {
            ShowBusyError("这个文件没能打开：" + ex.Message);
        }
    }

    private async Task LoadSheetAsync(int index, bool keepScroll)
    {
        var path = _filePath;
        if (path is null || index < 0 || index >= _sheetNames.Count)
        {
            return;
        }

        _loadCts?.Cancel();
        var cts = new CancellationTokenSource();
        _loadCts = cts;

        ShowBusy(
            _sheetNames.Count > 1 ? $"正在读取工作表「{_sheetNames[index]}」…" : "正在读取数据…",
            Path.GetFileName(path));

        try
        {
            var sheet = await Task.Run(
                () => WorkbookLoader.LoadSheet(path, index, _pool, cts.Token),
                cts.Token);
            if (cts.IsCancellationRequested)
            {
                return;
            }

            _sheet = sheet;
            _sheetIndex = index;
            _searchResult = SearchResult.Empty;
            Grid.Search = SearchResult.Empty;

            var header = HeaderDetection.Detect(sheet);
            Grid.DetectedHeaderRow = header.HeaderRowIndex;
            Grid.SetSheet(sheet, keepScroll);
            UpdateSheetTabSelection();
            HideBusy();

            // 冻结窗格的两个用途：
            //   1) 冻结块由 ExcelGrid 钉在顶部（表头一直在，往下翻也知道每列是什么）；
            //   2) 打开时直接看第一行数据，不用手动往下滚。
            // 没冻结的表沿用老办法：把 GDE_FIELD_NAMES 那一行滚到视口最上方。
            if (!keepScroll)
            {
                if (sheet.FreezeRows > 0)
                {
                    Grid.ScrollToRowOffset(0);
                    if (sheet.RowCount > sheet.FreezeRows)
                    {
                        Grid.SetActiveCell(sheet.FreezeRows, 0, ensureVisible: false);
                    }
                }
                else if (header.HeaderRowIndex > 0)
                {
                    var rowPx = Grid.Model.RowHeight;
                    Grid.ScrollToRowOffset(Math.Max(0, (header.HeaderRowIndex - 1) * rowPx));
                }
            }

            // 首屏就绪即记录耗时，用于对比 WPS/Excel 的打开速度
            var ms = Stopwatch.GetElapsedTime(_openStopwatchTicks).TotalMilliseconds;
            LoadTimeText.Text = $"打开耗时 {ms:F0} ms";
            FileInfoText.Text = $"{Path.GetFileName(path)} · {_sheetNames.Count} 个工作表";

            UpdateStatus();
            UpdatePreview();

            if (SearchBox.Text.Length > 0)
            {
                RunSearch();
            }
        }
        catch (OperationCanceledException)
        {
            // 切换工作表/文件导致的中断，属于正常流程
        }
        catch (Exception ex)
        {
            ShowBusyError("这个工作表没能读取：" + ex.Message);
        }
    }

    private void BuildSheetTabs()
    {
        SheetTabs.Children.Clear();
        if (_sheetNames.Count <= 1)
        {
            SheetTabBar.Visibility = Visibility.Collapsed;
            return;
        }

        SheetTabBar.Visibility = Visibility.Visible;
        for (var i = 0; i < _sheetNames.Count; i++)
        {
            var index = i;
            var tab = new ToggleButton
            {
                Content = $"{_sheetNames[i]}  ({index + 1})",
                Style = (Style)FindResource("SheetTab"),
                Tag = index,
                IsChecked = false,
            };
            tab.Click += async (_, _) =>
            {
                if (_sheetIndex == index)
                {
                    UpdateSheetTabSelection();
                    return;
                }

                await LoadSheetAsync(index, keepScroll: false);
            };
            SheetTabs.Children.Add(tab);
        }
    }

    private void UpdateSheetTabSelection()
    {
        foreach (var child in SheetTabs.Children)
        {
            if (child is ToggleButton tb && tb.Tag is int idx)
            {
                tb.IsChecked = idx == _sheetIndex;
            }
        }
    }

    // ================= 搜索 =================

    private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
    {
        SearchPlaceholder.Visibility = SearchBox.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;
        BtnClearSearch.IsEnabled = SearchBox.Text.Length > 0;

        _searchDebounce.Stop();
        if (_sheet is null)
        {
            return;
        }

        if (SearchBox.Text.Length == 0)
        {
            ClearSearchUi();
            _searchResult = SearchResult.Empty;
            Grid.Search = SearchResult.Empty;
            UpdateStatus();
            return;
        }

        _searchDebounce.Start();
    }

    private void RunSearch()
    {
        var sheet = _sheet;
        if (sheet is null)
        {
            return;
        }

        var query = SearchBox.Text;
        if (query.Length == 0)
        {
            return;
        }

        _searchCts?.Cancel();
        var cts = new CancellationTokenSource();
        _searchCts = cts;

        var options = new SearchOptions { Query = query };
        SearchSummary.Text = "搜索中…";

        // 搜索在后台线程跑，UI 不卡；同一查询有缓存，回车/上下条不会再扫一遍
        Task.Run(() => SearchService.Search(sheet, options, null, cts.Token), cts.Token)
            .ContinueWith(
                task =>
                {
                    if (task.IsCanceled || cts.IsCancellationRequested)
                    {
                        return;
                    }

                    if (task.IsFaulted)
                    {
                        SearchSummary.Text = "搜索没能完成，可以换个关键字再试";
                        return;
                    }

                    ApplySearchResult(task.Result);
                },
                TaskScheduler.FromCurrentSynchronizationContext());
    }

    private void ApplySearchResult(SearchResult result)
    {
        _searchResult = result;
        Grid.Search = result;

        if (result.IsEmpty)
        {
            SearchSummary.Text = string.Empty;
            UpdateSearchButtons();
            UpdateStatus();
            return;
        }

        if (!result.HasHits)
        {
            SearchSummary.Text = $"没有找到「{Trim(result.Query)}」（{result.ElapsedMs} ms）";
            UpdateSearchButtons();
            UpdateStatus();
            return;
        }

        SearchSummary.Text = result.Truncated
            ? $"命中 {result.RowHitCount} 行（{result.ElapsedMs} ms）"
            : $"命中 {result.RowHitCount} 行 / {result.CellHitCount} 个单元格（{result.ElapsedMs} ms）";

        UpdateSearchButtons();

        // 自动跳到第一个命中
        var first = result.NextHitRow(-1);
        if (first >= 0)
        {
            GoToHit(first);
        }
        else
        {
            UpdateStatus();
        }
    }

    private void UpdateSearchButtons()
    {
        var has = _searchResult.HasHits;
        BtnPrev.IsEnabled = has;
        BtnNext.IsEnabled = has;
    }

    private void ClearSearchUi()
    {
        SearchSummary.Text = string.Empty;
        UpdateSearchButtons();
    }

    private void OnSearchKeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Enter:
                _searchDebounce.Stop();
                if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
                {
                    GoPrevMatch();
                }
                else
                {
                    if (!_searchResult.HasHits || _searchResult.Query != SearchBox.Text)
                    {
                        RunSearch();
                    }
                    else
                    {
                        GoNextMatch();
                    }
                }

                e.Handled = true;
                break;
            case Key.Escape:
                OnClearSearch(sender, e);
                e.Handled = true;
                break;
            case Key.F3:
                GoNextMatch();
                e.Handled = true;
                break;
        }
    }

    private void OnNextMatch(object sender, RoutedEventArgs e) => GoNextMatch();

    private void OnPrevMatch(object sender, RoutedEventArgs e) => GoPrevMatch();

    private void GoNextMatch()
    {
        if (!_searchResult.HasHits)
        {
            return;
        }

        var next = _searchResult.NextHitRow(Grid.ActiveRow);
        if (next < 0)
        {
            // 走到末尾就回到第一条，避免用户以为"没了"
            next = _searchResult.NextHitRow(-1);
        }

        GoToHit(next);
    }

    private void GoPrevMatch()
    {
        if (!_searchResult.HasHits)
        {
            return;
        }

        var prev = _searchResult.PrevHitRow(Grid.ActiveRow);
        if (prev < 0)
        {
            prev = _searchResult.HitRowAt(_searchResult.RowHitCount);
        }

        GoToHit(prev);
    }

    private void GoToHit(int row)
    {
        var sheet = _sheet;
        if (sheet is null || row < 0)
        {
            return;
        }

        var col = SearchService.FirstHitColumn(sheet, _searchResult, row);
        Grid.GoToCell(row, Math.Max(0, col));
        UpdateStatus();
    }

    private void OnClearSearch(object sender, RoutedEventArgs e)
    {
        _searchDebounce.Stop();
        _searchCts?.Cancel();
        SearchBox.Text = string.Empty;
        SearchPlaceholder.Visibility = Visibility.Visible;
        _searchResult = SearchResult.Empty;
        Grid.Search = SearchResult.Empty;
        ClearSearchUi();
        UpdateStatus();
        Grid.Focus();
    }

    // ================= 文件操作 =================

    private async void OnOpenClick(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "选择要查看的表",
            Filter = "配表文件|*.xlsx;*.xlsm;*.xlsb;*.xls;*.csv|Excel 工作簿|*.xlsx;*.xlsm;*.xlsb;*.xls|CSV|*.csv|所有文件|*.*",
            CheckFileExists = true,
        };

        if (dlg.ShowDialog(this) == true)
        {
            await OpenFileAndWaitAsync(dlg.FileName);
        }
    }

    private async Task OpenFileAndWaitAsync(string path)
    {
        OpenFileAsync(path);
        await Task.CompletedTask;
    }

    private async void OnReloadClick(object sender, RoutedEventArgs e)
    {
        if (_filePath is null)
        {
            return;
        }

        _pool.Clear();
        SearchService.ClearCache();
        var keepIndex = _sheetIndex;
        await LoadSheetAsync(keepIndex < 0 ? 0 : keepIndex, keepScroll: false);
    }

    private void OnCopyCellClick(object sender, RoutedEventArgs e)
    {
        var text = Grid.GetActiveCellText();
        if (text.Length == 0)
        {
            return;
        }

        try
        {
            Clipboard.SetText(text);
        }
        catch (Exception)
        {
            // 剪贴板被占用时忽略
        }
    }

    private void OnExitClick(object sender, RoutedEventArgs e) => Close();

    private void OnEditCopyClick(object sender, RoutedEventArgs e)
    {
        if (_filePath is null)
        {
            OnOpenClick(sender, e);
            return;
        }

        StartEditCopy(_filePath);
    }

    /// <summary>手动选择另一个文件（通常是编辑后的副本）来和当前文件比对。</summary>
    private void OnDiffWithClick(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title = "选择要和当前文件比对的文件",
            Filter = "Excel 工作簿|*.xlsx;*.xlsm;*.xlsb;*.xls|CSV|*.csv|所有文件|*.*",
            CheckFileExists = true,
        };

        if (dlg.ShowDialog(this) != true)
        {
            return;
        }

        var oldPath = _filePath ?? dlg.FileName;
        ShowDiffAsync(oldPath, dlg.FileName);
    }

    private void OnWindowClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        _loadCts?.Cancel();
        _searchCts?.Cancel();
    }

    // ================= 拖放 =================

    private void OnWindowDragOver(object sender, DragEventArgs e)
    {
        var ok = e.Data.GetDataPresent(DataFormats.FileDrop);
        e.Effects = ok ? DragDropEffects.Copy : DragDropEffects.None;
        DropHint.Visibility = ok ? Visibility.Visible : Visibility.Collapsed;
        e.Handled = true;
    }

    private void OnWindowDrop(object sender, DragEventArgs e)
    {
        DropHint.Visibility = Visibility.Collapsed;
        if (!e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            return;
        }

        if (e.Data.GetData(DataFormats.FileDrop) is not string[] files || files.Length == 0)
        {
            return;
        }

        var file = files.FirstOrDefault(WorkbookLoader.IsSupported) ?? files[0];
        OpenFileAsync(file);
        e.Handled = true;
    }

    // ================= 视图 =================

    private void OnToggleRowNumbers(object sender, RoutedEventArgs e)
    {
        Grid.ShowRowNumbers = MenuRowNumbers.IsChecked;
        UpdateStatus();
    }

    private void OnToggleHeaderRow(object sender, RoutedEventArgs e)
    {
        _hideHeaderRow = !MenuShowHeaderRow.IsChecked;
        Grid.HideHeaderRow = _hideHeaderRow;
        Grid.KeepActiveCellVisible();
        UpdateStatus();
    }

    private void OnZoomIn(object sender, RoutedEventArgs e) => Grid.FontSizePt += 1;

    private void OnZoomOut(object sender, RoutedEventArgs e) => Grid.FontSizePt -= 1;

    private void OnZoomReset(object sender, RoutedEventArgs e) => Grid.FontSizePt = 13;

    private void OnAutoFitColumns(object sender, RoutedEventArgs e) => Grid.AutoFitAllColumns();

    // ================= 状态栏 =================

    private void UpdateStatus()
    {
        var sheet = _sheet;
        if (sheet is null)
        {
            StatusCell.Text = "—";
            StatusSize.Text = string.Empty;
            StatusRow.Text = string.Empty;
            StatusFilter.Text = string.Empty;
            return;
        }

        var row = Grid.ActiveRow;
        var col = Grid.ActiveCol;
        StatusCell.Text = $"{ExcelGrid.ColumnName(col)}{row + 1}";

        StatusSize.Text = $"{sheet.RowCount:N0} 行 × {sheet.ColCount} 列";
        if (Grid.FrozenRowCount > 0)
        {
            var frozen = $"表头第 1–{Grid.FrozenRowCount} 行（已冻结）";
            if (Grid.FrozenColCount > 0)
            {
                frozen += $" · 前 {Grid.FrozenColCount} 列冻结";
            }

            StatusSize.Text += " · " + frozen;
        }
        else if (Grid.DetectedHeaderRow > 0)
        {
            StatusSize.Text += $" · 表头在第 {Grid.DetectedHeaderRow + 1} 行";
        }

        if (sheet.SkippedTrailingRows > 0 || sheet.SkippedTrailingCols > 0)
        {
            StatusSize.Text += $"（已省略尾部 {sheet.SkippedTrailingRows:N0} 行空白）";
        }

        var top = Grid.TopVisibleRow;
        StatusRow.Text = $"当前显示第 {top + 1} – {Math.Min(sheet.RowCount, top + Grid.VisibleRowCount):N0} 行";

        if (_searchResult.HasHits)
        {
            var ordinal = _searchResult.OrdinalOfRow(Grid.ActiveRow);
            StatusFilter.Text = ordinal > 0
                ? $"搜索命中：第 {ordinal:N0} / {_searchResult.RowHitCount:N0} 条"
                : $"搜索命中 {_searchResult.RowHitCount:N0} 行";
        }
        else if (_searchResult.Options is not null && !_searchResult.HasHits)
        {
            StatusFilter.Text = "无搜索结果";
        }
        else
        {
            StatusFilter.Text = string.Empty;
        }

        var mem = GC.GetTotalMemory(false) / 1048576.0;
        StatusHint.Text = $"只读 · 内存 {mem:F0} MB";
    }

    // ================= 内容预览条 =================

    /// <summary>预览不出来的超长文本阈值：再长就只给个提示，避免一次排版几十万字符卡住界面。</summary>
    private const int PreviewLimit = 200_000;

    /// <summary>收起时这一行的高度：够放一行地址 + 摘要 + 「展开」，不会把表头压成半截。</summary>
    private const double PreviewCollapsedHeight = 26;

    /// <summary>预览条最少留的高度（再小就只剩半行字了）。</summary>
    private const double PreviewMinHeight = 72;

    /// <summary>
    /// 表格区的最小高度：预览条再高也不能把主表格挤没。
    /// 少了这个约束，窗口一矮（例如 900x360）预览条会把列标和数据行整块顶掉，
    /// 看起来就像"展开详情把上面的内容盖住了"。
    /// </summary>
    private const double GridMinHeight = 104;

    /// <summary>内容预览条默认就是展开的：选中格子直接看到完整内容，不用先按 Ctrl+P。</summary>
    private bool _previewPaneOpen = true;

    private double _previewHeight = 190;

    /// <summary>
    /// 展开/收起预览条。
    /// 展开时恢复上次拖出来的高度（但不能挤掉表格的最小高度），
    /// 收起时只留一行摘要（点一下就能再展开），不会再出现"被压成半截还拖不开"的状态。
    /// </summary>
    private void AppPreviewPane(bool open)
    {
        if (!open)
        {
            RememberPreviewHeight();
        }

        _previewPaneOpen = open;
        MenuPreviewPane.IsChecked = open;

        PreviewHeader.Visibility = open ? Visibility.Visible : Visibility.Collapsed;
        PreviewTextBox.Visibility = open ? Visibility.Visible : Visibility.Collapsed;
        PreviewCollapsedBar.Visibility = open ? Visibility.Collapsed : Visibility.Visible;
        // 收起时不留分隔线：整条就只剩那一行摘要
        PreviewSplitter.Visibility = open ? Visibility.Visible : Visibility.Collapsed;

        ApplyPreviewHeight();
        UpdatePreview();
    }

    /// <summary>
    /// 按当前窗口高度决定预览条行高：上限 = 可用高度 - 表格最小高度。
    /// 窗口被拉矮时预览条自动跟着收，放宽后再回到用户拖出来的高度。
    /// </summary>
    private void ApplyPreviewHeight()
    {
        // 收起时那一行只有 26px，最小高度必须跟着放开，否则收起状态下会被顶成 72
        PreviewRow.MinHeight = _previewPaneOpen ? PreviewMinHeight : 0;

        var available = ContentGrid.ActualHeight;
        if (available > 0)
        {
            var max = Math.Max(PreviewMinHeight, available - PreviewSplitter.Height - GridMinHeight);
            PreviewRow.MaxHeight = max;
        }

        var target = _previewPaneOpen
            ? Math.Clamp(_previewHeight, PreviewMinHeight, Math.Max(PreviewMinHeight, PreviewRow.MaxHeight))
            : PreviewCollapsedHeight;
        PreviewRow.Height = new GridLength(target);
    }

    /// <summary>把用户拖出来的高度记下来（不夹），收起再展开 / 窗口变大时按这个值恢复。</summary>
    private void RememberPreviewHeight()
    {
        var h = PreviewRow.Height;
        if (h.IsAbsolute && h.Value >= PreviewMinHeight)
        {
            _previewHeight = h.Value;
        }
    }

    /// <summary>内容区尺寸变了（窗口缩放 / 拖动分隔线）：重新夹一次预览条高度，并让选中格子保持可见。</summary>
    private void OnContentGridSizeChanged(object sender, SizeChangedEventArgs e)
    {
        ApplyPreviewHeight();

        // 布局还没走完时 ActualHeight 是旧值，等这一轮排完再校正可见性
        Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() => Grid.KeepActiveCellVisible()));
    }

    private void OnPreviewSplitterDragCompleted(object sender, DragCompletedEventArgs e)
    {
        RememberPreviewHeight();
        Grid.KeepActiveCellVisible();
    }

    private void OnTogglePreviewPane(object sender, RoutedEventArgs e) =>
        AppPreviewPane(MenuPreviewPane.IsChecked);

    private void OnPreviewClose(object sender, RoutedEventArgs e) => AppPreviewPane(false);

    private void OnPreviewExpand(object sender, RoutedEventArgs e)
    {
        AppPreviewPane(true);
        // 「展开 ▴」只负责展开，不让这次点击再冒泡去触发整条摘要的点击
        e.Handled = true;
    }

    private void OnPreviewCollapsedClick(object sender, MouseButtonEventArgs e) => AppPreviewPane(true);

    private void OnPreviewCopy(object sender, RoutedEventArgs e)
    {
        var sheet = _sheet;
        if (sheet is null)
        {
            return;
        }

        var text = sheet.Get(Grid.ActiveRow, Grid.ActiveCol);
        if (text.Length == 0)
        {
            return;
        }

        try
        {
            Clipboard.SetText(text);
        }
        catch (Exception)
        {
            // 剪贴板被别的程序占用时静默跳过，只读查看器不为此弹窗
        }
    }

    /// <summary>
    /// 把当前选中格子的完整内容填进预览条。
    /// 表格里格子宽度有限，长文案只能截断显示，这里给一份原文（自动换行 + 可滚动）。
    /// 收起状态那一行摘要也要跟着走，否则收起后看到的还是上一个格子。
    /// </summary>
    private void UpdatePreview()
    {
        var sheet = _sheet;

        if (sheet is null)
        {
            PreviewCollapsedAddress.Text = "—";
            PreviewCollapsedText.Text = "内容预览条已收起 — 点这里或按 Ctrl+P 展开看全文";
        }
        else
        {
            var collapsedRow = Grid.ActiveRow;
            var collapsedCol = Grid.ActiveCol;
            var collapsedText = sheet.Get(collapsedRow, collapsedCol);
            PreviewCollapsedAddress.Text = $"{ExcelGrid.ColumnName(collapsedCol)}{collapsedRow + 1}";
            PreviewCollapsedText.Text = collapsedText.Length == 0
                ? "（空单元格）"
                : OneLineSummary(collapsedText);
        }

        if (!_previewPaneOpen)
        {
            return;
        }

        if (sheet is null)
        {
            PreviewAddress.Text = "—";
            PreviewMeta.Text = string.Empty;
            PreviewWarn.Visibility = Visibility.Collapsed;
            PreviewTextBox.Text = string.Empty;
            return;
        }

        var row = Grid.ActiveRow;
        var col = Grid.ActiveCol;
        var text = sheet.Get(row, col);

        PreviewAddress.Text = $"{ExcelGrid.ColumnName(col)}{row + 1}";
        PreviewMeta.Text = text.Length == 0
            ? "（空单元格）"
            : $"{text.Length:N0} 字符 · {LineCount(text):N0} 行 · 第 {row + 1} 行 {ExcelGrid.ColumnName(col)} 列";

        if (text.Length > PreviewLimit)
        {
            PreviewWarn.Text = $"内容较长，这里只显示前 {PreviewLimit / 1000} 千字符（用「复制」取全文）";
            PreviewWarn.Visibility = Visibility.Visible;
            PreviewTextBox.Text = text[..PreviewLimit];
            return;
        }

        PreviewWarn.Visibility = Visibility.Collapsed;
        PreviewTextBox.Text = text;
    }

    /// <summary>收起状态那一行只取开头一段，太长的话设置文本本身也会变慢。</summary>
    private static string OneLineSummary(string text)
    {
        const int max = 300;
        var flat = DisplayWidth.Flatten(text);
        return flat.Length <= max ? flat : flat[..max] + "…";
    }


    private static int LineCount(string s)
    {
        var lines = 1;
        foreach (var ch in s)
        {
            if (ch == '\n')
            {
                lines++;
            }
        }

        return lines;
    }

    // ================= 改动比对 / 写回 =================

    /// <summary>
    /// 「用副本编辑」：把原文件复制到工作区 → 用 Excel/WPS 打开副本 → 等它关闭 → 自动进改动比对。
    /// 原文件全程只被读一次，不会被编辑器锁定，也不会被直接改写。
    /// </summary>
    public async void StartEditCopy(string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            ShowBusyError($"找不到文件：{sourcePath}");
            return;
        }

        var editor = EditSession.FindEditor();
        if (editor is null)
        {
            ShowBusyError("没有找到 Excel / WPS，先用只读方式打开这个文件。");
            OpenFileAsync(sourcePath);
            return;
        }

        string copy;
        try
        {
            copy = TempWorkspace.CreateCopy(sourcePath);
        }
        catch (Exception ex)
        {
            ShowBusyError("复制到工作区没能完成：" + ex.Message);
            return;
        }

        // 先把原文件只读显示出来，用户编辑时也有个参照
        OpenFileAsync(sourcePath);
        _editingPath = sourcePath;
        _editingCopy = copy;

        ShowBusy("已用副本打开编辑器…", $"编辑的是副本：{copy}\n保存后在 Excel 里关闭它，这里会自动弹出改动比对。");

        try
        {
            var proc = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(editor, $"\"{copy}\"")
            {
                UseShellExecute = false,
            });

            if (proc is null)
            {
                HideBusy();
                return;
            }

            // 等编辑器退出（不阻塞 UI），退出后再等文件写稳定再比对
            await Task.Run(() =>
            {
                proc.WaitForExit();
                WaitForFileSettled(copy, TimeSpan.FromSeconds(15));
            });

            HideBusy();
            ShowDiffAsync(sourcePath, copy);
        }
        catch (Exception ex)
        {
            HideBusy();
            ShowBusyError("启动编辑器没能完成：" + ex.Message);
        }
    }

    private string? _editingPath;
    private string? _editingCopy;

    /// <summary>等到文件大小与时间戳稳定（避免读到正在写入的半成品）。</summary>
    private static void WaitForFileSettled(string path, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        long lastSize = -1;
        DateTime lastWrite = DateTime.MinValue;
        var stable = 0;

        while (DateTime.UtcNow < deadline)
        {
            try
            {
                var info = new FileInfo(path);
                if (info.Exists && info.Length == lastSize && info.LastWriteTimeUtc == lastWrite)
                {
                    if (++stable >= 2)
                    {
                        return;
                    }
                }
                else
                {
                    stable = 0;
                    lastSize = info.Exists ? info.Length : -1;
                    lastWrite = info.Exists ? info.LastWriteTimeUtc : DateTime.MinValue;
                }
            }
            catch (IOException)
            {
                stable = 0;
            }

            Thread.Sleep(400);
        }
    }


    /// <summary>
    /// 打开 diff 视图：比对原文件与（temp 副本）新文件，列出改动的行与格子。
    /// <paramref name="oldLabel"/> / <paramref name="newLabel"/> 是给界面看的两侧名字
    /// （SVN 模式下是「文件 (revision N)」，比临时文件路径好认得多）。
    /// </summary>
    public async void ShowDiffAsync(string oldPath, string newPath, string? oldLabel = null, string? newLabel = null)
    {
        if (!File.Exists(oldPath) || !File.Exists(newPath))
        {
            ShowBusyError("要对比的文件找不到了，可能已被移动或删除。");
            return;
        }

        _diffTargetPath = oldPath;
        ShowBusy("正在比对改动…", $"{Path.GetFileName(oldPath)}  ←→  {Path.GetFileName(newPath)}");

        try
        {
            var report = await Task.Run(() => DiffService.Compare(oldPath, newPath));
            HideBusy();
            // 比对表是另一张表（盖在主表格上面），预览条在这里只会显示主表里过期的格子，
            // 所以先收起并换成一句说明；关掉对比后会自动恢复成展开状态。
            AppPreviewPane(false);
            PreviewCollapsedAddress.Text = string.Empty;
            PreviewCollapsedText.Text = "改动比对中 — 关闭对比后内容预览条会自动恢复";
            _diffView?.Show(report, oldLabel ?? oldPath, newLabel ?? newPath);
            Title = $"Excel 查看器 — 改动比对：{Path.GetFileName(newLabel ?? newPath)}";
        }
        catch (Exception ex)
        {
            ShowBusyError("比对没能完成：" + ex.Message);
        }
    }

    private void OnDiffWriteBack(object sender, RoutedEventArgs e) => _diffView?.RequestWriteBack();

    /// <summary>用户点了「写回原文件」：再确认一次，然后覆盖。</summary>
    private void OnWriteBackRequested(string newPath)
    {
        var target = _diffTargetPath;
        if (target is null)
        {
            return;
        }

        var answer = MessageBox.Show(
            this,
            $"会用编辑后的内容覆盖原文件：\n\n{target}\n\n原文件不会自动备份；如果拿不准，先选「另存为」。\n确认写回吗？",
            "写回原文件",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Warning,
            MessageBoxResult.Cancel);

        if (answer != MessageBoxResult.OK)
        {
            return;
        }

        try
        {
            // 覆盖前把原文件备份到工作目录，留一条后路
            var backupDir = TempWorkspace.BackupDir;
            var backup = Path.Combine(
                backupDir,
                $"{Path.GetFileNameWithoutExtension(target)}.{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(target)}");
            File.Copy(target, backup, overwrite: false);

            // 同名文件被 Excel 占用时用 File.Replace 更稳；失败则退回覆盖拷贝
            try
            {
                File.Copy(newPath, target, overwrite: true);
            }
            catch (IOException)
            {
                File.Replace(newPath, target, null);
            }

            MessageBox.Show(
                this,
                $"已写回原文件。\n\n覆盖前的副本备份在：\n{backup}",
                "写回完成",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            _diffView?.Hide();
            AppPreviewPane(true);
            OpenFileAsync(target);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                this,
                "写回没能完成：" + ex.Message + "\n\n原文件尚未被修改，可以用「另存为」保存到别处。",
                "写回失败",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void OnDiffSaveAs(object sender, RoutedEventArgs e)
    {
        var source = _diffView?.NewPath;
        if (source is null || !File.Exists(source))
        {
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title = "另存编辑后的文件",
            FileName = Path.GetFileName(_diffTargetPath ?? source),
            Filter = "Excel 工作簿|*.xlsx|所有文件|*.*",
        };

        if (dlg.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            File.Copy(source, dlg.FileName, overwrite: true);
            MessageBox.Show(this, "已另存到：\n" + dlg.FileName, "另存为", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "另存没能完成：" + ex.Message, "另存为", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void OnDiffClose(object sender, RoutedEventArgs e)
    {
        _diffView?.Hide();
        AppPreviewPane(true);
        if (_filePath is not null)
        {
            Title = $"Excel 查看器 — {Path.GetFileName(_filePath)}";
            OpenFileAsync(_filePath);
        }
    }

    // ================= 加载遮罩 =================

    private void ShowBusy(string text, string sub)
    {
        BusyText.Text = text;
        BusySubText.Text = sub;
        BusyCloseButton.Visibility = Visibility.Collapsed;
        // 加载中不拦截鼠标：用户仍可滚动已有内容
        BusyOverlay.IsHitTestVisible = false;
        BusyOverlay.Visibility = Visibility.Visible;
    }

    private void HideBusy() => BusyOverlay.Visibility = Visibility.Collapsed;

    private void ShowBusyError(string message)
    {
        BusyText.Text = message;
        BusySubText.Text = "关掉提示后可以继续打开其他文件。";
        BusyCloseButton.Visibility = Visibility.Visible;
        BusyOverlay.IsHitTestVisible = true;
        BusyOverlay.Visibility = Visibility.Visible;
    }

    private void OnBusyClose(object sender, RoutedEventArgs e) => HideBusy();

    // ================= 帮助 =================

    private void OnShowShortcuts(object sender, RoutedEventArgs e)
    {
        const string text =
            "快捷键\n\n" +
            "Ctrl+O        打开文件\n" +
            "F5            重新加载当前文件\n" +
            "Ctrl+F        定位到搜索框\n" +
            "F3 / 回车     下一个命中\n" +
            "Shift+F3      上一个命中\n" +
            "Tab           切换工作表\n" +
            "方向键        移动单元格\n" +
            "PageUp/Down   上下翻页\n" +
            "Ctrl+Home/End 跳到表格开头 / 末尾\n" +
            "Ctrl+C        复制选中单元格\n" +
            "Ctrl+ +/-     放大 / 缩小字号\n" +
            "Ctrl+P        收起 / 展开内容预览条\n" +
            "Esc           清除搜索\n\n" +
            "底部的内容预览条默认展开：选中格子就显示它的完整内容（长文本自动换行 + 可滚动）。\n" +
            "分隔线可以上下拖动改高度；想让出空间就按 Ctrl+P 收起，再按一次即可恢复。";

        MessageBox.Show(this, text, "快捷键说明", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OnShowAbout(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            this,
            "Excel 只读查看器\n\n" +
            "专为大配表快速浏览设计：\n" +
            "· 首屏只解析当前工作表，几万行也是秒开\n" +
            "· 自绘虚拟化表格，滚动到第几万行都不掉帧\n" +
            "· 全表关键字搜索，命中行整行高亮\n\n" +
            "只读打开，不会修改任何原始文件。",
            "关于",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    // ================= 其它 =================

    protected override void OnPreviewKeyDown(KeyEventArgs e)
    {
        base.OnPreviewKeyDown(e);
        var ctrl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);

        if (e.Key == Key.O && ctrl)
        {
            OnOpenClick(this, new RoutedEventArgs());
            e.Handled = true;
        }
        else if (e.Key == Key.F && ctrl)
        {
            SearchBox.Focus();
            SearchBox.SelectAll();
            e.Handled = true;
        }
        else if (e.Key == Key.F3)
        {
            if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
            {
                GoPrevMatch();
            }
            else
            {
                GoNextMatch();
            }

            e.Handled = true;
        }
        else if (e.Key == Key.F5)
        {
            OnReloadClick(this, new RoutedEventArgs());
            e.Handled = true;
        }
        else if (e.Key == Key.P && ctrl)
        {
            AppPreviewPane(!_previewPaneOpen);
            e.Handled = true;
        }
        else if (e.Key == Key.Escape && SearchBox.IsKeyboardFocusWithin)
        {
            OnClearSearch(this, new RoutedEventArgs());
            e.Handled = true;
        }
        else if (e.Key == Key.Tab && SheetTabs.Children.Count > 1 && !SearchBox.IsKeyboardFocusWithin)
        {
            var next = (Grid.ActiveRow >= 0 ? _sheetIndex + 1 : 0) % SheetTabs.Children.Count;
            if (next < 0)
            {
                next += SheetTabs.Children.Count;
            }

            _ = LoadSheetAsync(next, keepScroll: false);
            e.Handled = true;
        }
        else if (e.Key is Key.OemPlus or Key.Add && ctrl)
        {
            Grid.FontSizePt += 1;
            e.Handled = true;
        }
        else if (e.Key is Key.OemMinus or Key.Subtract && ctrl)
        {
            Grid.FontSizePt -= 1;
            e.Handled = true;
        }
    }

    private static string Trim(string s) => s.Length <= 24 ? s : s[..24] + "…";
}

