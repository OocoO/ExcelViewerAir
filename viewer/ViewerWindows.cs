namespace ExcelViewer;

/// <summary>
/// 「去重」规则眼里的一扇窗口。把 WPF 的东西挡在外面，规则本身就能在
/// <c>--dedup-selftest</c> 里用假窗口跑一遍，不必真的开界面。
/// </summary>
internal interface IViewerWindow
{
    /// <summary>这个窗口当前代表哪个文件（<see cref="FileKey"/> 归一后的路径）；还没打开过文件时是 null。</summary>
    string? DocumentKey { get; }

    /// <summary>正在显示改动比对（一张盖在表格上的清单表），不是普通的表格视图。</summary>
    bool IsShowingDiff { get; }

    /// <summary>只读打开一个文件（同一个窗口可以换文件）。</summary>
    void OpenPath(string path);

    /// <summary>「用副本编辑」：复制到工作区 → 用 Excel 打开副本 → 关掉后弹改动比对。</summary>
    void StartEdit(string path);

    /// <summary>提到前台（最小化时先还原）。</summary>
    void BringToFront();
}

/// <summary>
/// 记着当前进程里所有查看器窗口，并把一次「打开」请求落到正确的窗口上：
///   * 同一个表已经开着 → 只把那个窗口提到前台，不重新加载、不开新窗口；
///   * 还没打开过文件 → 先用那个空窗口（启动后直接拖文件进来时不必多一个空壳）；
///   * 否则开新窗口，一个表一个窗口。
///
/// 顺序也在这里维护：最后一个元素＝最近用过的窗口，不带文件启动时定位到的就是它。
/// </summary>
internal sealed class ViewerWindows
{
    private readonly List<IViewerWindow> _windows = new();

    public int Count => _windows.Count;

    /// <summary>最近用过的窗口（一个都没有时 null）。</summary>
    public IViewerWindow? MostRecent => _windows.Count == 0 ? null : _windows[^1];

    public void Add(IViewerWindow window) => _windows.Add(window);

    public void Remove(IViewerWindow window) => _windows.Remove(window);

    /// <summary>
    /// 已经在显示这个文件的窗口。同时有「表格视图」和「改动比对」两种窗口时优先表格视图 ——
    /// 比对视图只是盖在表格上的一层，把用户丢到比对上会让人以为表没打开。
    /// </summary>
    public IViewerWindow? FindByDocument(string key)
    {
        if (key.Length == 0)
        {
            return null;
        }

        var same = _windows
            .Where(w => w.DocumentKey is { } owned && string.Equals(owned, key, StringComparison.OrdinalIgnoreCase))
            .ToList();

        return same.FirstOrDefault(w => !w.IsShowingDiff) ?? same.FirstOrDefault();
    }

    /// <summary>还没打开过文件的空窗口（正在显示改动比对的那种不算，它里面已经有内容了）。</summary>
    public IViewerWindow? FindEmpty() =>
        _windows.FirstOrDefault(w => w.DocumentKey is null && !w.IsShowingDiff);

    /// <summary>
    /// 把一次请求落到窗口上。<paramref name="create"/> 负责真的建一个窗口（建好、显示出来），
    /// 登记进本集合由这里统一做，这样窗口关闭时的移除和它配得上对。
    /// </summary>
    public void Dispatch(OpenRequest request, Func<IViewerWindow> create)
    {
        if (request.Kind == OpenKind.None)
        {
            // 双击/命令行空跑：定位到最近用过的那个，不新开空窗口
            var target = MostRecent ?? Track(create());
            target.BringToFront();
            Touch(target);
            return;
        }

        if (FindByDocument(request.Path) is { } existing)
        {
            existing.BringToFront();
            Touch(existing);

            // 用副本编辑：比对视图不能就地重开，其余情况都只是在原窗口里再起一次编辑流程
            if (request.Kind == OpenKind.Edit)
            {
                existing.StartEdit(request.Path);
            }

            return;
        }

        var window = FindEmpty() ?? Track(create());
        window.BringToFront();
        Touch(window);

        if (request.Kind == OpenKind.Edit)
        {
            window.StartEdit(request.Path);
        }
        else
        {
            window.OpenPath(request.Path);
        }
    }

    private IViewerWindow Track(IViewerWindow window)
    {
        Add(window);
        return window;
    }

    /// <summary>记成「最近用过」：挪到末尾。</summary>
    private void Touch(IViewerWindow window)
    {
        if (_windows.Remove(window))
        {
            _windows.Add(window);
        }
    }
}
