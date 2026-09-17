using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using ExcelViewer.Model;
using ExcelViewer.Search;

namespace ExcelViewer;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        if (e.Args.Any(a => a.Equals("--selftest", StringComparison.OrdinalIgnoreCase)))
        {
            var exitCode = SelfTest.Run(e.Args);
            Shutdown(exitCode);
            return;
        }

        if (e.Args.Any(a => a.Equals("--layout", StringComparison.OrdinalIgnoreCase)))
        {
            Shutdown(LayoutProbe.Run(e.Args));
            return;
        }

        if (e.Args.Any(a => a.Equals("--layout-synth", StringComparison.OrdinalIgnoreCase)))
        {
            Shutdown(LayoutProbe.RunSynthetic(e.Args));
            return;
        }

        if (e.Args.Any(a => a.Equals("--textprobe", StringComparison.OrdinalIgnoreCase)))
        {
            Shutdown(LayoutProbe.TextProbe(e.Args));
            return;
        }

        if (e.Args.Any(a => a.Equals("--diff-selftest", StringComparison.OrdinalIgnoreCase)))
        {
            Shutdown(DiffSelfTest.Run());
            return;
        }

        if (e.Args.Any(a => a.Equals("--bench", StringComparison.OrdinalIgnoreCase)))
        {
            Shutdown(LayoutProbe.Bench(e.Args));
            return;
        }

        if (e.Args.Any(a => a.Equals("--diff-png", StringComparison.OrdinalIgnoreCase)))
        {
            Shutdown(DiffProbe.Run(e.Args));
            return;
        }

        if (e.Args.Any(a => a.Equals("--probe", StringComparison.OrdinalIgnoreCase)))
        {
            Shutdown(MainProbe.Run(e.Args));
            return;
        }

        if (e.Args.Any(a => a.Equals("--diff-json", StringComparison.OrdinalIgnoreCase)))
        {
            var pair = e.Args.Where(a => !a.StartsWith('-') && File.Exists(a)).Take(2).ToArray();
            var outPath = e.Args.FirstOrDefault(a => a.StartsWith("--out=", StringComparison.OrdinalIgnoreCase))?[6..];
            Shutdown(pair.Length == 2 ? EditSession.DumpJson(pair[0], pair[1], outPath) : 1);
            return;
        }

        // --edit 需要窗口来跑"启动编辑器 → 等关闭 → 比对"的流程，交给 MainWindow
        DispatcherUnhandledException += (_, args) =>
        {
            MessageBox.Show(
                "查看器遇到一个内部问题：\n\n" + args.Exception.Message + "\n\n可以重新打开文件再试；不影响原文件内容。",
                "Excel 查看器",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            args.Handled = true;
        };

        var window = new MainWindow();
        MainWindow = window;
        window.Show();

        // --edit 原文件：复制到工作区 → 用 Excel 打开副本 → 关闭后自动进改动比对
        var editIndex = Array.FindIndex(e.Args, a => a.Equals("--edit", StringComparison.OrdinalIgnoreCase));
        if (editIndex >= 0)
        {
            var target = e.Args.Skip(editIndex + 1).FirstOrDefault(a => !a.StartsWith('-'));
            if (target is not null && File.Exists(target))
            {
                window.StartEditCopy(target);
                return;
            }
        }

        // --diff-svn：SVN 的外部 diff 入口（TortoiseSVN 与 svn --diff-cmd 都走这里）
        if (e.Args.Any(a => a.Equals("--diff-svn", StringComparison.OrdinalIgnoreCase)))
        {
            SvnDiffInput.Pair? svnPair;
            try
            {
                svnPair = SvnDiffInput.Resolve(e.Args);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    window,
                    "比对前的准备工作没能完成：" + ex.Message + "\n\n"
                    + "（SVN 传来的旧版本文件没有扩展名，查看器需要先复制一份带扩展名的临时副本，"
                    + "这一步写临时目录失败时会看到这条提示。）",
                    "Excel 查看器",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (svnPair is null)
            {
                MessageBox.Show(
                    window,
                    "没有从命令行里认出要比对的两个文件。\n\n"
                    + "TortoiseSVN：设置 → 差异查看器 → 高级 → .xlsx →\n"
                    + "    \"…\\viewer\\publish\\ExcelViewer.exe\" --diff-svn %base %mine\n\n"
                    + "命令行：svn diff --force --diff-cmd \"…\\viewer\\publish\\ExcelViewer.exe\" -x \"--diff-svn\"",
                    "Excel 查看器",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            window.ShowDiffAsync(svnPair.OldPath, svnPair.NewPath, svnPair.OldLabel, svnPair.NewLabel);
            return;
        }

        // --diff 老文件 新文件：直接进改动比对视图
        if (e.Args.Any(a => a.Equals("--diff", StringComparison.OrdinalIgnoreCase)))
        {
            var pair = e.Args.Where(a => !a.StartsWith('-') && File.Exists(a)).Take(2).ToArray();
            if (pair.Length == 2)
            {
                window.ShowDiffAsync(pair[0], pair[1]);
                return;
            }
        }

        // 支持命令行 / 右键“用查看器打开”传入的文件
        var files = e.Args.Where(a => !a.StartsWith('-')).ToArray();
        if (files.Length > 0)
        {
            window.OpenFileAsync(files[0]);
        }
    }
}

/// <summary>
/// 改动比对的无界面探针：把改动清单输出成 CSV 与 PNG，方便验证 diff 渲染。
/// 用法: --diff-png 原文件 新文件 --out=清单.csv --png=预览.png
/// </summary>
internal static class DiffProbe
{
    public static int Run(string[] args)
    {
        var pair = args.Where(a => !a.StartsWith('-') && File.Exists(a)).Take(2).ToArray();
        if (pair.Length != 2)
        {
            Console.Error.WriteLine("--diff-png 需要两个文件路径");
            return 1;
        }

        var report = DiffService.Compare(pair[0], pair[1]);
        Console.WriteLine(report.Summary);
        foreach (var s in report.Sheets.Where(s => s.HasChanges))
        {
            Console.WriteLine($"  [{s.SheetName}] {s.ChangeSummary}");
            foreach (var row in s.Rows.Take(20))
            {
                if (row.Kind == RowChangeKind.Modified)
                {
                    foreach (var c in row.Cells.Take(6))
                    {
                        Console.WriteLine($"    {row.KindText} 第{c.Row + 1}行 {Controls.ExcelGrid.ColumnName(c.Col)}列: "
                            + $"\"{Short(c.OldValue)}\" -> \"{Short(c.NewValue)}\"");
                    }

                    continue;
                }

                Console.WriteLine($"    {row.KindText} 第{row.DisplayRow + 1}行 (ID {Short(row.Key)}): \"{Short(row.RowText)}\"");
            }
        }

        var pool = new StringPool();
        var changed = report.Sheets.FirstOrDefault(s => s.HasChanges);
        if (changed is null)
        {
            Console.WriteLine("没有改动，跳过清单渲染。");
            return 0;
        }

        var sheet = DiffService.BuildChangeSheet(changed, pool, Controls.ExcelGrid.ColumnName);

        var csvPath = args.FirstOrDefault(a => a.StartsWith("--out=", StringComparison.OrdinalIgnoreCase))?[6..];
        if (!string.IsNullOrEmpty(csvPath))
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.Join(',', DiffService.ChangeSheetHeaders));
            for (var r = 0; r < sheet.RowCount; r++)
            {
                sb.AppendLine(string.Join(',', Enumerable.Range(0, sheet.ColCount).Select(c => Csv(sheet.Get(r, c)))));
            }

            File.WriteAllText(csvPath, sb.ToString(), new UTF8Encoding(true));
            Console.WriteLine("改动清单已写出: " + csvPath);
        }

        var pngPath = args.FirstOrDefault(a => a.StartsWith("--png=", StringComparison.OrdinalIgnoreCase))?[6..];
        if (!string.IsNullOrEmpty(pngPath))
        {
            const int W = 1100;
            const int H = 420;
            var grid = new Controls.ExcelGrid();
            grid.ShowVirtualSheet(sheet);
            var host = new System.Windows.Controls.Border
            {
                Child = grid,
                Width = W,
                Height = H,
                Background = System.Windows.Media.Brushes.White,
            };
            host.Measure(new Size(W, H));
            host.Arrange(new Rect(0, 0, W, H));
            host.UpdateLayout();

            var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap(W, H, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            rtb.Render(host);
            var enc = new System.Windows.Media.Imaging.PngBitmapEncoder();
            enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(rtb));
            using var fs = File.Create(pngPath);
            enc.Save(fs);
            Console.WriteLine("改动清单预览已保存: " + pngPath);
        }

        return 0;
    }

    private static string Short(string s) => s.Length <= 24 ? s : s[..24] + "…";

    private static string Csv(string s) =>
        s.Contains(',') || s.Contains('"') || s.Contains('\n')
            ? "\"" + s.Replace("\"", "\"\"") + "\""
            : s;
}

/// <summary>
/// 界面探针：真的建出 MainWindow、真的走一遍打开文件的流程，再把界面上的状态打印出来。
/// 用来定位"数据加载成功但表格是空的"这类只在真实窗口里出现的问题。
/// 用法: --probe 文件.xlsx [--png=out.png]
/// </summary>
internal static class MainProbe
{
    public static int Run(string[] args)
    {
        var path = args.FirstOrDefault(a => !a.StartsWith('-') && File.Exists(a));
        if (path is null)
        {
            Console.Error.WriteLine("--probe 需要一个文件路径");
            return 1;
        }

        var window = new MainWindow();
        window.Width = 1280;
        window.Height = 820;
        window.Left = -4000;
        window.Top = -4000;
        window.Show();

        Pump(300);
        Console.WriteLine($"窗口已显示: ActualWidth={window.ActualWidth} ActualHeight={window.ActualHeight}");

        window.OpenFileAsync(path);

        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(20);
        var ready = false;
        while (DateTime.UtcNow < deadline)
        {
            Pump(100);
            var sheet = Field(window, "_sheet");
            if (sheet is not null)
            {
                ready = true;
                break;
            }
        }

        var grid = window.Grid;
        var model = grid.Model;
        Console.WriteLine($"加载完成: {ready}");
        Console.WriteLine($"MainWindow._sheet      = {Describe(Field(window, "_sheet"))}");
        Console.WriteLine($"ExcelGrid.Model.Sheet  = {Describe(model.Sheet)}");
        Console.WriteLine($"Model.RowCount/ColCount= {model.RowCount} / {model.ColCount}");
        Console.WriteLine($"ActualWidth/Height     = {grid.ActualWidth} / {grid.ActualHeight}");
        Console.WriteLine($"RowHeight/HeaderHeight = {model.RowHeight} / {model.HeaderHeight}");
        Console.WriteLine($"GutterWidth            = {model.GutterWidth}");
        Console.WriteLine($"TotalWidth/TotalHeight = {model.TotalWidth} / {model.TotalHeight}");
        Console.WriteLine($"DetectedHeaderRow      = {grid.DetectedHeaderRow}");
        Console.WriteLine($"VisibleRowCount        = {grid.VisibleRowCount}");
        Console.WriteLine($"TopVisibleRow          = {grid.TopVisibleRow}");
        Console.WriteLine($"StatusSize/Row/Hint    = {window.StatusSize.Text} | {window.StatusRow.Text} | {window.StatusHint.Text}");
        Console.WriteLine($"LoadTimeText           = {window.LoadTimeText.Text}");

        // --diff-svn / --diff：把比对视图也探一遍（顺便验证 SVN 传来的参数能被正确认出来）
        var diffRequested = false;
        if (args.Any(a => a.Equals("--diff-svn", StringComparison.OrdinalIgnoreCase)))
        {
            var pair = SvnDiffInput.Resolve(args);
            Console.WriteLine($"SVN 参数解析           = {(pair is null ? "(失败)" : pair.OldLabel + "  ←→  " + pair.NewLabel)}");
            if (pair is not null)
            {
                window.ShowDiffAsync(pair.OldPath, pair.NewPath, pair.OldLabel, pair.NewLabel);
                diffRequested = true;
            }
        }
        else if (args.Any(a => a.Equals("--diff", StringComparison.OrdinalIgnoreCase)))
        {
            var files = args.Where(a => !a.StartsWith('-') && File.Exists(a)).Take(2).ToArray();
            if (files.Length == 2)
            {
                window.ShowDiffAsync(files[0], files[1]);
                diffRequested = true;
            }
        }

        if (diffRequested)
        {
            var diffDeadline = DateTime.UtcNow + TimeSpan.FromSeconds(120);
            while (DateTime.UtcNow < diffDeadline)
            {
                Pump(200);
                if (window.DiffBanner.Visibility == Visibility.Visible)
                {
                    break;
                }
            }

            Console.WriteLine($"改动比对横幅           = {window.DiffBannerText.Text}");
            Console.WriteLine($"改动比对明细           = {window.DiffDetailText.Text}");
        }

        var cellArg = args.FirstOrDefault(a => a.StartsWith("--cell=", StringComparison.OrdinalIgnoreCase))?[7..];        if (cellArg is not null)
        {
            var parts = cellArg.Split(',', 2);
            if (parts.Length == 2 && int.TryParse(parts[0], out var pr) && int.TryParse(parts[1], out var pc))
            {
                grid.GoToCell(pr, pc);
            }
        }
        else
        {
            grid.SetActiveCell(0, 0, ensureVisible: false);
        }

        if (args.Any(a => a.Equals("--preview", StringComparison.OrdinalIgnoreCase))
            || args.Any(a => a.Equals("--preview-off", StringComparison.OrdinalIgnoreCase)))
        {
            var open = !args.Any(a => a.Equals("--preview-off", StringComparison.OrdinalIgnoreCase));
            var toggle = window.GetType()
                .GetMethod("AppPreviewPane", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;

            toggle.Invoke(window, new object[] { open });
            Pump(300);
            Console.WriteLine($"预览条: 展开={Field(window, "_previewPaneOpen")} 行高={window.PreviewRow.Height} "
                + $"地址={window.PreviewAddress.Text} 概要={window.PreviewMeta.Text} "
                + $"正文长度={window.PreviewTextBox.Text.Length} 提示={window.PreviewWarn.Visibility} "
                + $"收起行={window.PreviewCollapsedBar.Visibility} 收起摘要={window.PreviewCollapsedText.Text}");
        }

        Pump(400);

        var pngWindow = args.FirstOrDefault(a => a.StartsWith("--png-window=", StringComparison.OrdinalIgnoreCase))?[13..];
        if (!string.IsNullOrEmpty(pngWindow))
        {
            var root = (System.Windows.FrameworkElement)window.Content;
            var w = (int)Math.Ceiling(root.ActualWidth);
            var h = (int)Math.Ceiling(root.ActualHeight);
            var rtbW = new System.Windows.Media.Imaging.RenderTargetBitmap(w, h, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            rtbW.Render(root);
            var encW = new System.Windows.Media.Imaging.PngBitmapEncoder();
            encW.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(rtbW));
            using (var fs = File.Create(pngWindow))
            {
                encW.Save(fs);
            }

            Console.WriteLine($"整窗内容已渲染: {pngWindow} ({w}x{h})");
        }

        var png = args.FirstOrDefault(a => a.StartsWith("--png=", StringComparison.OrdinalIgnoreCase))?[6..];
        if (!string.IsNullOrEmpty(png))
        {
            const int W = 1100;
            const int H = 700;
            var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap(W, H, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            rtb.Render(grid);
            var enc = new System.Windows.Media.Imaging.PngBitmapEncoder();
            enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(rtb));
            using (var fs = File.Create(png))
            {
                enc.Save(fs);
            }

            Console.WriteLine($"表格区域已渲染: {png}");
        }

        window.Close();
        return 0;
    }

    private static object? Field(object target, string name) =>
        target.GetType().GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?.GetValue(target);

    private static string Describe(object? sheet) => sheet is null
        ? "(null)"
        : sheet is ExcelViewer.Model.Sheet s
            ? $"Sheet '{s.Name}' {s.RowCount} 行 x {s.ColCount} 列"
            : sheet.GetType().Name;

    private static void Pump(int ms)
    {
        var frame = new System.Windows.Threading.DispatcherFrame();
        var timer = new System.Windows.Threading.DispatcherTimer(
            TimeSpan.FromMilliseconds(ms),
            System.Windows.Threading.DispatcherPriority.Background,
            (_, _) => frame.Continue = false,
            System.Windows.Threading.Dispatcher.CurrentDispatcher);
        timer.Start();
        System.Windows.Threading.Dispatcher.PushFrame(frame);
        timer.Stop();
    }
}

/// <summary>
/// 自检：不打开窗口，直接用真实的加载器 + 搜索器跑一遍，验证数据管线。
/// 用法: dotnet ExcelViewer.dll --selftest [文件...]
/// </summary>
internal static class SelfTest
{
    public static int Run(string[] args)
    {
        var output = new StringBuilder();
        var failures = 0;

        void Log(string s)
        {
            output.AppendLine(s);
            Console.WriteLine(s);
        }

        var files = args.Where(a => !a.StartsWith('-') && File.Exists(a)).ToArray();
        if (files.Length == 0)
        {
            Log("!! 没有可用的测试文件");
            return 1;
        }

        foreach (var path in files)
        {
            Log($"=== {Path.GetFileName(path)} ===");
            var pool = new StringPool();
            var names = WorkbookLoader.ReadSheetNames(path);
            Log($"  工作表: {string.Join(" | ", names)}");

            foreach (var (name, index) in names.Select((n, i) => (n, i)))
            {
                var sw = Stopwatch.StartNew();
                var sheet = WorkbookLoader.LoadSheet(path, index, pool);
                sw.Stop();

                Log($"  [{name}] {sheet.RowCount} 行 x {sheet.ColCount} 列, 非空 {sheet.NonEmptyCellCount:N0} 单元格, "
                    + $"读取 {sw.ElapsedMilliseconds} ms, 省略尾部空行 {sheet.SkippedTrailingRows}");

                // 抽查：任意一行的取值不应抛异常，且表头行有内容
                if (sheet.RowCount == 0)
                {
                    Log("   !! 空表");
                    failures++;
                    continue;
                }

                var headerSample = string.Join(" | ", Enumerable.Range(0, Math.Min(6, sheet.ColCount)).Select(c => sheet.Get(0, c)));
                Log($"    首行前几列: {Truncate(headerSample, 90)}");

                var lastRow = sheet.RowCount - 1;
                var lastSample = string.Join(" | ", Enumerable.Range(0, Math.Min(4, sheet.ColCount)).Select(c => sheet.Get(lastRow, c)));
                Log($"    末行前几列: {Truncate(lastSample, 90)}");

                // 搜索：取首行第一个非空单元格里的关键字，应当至少命中它自己
                var probe = Enumerable.Range(0, sheet.ColCount)
                    .Select(c => sheet.Get(0, c))
                    .FirstOrDefault(v => v.Length >= 2);
                if (probe is not null)
                {
                    var needle = probe.Length > 12 ? probe[..12] : probe;
                    var hits = SearchService.Search(sheet, new SearchOptions { Query = needle });
                    Log($"    搜索「{Truncate(needle, 20)}」命中 {hits.RowHitCount} 行 / {hits.CellHitCount} 单元格, {hits.ElapsedMs} ms");
                    if (!hits.HasHits)
                    {
                        Log("   !! 搜索自己没有命中，可能有问题");
                        failures++;
                    }

                    var firstHit = hits.NextHitRow(-1);
                    if (firstHit < 0 || SearchService.FirstHitColumn(sheet, hits, firstHit) < 0)
                    {
                        Log("   !! 定位命中列失败");
                        failures++;
                    }
                }

                // 列式存储一致性：Get 越界应返回空串而不是抛异常
                if (sheet.Get(-1, 0).Length != 0 || sheet.Get(0, -1).Length != 0 || sheet.Get(sheet.RowCount, 0).Length != 0)
                {
                    Log("   !! 越界读取没有返回空值");
                    failures++;
                }
            }

            Log($"  字符串池: 唯一值 {pool.UniqueCount:N0}（池命中 {pool.HitCount:N0} / 新增 {pool.MissCount:N0}）");
        }

        var dumpPath = args.FirstOrDefault(a => a.StartsWith("--out=", StringComparison.OrdinalIgnoreCase))?[6..];
        if (!string.IsNullOrEmpty(dumpPath))
        {
            File.WriteAllText(dumpPath, output.ToString(), new UTF8Encoding(false));
            Console.WriteLine($"结果已写入 {dumpPath}");
        }

        Console.WriteLine(failures == 0 ? "SELFTEST OK" : $"SELFTEST FAILED ({failures})");
        return failures == 0 ? 0 : 1;
    }

    private static string Truncate(string s, int max) => s.Length <= max ? s : s[..max] + "…";

}

/// <summary>
/// 布局探针：在真实 WPF 环境里创建窗口与表格，把列偏移量打印出来核对（不显示窗口）。
/// 用法: dotnet ExcelViewer.dll --layout 文件.xlsx
/// </summary>
internal static class LayoutProbe
{
    public static int Run(string[] args)
    {
        var path = args.FirstOrDefault(a => !a.StartsWith('-') && File.Exists(a));
        if (path is null)
        {
            Console.Error.WriteLine("--layout 需要一个文件路径（或用 --layout-synth 测合成表）");
            return 1;
        }

        var pool = new StringPool();
        var sheet = WorkbookLoader.LoadSheet(path, 0, pool);
        return Probe(sheet, args);
    }

    /// <summary>合成一张内容完全可控的小表，用来隔离渲染问题。</summary>
    public static int RunSynthetic(string[] args)
    {
        var pool = new StringPool();
        var sheet = new Sheet("synth", pool);
        var useLongValues = args.Any(a => a.Equals("--long", StringComparison.OrdinalIgnoreCase));
        for (var r = 0; r < 30; r++)
        {
            sheet.BeginRow();
            // 用与真实配表同样长度的值，复现"超宽列"场景
            sheet.AddCell(useLongValues ? "GDE_IGNORE" : "A" + r.ToString("D3"));
            sheet.AddCell(useLongValues ? "字段名" : "B" + r.ToString("D3"));
            sheet.AddCell(useLongValues ? "中文描述字段" : "中文字段" + r);
            sheet.AddCell(useLongValues ? r.ToString() : r.ToString());
            sheet.EndRow();
        }

        sheet.Finish();
        return Probe(sheet, args);
    }

    private static int Probe(Sheet sheet, string[] args)
    {
        var grid = new ExcelViewer.Controls.ExcelGrid();
        grid.DetectedHeaderRow = HeaderDetection.Detect(sheet).HeaderRowIndex;
        if (args.Any(a => a.Equals("--no-gutter", StringComparison.OrdinalIgnoreCase)))
        {
            grid.ShowRowNumbers = false;
        }

        grid.SetSheet(sheet, keepScroll: false);

        // 走一遍真实布局：Measure/Arrange 之后列偏移量才有意义
        grid.Measure(new Size(1280, 560));
        grid.Arrange(new Rect(0, 0, 1280, 560));

        var model = grid.Model;
        Console.WriteLine($"sheet={sheet.Name} rows={sheet.RowCount} cols={sheet.ColCount}");
        Console.WriteLine($"GutterWidth={model.GutterWidth} HeaderHeight={model.HeaderHeight} RowHeight={model.RowHeight}");
        Console.WriteLine($"TotalWidth={model.TotalWidth} ContentWidth={model.ContentWidth} ShowRowNumbers={model.ShowRowNumbers}");
        Console.WriteLine($"DetectedHeaderRow={grid.DetectedHeaderRow}");
        for (var c = 0; c < Math.Min(12, sheet.ColCount); c++)
        {
            Console.WriteLine($"  col{c,2} [{ExcelViewer.Controls.ExcelGrid.ColumnName(c)}] left={model.ColumnLeft(c),8:F1} right={model.ColumnRight(c),8:F1} width={model.GetColumnWidth(c),7:F1}");
        }

        var probeX = new[] { 0.0, 10, 30, 55, 56, 57, 80, 120, 200, 300 };
        foreach (var x in probeX)
        {
            Console.WriteLine($"  ColumnAt({x,5:F0}) = {model.ColumnAt(x)}");
        }

        // 输出前几行的首列值，确认渲染取到的内容
        Console.WriteLine("首列前 6 行: " + string.Join(" | ", Enumerable.Range(0, Math.Min(6, sheet.RowCount)).Select(r => sheet.Get(r, 0))));

        // 离屏渲染成位图，用像素证据核对裁剪与对齐
        const int W = 1280;
        const int H = 560;
        var host = new System.Windows.Controls.Border
        {
            Child = grid,
            Width = W,
            Height = H,
            Background = System.Windows.Media.Brushes.White,
        };
        host.Measure(new Size(W, H));
        host.Arrange(new Rect(0, 0, W, H));
        host.UpdateLayout();

        var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap(W, H, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
        rtb.Render(host);        var stride = W * 4;
        var pixels = new byte[stride * H];
        rtb.CopyPixels(pixels, stride, 0);

        var outPath = args.FirstOrDefault(a => a.StartsWith("--png=", StringComparison.OrdinalIgnoreCase))?[6..];
        if (!string.IsNullOrEmpty(outPath))
        {
            var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(rtb));
            using var fs = File.Create(outPath);
            encoder.Save(fs);
            Console.WriteLine($"离屏渲染已保存: {outPath}");
        }

        Console.WriteLine("每行最左侧深色像素 x（用于验证是否被行号栏裁掉）：");
        for (var r = 0; r < 12; r++)
        {
            var y = (int)(model.HeaderHeight + (r * model.RowHeight) + (model.RowHeight / 2));
            if (y >= H)
            {
                break;
            }

            var leftmost = -1;
            for (var x = 0; x < W; x++)
            {
                var i = (y * stride) + (x * 4);
                var b = pixels[i];
                var g = pixels[i + 1];
                var rr = pixels[i + 2];
                if (rr < 0x90 && g < 0x90 && b < 0x90)
                {
                    leftmost = x;
                    break;
                }
            }

            Console.WriteLine($"  row{r,2} y={y,4} 最左深色像素 x={leftmost}  (行号栏右边界={model.GutterWidth:F0})");
        }

        Console.WriteLine("每行深色像素的连续区间（列边界用）：");        for (var r = 0; r < 8; r++)
        {
            var y0 = (int)(model.HeaderHeight + (r * model.RowHeight)) + 1;
            var y1 = Math.Min(H - 1, (int)(model.HeaderHeight + ((r + 1) * model.RowHeight)) - 1);
            if (y0 >= H)
            {
                break;
            }

            // 逐列判断该列在整行高度内是否出现深色像素
            var runs = new List<(int Start, int End)>();
            var inRun = false;
            var start = 0;
            for (var x = 0; x < W; x++)
            {
                var dark = false;
                for (var y = y0; y <= y1 && !dark; y++)
                {
                    var i = (y * stride) + (x * 4);
                    if (pixels[i] < 0x90 && pixels[i + 1] < 0x90 && pixels[i + 2] < 0x90)
                    {
                        dark = true;
                    }
                }

                if (dark && !inRun)
                {
                    inRun = true;
                    start = x;
                }
                else if (!dark && inRun)
                {
                    inRun = false;
                    runs.Add((start, x - 1));
                }
            }

            if (inRun)
            {
                runs.Add((start, W - 1));
            }

            Console.WriteLine($"  row{r,2}: " + string.Join("  ", runs.Take(12).Select(t => $"[{t.Start}-{t.End}]")));
        }

        return 0;
    }

    /// <summary>
    /// 性能基准：加载耗时、可视区重绘耗时（模拟滚动）、搜索耗时、内存占用。
    /// 用法: dotnet ExcelViewer.dll --bench 文件.xlsx [--frames=200]
    /// </summary>
    public static int Bench(string[] args)
    {
        var path = args.FirstOrDefault(a => !a.StartsWith('-') && File.Exists(a));
        if (path is null)
        {
            Console.Error.WriteLine("--bench 需要一个文件路径");
            return 1;
        }

        var frames = 200;
        var framesArg = args.FirstOrDefault(a => a.StartsWith("--frames=", StringComparison.OrdinalIgnoreCase));
        if (framesArg is not null)
        {
            frames = int.Parse(framesArg[9..]);
        }

        const int W = 1400;
        const int H = 700;

        Console.WriteLine($"### {Path.GetFileName(path)}  ({new FileInfo(path).Length / 1048576.0:F2} MB)");

        // 1) 只读首 sheet 的耗时（首屏真实成本）
        var pool = new StringPool();
        var swLoad = Stopwatch.StartNew();
        var sheet = WorkbookLoader.LoadSheet(path, 0, pool);
        swLoad.Stop();

        Console.WriteLine($"  首 sheet 加载      : {swLoad.Elapsed.TotalMilliseconds,8:F0} ms  "
            + $"({sheet.RowCount:N0} 行 x {sheet.ColCount} 列, 非空 {sheet.NonEmptyCellCount:N0})");
        Console.WriteLine($"  字符串池           : 唯一 {pool.UniqueCount:N0}  （命中 {pool.HitCount:N0}）");

        var grid = new ExcelViewer.Controls.ExcelGrid();
        var host = new System.Windows.Controls.Border
        {
            Child = grid,
            Width = W,
            Height = H,
            Background = System.Windows.Media.Brushes.White,
        };
        host.Measure(new Size(W, H));
        host.Arrange(new Rect(0, 0, W, H));
        host.UpdateLayout();

        // 2) 首次上屏（含列宽自适应 + 首次排版）
        var swFirst = Stopwatch.StartNew();
        grid.SetSheet(sheet, keepScroll: false);
        host.UpdateLayout();
        var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap(W, H, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
        rtb.Render(host);
        swFirst.Stop();
        Console.WriteLine($"  首次上屏（含列宽） : {swFirst.Elapsed.TotalMilliseconds,8:F0} ms");

        // 3) 模拟滚动：jump = 随机远跳（最坏情况，缓存基本失效）；smooth = 连续滚动（真实使用情况）
        var mode = args.Any(a => a.Equals("--scroll-jump", StringComparison.OrdinalIgnoreCase)) ? "jump" : "smooth";
        var visibleRows = grid.VisibleRowCount;
        var times = new List<double>(frames);
        var swFrame = new Stopwatch();
        var row = 0;
        for (var i = 0; i < frames; i++)
        {
            if (mode == "jump")
            {
                row = (i * 97) % Math.Max(1, sheet.RowCount - visibleRows);
            }
            else
            {
                row = (row + 2) % Math.Max(1, sheet.RowCount - visibleRows);
            }

            swFrame.Restart();
            grid.GoToCell(row, 0);
            host.UpdateLayout();
            rtb.Render(host);
            swFrame.Stop();
            times.Add(swFrame.Elapsed.TotalMilliseconds);
        }

        times.Sort();
        var avg = times.Average();
        Console.WriteLine($"  滚动重绘 x{frames,-4} ({mode,-6}): 平均 {avg,7:F2} ms  中位 {times[times.Count / 2],7:F2} ms  "
            + $"P95 {times[(int)(times.Count * 0.95)],7:F2} ms  最差 {times[^1],7:F2} ms  （{1000.0 / avg:F0} 帧/秒）");

        // 4) 搜索：全表关键字
        foreach (var needle in new[] { "a", "1", "GDE", "GDE_IGNORE" })
        {
            var swSearch = Stopwatch.StartNew();
            var result = SearchService.Search(sheet, new SearchOptions { Query = needle });
            swSearch.Stop();
            Console.WriteLine($"  搜索「{needle,-12}」: {swSearch.Elapsed.TotalMilliseconds,8:F1} ms  "
                + $"命中 {result.RowHitCount:N0} 行 / {result.CellHitCount:N0} 单元格");
        }

        // 5) 内存
        var managed = GC.GetTotalMemory(false) / 1048576.0;
        var peak = Process.GetCurrentProcess().PeakWorkingSet64 / 1048576.0;
        var current = Process.GetCurrentProcess().WorkingSet64 / 1048576.0;
        Console.WriteLine($"  内存               : 托管 {managed:F0} MB, 进程当前 {current:F0} MB, 峰值 {peak:F0} MB");
        Console.WriteLine($"  排版缓存条目       : {grid.TextLayoutCacheCount}");

        return 0;
    }

    /// <summary>
    /// 直接测 CellTextRenderer：把几条已知文本画在已知 x 上，再看像素落在哪里。
    /// 用来定位"文字被画到左边"到底是渲染器还是表格布局的问题。
    /// </summary>
    public static int TextProbe(string[] args)
    {
        const int W = 700;
        // 两组样本（TF 一遍 + FT 一遍）各 4 行 * 24px，再加分隔线，200 高会越界读到画布外
        const int H = 260;
        var visual = new System.Windows.Media.DrawingVisual();
        var typeface = new System.Windows.Media.Typeface(
            new System.Windows.Media.FontFamily("Microsoft YaHei UI, Segoe UI"),
            System.Windows.FontStyles.Normal,
            System.Windows.FontWeights.Normal,
            System.Windows.FontStretches.Normal);
        var black = System.Windows.Media.Brushes.Black;

        var renderer = new ExcelViewer.Controls.CellTextRenderer(typeface, 13, 1.0, black);

        var samples = new (string Text, double X, double W, int Align)[]
        {
            ("GDE_IGNORE", 56, 132.9, ExcelViewer.Controls.CellTextRenderer.AlignLeft),
            ("GDE_IGNORE", 56, 45.0, ExcelViewer.Controls.CellTextRenderer.AlignLeft),
            ("AA", 200, 40, ExcelViewer.Controls.CellTextRenderer.AlignLeft),
            ("中文测试", 300, 120, ExcelViewer.Controls.CellTextRenderer.AlignLeft),
        };

        using (var dc = visual.RenderOpen())
        {
            dc.DrawRectangle(System.Windows.Media.Brushes.White, null, new Rect(0, 0, W, H));
            var y = 6.0;
            foreach (var s in samples)
            {
                renderer.Draw(dc, s.Text, s.Align, s.X, y, s.W, 22);
                y += 24;
            }

            dc.DrawLine(new System.Windows.Media.Pen(System.Windows.Media.Brushes.LightGray, 1), new Point(0, y), new Point(W, y));
            y += 6;

            // 同参数的 FormattedText 版本，作为对照
            foreach (var s in samples)
            {
                var ft = new System.Windows.Media.FormattedText(
                    s.Text,
                    System.Globalization.CultureInfo.CurrentUICulture,
                    FlowDirection.LeftToRight,
                    typeface,
                    13,
                    black,
                    1.0)
                {
                    MaxTextWidth = s.W,
                    Trimming = TextTrimming.CharacterEllipsis,
                };
                dc.DrawText(ft, new Point(s.X, y));
                y += 24;
            }
        }

        var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap(W, H, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
        rtb.Render(visual);
        var stride = W * 4;
        var px = new byte[stride * H];
        rtb.CopyPixels(px, stride, 0);

        void Scan(string tag, int y0, int y1, double expectedX, string text)
        {
            // 采样区间可能压到画布边缘，先夹一下，避免越界读像素把探针自己搞崩
            y0 = Math.Clamp(y0, 0, H - 1);
            y1 = Math.Clamp(y1, y0 + 1, H);

            var left = -1;
            var right = -1;
            for (var x = 0; x < W; x++)
            {
                var dark = false;
                for (var y = y0; y < y1 && !dark; y++)
                {
                    var idx = (y * stride) + (x * 4);
                    if (px[idx] < 0x80 && px[idx + 1] < 0x80 && px[idx + 2] < 0x80)
                    {
                        dark = true;
                    }
                }

                if (dark)
                {
                    if (left < 0)
                    {
                        left = x;
                    }

                    right = x;
                }
            }

            Console.WriteLine($"  [{tag}] {text,-12} 期望x={expectedX,6:F0} 实测[{left},{right}] 宽度={right - left + 1}");
        }

        Console.WriteLine("TextProbe A/B（上=TextFormatter，下=FormattedText）：");
        for (var i = 0; i < samples.Length; i++)
        {
            Scan("TF", (int)(6 + (i * 24)), (int)(6 + (i * 24)) + 22, samples[i].X, samples[i].Text);
        }

        var yBase = (int)(6 + (samples.Length * 24) + 6);
        for (var i = 0; i < samples.Length; i++)
        {
            Scan("FT", yBase + (i * 24), yBase + (i * 24) + 22, samples[i].X, samples[i].Text);
        }

        var outPath = args.FirstOrDefault(a => a.StartsWith("--png=", StringComparison.OrdinalIgnoreCase))?[6..];
        if (!string.IsNullOrEmpty(outPath))
        {
            var enc = new System.Windows.Media.Imaging.PngBitmapEncoder();
            enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(rtb));
            using var fs = File.Create(outPath);
            enc.Save(fs);
            Console.WriteLine($"已保存 {outPath}");
        }

        return 0;
    }
}

/// <summary>
/// 改动比对的自检：用合成数据把"行对齐"的几种典型情况跑一遍，并检查 SVN 参数解析/格式嗅探。
/// 用法: ExcelViewer.exe --diff-selftest
/// 这些用例都是曾经真出过问题的地方（插入一行报出几百处改动、空行刷屏、无扩展名的 pristine 文件）。
/// </summary>
internal static class DiffSelfTest
{
    private static int _failures;

    public static int Run()
    {
        _failures = 0;

        CaseCellEdit();
        CaseRowInsert();
        CaseRowDelete();
        CaseMixed();
        CaseDuplicateRows();
        CaseNoKeyColumn();
        CaseSvnArguments();

        if (_failures == 0)
        {
            Console.WriteLine("DIFF SELFTEST OK");
            return 0;
        }

        Console.WriteLine($"DIFF SELFTEST FAILED: {_failures} 项不通过");
        return 1;
    }

    // ---------------- 用例 ----------------

    private static void CaseCellEdit()
    {
        var report = Compare(Table(100), Edit(Table(100), 50, 2, "改过了"));
        Check("改一个格子只报 1 处", report.TotalChangedCells == 1 && report.TotalAddedRows == 0 && report.TotalRemovedRows == 0, report.Summary);
    }

    private static void CaseRowInsert()
    {
        // 曾经的 bug：200 行表中间插一行会报出 300+ 处"改动"
        var newRows = Table(200).ToList();
        newRows.Insert(100, new[] { "GDE_IGNORE", "ID9999", "新加的一行", "新说明" });
        var report = Compare(Table(200), newRows.ToArray());
        Check("中间插一行只报 1 行新增", report.TotalAddedRows == 1 && report.TotalRemovedRows == 0 && report.TotalChangedCells == 0, report.Summary);
    }

    private static void CaseRowDelete()
    {
        var rows = Table(100).ToList();
        rows.RemoveAt(30);
        var report = Compare(Table(100), rows.ToArray());
        Check("删一行只报 1 行删除", report.TotalRemovedRows == 1 && report.TotalAddedRows == 0 && report.TotalChangedCells == 0, report.Summary);
    }

    private static void CaseMixed()
    {
        var oldRows = Table(100);
        var newRows = Table(100);
        newRows = Edit(newRows, 10, 2, "改了A");
        newRows = Edit(newRows, 80, 3, "改了B");
        var list = newRows.ToList();
        list.Insert(50, new[] { "GDE_IGNORE", "ID7777", "插入行", "X" });
        list.RemoveAt(20);
        var report = Compare(oldRows, list.ToArray());
        Check(
            "混合改动：2 格 + 1 增设 + 1 删除",
            report.TotalChangedCells == 2 && report.TotalAddedRows == 1 && report.TotalRemovedRows == 1,
            report.Summary);
    }

    private static void CaseDuplicateRows()
    {
        // 重复行不是唯一锚，会整段落进"区间"；两边一模一样时必须报 0 改动（曾经会刷出几千条假增删）
        var rows = new List<string[]> { new[] { "GDE_IGNORE", "同一条内容" } };
        for (var i = 0; i < 300; i++)
        {
            rows.Add(new[] { "同一条内容", "同一条内容" });
        }

        var report = Compare(rows.ToArray(), rows.ToArray());
        Check("两边完全一样（含大量重复行）→ 0 改动", !report.HasChanges, report.Summary);
    }

    private static void CaseNoKeyColumn()
    {
        // 没有唯一列的表：只能靠整行文本 + 相似度配对
        var oldRows = NoKeyTable(60);
        var newRows = NoKeyTable(60).ToList();
        newRows[10] = new[] { "类型A", "内容10改了" };
        newRows.Insert(31, new[] { "类型A", "插入内容" });
        var report = Compare(oldRows, newRows.ToArray());
        Check(
            "无主键表：1 格修改 + 1 行新增",
            report.TotalChangedCells == 1 && report.TotalAddedRows == 1 && report.TotalRemovedRows == 0,
            report.Summary);
    }

    private static void CaseSvnArguments()
    {
        var dir = Path.Combine(Path.GetTempPath(), "ExcelViewer-diffselftest-" + Guid.NewGuid().ToString("N")[..6]);
        Directory.CreateDirectory(dir);
        try
        {
            // 造一个 zip 头 + xl/workbook.xml 的假 xlsx，但故意用 pristine 的 .svn-base 文件名
            var pristine = Path.Combine(dir, "29e3e4770d9b96cd165b9d261c51fda557a78d19.svn-base");
            using (var fs = File.Create(pristine))
            using (var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create))
            {
                using var w = new StreamWriter(zip.CreateEntry("xl/workbook.xml").Open());
                w.Write("<workbook/>");
            }

            var work = Path.Combine(dir, "T.xlsx");
            File.Copy(pristine, work, overwrite: true);

            var args = new[]
            {
                "--diff-svn",
                "-L", work + "\t(revision 7)",
                "-L", work + "\t(working copy)",
                pristine,
                work,
            };

            var pair = SvnDiffInput.Resolve(args);
            Check("SVN 参数解析成功", pair is not null);
            if (pair is null)
            {
                return;
            }

            Check("无扩展名的 pristine 被补成 .xlsx", pair.OldPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase), pair.OldPath);
            Check("带扩展名的文件不复制", string.Equals(pair.NewPath, work, StringComparison.OrdinalIgnoreCase), pair.NewPath);
            Check("标签保留了 (revision 7)", pair.OldLabel.Contains("(revision 7)"), pair.OldLabel);
            Check("标签里没有制表符", !pair.OldLabel.Contains('\t'), pair.OldLabel);
        }
        finally
        {
            try
            {
                Directory.Delete(dir, recursive: true);
            }
            catch (IOException)
            {
            }
        }
    }

    // ---------------- 工具 ----------------

    private static DiffReport Compare(string[][] oldRows, string[][] newRows)
    {
        var pool = new StringPool();
        var oldSheet = Build(pool, "S", oldRows);
        var newSheet = Build(pool, "S", newRows);
        var diff = DiffService.CompareSheets(oldSheet, newSheet, "S");
        return new DiffReport
        {
            OldPath = "old",
            NewPath = "new",
            Sheets = new List<SheetDiff> { diff },
            RemovedSheets = new List<string>(),
            AddedSheets = new List<string>(),
        };
    }

    private static Sheet Build(StringPool pool, string name, string[][] rows)
    {
        var sheet = new Sheet(name, pool);
        foreach (var row in rows)
        {
            sheet.BeginRow();
            foreach (var cell in row)
            {
                sheet.AddCell(cell);
            }

            sheet.EndRow();
        }

        sheet.Finish();
        return sheet;
    }

    private static string[][] Table(int rows)
    {
        var list = new List<string[]> { new[] { "GDE_FIELD_NAMES", "ID", "Name", "Desc" } };
        for (var i = 1; i <= rows; i++)
        {
            list.Add(new[] { "GDE_IGNORE", $"ID{i:0000}", $"名称{i}", $"说明{i}" });
        }

        return list.ToArray();
    }

    private static string[][] NoKeyTable(int rows)
    {
        var list = new List<string[]>();
        for (var i = 1; i <= rows; i++)
        {
            list.Add(new[] { "类型A", $"内容{i}" });
        }

        return list.ToArray();
    }

    /// <summary>改一个格子（1-based 行号，模拟策划在 Excel 里改值）。</summary>
    private static string[][] Edit(string[][] rows, int row1Based, int col, string value)
    {
        var copy = rows.Select(r => r.ToArray()).ToArray();
        copy[row1Based - 1][col] = value;
        return copy;
    }

    private static void Check(string what, bool ok, string detail = "")
    {
        Console.WriteLine($"  [{(ok ? "OK" : "FAIL")}] {what}{(ok || detail.Length == 0 ? string.Empty : "  → " + detail)}");
        if (!ok)
        {
            _failures++;
        }
    }
}





