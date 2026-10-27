using System.IO;
using System.Text;
using System.Windows.Threading;

namespace ExcelViewer;

/// <summary>
/// 「同一个表只开一个窗口」的自检：路径归一、命令行解析、请求编码、进程间转交、窗口落点策略。
/// 用法: ExcelViewer.exe --dedup-selftest
///
/// 这里用的都是真代码路径（真的建互斥体、真的开命名管道、真的走一遍编码解码），
/// 只有"窗口"换成了假窗口 —— 去重规则不依赖 WPF，正是为了能这样单独验。
/// </summary>
internal static class DedupSelfTest
{
    private static int _failures;

    public static int Run()
    {
        _failures = 0;

        CaseFileKey();
        CaseParse();
        CaseWire();
        CaseDispatch();
        CasePipe();

        Console.WriteLine(_failures == 0 ? "DEDUP SELFTEST OK" : $"DEDUP SELFTEST FAILED: {_failures} 项不通过");
        return _failures == 0 ? 0 : 1;
    }

    // ---------------- 路径归一 ----------------

    private static void CaseFileKey()
    {
        Console.WriteLine("路径归一：");

        Check("正斜杠和反斜杠是同一个键",
            FileKey.Of(@"C:/配置/Item.xlsx") == FileKey.Of(@"C:\配置\Item.xlsx"),
            FileKey.Of(@"C:/配置/Item.xlsx"));

        Check("相对路径按当前目录补成完整路径",
            FileKey.Of("Item.xlsx") == FileKey.Of(Path.Combine(Environment.CurrentDirectory, "Item.xlsx")),
            FileKey.Of("Item.xlsx"));

        Check("路径里的 . 和 .. 被抹平",
            FileKey.Of(@"C:\配置\.\子目录\..\Item.xlsx") == FileKey.Of(@"C:\配置\Item.xlsx"),
            FileKey.Of(@"C:\配置\.\子目录\..\Item.xlsx"));

        Check("带首尾引号的路径也认",
            FileKey.Of("\"C:\\配置\\Item.xlsx\"") == FileKey.Of(@"C:\配置\Item.xlsx"),
            FileKey.Of("\"C:\\配置\\Item.xlsx\""));

        Check("不同的文件是不同的键",
            FileKey.Of(@"C:\配置\Item.xlsx") != FileKey.Of(@"C:\配置\Level.xlsx"));

        Check("空路径得到空键", FileKey.Of("   ").Length == 0);
    }

    // ---------------- 命令行解析 ----------------

    private static void CaseParse()
    {
        Console.WriteLine("命令行解析：");

        var view = OpenRequest.FromCommandLine(new[] { @"test\Item.xlsx" });
        Check("单个文件 → 查看请求", view is { Kind: OpenKind.View });
        Check("请求里带的是完整路径",
            view is not null && view.Path == FileKey.Of(Path.Combine(Environment.CurrentDirectory, @"test\Item.xlsx")),
            view?.Path);

        Check("没有参数 → 只定位窗口",
            OpenRequest.FromCommandLine(Array.Empty<string>()) is { Kind: OpenKind.None });

        Check("--edit → 用副本编辑",
            OpenRequest.FromCommandLine(new[] { "--edit", @"test\Item.xlsx" }) is { Kind: OpenKind.Edit });

        Check("--diff 不参与去重",
            OpenRequest.FromCommandLine(new[] { "--diff", @"test\Item.xlsx", @"test\Level.xlsx" }) is null);

        Check("--diff-svn 不参与去重",
            OpenRequest.FromCommandLine(new[] { "--diff-svn", "-L", "a", "-L", "b", "c" }) is null);

        Check("--selftest 不参与去重",
            OpenRequest.FromCommandLine(new[] { "--selftest", @"test\Item.xlsx" }) is null);

        Check("带空格的中文路径不被拆开",
            OpenRequest.FromCommandLine(new[] { @"D:\配表\Item 列表.xlsx" }) is { Kind: OpenKind.View, Path: var spaced }
            && spaced == FileKey.Of(@"D:\配表\Item 列表.xlsx"));
    }

    // ---------------- 请求编码 ----------------

    private static void CaseWire()
    {
        Console.WriteLine("请求编码：");

        const string tricky = @"D:\配表\Item 列表 (1).xlsx";
        var decoded = OpenRequest.Decode(OpenRequest.View(tricky).Encode());
        Check("查看请求能原样解回来",
            decoded is { Kind: OpenKind.View } && decoded.Path == FileKey.Of(tricky),
            decoded?.Path);

        Check("编辑请求也能往返",
            OpenRequest.Decode(OpenRequest.Edit(tricky).Encode()) is { Kind: OpenKind.Edit });

        Check("空请求能往返",
            OpenRequest.Decode(OpenRequest.None.Encode()) is { Kind: OpenKind.None });

        Check("坏数据解出来是 null（不抛异常）",
            OpenRequest.Decode("view\t这不是 base64") is null
            && OpenRequest.Decode("") is null
            && OpenRequest.Decode(null) is null
            && OpenRequest.Decode("what\t" + Convert.ToBase64String(Encoding.UTF8.GetBytes("x"))) is null);
    }

    // ---------------- 窗口落点 ----------------

    private static void CaseDispatch()
    {
        Console.WriteLine("窗口落点：");

        var tracker = new ViewerWindows();
        var created = new List<FakeWindow>();
        FakeWindow Create()
        {
            var window = new FakeWindow();
            created.Add(window);
            return window;
        }

        tracker.Dispatch(OpenRequest.View(@"C:\配置\Item.xlsx"), Create);
        Check("第一次打开会开出一个窗口", created.Count == 1 && tracker.Count == 1);

        tracker.Dispatch(OpenRequest.View(@"C:\配置\Item.xlsx"), Create);
        Check("再打开同一个表不再开窗口", created.Count == 1 && tracker.Count == 1);
        Check("再打开同一个表会把原窗口提到前台", created[0].BringToFrontCount == 2, created[0].BringToFrontCount.ToString());
        Check("再打开同一个表不会重新加载", created[0].Calls.Count == 1, string.Join(" / ", created[0].Calls));

        tracker.Dispatch(OpenRequest.View(@"c:\配置\item.XLSX"), Create);
        Check("大小写不同的同一个表也算同一个窗口", created.Count == 1);

        tracker.Dispatch(OpenRequest.View(@"C:\配置\Level.xlsx"), Create);
        Check("换一个表会开新窗口", created.Count == 2 && tracker.Count == 2);
        Check("最近用过的窗口是最后打开的那个", ReferenceEquals(tracker.MostRecent, created[1]));

        tracker.Dispatch(OpenRequest.None, Create);
        Check("不带文件启动会定位到最近用过的窗口，不再开空窗口",
            created.Count == 2 && created[1].BringToFrontCount == 2);

        tracker.Remove(created[0]);
        Check("窗口关掉后不再参与去重", tracker.Count == 1);
        tracker.Dispatch(OpenRequest.View(@"C:\配置\Item.xlsx"), Create);
        Check("关掉的窗口不会复活，会重新开一个", created.Count == 3);

        // 启动时先开了个空窗口（拖放提示），随后双击文件：应该直接用这个空窗口
        var emptyTracker = new ViewerWindows();
        var emptyCreated = new List<FakeWindow>();
        FakeWindow CreateEmpty()
        {
            var window = new FakeWindow();
            emptyCreated.Add(window);
            return window;
        }

        emptyTracker.Dispatch(OpenRequest.None, CreateEmpty);
        emptyTracker.Dispatch(OpenRequest.View(@"C:\配置\Item.xlsx"), CreateEmpty);
        Check("启动时的空窗口会被直接用来显示文件，不再多开一个空壳",
            emptyCreated.Count == 1 && emptyCreated[0].Calls.Count == 1,
            string.Join(" / ", emptyCreated[0].Calls));

        // 正在看改动比对的窗口（DocumentKey 为空）不能当成空窗口被占用
        var diffTracker = new ViewerWindows();
        var diffWindow = new FakeWindow { IsShowingDiff = true };
        diffTracker.Add(diffWindow);
        var diffCreated = new List<FakeWindow>();
        FakeWindow CreateDiff()
        {
            var window = new FakeWindow();
            diffCreated.Add(window);
            return window;
        }

        diffTracker.Dispatch(OpenRequest.View(@"C:\配置\Item.xlsx"), CreateDiff);
        Check("正在看改动比对的窗口不会被当成空窗口占用",
            diffCreated.Count == 1 && diffWindow.Calls.Count == 0);

        // 同一个表的「用副本编辑」要落在原窗口上，不能再复制一份副本出去
        var editTracker = new ViewerWindows();
        var editCreated = new List<FakeWindow>();
        FakeWindow CreateEdit()
        {
            var window = new FakeWindow();
            editCreated.Add(window);
            return window;
        }

        editTracker.Dispatch(OpenRequest.View(@"C:\配置\Item.xlsx"), CreateEdit);
        editTracker.Dispatch(OpenRequest.Edit(@"C:\配置\Item.xlsx"), CreateEdit);
        Check("同一个表的「用副本编辑」落在原窗口上",
            editCreated.Count == 1
            && editCreated[0].Calls.Count == 2
            && editCreated[0].Calls[1].StartsWith("edit", StringComparison.Ordinal),
            string.Join(" / ", editCreated[0].Calls));
    }

    // ---------------- 管道转交（端到端） ----------------

    private static void CasePipe()
    {
        Console.WriteLine("进程间转交：");

        var suffix = Guid.NewGuid().ToString("N");
        var mutexName = @"Local\ExcelViewer.SelfTest." + suffix;
        var pipeName = "ExcelViewer.SelfTest." + suffix;

        using var host = new SingleInstance(mutexName, pipeName);
        Check("第一个进程抢到主持权", host.TryBecomeHost());

        // 互斥体的所有权是按线程算的（同一个线程自己再抢会成功），所以要另起一个线程来验
        var otherOwns = true;
        var other = new Thread(() =>
        {
            using var secondHost = new SingleInstance(mutexName, pipeName);
            otherOwns = secondHost.TryBecomeHost();
        });
        other.Start();
        other.Join();
        Check("主持权在会话内唯一", !otherOwns);

        OpenRequest? received = null;
        var got = new ManualResetEventSlim();
        host.StartHosting(request =>
        {
            received = request;
            got.Set();
        });

        using var client = new SingleInstance(mutexName, pipeName);
        var first = OpenRequest.View(@"C:\配置\Item.xlsx");
        Check("请求能交给主持进程", client.TryHandoff(first, attempts: 20, connectMs: 200));

        // 主持进程把请求转到 UI 线程处理；这里泵一下消息循环，让那一步真的跑到
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(3);
        while (!got.IsSet && DateTime.UtcNow < deadline)
        {
            Pump(50);
        }

        Check("主持进程收到了请求", received is not null);
        Check("收到的是同一个文件", received is not null && received.Kind == OpenKind.View && received.Path == first.Path,
            received?.Path);

        // 一条连接收完就应该回到监听状态：第二条请求也要能进来
        got.Reset();
        received = null;
        var second = OpenRequest.Edit(@"C:\配置\Level.xlsx");
        Check("第二条请求也能交过去（监听没有停在第一条上）",
            client.TryHandoff(second, attempts: 20, connectMs: 200));

        deadline = DateTime.UtcNow + TimeSpan.FromSeconds(3);
        while (!got.IsSet && DateTime.UtcNow < deadline)
        {
            Pump(50);
        }

        Check("第二条请求的类型和路径都对",
            received is not null && received.Kind == OpenKind.Edit && received.Path == second.Path,
            received?.Path);

        using var lonely = new SingleInstance(
            @"Local\ExcelViewer.SelfTest." + Guid.NewGuid().ToString("N"),
            "ExcelViewer.SelfTest." + Guid.NewGuid().ToString("N"));
        Check("没有主持进程时转交失败（不会挂住）",
            !lonely.TryHandoff(first, attempts: 2, connectMs: 50));
    }

    // ---------------- 工具 ----------------

    /// <summary>假窗口：只记下"被要求做什么"，不碰 WPF。</summary>
    private sealed class FakeWindow : IViewerWindow
    {
        private readonly List<string> _calls = new();

        public string? DocumentKey { get; private set; }

        public bool IsShowingDiff { get; set; }

        public int BringToFrontCount { get; private set; }

        public IReadOnlyList<string> Calls => _calls;

        public void OpenPath(string path)
        {
            DocumentKey = path;
            _calls.Add("open " + Path.GetFileName(path));
        }

        public void StartEdit(string path)
        {
            DocumentKey = path;
            _calls.Add("edit " + Path.GetFileName(path));
        }

        public void BringToFront() => BringToFrontCount++;
    }

    /// <summary>泵一下消息循环，好让 BeginInvoke 过来的请求真的被执行。</summary>
    private static void Pump(int ms)
    {
        var frame = new DispatcherFrame();
        var timer = new DispatcherTimer(
            TimeSpan.FromMilliseconds(ms),
            DispatcherPriority.Background,
            (_, _) => frame.Continue = false,
            Dispatcher.CurrentDispatcher);
        timer.Start();
        Dispatcher.PushFrame(frame);
        timer.Stop();
    }

    private static void Check(string what, bool ok, string? detail = null)
    {
        var note = ok || string.IsNullOrEmpty(detail) ? string.Empty : "  → " + detail;
        Console.WriteLine($"  [{(ok ? "OK" : "FAIL")}] {what}{note}");
        if (!ok)
        {
            _failures++;
        }
    }
}
