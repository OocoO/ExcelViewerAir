using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Windows;

namespace ExcelViewer;

/// <summary>
/// 文件身份：把「同一个表」的各种写法归一成同一个键。
///
/// 归一化做三件事：去掉首尾引号与空格、按当前目录补成完整路径、按文件自身的符号链接指向真实文件。
/// 不做的事：8.3 短文件名（<c>PROGRA~1</c>）与父目录上的 junction 都不展开 ——
/// 那要按文件句柄取「卷 + 文件 ID」才准；配表场景里同一个文件基本就是同一条路径，
/// 先用路径比，判不出来的情况只是"没去掉重"，不会张冠李戴。
/// </summary>
internal static class FileKey
{
    /// <summary>归一路径。同一个文件的两种写法会得到相同的字符串（大小写由比较方决定）。</summary>
    public static string Of(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        var full = path.Trim().Trim('"');

        try
        {
            full = Path.GetFullPath(full);
        }
        catch (Exception)
        {
            // 非法路径（超长、含非法字符）就按原样当键：至少它和自己能对上
        }

        try
        {
            if (File.Exists(full) && File.ResolveLinkTarget(full, returnFinalTarget: true) is { } target)
            {
                full = Path.GetFullPath(target.FullName);
            }
        }
        catch (Exception)
        {
            // 解析链接失败不影响用路径本身当键
        }

        return full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}

/// <summary>一次「打开」请求的种类。</summary>
internal enum OpenKind
{
    /// <summary>没带文件（双击/命令行空跑）：定位到最近用过的窗口。</summary>
    None,

    /// <summary>只读查看。</summary>
    View,

    /// <summary>用副本编辑（复制到工作区 → 用 Excel 打开副本 → 关掉后弹改动比对）。</summary>
    Edit,
}

/// <summary>
/// 一次「打开」请求。既用于启动时决定开什么窗口，也是进程之间转交的内容。
/// 路径在构造时就按 <see cref="FileKey"/> 归一：转交给别的进程后，那边的工作目录未必和这里一样，
/// 传过去必须是完整路径，否则两边算出来的键对不上，去重也就失效了。
/// </summary>
internal sealed class OpenRequest
{
    private OpenRequest(OpenKind kind, string path)
    {
        Kind = kind;
        Path = path;
    }

    public OpenKind Kind { get; }

    /// <summary>要打开的文件（完整路径）；<see cref="OpenKind.None"/> 时是空串。</summary>
    public string Path { get; }

    /// <summary>空请求：只定位窗口，不打开任何东西。</summary>
    public static OpenRequest None { get; } = new(OpenKind.None, string.Empty);

    public static OpenRequest View(string path) => new(OpenKind.View, FileKey.Of(path));

    public static OpenRequest Edit(string path) => new(OpenKind.Edit, FileKey.Of(path));

    /// <summary>
    /// 从命令行认出「这次要打开哪个表」。
    ///
    /// 返回 null 表示这次启动不参与去重：自检 / 探针 / 改动比对都各开各的进程。
    /// 其中 <c>--diff-svn</c> 尤其不能转交 —— svn 和 TortoiseSVN 会等外部 diff 程序退出，
    /// 转手给一个常驻进程会让它们以为比对瞬间就结束了。
    /// </summary>
    public static OpenRequest? FromCommandLine(string[] args)
    {
        for (var i = 0; i < args.Length; i++)
        {
            if (!args[i].StartsWith('-'))
            {
                continue;
            }

            if (!args[i].Equals("--edit", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var target = args.Skip(i + 1).FirstOrDefault(a => !a.StartsWith('-'));
            return target is null ? None : Edit(target);
        }

        var file = args.FirstOrDefault(a => !a.StartsWith('-') && a.Length > 0);
        return file is null ? None : View(file);
    }

    /// <summary>编码成一行文本。路径走 base64，省得为制表符/换行做转义。</summary>
    public string Encode() => Kind switch
    {
        OpenKind.None => "none",
        OpenKind.Edit => "edit\t" + Base64(Path),
        _ => "view\t" + Base64(Path),
    };

    /// <summary>解一行文本；认不出来就返回 null（管道另一头是别的版本/半截数据时不要崩）。</summary>
    public static OpenRequest? Decode(string? line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return null;
        }

        var parts = line.Split('\t');
        var kind = parts[0].Trim();
        if (kind.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return None;
        }

        if (parts.Length < 2)
        {
            return null;
        }

        string path;
        try
        {
            path = Encoding.UTF8.GetString(Convert.FromBase64String(parts[1]));
        }
        catch (FormatException)
        {
            return null;
        }

        if (kind.Equals("view", StringComparison.OrdinalIgnoreCase))
        {
            return new OpenRequest(OpenKind.View, FileKey.Of(path));
        }

        return kind.Equals("edit", StringComparison.OrdinalIgnoreCase)
            ? new OpenRequest(OpenKind.Edit, FileKey.Of(path))
            : null;
    }

    private static string Base64(string text) => Convert.ToBase64String(Encoding.UTF8.GetBytes(text));
}

/// <summary>
/// 单实例：让「同一个表重复打开」落到同一个窗口上。
///
/// 做法是一个命名互斥体加一条命名管道：
///   * 第一个启动的进程抢到互斥体，成为「主持进程」，开一条管道开始接活；
///   * 后来启动的进程抢不到互斥体，就把「要打开哪个文件」写成一行交给主持进程，然后自己退出。
/// 于是窗口都活在主持进程里：同一个文件再次打开时只做「提到前台」，不同文件才新开窗口。
///
/// 三个必须守住的点：
///   1. 抢不到主持权又连不上通道（对方正在退出/卡住）时必须能自己开窗，不能什么都不弹；
///   2. 请求必须在 UI 线程上处理（开窗、定位都得在 UI 线程）；
///   3. 自检 / 探针 / 改动比对一律不参与，否则 <c>--probe</c> 之类会把自己的窗口塞给别人才跑不下去。
/// </summary>
internal sealed class SingleInstance : IDisposable
{
    /// <summary>主持进程的互斥体。Local\ 作用域＝按登录会话隔离，多用户各开各的互不干扰。</summary>
    public const string HostMutexName = @"Local\ExcelViewer.Host.v1";

    /// <summary>请求通道的名字。</summary>
    public const string RequestPipeName = "ExcelViewer.Open.v1";

    private readonly Mutex _mutex;
    private readonly string _pipeName;
    private readonly CancellationTokenSource _stop = new();

    private bool _ownsMutex;
    private bool _disposed;

    public SingleInstance(string mutexName = HostMutexName, string pipeName = RequestPipeName)
    {
        _pipeName = pipeName;
        _mutex = new Mutex(initiallyOwned: false, mutexName);
    }

    /// <summary>本进程是不是主持进程。</summary>
    public bool IsHost => _ownsMutex;

    /// <summary>
    /// 抢主持进程的身份。抢到就由本进程开通道、接收后来者交过来的请求；抢不到返回 false。
    /// </summary>
    public bool TryBecomeHost()
    {
        if (_ownsMutex)
        {
            return true;
        }

        if (_disposed)
        {
            return false;
        }

        try
        {
            _ownsMutex = _mutex.WaitOne(0);
        }
        catch (AbandonedMutexException)
        {
            // 上一个主持进程被强杀：互斥体是「被遗弃」状态，所有权已经转过来了
            _ownsMutex = true;
        }

        return _ownsMutex;
    }

    /// <summary>开始接收请求。<paramref name="handle"/> 在 UI 线程上被调用。</summary>
    public void StartHosting(Action<OpenRequest> handle)
    {
        if (!_ownsMutex || _disposed)
        {
            return;
        }

        var thread = new Thread(() => AcceptLoop(handle))
        {
            IsBackground = true,
            Name = "ExcelViewer 单实例请求",
        };
        thread.Start();
    }

    /// <summary>
    /// 把请求交给已经在跑的那个查看器。成功返回 true，本进程随即应该退出。
    /// 连不上（对方还没起来 / 正在退出 / 卡住）返回 false，由调用方决定怎么降级。
    /// </summary>
    public bool TryHandoff(OpenRequest request, int attempts = 12, int connectMs = 150)
    {
        var payload = Encoding.UTF8.GetBytes(request.Encode() + "\n");

        for (var i = 0; i < attempts; i++)
        {
            try
            {
                using var pipe = new NamedPipeClientStream(".", _pipeName, PipeDirection.Out);
                pipe.Connect(connectMs);
                pipe.Write(payload, 0, payload.Length);
                pipe.Flush();
                return true;
            }
            catch (TimeoutException)
            {
                // 主持进程还在启动：等一下再连
            }
            catch (UnauthorizedAccessException)
            {
                // 权限层面就打不开（例如对方以管理员身份在跑）：重试也没用，直接降级
                return false;
            }
            catch (IOException)
            {
                // 通道实例被占满 / 主持进程正在退出
            }
            catch (Exception)
            {
                // 通道上的任何意外都不能把启动搞崩：降级成"自己开一个窗口"
                return false;
            }

            Thread.Sleep(60);
        }

        return false;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        try
        {
            _stop.Cancel();
        }
        catch (ObjectDisposedException)
        {
        }

        if (_ownsMutex)
        {
            try
            {
                _mutex.ReleaseMutex();
            }
            catch (ApplicationException)
            {
                // 不是持锁线程（或已经释放）：进程就要退了，剩下的交给系统
            }
            catch (ObjectDisposedException)
            {
            }

            _ownsMutex = false;
        }

        _mutex.Dispose();
    }

    /// <summary>
    /// 一条连接只收一行请求。每条连接都新建一个管道实例，
    /// 这样上一条连接出什么错都不会把下一条带坏。
    /// </summary>
    private void AcceptLoop(Action<OpenRequest> handle)
    {
        while (!_stop.IsCancellationRequested)
        {
            try
            {
                using var pipe = new NamedPipeServerStream(
                    _pipeName,
                    PipeDirection.In,
                    maxNumberOfServerInstances: 1,
                    PipeTransmissionMode.Byte,
                    // 默认 DACL 就是「当前用户 + SYSTEM + 管理员」，别的登录用户连不进来，
                    // 不需要再加 PipeOptions.CurrentUserOnly
                    PipeOptions.None);

                pipe.WaitForConnection();

                string? line;
                using (var reader = new StreamReader(pipe, new UTF8Encoding(false)))
                {
                    // 客户端写完一行就关；读到一半断掉时拿到的半截数据由 Decode 兜住
                    line = reader.ReadLine();
                }

                if (OpenRequest.Decode(line) is { } request)
                {
                    RunOnUiThread(handle, request);
                }
            }
            catch (ObjectDisposedException)
            {
                break;
            }
            catch (IOException)
            {
                // 客户端中途断开：换一条新连接继续接活
            }
            catch (Exception)
            {
                // 单个连接出问题不能让整条通道停摆
                Thread.Sleep(50);
            }
        }
    }

    private static void RunOnUiThread(Action<OpenRequest> handle, OpenRequest request)
    {
        var app = Application.Current;
        if (app is null)
        {
            handle(request);
            return;
        }

        // 请求来自后台线程，开窗与定位都必须在 UI 线程上做
        app.Dispatcher.BeginInvoke(new Action(() => handle(request)));
    }
}
