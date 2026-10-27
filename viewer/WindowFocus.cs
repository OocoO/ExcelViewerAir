using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace ExcelViewer;

/// <summary>
/// 把窗口提到前台。
///
/// 为什么不能只调 <see cref="Window.Activate"/>：Windows 有一条「前台锁」，后台进程调用
/// <c>SetForegroundWindow</c>（WPF 的 Activate 底下就是它）会被静默忽略。
/// 于是接手请求的进程即使把窗口找对了，用户也可能什么都看不到——窗口确实被提到了 Z 序最前，
/// 但焦点还在别的程序上，任务栏按钮也不闪，看起来就像"双击了文件却没反应"。
///
/// 两道保险：
///   1. 转交请求的那个进程（用户刚点过、有前台权限）退出前用
///      <see cref="AllowAnyProcessForeground"/> 把权限让出来；
///   2. 还抢不到就借当前前台窗口的输入队列再试一次；再不行就闪一下任务栏按钮，
///      至少让人知道窗口在哪。
/// </summary>
internal static class WindowFocus
{
    /// <summary>ASFW_ANY：允许任意进程设置前台窗口。</summary>
    private const int AsfwAny = -1;

    private const uint FlashwTray = 0x00000002;

    /// <summary>一直闪，直到窗口被切到前台（否则闪一下没人看见就白搭）。</summary>
    private const uint FlashwTimerNoForeground = 0x0000000C;

    /// <summary>
    /// 让出「设置前台窗口」的权限。只有本身有前台权限的进程（用户刚点它启动的那个）调用才有效，
    /// 所以这步放在「转交请求之后、退出之前」。
    /// </summary>
    internal static void AllowAnyProcessForeground()
    {
        try
        {
            AllowSetForegroundWindow(AsfwAny);
        }
        catch (DllNotFoundException)
        {
        }
        catch (EntryPointNotFoundException)
        {
        }
    }

    /// <summary>提到前台并尽量拿到焦点，返回是否真的成了前台窗口。</summary>
    internal static bool BringToFront(Window window)
    {
        if (!window.IsVisible)
        {
            window.Show();
        }

        if (window.WindowState == WindowState.Minimized)
        {
            window.WindowState = WindowState.Normal;
        }

        var hwnd = new WindowInteropHelper(window).Handle;

        // 先让 WPF 自己排一次 Z 序（同进程内一定生效），跨进程的前台限制靠下面的 Win32 调用
        window.Activate();

        if (hwnd == IntPtr.Zero)
        {
            return false;
        }

        try
        {
            if (SetForegroundWindow(hwnd))
            {
                return true;
            }

            if (TryWithForegroundInputQueue(hwnd))
            {
                return true;
            }

            return Flash(hwnd) && GetForegroundWindow() == hwnd;
        }
        catch (DllNotFoundException)
        {
            return false;
        }
        catch (EntryPointNotFoundException)
        {
            return false;
        }
    }

    /// <summary>
    /// 借当前前台窗口的输入队列来调 SetForegroundWindow：共享输入队列后，前台锁不再拦我们。
    /// 用不成（拿不到前台窗口、Attach 失败）就算了，外面还有闪任务栏兜底。
    /// </summary>
    private static bool TryWithForegroundInputQueue(IntPtr hwnd)
    {
        var foreground = GetForegroundWindow();
        if (foreground == IntPtr.Zero)
        {
            return false;
        }

        var foregroundThread = GetWindowThreadProcessId(foreground, out _);
        var selfThread = GetCurrentThreadId();
        if (foregroundThread == 0 || foregroundThread == selfThread)
        {
            return false;
        }

        if (!AttachThreadInput(selfThread, foregroundThread, true))
        {
            return false;
        }

        try
        {
            return SetForegroundWindow(hwnd);
        }
        finally
        {
            AttachThreadInput(selfThread, foregroundThread, false);
        }
    }

    private static bool Flash(IntPtr hwnd)
    {
        var info = new FlashWindowInfo
        {
            Size = (uint)Marshal.SizeOf<FlashWindowInfo>(),
            Hwnd = hwnd,
            Flags = FlashwTray | FlashwTimerNoForeground,
            Count = 3,
        };

        return FlashWindowEx(ref info);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct FlashWindowInfo
    {
        public uint Size;
        public IntPtr Hwnd;
        public uint Flags;
        public uint Count;
        public uint Timeout;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AllowSetForegroundWindow(int processId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AttachThreadInput(uint attach, uint attachTo, [MarshalAs(UnmanagedType.Bool)] bool doAttach);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool FlashWindowEx(ref FlashWindowInfo info);
}
