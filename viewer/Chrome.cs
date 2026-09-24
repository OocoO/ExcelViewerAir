using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.Win32;

namespace ExcelViewer;

/// <summary>
/// Win11 窗口外观：圆角 + 云母（Mica）背景 + 标题栏配色。
///
/// 这一层只碰「窗口外壳」，不碰任何内容绘制 —— 表格仍然是 WPF 自绘（ExcelGrid），
/// 性能和自检体系（--probe / --layout / --png-window）都不受影响。
///
/// 三个必须守住的点：
///   1. 系统版本不够（&lt; Win11 22000）直接返回，不要留下半套状态；
///   2. 用户把「透明效果」关掉时不能挂 Mica —— 否则窗口底色让出去以后
///      拿不到材质，会得到一块黑窗口（这个失败模式比"没有云母"难看得多）；
///   3. 任何一个 DWM 调用失败都只回退自己那一项，绝不抛出去影响启动。
/// </summary>
internal static class Chrome
{
    /// <summary>DwmSetWindowAttribute: 深色标题栏。</summary>
    private const int DwmwaUseImmersiveDarkMode = 20;

    /// <summary>DwmSetWindowAttribute: 窗口圆角偏好。</summary>
    private const int DwmwaWindowCornerPreference = 33;

    /// <summary>DwmSetWindowAttribute: 系统背景材质（Mica / Acrylic / Tabbed）。</summary>
    private const int DwmwaSystemBackdropType = 38;

    /// <summary>DWMWCP_ROUND：Win11 标准圆角。</summary>
    private const int DwmwcpRound = 2;

    /// <summary>DWMSBT_MAINWINDOW：Mica（主窗口用的那档）。</summary>
    private const int DwmsbtMainWindow = 2;

    /// <summary>圆角是否已生效（纯诊断用，--probe 会打印）。</summary>
    public static bool Rounded { get; private set; }

    /// <summary>云母是否已生效。为 true 时内容区底色才会让出去透出材质。</summary>
    public static bool MicaApplied { get; private set; }

    /// <summary>
    /// 给窗口套上 Win11 外观。必须在 HWND 建好之后调用（<see cref="Window.SourceInitialized"/>）。
    /// </summary>
    public static void Apply(Window window, bool dark = false)
    {
        Rounded = false;
        MicaApplied = false;

        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        // Win10 及更早：既没有圆角也没有系统背景材质，整套跳过
        if (Environment.OSVersion.Version.Build < 22000)
        {
            return;
        }

        var hwnd = new WindowInteropHelper(window).Handle;
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        Rounded = TrySet(hwnd, DwmwaWindowCornerPreference, DwmwcpRound);

        // 深色标题栏跟随应用主题（浅色主题下这个调用是无害的）
        TrySet(hwnd, DwmwaUseImmersiveDarkMode, dark ? 1 : 0);

        // 只有「Win11 22H2+ 且系统透明效果开着」才挂云母，否则让出底色的那一刻就会露黑
        if (MicaEnabled && Environment.OSVersion.Version.Build >= 22621 && TransparencyEnabled())
        {
            MicaApplied = TrySet(hwnd, DwmwaSystemBackdropType, DwmsbtMainWindow);
        }

        if (MicaApplied)
        {
            // 让出窗口底色，材质才透得出来。内容层自己有底色，表格区不会跟着变透明。
            window.Background = Brushes.Transparent;
        }
    }

    /// <summary>把外壳状态恢复成"没有材质"，并把底色还回来（回退路径用）。</summary>
    public static void ResetToSolid(Window window, Brush solidBackground)
    {
        MicaApplied = false;
        Rounded = false;
        window.Background = solidBackground;
    }

    private static bool TrySet(IntPtr hwnd, int attribute, int value)
    {
        try
        {
            // HRESULT == 0 才算真的生效；属性不被识别时会返回 E_INVALIDARG
            return DwmSetWindowAttribute(hwnd, attribute, ref value, sizeof(int)) == 0;
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
    /// 云母材质开关，**默认关**。
    ///
    /// 关掉的原因：本应用的形态是「内容铺满 + 上下两条外壳带承载文字」，
    /// 而外壳带为了让材质透上来做成了半透明 —— 于是它的实际颜色变成
    /// `0.6 × 浅灰 + 0.4 × 桌面壁纸`：壁纸深的地方这两条就发灰发暗，
    /// 而且窗口在屏幕上挪个位置颜色还会跟着变，同一份界面看着不一样。
    /// 实测下来这个不确定性不值得，稳定的不透明底色观感更好。
    ///
    /// 想试材质：设环境变量 <c>EXCELVIEWER_MICA=1</c> 再启动。
    /// </summary>
    private static bool MicaEnabled =>
        Environment.GetEnvironmentVariable("EXCELVIEWER_MICA") is "1" or "true" or "TRUE";

    /// <summary>
    /// 读用户的「透明效果」开关（设置 → 个性化 → 颜色）。
    /// 键不存在时系统默认是开的，所以只有明确读到 0 才判定为关。
    /// </summary>
    private static bool TransparencyEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");

            return key?.GetValue("EnableTransparency") is not int value || value != 0;
        }
        catch
        {
            // 读不到就当作关闭：宁可没有云母，也不要一块黑窗口
            return false;
        }
    }

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);
}
