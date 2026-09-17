// 截图工具：抓取指定进程主窗口的位图，用于验证自绘表格的真实渲染效果。
// 用法: dotnet run -c Release -- <进程名或PID> <输出png>
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

internal static class Shot
{
    private const int PW_RENDERFULLCONTENT = 0x00000002;

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    private static int Main(string[] args)
    {
        if (args.Length >= 6 && args[0] == "crop")
        {
            return Crop(
                args[1],
                args[2],
                int.Parse(args[3]),
                int.Parse(args[4]),
                int.Parse(args[5]),
                int.Parse(args[6]),
                args.Length > 7 ? int.Parse(args[7]) : 2);
        }

        if (args.Length < 2)
        {
            Console.Error.WriteLine("usage: Shot <processNameOrPid> <out.png>");
            Console.Error.WriteLine("       Shot crop <in.png> <out.png> <x> <y> <w> <h> [scale]");
            return 2;
        }

        Process? proc = int.TryParse(args[0], out var pid)
            ? Process.GetProcessById(pid)
            : Process.GetProcessesByName(args[0]).FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);

        if (proc is null)
        {
            Console.Error.WriteLine($"process not found: {args[0]}");
            return 1;
        }

        var hwnd = proc.MainWindowHandle;
        if (hwnd == IntPtr.Zero)
        {
            Console.Error.WriteLine("no main window handle");
            return 1;
        }

        if (!GetWindowRect(hwnd, out var rect))
        {
            Console.Error.WriteLine("GetWindowRect failed");
            return 1;
        }

        var w = rect.Right - rect.Left;
        var h = rect.Bottom - rect.Top;
        Console.WriteLine($"hwnd={hwnd} size={w}x{h} title={proc.MainWindowTitle}");

        using var bmp = new Bitmap(w, h);
        using (var g = Graphics.FromImage(bmp))
        {
            // WPF 默认走硬件加速，PrintWindow 常常只能拿到空白；
            // 直接按窗口屏幕坐标抓屏最可靠（需要窗口没有被遮挡）。
            g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(w, h), CopyPixelOperation.SourceCopy);
        }

        bmp.Save(args[1], ImageFormat.Png);
        Console.WriteLine($"saved {args[1]}");
        return 0;
    }

    /// <summary>放大裁剪：把某个区域放大保存，便于细看对齐与裁切问题。</summary>
    private static int Crop(string input, string output, int x, int y, int w, int h, int scale)
    {
        using var src = new Bitmap(input);
        x = Math.Clamp(x, 0, Math.Max(0, src.Width - 1));
        y = Math.Clamp(y, 0, Math.Max(0, src.Height - 1));
        w = Math.Min(w, src.Width - x);
        h = Math.Min(h, src.Height - y);

        using var dst = new Bitmap(w * scale, h * scale);
        using (var g = Graphics.FromImage(dst))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
            g.DrawImage(src, new Rectangle(0, 0, w * scale, h * scale), new Rectangle(x, y, w, h), GraphicsUnit.Pixel);

            // 每 50 像素画一条参考线，方便读数
            using var pen = new Pen(Color.FromArgb(120, 255, 0, 0));
            for (var gx = 0; gx < w; gx += 50)
            {
                g.DrawLine(pen, gx * scale, 0, gx * scale, h * scale);
            }

            for (var gy = 0; gy < h; gy += 50)
            {
                g.DrawLine(pen, 0, gy * scale, w * scale, gy * scale);
            }
        }

        dst.Save(output, ImageFormat.Png);
        Console.WriteLine($"saved {output} ({w}x{h} @{scale}x, 参考线每 50 源像素)");
        return 0;
    }
}
