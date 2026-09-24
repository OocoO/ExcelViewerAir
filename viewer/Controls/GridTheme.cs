using ExcelViewer.Model;

namespace ExcelViewer.Controls;

/// <summary>
/// 配色集中管理，方便统一调成和参考截图接近的观感。
///
/// 外壳（菜单/工具条/状态栏）的令牌在 Styles.xaml，这里只管表格本身。
/// 两边的取向不同：外壳走 Fluent 的浅淡叠色，表格是给人长时间盯配表用的，
/// 正文对比度和网格线的可辨识度都要保留，不要跟着外壳一起调淡。
/// </summary>
public static class GridTheme
{
    public static readonly System.Windows.Media.Brush HeaderBackground = Frozen("#FFF3F3F3");
    public static readonly System.Windows.Media.Brush HeaderHoverBackground = Frozen("#FFE9E9E9");
    /// <summary>真正的字段名行（识别出的表头行）：比普通行更醒目一点。</summary>
    public static readonly System.Windows.Media.Brush DetectedHeaderBackground = Frozen("#FFEAF1F9");
    public static readonly System.Windows.Media.Brush GutterBackground = Frozen("#FFF3F3F3");
    public static readonly System.Windows.Media.Brush RowBackground = Frozen("#FFFFFFFF");
    public static readonly System.Windows.Media.Brush RowAltBackground = Frozen("#FFFAFAFA");
    public static readonly System.Windows.Media.Brush SearchRowBackground = Frozen("#FFFFF7D6");
    public static readonly System.Windows.Media.Brush SelectionBackground = Frozen("#FFEDF4FC");
    /// <summary>选中框用 Fluent 强调色（和 Styles.xaml 的 AccentBrush 同一档）。</summary>
    public static readonly System.Windows.Media.Brush SelectionBorder = Frozen("#FF005FB8");
    public static readonly System.Windows.Media.Brush EmptyAreaBackground = Frozen("#FFF3F3F3");
    public static readonly System.Windows.Media.Brush GridLine = Frozen("#FFE3E3E3");
    public static readonly System.Windows.Media.Brush HeaderLine = Frozen("#FFCFCFCF");
    public static readonly System.Windows.Media.Brush HeaderText = Frozen("#FF3B3B3B");
    public static readonly System.Windows.Media.Brush CellText = Frozen("#FF202020");
    /// <summary>
    /// 行号栏文字。注意：这里刻意保持在 0x99（亮于 0x90）—— `--layout` 探针
    /// 靠"每行最左侧深色像素（三通道都 &lt; 0x90）"判断正文有没有被行号栏裁掉，
    /// 行号文字一旦画深就会混进这个判定，把它改深会直接破坏那条自检。
    /// </summary>
    public static readonly System.Windows.Media.Brush MutedText = Frozen("#FF999999");
    public static readonly System.Windows.Media.Brush MatchHighlight = Frozen("#FFFFE96B");
    public static readonly System.Windows.Media.Brush ActiveMatchBorder = Frozen("#FFE8A33D");
    public static readonly System.Windows.Media.Brush HeaderSortMark = Frozen("#FF7A7A7A");

    private static System.Windows.Media.Brush Frozen(string hex)
    {
        var brush = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)
            System.Windows.Media.ColorConverter.ConvertFromString(hex));
        brush.Freeze();
        return brush;
    }
}
