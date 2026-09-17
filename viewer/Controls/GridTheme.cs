using ExcelViewer.Model;

namespace ExcelViewer.Controls;

/// <summary>配色集中管理，方便统一调成和参考截图接近的观感。</summary>
public static class GridTheme
{
    public static readonly System.Windows.Media.Brush HeaderBackground = Frozen("#FFECECEC");
    public static readonly System.Windows.Media.Brush HeaderHoverBackground = Frozen("#FFE0E0E0");
    /// <summary>真正的字段名行（识别出的表头行）：比普通行更醒目一点。</summary>
    public static readonly System.Windows.Media.Brush DetectedHeaderBackground = Frozen("#FFE8EEF7");
    public static readonly System.Windows.Media.Brush GutterBackground = Frozen("#FFF6F6F6");
    public static readonly System.Windows.Media.Brush RowBackground = Frozen("#FFFFFFFF");
    public static readonly System.Windows.Media.Brush RowAltBackground = Frozen("#FFFAFAFA");
    public static readonly System.Windows.Media.Brush SearchRowBackground = Frozen("#FFFFF7D6");
    public static readonly System.Windows.Media.Brush SelectionBackground = Frozen("#FFEAF3FF");
    public static readonly System.Windows.Media.Brush SelectionBorder = Frozen("#FF2E6CC4");
    public static readonly System.Windows.Media.Brush EmptyAreaBackground = Frozen("#FFF3F3F3");
    public static readonly System.Windows.Media.Brush GridLine = Frozen("#FFDCDCDC");
    public static readonly System.Windows.Media.Brush HeaderLine = Frozen("#FFC0C0C0");
    public static readonly System.Windows.Media.Brush HeaderText = Frozen("#FF444444");
    public static readonly System.Windows.Media.Brush CellText = Frozen("#FF202020");
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
