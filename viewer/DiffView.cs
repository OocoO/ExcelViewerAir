using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using ExcelViewer.Model;

namespace ExcelViewer;

/// <summary>
/// 改动比对视图：两个工作簿比一轮，把"改动的格子"摊成一张清单表显示。
///
/// 编辑走 temp 副本 → 保存 → 在这里看改了哪些格子 → 再决定是否写回原文件，
/// 这样"覆盖原文件"永远是一次显式确认，而不是保存时的副作用。
/// </summary>
internal sealed class DiffView
{
    private readonly Window _owner;
    private readonly Grid _gridHost;
    private readonly Border _banner;
    private readonly TextBlock _bannerText;
    private readonly TextBlock _detail;
    private readonly StackPanel _tabs;
    private readonly Button _writeBackButton;
    private readonly Button _closeButton;
    private readonly Action<string> _onWriteBack;

    private DiffReport? _report;
    private List<Sheet> _changeSheets = new();
    private int _current = -1;
    private readonly Controls.ExcelGrid _grid = new();
    private readonly StringPool _pool = new();

    public DiffView(
        Window owner,
        Grid gridHost,
        Border banner,
        TextBlock bannerText,
        TextBlock detail,
        StackPanel tabs,
        Button writeBackButton,
        Button closeButton,
        Action<string> onWriteBack)
    {
        _owner = owner;
        _gridHost = gridHost;
        _banner = banner;
        _bannerText = bannerText;
        _detail = detail;
        _tabs = tabs;
        _writeBackButton = writeBackButton;
        _closeButton = closeButton;
        _onWriteBack = onWriteBack;

        // 比对表格是盖在主表格上面的"第二层"，默认必须是隐藏的：
        // 否则它（一张空表）会一直挡在正常查看的表格前面，表现为"数据加载了但界面是空的"。
        _grid.Visibility = Visibility.Collapsed;
        _gridHost.Children.Add(_grid);
    }

    public bool IsActive { get; private set; }

    /// <summary>把值写回后用于显示的文件路径（新文件）。</summary>
    public string? NewPath => _report?.NewPath;

    public void Show(DiffReport report, string oldLabel, string newLabel)
    {
        _report = report;
        IsActive = true;
        _pool.Clear();

        _gridHost.Visibility = Visibility.Visible;
        _grid.Visibility = Visibility.Visible;
        _banner.Visibility = Visibility.Visible;
        _writeBackButton.Visibility = report.HasChanges ? Visibility.Visible : Visibility.Collapsed;
        _closeButton.Visibility = Visibility.Visible;

        _bannerText.Text = report.HasChanges
            ? $"发现改动：{report.Summary}"
            : "两个文件内容一致，没有发现改动。";
        _bannerText.Foreground = report.HasChanges
            ? new SolidColorBrush(Color.FromRgb(0xB4, 0x5A, 0x00))
            : new SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0x32));

        var lines = new List<string>
        {
            $"原文件：{Path.GetFileName(oldLabel)}",
            $"新文件：{Path.GetFileName(newLabel)}",
        };

        if (report.RemovedSheets.Count > 0)
        {
            lines.Add("新文件缺少工作表：" + string.Join("、", report.RemovedSheets));
        }

        if (report.AddedSheets.Count > 0)
        {
            lines.Add("新文件新增工作表：" + string.Join("、", report.AddedSheets));
        }

        lines.Add($"比对耗时 {report.ElapsedMs} ms");
        _detail.Text = string.Join("    |    ", lines);

        BuildChangeSheets();
        BuildTabs();
        Show(0);
    }

    private void BuildChangeSheets()
    {
        _changeSheets = new List<Sheet>();
        var report = _report;
        if (report is null)
        {
            return;
        }

        foreach (var d in report.Sheets.Where(s => s.HasChanges))
        {
            if (d.Changes.Count > 0)
            {
                var sheet = DiffService.BuildChangeSheet(d.SheetName, d.Changes, _pool, Controls.ExcelGrid.ColumnName);
                sheet.Title = $"{d.SheetName} · {d.Changes.Count} 处";
                _changeSheets.Add(sheet);
            }

            // 列数变化（新增/缺失整列）不在逐格比对里体现，单独给一条提示
            if (d.RemovedColumns.Count > 0 || d.AddedColumns.Count > 0)
            {
                var note = new Sheet($"{d.SheetName} · 列变化", _pool);
                note.BeginRow();
                note.AddCell("列变化");
                note.AddCell("说明");
                note.EndRow();
                if (d.AddedColumns.Count > 0)
                {
                    note.BeginRow();
                    note.AddCell("新增列");
                    note.AddCell(string.Join(", ", d.AddedColumns.Select(Controls.ExcelGrid.ColumnName)));
                    note.EndRow();
                }

                if (d.RemovedColumns.Count > 0)
                {
                    note.BeginRow();
                    note.AddCell("缺失列");
                    note.AddCell(string.Join(", ", d.RemovedColumns.Select(Controls.ExcelGrid.ColumnName)));
                    note.EndRow();
                }

                note.Finish();
                note.Title = $"{d.SheetName} · 列变化";
                _changeSheets.Add(note);
            }
        }
    }

    private void BuildTabs()
    {
        _tabs.Children.Clear();
        for (var i = 0; i < _changeSheets.Count; i++)
        {
            var index = i;
            var tab = new ToggleButton
            {
                Content = _changeSheets[i].Title,
                Style = (Style)_owner.FindResource("SheetTab"),
                Tag = index,
                IsChecked = i == 0,
            };
            tab.Click += (_, _) => Show(index);
            _tabs.Children.Add(tab);
        }
    }

    public void Show(int index)
    {
        if (index < 0 || index >= _changeSheets.Count)
        {
            _grid.SetSheet(null, keepScroll: false);
            return;
        }

        _current = index;
        foreach (var child in _tabs.Children)
        {
            if (child is ToggleButton tb && tb.Tag is int i)
            {
                tb.IsChecked = i == index;
            }
        }

        var sheet = _changeSheets[index];
        _grid.SetSheet(sheet, keepScroll: false);
        _grid.DetectedHeaderRow = 0;
        _grid.Focus();
    }

    public void ScrollTab(int delta) => Show(Math.Clamp(_current + delta, 0, _changeSheets.Count - 1));

    public void Hide()
    {
        IsActive = false;
        _grid.Visibility = Visibility.Collapsed;
        // 只藏比对表；GridHost 是主表格的宿主，藏了它会让正常查看的表格一起消失
        _gridHost.Visibility = Visibility.Visible;
        _banner.Visibility = Visibility.Collapsed;
        _writeBackButton.Visibility = Visibility.Collapsed;
        _closeButton.Visibility = Visibility.Collapsed;
    }

    public void RequestWriteBack()
    {
        if (_report is null || !_report.HasChanges)
        {
            return;
        }

        _onWriteBack(_report.NewPath);
    }
}
