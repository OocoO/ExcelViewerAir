using System.IO;
using System.Text.Json;
using ExcelViewer.Model;

namespace ExcelViewer;

/// <summary>
/// 「用副本编辑」流程的辅助逻辑。
///
/// 流程：把原文件复制到工作区 → 用 Excel/WPS 打开**副本** → 关闭后比对改动 → 用户确认才写回。
/// 原文件全程只被读一次，不会被编辑器锁定，也不会被直接改写。
/// （编排在 <see cref="MainWindow.StartEditCopy"/>，这里只放查找编辑器与 JSON 输出这类无 UI 逻辑。）
/// </summary>
public static class EditSession
{
    /// <summary>常见表格编辑器的可执行文件名。</summary>
    private static readonly string[] EditorCandidates = { "EXCEL.EXE", "et.exe", "wps.exe" };

    /// <summary>找系统里的表格编辑器，找不到返回 null。</summary>
    public static string? FindEditor()
    {
        foreach (var name in EditorCandidates)
        {
            var path = FindOnPath(name);
            if (path is not null)
            {
                return path;
            }
        }

        // App Paths 注册表（不依赖 PATH）
        foreach (var sub in new[]
                 {
                     @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
                     @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
                 })
        {
            try
            {
                using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(sub);
                if (key?.GetValue(null) is string p && File.Exists(p))
                {
                    return p;
                }
            }
            catch (Exception)
            {
                // 注册表不可读就跳过
            }
        }

        foreach (var guess in new[]
                 {
                     @"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE",
                     @"C:\Program Files (x86)\Microsoft Office\Root\Office16\EXCEL.EXE",
                 })
        {
            if (File.Exists(guess))
            {
                return guess;
            }
        }

        return null;
    }

    private static string? FindOnPath(string exeName)
    {
        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var dir in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            try
            {
                var candidate = Path.Combine(dir.Trim('"'), exeName);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            catch (Exception)
            {
                // 非法路径项直接跳过
            }
        }

        return null;
    }

    /// <summary>
    /// 无界面比对：把结果输出成 JSON，供脚本判断"到底有没有改动"。
    /// 因为这是 WinExe，stdout 在有些启动方式下拿不到，所以支持 --out= 写文件。
    /// 退出码 1 表示有改动，0 表示一致，2 表示出错。
    /// </summary>
    public static int DumpJson(string oldPath, string newPath, string? outPath = null)
    {
        string json;
        int exitCode;
        try
        {
            var report = DiffService.Compare(oldPath, newPath);
            json = JsonSerializer.Serialize(
                new
                {
                    hasChanges = report.HasChanges,
                    summary = report.Summary,
                    totalChangedCells = report.TotalChangedCells,
                    totalAddedRows = report.TotalAddedRows,
                    totalRemovedRows = report.TotalRemovedRows,
                    elapsedMs = report.ElapsedMs,
                    addedSheets = report.AddedSheets,
                    removedSheets = report.RemovedSheets,
                    sheets = report.Sheets.Where(s => s.HasChanges).Select(s => new
                    {
                        name = s.SheetName,
                        summary = s.ChangeSummary,
                        changes = s.ChangedCellCount,
                        addedRows = s.AddedRowCount,
                        removedRows = s.RemovedRowCount,
                        oldRows = s.OldRowCount,
                        newRows = s.NewRowCount,
                        addedColumns = s.AddedColumns,
                        removedColumns = s.RemovedColumns,
                        sample = s.Rows.Take(30).Select(r => new
                        {
                            kind = r.KindText,
                            row = r.DisplayRow + 1,
                            key = r.Key,
                            col = r.Kind == RowChangeKind.Modified && r.Cells.Count > 0 ? r.Cells[0].Col : -1,
                            oldValue = r.Kind == RowChangeKind.Modified ? string.Join(" | ", r.Cells.Select(c => c.OldValue)) : r.RowText,
                            newValue = r.Kind == RowChangeKind.Modified ? string.Join(" | ", r.Cells.Select(c => c.NewValue)) : string.Empty,
                        }),
                    }),
                },
                new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                });

            exitCode = report.HasChanges ? 1 : 0;
        }
        catch (Exception ex)
        {
            json = "{\"error\":\"" + ex.Message.Replace("\"", "'") + "\"}";
            exitCode = 2;
        }

        if (!string.IsNullOrEmpty(outPath))
        {
            File.WriteAllText(outPath, json, new System.Text.UTF8Encoding(false));
        }

        try
        {
            Console.WriteLine(json);
        }
        catch (Exception)
        {
            // 无控制台时忽略
        }

        return exitCode;
    }
}
