using System.IO;

namespace ExcelViewer;

/// <summary>
/// 临时工作区：编辑流程用的"副本目录"。
///
/// 设计目的：**绝不让原始配表被锁**。编辑时把原文件复制到工作区，Excel 打开的是副本，
/// 原文件自始至终只被只读读一次；改完再比对、确认后才写回。
/// 工作区默认放在 %LOCALAPPDATA%\ExcelViewer\ 下（而不是 TEMP），
/// 这样即使用户没及时处理，副本也不会被系统清理掉导致改动丢失。
/// 可用环境变量 EXCELVIEWER_WORKSPACE 覆盖（便携使用/测试用）。
/// </summary>
public static class TempWorkspace
{
    private const string RootName = "ExcelViewer";

    private const string WorkspaceEnvVar = "EXCELVIEWER_WORKSPACE";

    /// <summary>工作区根目录。</summary>
    public static string Root
    {
        get
        {
            var custom = Environment.GetEnvironmentVariable(WorkspaceEnvVar);
            var dir = !string.IsNullOrWhiteSpace(custom)
                ? custom
                : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), RootName);
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    /// <summary>每次编辑会话一个子目录，便于整体清理和事后查找。</summary>
    public static string SessionDir
    {
        get
        {
            var dir = Path.Combine(Root, "sessions", DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N")[..6]);
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    /// <summary>写回前的自动备份目录。</summary>
    public static string BackupDir
    {
        get
        {
            var dir = Path.Combine(Root, "backup");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    /// <summary>把原文件复制到工作区，返回副本路径。原文件只被读，不会被写。</summary>
    public static string CreateCopy(string sourcePath)
    {
        var dir = SessionDir;
        var target = Path.Combine(dir, Path.GetFileName(sourcePath));

        // 以共享只读方式打开原文件再复制，避免 Excel 正开着时复制失败
        using var src = new FileStream(
            sourcePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            1 << 16,
            FileOptions.SequentialScan);
        using var dst = new FileStream(target, FileMode.Create, FileAccess.Write, FileShare.None);
        src.CopyTo(dst);
        return target;
    }
}
