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

    /// <summary>
    /// SVN 外部 diff 用的临时副本目录（`--diff-svn`）。
    /// svn 传进来的旧版本文件常常是 `.svn\pristine\XX\&lt;sha1&gt;.svn-base`，**没有扩展名**，
    /// 而查看器是按扩展名挑解析器的，所以要先按文件头嗅探格式、复制成带扩展名的副本。
    /// </summary>
    public static string SvnDiffDir
    {
        get
        {
            var dir = Path.Combine(Root, "svn-diff");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    /// <summary>
    /// 清掉过期的临时副本。这些副本在查看器关掉后就没用了，但当时可能正被读取，
    /// 所以不在退出时删，而是每次新建副本前顺手清理几天前的。
    /// </summary>
    public static void CleanupSvnDiff(TimeSpan olderThan)
    {
        try
        {
            var root = SvnDiffDir;
            var deadline = DateTime.Now - olderThan;
            foreach (var dir in Directory.EnumerateDirectories(root))
            {
                try
                {
                    if (Directory.GetLastWriteTime(dir) < deadline)
                    {
                        Directory.Delete(dir, recursive: true);
                    }
                }
                catch (IOException)
                {
                    // 还在被别的查看器窗口读着，下次再说
                }
                catch (UnauthorizedAccessException)
                {
                }
            }
        }
        catch (IOException)
        {
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
