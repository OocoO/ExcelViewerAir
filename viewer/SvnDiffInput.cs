using System.IO;
using System.IO.Compression;
using ExcelViewer.Model;

namespace ExcelViewer;

/// <summary>
/// SVN 外部 diff 的参数处理（`--diff-svn`）。
///
/// 两种客户端拼出来的命令行不一样：
///   * TortoiseSVN（设置 → 差异查看器 → 高级）：`ExcelViewer.exe --diff-svn %base %mine`
///   * 命令行 svn：`svn diff --force --diff-cmd "…\ExcelViewer.exe" -x "--diff-svn"`，
///     svn 会拼成 `<程序> --diff-svn -L "标签1" -L "标签2" 文件1 文件2`。
///
/// 坑在文件1：它是 `.svn\pristine\XX\&lt;sha1&gt;.svn-base` —— **没有扩展名**。
/// 查看器是按扩展名挑解析器的，直接传进去只会得到"这个文件类型暂时看不了"。
/// 所以这里：
///   1. 跳过 `-L` 标签，取最后两个真实文件（不能傻按 %1 %2 取）；
///   2. 按文件头嗅探真实格式（zip / OLE / 其它当文本），扩展名不被支持时复制成带扩展名的副本；
///   3. 标签里的 `\t(revision N)` 换成空格，界面上直接显示成「T.xlsx (revision 12)」。
/// </summary>
internal static class SvnDiffInput
{
    /// <summary>两个待比对的文件，外加给界面显示的标签。</summary>
    internal sealed record Pair(string OldPath, string NewPath, string OldLabel, string NewLabel);

    /// <summary>超过这个天数的临时副本会被清掉。</summary>
    private static readonly TimeSpan KeepCopies = TimeSpan.FromDays(3);

    public static Pair? Resolve(string[] args)
    {
        var labels = new List<string>();
        var paths = new List<string>();

        for (var i = 0; i < args.Length; i++)
        {
            var a = args[i];
            if (a.Equals("--diff-svn", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (a == "-L")
            {
                if (i + 1 < args.Length)
                {
                    labels.Add(args[++i]);
                }

                continue;
            }

            if (a.StartsWith('-'))
            {
                continue;
            }

            paths.Add(a);
        }

        // 只有真实存在的文件才算数（-x 里的其它参数可能也被当成路径）
        var files = paths.Where(File.Exists).ToList();
        if (files.Count < 2)
        {
            return null;
        }

        var oldFile = files[^2];
        var newFile = files[^1];
        var oldLabel = CleanLabel(labels.Count >= 2 ? labels[^2] : oldFile);
        var newLabel = CleanLabel(labels.Count >= 1 ? labels[^1] : newFile);

        return new Pair(EnsureLoadable(oldFile), EnsureLoadable(newFile), oldLabel, newLabel);
    }

    /// <summary>`"D:\x\T.xlsx\t(revision 12)"` → `D:\x\T.xlsx (revision 12)`。</summary>
    private static string CleanLabel(string label)
    {
        var text = label.Replace('\t', ' ').Trim();
        return text.Length == 0 ? "(未知版本)" : text;
    }

    /// <summary>扩展名不被支持时，按内容嗅探格式并复制一份带正确扩展名的副本。</summary>
    private static string EnsureLoadable(string path)
    {
        if (WorkbookLoader.IsSupported(path))
        {
            return path;
        }

        var ext = SniffExtension(path);
        TempWorkspace.CleanupSvnDiff(KeepCopies);

        var dir = Path.Combine(
            TempWorkspace.SvnDiffDir,
            $"{DateTime.Now:yyyyMMdd_HHmmss}_{Environment.ProcessId}");
        Directory.CreateDirectory(dir);

        var target = Path.Combine(dir, SafeFileName(path) + ext);
        File.Copy(path, target, overwrite: true);
        return target;
    }

    /// <summary>按文件头判断真实格式；认不出来就当 csv（纯文本）试。</summary>
    private static string SniffExtension(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            var head = new byte[8];
            var read = fs.Read(head, 0, head.Length);

            // zip 容器：xlsx / xlsm / xlsb 都是，靠内部条目再分一次
            if (read >= 4 && head[0] == 'P' && head[1] == 'K')
            {
                fs.Position = 0;
                using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
                if (zip.GetEntry("xl/workbook.bin") is not null)
                {
                    return ".xlsb";
                }

                return zip.GetEntry("xl/vbaProject.bin") is not null ? ".xlsm" : ".xlsx";
            }

            // OLE 复合文档：老 .xls
            if (read >= 8 && head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0)
            {
                return ".xls";
            }
        }
        catch (Exception)
        {
            // 嗅探失败就按文本处理，交给加载器去报错
        }

        return ".csv";
    }

    private static string SafeFileName(string path)
    {
        var name = Path.GetFileNameWithoutExtension(path);
        if (string.IsNullOrWhiteSpace(name))
        {
            name = "svn";
        }

        foreach (var bad in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(bad, '_');
        }

        return name.Length <= 40 ? name : name[..40];
    }
}
