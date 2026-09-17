namespace ExcelViewer.Model;

/// <summary>
/// 估算字符串在表格里占用的"显示宽度"（以半角字符为单位）。
/// 中日韩全角字符按 2 计，其余按 1 计。用于自动列宽，避免逐格做字体测量。
/// </summary>
public static class DisplayWidth
{
    public static int Measure(string s)
    {
        var w = 0;
        foreach (var ch in s)
        {
            w += IsWide(ch) ? 2 : 1;
        }

        return w;
    }

    /// <summary>
    /// 估算显示宽度，超过 <paramref name="capUnits"/> 就直接返回"很宽"的哨兵值。
    /// 只需要知道"这一列很宽"，具体多宽没意义（列宽本来就有上限），因此提前退出。
    /// </summary>
    public static int MeasureCapped(string s, int capUnits)
    {
        var w = 0;
        for (var i = 0; i < s.Length; i++)
        {
            w += IsWide(s[i]) ? 2 : 1;
            if (w > capUnits)
            {
                return capUnits + 1;
            }
        }

        return w;
    }

    private static bool IsWide(char c)
    {
        // 控制字符/制表符不参与宽度计算
        if (c < 0x20)
        {
            return false;
        }

        return c switch
        {
            >= '\u1100' and <= '\u115F' => true,  // 谚文字母
            >= '\u2E80' and <= '\u303E' => true,  // 中日韩部首、标点
            >= '\u3041' and <= '\u33FF' => true,  // 假名、注音、兼容字符
            >= '\u3400' and <= '\u4DBF' => true,  // 扩展 A
            >= '\u4E00' and <= '\u9FFF' => true,  // 基本汉字
            >= '\uA000' and <= '\uA4CF' => true,  // 彝文
            >= '\uAC00' and <= '\uD7A3' => true,  // 谚文音节
            >= '\uF900' and <= '\uFAFF' => true,  // 兼容汉字
            >= '\uFE30' and <= '\uFE4F' => true,  // 兼容形式
            >= '\uFF00' and <= '\uFF60' => true,  // 全角形式
            >= '\uFFE0' and <= '\uFFE6' => true,  // 全角符号
            _ => false,
        };
    }

    /// <summary>把单元格值里的换行压成空格用于单行表格显示。</summary>
    public static string Flatten(string s)
    {
        if (s.IndexOfAny(FlattenChars) < 0)
        {
            return s;
        }

        var buf = s.ToCharArray();
        for (var i = 0; i < buf.Length; i++)
        {
            if (buf[i] is '\r' or '\n' or '\t')
            {
                buf[i] = ' ';
            }
        }

        return new string(buf);
    }

    private static readonly char[] FlattenChars = { '\r', '\n', '\t' };
}
