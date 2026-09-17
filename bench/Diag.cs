using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using System.Xml;
using ExcelDataReader;

namespace Bench;

/// <summary>诊断：逐 sheet 打印行数，并直接看 sheet XML 的真实范围。</summary>
internal static class Diag
{
    public static void Run(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Console.WriteLine($"### {Path.GetFileName(path)}");

        using var fs = File.OpenRead(path);
        using var reader = ExcelReaderFactory.CreateReader(fs);
        var sw = Stopwatch.StartNew();
        var sheetNo = 0;
        do
        {
            var t0 = sw.ElapsedMilliseconds;
            var rows = 0;
            var maxCol = 0;
            while (reader.Read())
            {
                rows++;
                maxCol = Math.Max(maxCol, reader.FieldCount);
            }

            Console.WriteLine($"  [{sheetNo++}] {reader.Name,-24} rows={rows,7} fieldCount={reader.FieldCount,3} maxCol={maxCol,3} took={sw.ElapsedMilliseconds - t0,6}ms");
        } while (reader.NextResult());

        Console.WriteLine($"  reader total = {sw.ElapsedMilliseconds}ms");

        using var zip = ZipFile.OpenRead(path);
        foreach (var e in zip.Entries)
        {
            if (e.FullName.StartsWith("xl/worksheets/") && e.FullName.EndsWith(".xml"))
            {
                using var s = e.Open();
                using var xml = XmlReader.Create(s, new XmlReaderSettings { IgnoreWhitespace = true });
                var rows = 0;
                var maxColIdx = 0;
                string? firstRef = null;
                string? lastRef = null;
                var attrCount = 0;
                while (xml.Read())
                {
                    if (xml.NodeType != XmlNodeType.Element)
                    {
                        continue;
                    }

                    if (xml.LocalName == "dimension")
                    {
                        Console.WriteLine($"      dimension={xml.GetAttribute("ref")}");
                    }
                    else if (xml.LocalName == "row")
                    {
                        var r = xml.GetAttribute("r");
                        firstRef ??= r;
                        lastRef = r;
                        rows++;
                        if (xml.HasAttributes)
                        {
                            attrCount += xml.AttributeCount;
                        }
                    }
                    else if (xml.LocalName == "c")
                    {
                        var r = xml.GetAttribute("r");
                        if (r is not null)
                        {
                            var n = 0;
                            foreach (var ch in r)
                            {
                                if (ch is >= 'A' and <= 'Z')
                                {
                                    n = (n * 26) + (ch - 'A' + 1);
                                }
                                else
                                {
                                    break;
                                }
                            }

                            maxColIdx = Math.Max(maxColIdx, n);
                        }
                    }
                }

                Console.WriteLine($"      {e.FullName,-32} comp={e.CompressedLength / 1024,7}KB raw={e.Length / 1024,7}KB <row>={rows,7} first={firstRef} last={lastRef} maxCol={maxColIdx} attrSum={attrCount}");
            }
        }
    }
}
