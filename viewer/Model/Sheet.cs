namespace ExcelViewer.Model;

/// <summary>
/// 一个工作表的列式只读存储。
///
/// 用「列数组」而不是「行对象」的原因：
///   * 配表列数固定（最多几十列），列数组能精确按需增长，不浪费行对象头；
///   * 搜索是按列顺序扫，列式布局的缓存局部性更好；
///   * 渲染取某列区间时不需要跨行对象间接跳转。
/// 单元格一律存 string（数字与日期在加载时已格式化成文本），
/// 相同值经 StringPool 去重，因此数组里放的是共享引用。
/// </summary>
public sealed class Sheet
{
    private const int InitialRows = 256;

    private readonly List<string[]> _cols = new();
    private readonly List<int> _rowWidth = new(InitialRows);
    private readonly List<bool> _rowHasData = new(InitialRows);
    private readonly List<int> _lastRowWithDataByCol;

    private readonly StringPool _pool;
    private string[] _curRow = new string[16];
    private int _curRowLen;

    public Sheet(string name, StringPool pool)
    {
        Name = name;
        Title = name;
        _pool = pool;
        _lastRowWithDataByCol = new List<int>();
    }

    public string Name { get; }

    /// <summary>界面上显示的名字（diff 视图里写成"字段名（原）→（新）"）。</summary>
    public string Title { get; set; }

    /// <summary>表格里显示的行数（不含尾部整行为空的记录）。</summary>
    public int RowCount { get; private set; }

    /// <summary>表格里显示的列数（不含尾部整列为空的记录）。</summary>
    public int ColCount { get; private set; }

    /// <summary>非空单元格总数，用于状态栏与搜索进度。</summary>
    public long NonEmptyCellCount { get; private set; }

    /// <summary>加载时读到的最大行号（含被裁掉的尾部空行），用于说明"省略了 N 行空白"。</summary>
    public int SourceRowCount { get; private set; }

    /// <summary>加载时读到的最大列号（含被裁掉的尾部空列）。</summary>
    public int SourceColCount { get; private set; }

    /// <summary>加载耗时（毫秒），状态栏展示用。</summary>
    public long LoadMs { get; set; }

    /// <summary>被省略的尾部空白行数。</summary>
    public int SkippedTrailingRows => Math.Max(0, SourceRowCount - RowCount);

    /// <summary>被省略的尾部空白列数。</summary>
    public int SkippedTrailingCols => Math.Max(0, SourceColCount - ColCount);

    // ---------------- 加载期写入 ----------------

    private int _sourceRowIndex;

    public void BeginRow() => _curRowLen = 0;

    public void AddCell(string? rawValue)
    {
        var v = _pool.Intern(rawValue);
        if (_curRowLen >= _curRow.Length)
        {
            Array.Resize(ref _curRow, _curRow.Length * 2);
        }

        _curRow[_curRowLen++] = v;
    }

    /// <summary>结束当前行，把行缓冲追加到列数组。</summary>
    public void EndRow()
    {
        var width = _curRowLen;
        _sourceRowIndex++;
        SourceRowCount = _sourceRowIndex;
        if (width > SourceColCount)
        {
            SourceColCount = width;
        }

        if (width == 0)
        {
            _rowWidth.Add(0);
            _rowHasData.Add(false);
            return;
        }

        while (_cols.Count < width)
        {
            _cols.Add(new string[InitialRows]);
            _lastRowWithDataByCol.Add(-1);
        }

        var rowIndex = _rowWidth.Count;
        var hasData = false;
        for (var c = 0; c < width; c++)
        {
            var v = _curRow[c];
            var col = _cols[c];
            if (rowIndex >= col.Length)
            {
                Array.Resize(ref col, col.Length * 2);
                _cols[c] = col;
            }

            col[rowIndex] = v;
            if (v.Length != 0)
            {
                hasData = true;
                _lastRowWithDataByCol[c] = rowIndex;
                NonEmptyCellCount++;
            }
        }

        _rowWidth.Add(width);
        _rowHasData.Add(hasData);
    }

    /// <summary>加载完成后收尾：裁掉尾部空行/空列，算出列宽。</summary>
    public void Finish()
    {
        // 尾部空行：从最后一行往前找到第一个有数据的行
        var last = _rowWidth.Count - 1;
        while (last >= 0 && !_rowHasData[last])
        {
            last--;
        }

        RowCount = last + 1;

        // 尾部空列
        var lastCol = -1;
        for (var c = 0; c < _cols.Count; c++)
        {
            if (_lastRowWithDataByCol[c] >= 0)
            {
                lastCol = c;
            }
        }

        ColCount = lastCol + 1;

        // 列宽由 GridModel 用真实字体测量并按需缓存，这里不再预先全表扫一遍
        // （那是几万行 x 几十列的无谓开销，且会拖慢首屏）。
        TrimColumnCapacity();
        _curRow = Array.Empty<string>();
    }

    private void TrimColumnCapacity()
    {
        for (var c = 0; c < ColCount; c++)
        {
            var col = _cols[c];
            if (col.Length > RowCount)
            {
                Array.Resize(ref col, Math.Max(RowCount, 1));
                _cols[c] = col;
            }
        }
    }

    // ---------------- 读取 ----------------

    public string Get(int row, int col)
    {
        if ((uint)col >= (uint)ColCount)
        {
            return string.Empty;
        }

        var c = _cols[col];
        return (uint)row < (uint)c.Length ? c[row] : string.Empty;
    }

    /// <summary>整行是否全空。只扫到该行的实际宽度，比逐列扫更快。</summary>
    public bool IsRowEmpty(int row)
    {
        if ((uint)row >= (uint)_rowWidth.Count)
        {
            return true;
        }

        if (!_rowHasData[row])
        {
            return true;
        }

        var width = Math.Min(_rowWidth[row], ColCount);
        for (var c = 0; c < width; c++)
        {
            if (Get(row, c).Length != 0)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>列数组只读访问，供搜索器使用以避免逐格边界检查。</summary>
    public string[] ColumnBuffer(int col) => col >= 0 && col < ColCount ? _cols[col] : Array.Empty<string>();

    /// <summary>预估常驻内存（字节），状态栏展示用。</summary>
    public long EstimateBytes()
    {
        long bytes = 0;
        for (var c = 0; c < _cols.Count; c++)
        {
            bytes += (long)_cols[c].Length * IntPtr.Size;
        }

        bytes += (long)_rowWidth.Count * sizeof(int);
        return bytes;
    }
}

