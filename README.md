# Excel 只读查看器（ExcelViewer）

给「大配表看一眼」这个场景做的 Windows 查看器：**只读、秒开、能全表搜**，性能取向优先，
界面刻意做得和 Excel/WPS 接近，但查看本身没有任何编辑能力。

打开方式（三选一，可同时用）：

* 双击 `ExcelViewer.cmd`，或把 xlsx / csv 拖到它上面
* 跑 `viewer\publish\ExcelViewer.exe "某个表.xlsx"`
* 导入 `注册双击打开.reg` 后，右键 xlsx 会出现「用查看器打开（只读）」和「用副本编辑」

---

## 编辑流程：不锁原文件

配表要改的时候，**不要**让查看器和 Excel 去抢同一个文件。这里的做法是：

```
原文件 ──只读复制一次──▶ 工作区副本 ──Excel 打开──▶ 你编辑 ──保存关闭──▶ 自动弹出改动比对
   ▲                                                                        │
   └──────────────── 只有你点「写回原文件」并二次确认，才会被覆盖 ◀──────────┘
```

具体保证：

* **原文件全程只被读一次**，不会被 Excel 锁定 —— 实测在 Excel 开着副本时，
  原文件仍能以独占模式打开（`FileShare.None`）。
* 工作区在 `%LOCALAPPDATA%\ExcelViewer\sessions\<时间戳>\`，不在系统 TEMP 里，
  不会被系统清理掉导致改动丢失；可用环境变量 `EXCELVIEWER_WORKSPACE` 改位置。
* Excel 保存是「写临时文件再替换」，所以比对**等编辑器进程退出 + 文件大小/时间戳稳定**之后才做，
  不会读到半个文件。
* 写回前会自动在 `%LOCALAPPDATA%\ExcelViewer\backup\` 留一份带时间戳的备份。
* 不喜欢写回就用「另存为…」，全程不碰原文件。

改动比对是把**改动过的格子**摊成一张清单表（行号 / 列 / 原值 / 新值），可以直接搜索、复制。

### 进入编辑流程的三种方式

| 方式 | 操作 |
| --- | --- |
| 右键菜单 | 导入 `注册双击打开.reg` 后，xlsx 右键 →「用副本编辑（不锁原文件）」 |
| 查看器内 | 打开文件后点工具栏「用副本编辑」，或菜单 文件 → 用副本编辑 |
| 命令行 | `ExcelViewer.exe --edit "某个表.xlsx"` |

改完想手动比对：菜单 文件 →「与原文件比对改动…」，或
`ExcelViewer.exe --diff "原文件" "改过的文件"`。

## 注册双击打开

Windows 10/11 的默认打开方式存在 `UserChoice` 键里，**带系统哈希保护，任何程序都无法静默改成默认**
（这是系统防止程序篡改关联的机制，不是没做）。所以流程是：导入注册表 → 手动确认一次。

```powershell
# 1) 双击导入（只写 HKCU，不需要管理员权限）
注册双击打开.reg

# 2) 手动把默认打开方式设成查看器，任选一种：
#    a. 双击 设为默认打开方式.cmd，跳到系统设置页，搜索 .xlsx 后选「Excel 只读查看器」
#    b. 右键任意 .xlsx → 打开方式 → 选择其他应用 → 选「Excel 只读查看器」→ 勾选「始终」
```

注册表文件里做了什么：

* 注册 ProgId `ExcelViewer.Sheet`，让它出现在「打开方式」列表里
* 给 `.xlsx / .xlsm / .xlsb / .csv` 的 `OpenWithProgids` 加上它（**不改默认值**）
* 给 `.xlsx` 右键菜单加两项：只读打开、用副本编辑

> 早先那套「转 CSV 查看」（`xlsx2csv.py` / `convert_open.vbs` / `转换CSV.bat` /
> `装右键菜单.reg` / `设双击打开.reg` / `恢复Excel双击.reg`）已经全部移除，
> 注册表里的 `excel2csv` ProgId、`.xlsx` 默认值和那条右键项也一并清掉了。
> 现在 `.xlsx` 双击走系统默认（本机是 Excel），要只读看就用右键那两项或 `ExcelViewer.cmd`。

**注意**：`注册双击打开.reg` 里的路径是按当前安装位置写死的（共 5 处）。
如果之后移动了工程目录，把里面的
`C:\\Users\\18223\\Programs\\excel2csv\\viewer\\publish\\ExcelViewer.exe`
整体替换成新路径即可。

卸载：删掉 `HKCU\Software\Classes\ExcelViewer.Sheet`、
`HKCU\Software\Classes\.xlsx\shell\excelviewer`、`...\excelvieweredit`，
以及各扩展名 `OpenWithProgids` 下的 `ExcelViewer.Sheet` 值。

---

## 为什么比 WPS 快

WPS/Excel 打开配表慢，主要慢在「一次性把整个工作簿全部解析、全部建单元格对象、全部排版」。
这个查看器把这三件事全部拆开：

| 做法 | 效果 |
| --- | --- |
| **只解析当前工作表** | 打开 `Languages.xlsx`（3 个 sheet / 4.2 万行）只读首 sheet，切换 sheet 时才读下一个 |
| **列式存储 + 字符串池** | 相同值（`GDE_IGNORE`、空值、重复文案）全表只存一份，重复率高的配表内存省一大截 |
| **自绘虚拟化表格** | 只画可视区那几十行，滚动到第 3 万行与第 3 行的开销相同 |
| **不做公式计算 / 不建样式对象** | 只读文本，跳过 Excel 的公式引擎与格式系统 |

实测（本机，Release 构建，1400x700 视口，见下方「复现方法」）：

| 文件 | 规模 | 首 sheet 加载 | 首次上屏 | 滚动重绘 | 全表搜索 |
| --- | --- | --- | --- | --- | --- |
| Languages.xlsx | 31363 行 x 9 列 | 927 ms | 38 ms | 中位 ~23 ms | 4–10 ms |
| Item.xlsx | 3668 行 x 13 列 | 184 ms | 31 ms | 中位 ~16 ms | 0.5–5 ms |
| Level.xlsx | 2305 行 x 71 列 | 284 ms | 26 ms | 中位 ~9 ms | 0.5–4 ms |

> 界面右上角会显示真实「打开耗时」，可以拿去和 WPS 直接对比。

## 功能

* **只读**：查看时不写任何东西回原文件，也不往用户目录写配置。
* **内容预览条（默认展开）**：配表里有些说明列是几千字的 JSON / 多行文案，格子宽度装不下。
  底部常驻一条 Excel 风格的预览条，**选中哪个格子就显示它的完整内容**：自动换行 + 纵向滚动，
  可以直接选中或用「复制」取原文；标题栏给出「N 字符 · M 行」和所在位置。
  上面的分隔线**拖动就能改高度**（拖出来的高度会被记住）；按 `Ctrl+P` 收起后，
  预览条会变成一行摘要（地址 + 内容开头 + 「展开 ▴」），点一下即可恢复，不会留下半截按钮。
* **格子只画一行**：列宽不会再被长文案撑开，放不下就用 `…` 截断（和 Excel 默认行为一致），
  不会出现"多行文字叠在一起"；要看全文就选中它，底部预览条里有完整内容。
* **全表搜索**：输入即搜（后台线程 + 220ms 防抖），命中行整行淡黄、命中单元格亮黄，
  当前命中行有描边；`F3` / `Shift+F3` 上下条跳转，状态栏显示「第 n / m 条」。
* **表格操作**：冻结列标与表头行、行号栏、列宽拖拽、双击列边界自适应、键盘导航、
  `Ctrl+C` 复制选中区域、`Ctrl+ +/-` 缩放字号。
* **自动识别表头行**：项目配表的结构是 `GDE_IGNORE` / `GDE_FIELD_NAMES` / 中文描述 / `GDE_FIELD_TYPES`，
  查看器会在前 8 行里找 `GDE_FIELD_NAMES`，把那一行标成表头（加粗 + 淡蓝底），
  并在打开时自动滚到它附近。找不到标记就按普通表格处理（第一行当表头）。
* **改动比对**：把两个工作簿的逐格差异摊成清单表（行号 / 列 / 原值 / 新值）。
* 支持格式：`xlsx` / `xlsm` / `xlsb` / `xls` / `csv`。

## 快捷键

| 按键 | 作用 |
| --- | --- |
| `Ctrl+O` / `F5` | 打开文件 / 重新加载当前文件（重新导出后可刷新） |
| `Ctrl+F` | 定位到搜索框 |
| `回车` / `F3` | 下一个命中 |
| `Shift+F3` | 上一个命中 |
| `Tab` | 切换工作表 |
| 方向键 / `PageUp` `PageDown` | 移动 / 翻页 |
| `Ctrl+Home` / `Ctrl+End` | 跳到表格开头 / 末尾 |
| `Ctrl+C` | 复制选中单元格（制表符分隔，可直接粘回 Excel） |
| `Ctrl+P` | 收起 / 展开底部内容预览条（默认展开；收起后点一下那一行也能展开） |
| `Ctrl+ +` / `Ctrl+ -` | 放大 / 缩小字号 |
| `Esc` | 清除搜索 |

## 目录结构

```
excel2csv/
├─ ExcelViewer.cmd                  便捷启动器（可拖文件到它上面）
├─ ExcelViewerEdit.vbs              「用副本编辑」入口（右键菜单调用它）
├─ 注册双击打开.reg                 注册 ProgId / 打开方式 / 右键菜单
├─ 设为默认打开方式.cmd             注册 + 跳到系统设置页设默认
├─ NuGet.config                     离线还原配置（本机无法访问 nuget.org）
├─ viewer/                          查看器源码（C# / WPF）
│  ├─ publish/ExcelViewer.exe       ← 直接运行这个
│  ├─ Model/                        数据模型：字符串池、列式 Sheet、加载器、表头识别、改动比对
│  ├─ Controls/                     自绘虚拟化表格、布局模型、文本绘制、配色
│  ├─ Search/                       全表搜索
│  ├─ DiffView.cs                   改动比对视图
│  ├─ EditSession.cs                编辑器查找、无界面比对输出
│  └─ TempWorkspace.cs              工作区 / 副本 / 备份目录
├─ bench/                           解析引擎基准与数据核对（复用 viewer 的加载器）
└─ tools/
   ├─ launch.ps1                    启动器逻辑（中文提示放这里）
   ├─ install-open-with.ps1          写注册表 + 打开默认应用设置页
   ├─ Shot.csproj / Program.cs       窗口截图 / 图片放大（独立小程序）
   └─ make_diff_fixture.py           造一份「改过几处」的测试数据
```

> **`.cmd` 文件一律保持纯 ASCII**（见下方「已修的坑」）。所以上面几个 `.cmd` 都很短，
> 真正的内容和中文提示都在 `tools\*.ps1` 里。

## 构建

```powershell
# 构建
dotnet build viewer\ExcelViewer.csproj -c Release

# 发布（生成 viewer\publish\ExcelViewer.exe）
dotnet publish viewer\ExcelViewer.csproj -c Release -o viewer\publish
```

依赖：.NET 10 SDK + .NET 10 桌面运行时（本机已装）。发布物是**框架依赖**的独立 exe，
换机器需要装 .NET 10 Desktop Runtime；本机直接可用。

> 关于离线：这台机器访问 nuget.org 会 SSL 失败，所以 `NuGet.config` 指向本地缓存。
> `ExcelViewer.csproj` 里还把 apphost 指向了 SDK 自带的 packs 目录
> （NuGet 缓存里没有 10.x 的 `Microsoft.NETCore.App.Host.win-x64`），
> 这样 publish 不需要联网。

## 自检与基准（改代码后用来自证没坏）

```powershell
$exe = 'viewer\publish\ExcelViewer.exe'

# 数据管线自检：逐 sheet 加载、抽查首末行、搜索命中、字符串池统计
& $exe --selftest test\Languages.xlsx test\Item.xlsx

# 性能基准：加载 / 首屏 / 滚动重绘 / 搜索 / 内存
& $exe --bench test\Languages.xlsx --frames=200

# 布局探针：打印列偏移、离屏渲染成 PNG，并用像素断言检查文字有没有被行号栏裁掉
& $exe --layout test\Item.xlsx --png=out.png

# 界面探针：真的建出主窗口、真的打开一个文件，打印窗口里的实际状态并整窗离屏渲染
& $exe --probe 某个表.xlsx --png-window=整窗.png --png=表格.png

# 预览条两种状态各渲染一张（不加参数时就是默认的展开状态）
& $exe --probe 某个表.xlsx --preview      --png-window=预览条展开.png
& $exe --probe 某个表.xlsx --preview-off  --png-window=预览条收起.png

# 文本渲染器 A/B（TextFormatter 与 FormattedText 的落点对比）
& $exe --textprobe

# 改动比对：无界面输出 JSON（退出码 1 = 有改动）与改动清单
& $exe --diff-json 原.xlsx 新.xlsx --out=diff.json
& $exe --diff-png  原.xlsx 新.xlsx --out=清单.csv --png=预览.png

# 走一遍完整编辑流程（复制副本 + 打开 Excel + 关闭后自动比对）
& $exe --edit test\Item.xlsx
```

造一份「改过几处」的测试数据：

```powershell
python tools\make_diff_fixture.py test\Item.xlsx $env:TEMP\edited\Item_edited.xlsx
```

## 已修的坑（留给后来人）

* **`.cmd` 文件里不要出现任何非 ASCII 字符**。cmd.exe 用**控制台当前代码页**（本机 936）解析
  `.cmd`，而文件是 UTF-8；`chcp 65001` 只对**它执行之后**的行生效。所以哪怕只是注释里的一串中文，
  字节也会被错误解码、切断后续行，症状非常迷惑：
  `'xxx' is not recognized as an internal or external command`、`else was unexpected at this time`，
  甚至把 `reg import` 的参数吃掉只剩 `resolves`。
  **正确做法**：`.cmd` 只留纯 ASCII，中文提示与中文路径都交给 `tools\*.ps1`
  （PowerShell 按 UTF-8 读 `.ps1`，中文安全）。
  现在根目录只剩 `ExcelViewer.cmd` 和 `设为默认打开方式.cmd` 两个 `.cmd`，
  它们都只有一行 `powershell -File tools\*.ps1`，实测没有解析问题。
* **`reg import` 的退出码不可靠**：即使导入成功也可能返回 1，别用它判断成败，
  导入后用 `reg query` / `Test-Path` 复核。
* **不要用 `TextFormatter` 画单元格**。它的 `TextRunCache` 会按 `TextSource` 实例复用上一次的
  排版结果，即使显式传入 cache 参数也一样：同一个实例连续画不同字符串时，宽度和字形对不上，
  表现为「文字被画到别的列去了」。现在用 `FormattedText` + 按 (字符串, 对齐, 宽度) 的缓存，
  既正确又快（滚动时基本命中缓存）。
* **列宽不能直接用全表最大文本宽度**。说明性长文本列会把列宽顶到上限，导致所有列一样宽、
  一屏看不到几个字段。现在只采样前 80 行（外加保证表头行被采样），且优先用真实字体测量。
* **`PushClip` 之外还要注意绘制顺序**：行号栏、列标是在单元格之后画的，改这块时别把
  单元格文字画到行号栏左边去（用 `--layout --png` 的像素断言可以立刻发现）。
* **`UserChoice` 改不了**：想静默把程序设成默认打开方式是不可能的，别在这上面浪费时间，
  直接引导用户去系统设置页确认。
* **不要在编辑器还开着时就比对文件**：Excel 保存是「写临时文件再替换」，
  必须等进程退出 + 文件稳定，否则会读到半个文件。
* **叠在主表格上面的第二张表，默认必须 `Visibility.Collapsed`**：改动比对用的那张
  自绘表格（`DiffView._grid`）是 `GridHost.Children.Add` 上去的第二层，Z 序在主表格之后。
  它构造时如果是默认的 `Visible`，就会用一张空白表把正常查看的表格整个盖住，
  症状非常有迷惑性——**状态栏显示「11 行 × 22 列 · 打开耗时 90 ms」，数据、搜索、内存全部正常，
  但界面中间只有一句「把 Excel / CSV 拖到这里」，列标、行号栏、单元格一个都看不见**
  （`ExcelGrid.OnRender` 里 `sheet is null` 只有它自己知道）。
  定位方法：`--probe` 会打印 `MainWindow._sheet` 与 `ExcelGrid.Model.Sheet`，
  两个都不为空却仍然画空提示，就是被上层表格盖住了；
  `--png-window` 把整窗离屏渲染成 PNG，可以直接确认。
  修法：构造时 `Collapsed`，`DiffView.Show()` 里 `Visible`，`DiffView.Hide()` 里再 `Collapsed`。
* **`DiffView.Hide()` 不能顺手把 `GridHost` 一起藏掉**：`GridHost` 是**主表格**的宿主，
  比对表只是它里面的第二层。藏了它，关掉改动比对后整个数据区会一片空白
  （看着像"文件没加载"）。只藏 `DiffView._grid` 就够。
* **单元格文本必须限制成一行**（`FormattedText.MaxLineCount = 1` + `Trimming = CharacterEllipsis`）：
  行高是固定的 22px，而 `FormattedText` 只要设了 `MaxTextWidth` 就会**按列宽自动折行**，
  多行文本块再被 `Draw` 垂直居中，于是第二行开始会画到上下相邻的行上去——
  症状是"格子里几行字叠成一团、糊得看不清"（说明性长文案列最明显）。
  注意 `WorkbookLoader` 已经把值里的换行压成空格了，所以叠字不是换行引起的，是折行引起的；
  网上常见的"把双行合并"的排查方向会白跑。
* **`GridSplitter` 要拖的是"外面那一行"**：预览条的分隔线原先放在预览条内部的 `Grid` 里，
  它拖的是那个内层 Grid 自己的行（`Auto` / `*`），而预览条的高度其实由**外层**行定义决定，
  所以拖动完全没反应。正确做法：外层 Grid 开三行（表格 `*` / 分隔线 `Auto` / 预览 `Height`），
  `GridSplitter` 单独占中间那行，并显式写 `ResizeBehavior="PreviousAndNext"`、`ResizeDirection="Rows"`。
* **`--textprobe` 的画布高度要留够**：两组样本（TF/FT 各 4 行 × 24px）加分隔线后超过 200px，
  `Scan` 会越界读像素数组直接抛 `IndexOutOfRangeException`；现在画布 260px，并且 `Scan` 自己会夹住 y 范围。
