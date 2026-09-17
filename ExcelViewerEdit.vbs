' ExcelViewerEdit.vbs — "用副本编辑"入口
'
' 作用：把选中的 xlsx 复制到工作区，用 Excel/WPS 打开【副本】，
'       原文件全程只读、不会被锁；在 Excel 里保存并关闭后，
'       查看器会自动弹出"改动比对"，让你看着 diff 决定是否写回。
'
' 被注册表以 wscript.exe 调用：wscript "...\ExcelViewerEdit.vbs" "%1"
Option Explicit

Dim fso, here, viewer, arg, cmd, sh
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

If WScript.Arguments.Count < 1 Then
    MsgBox "请把 .xlsx 文件（或它的快捷方式）拖到本脚本上。", 64, "用副本编辑"
    WScript.Quit 1
End If

here = fso.GetParentFolderName(WScript.ScriptFullName)
viewer = here & "\viewer\publish\ExcelViewer.exe"

If Not fso.FileExists(viewer) Then
    MsgBox "没有找到查看器：" & vbCrLf & viewer & vbCrLf & vbCrLf & _
           "请先在 excel2csv 目录执行：" & vbCrLf & _
           "    dotnet publish viewer\ExcelViewer.csproj -c Release -o viewer\publish", _
           16, "用副本编辑"
    WScript.Quit 1
End If

arg = WScript.Arguments(0)

' --edit：复制到工作区 -> 用 Excel 打开副本 -> 等关闭 -> 自动比对改动
cmd = """" & viewer & """ --edit """ & arg & """"
sh.Run cmd, 0, False
