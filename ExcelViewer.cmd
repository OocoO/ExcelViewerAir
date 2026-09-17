@echo off
rem ===========================================================================
rem  ExcelViewer.cmd - launcher. Drag an xlsx/csv onto this file to open it.
rem
rem  ASCII-ONLY ON PURPOSE:
rem    cmd.exe parses a .cmd file using the console's current codepage, and
rem    "chcp 65001" only affects lines AFTER it runs. So a UTF-8 non-ASCII byte
rem    anywhere in this file - even inside a quoted argument - can be
rem    mis-decoded and desync the parser. Symptoms: "else was unexpected at
rem    this time", chopped-up commands, "'xxx' is not recognized".
rem    All Chinese text lives in tools\launch.ps1 instead.
rem ===========================================================================

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\launch.ps1" %*
