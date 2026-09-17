@echo off
rem ===========================================================================
rem  set-default-app.cmd   (file name in Chinese; contents are ASCII-only)
rem
rem  Registers ExcelViewer as an "open with" choice and sends you to the
rem  Windows "Default apps" page.
rem
rem  Why you must confirm once: Windows 10/11 stores the default handler in the
rem  UserChoice key, protected by a system hash. No program can set it silently.
rem
rem  WHY NO CHINESE IN THIS FILE:
rem    cmd.exe parses a .cmd using the console's current codepage, and
rem    "chcp 65001" only takes effect for later lines. Even a handful of UTF-8
rem    Chinese bytes (a comment, or a Chinese file name in a path) can desync
rem    that parser - typical symptoms are "'xxx' is not recognized" and
rem    "else was unexpected at this time". All Chinese text, and the Chinese
rem    .reg file name, are handled by tools\install-open-with.ps1 instead.
rem ===========================================================================

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\install-open-with.ps1" %*
