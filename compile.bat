@rem compile.bat
@rem AHK v2

@echo off
cd %~dp0

set appname=kontaktSupporter

set autohotkeyExe=C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe
set autohotkeyCompiler=C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe

call %appname%.exe remove

call "%autohotkeyExe%" /in %appname%.ahk /out %appname%.exe /icon %appname%.ico /base "%autohotkeyCompiler%

call upx --best %appname%.exe
