@echo off
SET LIB=d:\VisualStudio2015\VC\LIB;E:\GIT_Repositories\Python-CodeBase-Tools\Comp_V313;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.22621.0\ucrt\x86;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.22621.0\um\x86
set PATH=D:\VisualStudio2022\VC\Tools\MSVC\14.44.35207\bin\Hostx64\x86;%PATH%
set PATH=D:\VisualStudio2022\Common7\IDE;%PATH%
title Python C Modules
set TMP=C:\temp
echo Python C Modules for Python 3.13

set CMODULESHOME=E:\GIT_Repositories\Python-CodeBase-Tools\Comp_V313\CMODULES

rem copy "D:\MSWIN7SDK\Lib\user32.lib"
rem copy "D:\MSWIN7SDK\Lib\kernel32.lib"
copy "C:\Program Files (x86)\Windows Kits\10\Lib\10.0.22621.0\um\x86\kernel32.lib"
copy "C:\Program Files (x86)\Windows Kits\10\Lib\10.0.22621.0\um\x86\user32.lib"
cmd