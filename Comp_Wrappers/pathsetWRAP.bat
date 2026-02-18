@echo off
SET LIB=d:\VisualStudio2015\VC\LIB;d:\codebase\lib\sa\msvc;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\ucrt\x86;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\um\x86
set PATH=D:\VisualStudio2015\VC\bin;%PATH%
set PATH=D:\VisualStudio2015\Common7\IDE;%PATH%
title Wrapper C Modules
set TMP=C:\temp
echo Wrapper C Modules For All PRograms

set CMODULESHOME=c:\mpsspythonscripts\cmodules

copy "C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\um\x86\kernel32.lib"
copy "C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\um\x86\user32.lib"
cmd