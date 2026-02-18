@echo off
SET LIB=d:\VisualStudio2015\VC\LIB;d:\codebase\lib\sa\msvc;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\ucrt\x64;C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\um\x64
set PATH=D:\VisualStudio2015\VC\bin;%PATH%
set PATH=D:\VisualStudio2015\Common7\IDE;%PATH%
title Wrapper C Modules
set TMP=C:\temp
set PreferredToolArchitecture=x64
echo Wrapper C Modules For All PRograms

set CMODULESHOME=c:\mpsspythonscripts\cmodules

copy "C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\um\x64\kernel32.lib"
copy "C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\um\x64\user32.lib"
cmd
