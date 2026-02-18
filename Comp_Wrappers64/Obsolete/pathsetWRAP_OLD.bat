@echo off
SET LIB=d:\VisualStudio2010\VC\LIB;d:\codebase\lib\sa\msvc;D:\MSWIN7SDK\Lib
set PATH=D:\VisualStudio2010\VC\bin;%PATH%
set PATH=D:\VisualStudio2010\Common7\IDE;%PATH%
title Python C Modules
set TMP=C:\temp
echo Python C Modules

set CMODULESHOME=c:\mpsspythonscripts\cmodules\Comp_V27

copy "D:\MSWIN7SDK\Lib\user32.lib"
copy "D:\MSWIN7SDK\Lib\kernel32.lib"
cmd