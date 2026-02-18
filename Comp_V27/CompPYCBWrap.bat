cls
cl /WX /ID:\codebase\WorkingSource /ID:\VisualStudio2010\VC\include /IC:\Python27\include /ID:\MSWin7SDK\Include /ID:\MSWin7SDK\Include\um c4dll.lib c4LIB.LIB zLib.Lib c:\python27\libs\python27.lib /LD ..\C_Source\CodeBasePYWrapper.c /FoCodeBasePYWrapper.obj
copy CodeBasePYWrapper.dll CodeBasePYWrapper27.pyd
copy CodeBasePYWrapper27.pyd ..\CBToolsInstallDir\codebasetools
