cls
REM cl /w /Id:\codebase\source /Id:\VisualStudio2010\VC\include /ID:\MSWIN7SDK\Include /ID:\MSWIN8SDK\Include\shared /I"D:\PGI1303\Microsoft Open Tools 11\PlatformSDK\include" c4dll.lib c4LIB.LIB zLib.Lib /LD ..\C_Source\CodeBaseWrapper.c /FoCodeBaseWrapper.obj
cl /w /Id:\codebase\source /Id:\VisualStudio2010\VC\include /ID:\MSWIN7SDK\Include c4dll.lib c4LIB.LIB zLib.Lib /LD ..\C_Source\CodeBaseWrapper.c /FoCodeBaseWrapper.obj

copy CodeBaseWrapper.dll CodeBaseWrapperLarge.dll
