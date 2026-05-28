cls
cl /WX /IE:\GIT_Repositories\CodeBase-for-DBF\WorkingSource /ID:\VisualStudio2022\VC\Tools\MSVC\14.44.35207\include /IC:\Python313\include /I"C:\Program Files (x86)\Windows Kits\10\Include\10.0.22621.0\shared" /I"C:\Program Files (x86)\Windows Kits\10\Include\10.0.22621.0\um" /I"C:\Program Files (x86)\Windows Kits\10\Include\10.0.22621.0\ucrt" c4dll.lib zLib.Lib c:\python313\libs\python313.lib /LD ..\C_Source\CodeBasePYWrapper313.c /FoCodeBasePYWrapper.obj
copy CodeBasePYWrapper.dll CodeBasePYWrapper313.pyd
REM copy CodeBasePYWrapper313.pyd ..\CBToolsInstallDir\codebasetools
REM ONLY DO THE ABOVE WHEN KNOWN TO BE WORKING...

