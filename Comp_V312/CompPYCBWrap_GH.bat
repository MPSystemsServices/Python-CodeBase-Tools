cls
cl /WX /IE:\GIT_Repositories\CodeBase-for-DBF\WorkingSource /ID:\VisualStudio2015\VC\include /IC:\Python312\include /I"C:\Program Files (x86)\Windows Kits\10\Include\10.0.22621.0\shared" /I"C:\Program Files (x86)\Windows Kits\10\Include\10.0.22621.0\um" /I"C:\Program Files (x86)\Windows Kits\10\Include\10.0.22621.0\ucrt" c4dll.lib zLib.Lib c:\python312\libs\python312.lib /LD ..\C_Source\CodeBasePYWrapper.c /FoCodeBasePYWrapper.obj
copy CodeBasePYWrapper.dll CodeBasePYWrapper312.pyd
REM copy CodeBasePYWrapper312.pyd ..\CBToolsInstallDir\codebasetools
REM ONLY DO THE ABOVE WHEN KNOWN TO BE WORKING...

