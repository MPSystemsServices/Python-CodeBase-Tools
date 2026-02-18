cls
cl /WX /Id:\codebase\WorkingSource /ID:\VisualStudio2015\VC\include /IC:\Python37\include /I"C:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\shared" /I"C:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\um" /I"C:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\ucrt" c4dll.lib c4LIB.LIB zLib.Lib c:\python37\libs\python37.lib /LD ..\C_Source\CodeBasePYWrapper.c /FoCodeBasePYWrapper.obj
copy CodeBasePYWrapper.dll CodeBasePYWrapper37.pyd
copy CodeBasePYWrapper37.pyd ..\CBToolsInstallDir\codebasetools

