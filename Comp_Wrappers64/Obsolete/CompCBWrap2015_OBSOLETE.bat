cls
REM cl /w /Id:\codebase\WorkingSource /Id:\VisualStudio2015\VC\include /I"c:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\ucrt" /I"c:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\shared" /I"c:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\um" C4dll64.lib c4LIB.LIB zLib.Lib /LD ..\C_Source\CodeBaseWrapper.c /FoCodeBaseWrapper.obj
cl /w /Id:\codebase\WorkingSource /Id:\VisualStudio2015\VC\include /I"c:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\ucrt" /I"c:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\shared" /I"c:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\um" /LD ..\C_Source\CodeBaseWrapper.c  /link /LIBPATH:"C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\um\x64" /LIBPATH:"C:\Program Files (x86)\Windows Kits\10\Lib\10.0.15063.0\ucrt\x64" /FoCodeBaseWrapper.obj

copy CodeBaseWrapper.dll CodeBaseWrapperLarge.dll
