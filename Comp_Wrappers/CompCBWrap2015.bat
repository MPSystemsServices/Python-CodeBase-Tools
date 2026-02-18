cls
cl /w /Id:\codebase\WorkingSource /Id:\VisualStudio2015\VC\include /I"c:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\ucrt" /I"c:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\shared" /I"c:\Program Files (x86)\Windows Kits\10\Include\10.0.15063.0\um" c4dll.lib c4LIB.LIB zLib.Lib /LD ..\C_Source\CodeBaseWrapper.c /FoCodeBaseWrapper.obj

copy CodeBaseWrapper.dll CodeBaseWrapperLarge.dll
