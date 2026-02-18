As of Sept. 19, 2023, this directory is set up to compile the CodeBaseWrapper64.DLL and its companion CodeBaseWrapperLarge64.DLL

These modules provide VFP-like functions based on the newly compiled CodeBase C4DLL64.DLL which is the 64-bit version of the CodeBase
product (open source)

Use pathsetWrap.bat to make sure pathing is correct.

To recompile CodeBaseWrapper64.DLL, open CodeBaseWrapper64.vcxproj with Visual Studio 2022.  This should then allow you to 
"Build" the project.  This results in the finished DLL being written to the x64\release directory.  The batch file
MakeCBWrapDLLs.bat will copy over the required .DLL and .LIB files into this directory for subsequent deployment in the system.

To test the new compile, you can run cb64wraptest.py using a 64-bit instance of Python (a 32-bit version WILL NOT WORK!)  This
application is intended ONLY for Python 3.x versions, as it requires all text strings to be converted to bytes type variables
before passing to the .DLL.

The source code for this build is found in c:\mpsspythonscripts\cmodules\c_source.  The source code is identical for 64-bit and
32-bit compilation.  However, note that the d4all.h file, found in the d:\codebase\WorkingSource directory must be set to 64-bit
configuration by using the appropriate .BAT file in the WorkingSource directory before any compilation here is attempted.

As of this date, the 64-bit version of LibXL and its wrapper has NOT been developed.
