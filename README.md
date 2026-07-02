# Python-CodeBase-Tools
Python binding for the CodeBase-for-DBF data table engine, plus tools for spreadsheet access and other functions
# Software Description
CodeBase Tools is a Python binding that encapsulates the very low-level functioning of the CodeBase(tm) product with simple powerful functions, making DBF format data table access convenient to use for Python programmers using all recent versions of Python.  CodeBase(tm) is a formerly proprietary product released to Open Source as of September, 2018 and updated extensively in 2026.  The resulting DBF table access is dramatically faster than previously available DBF table access modules for Python and capable of sophisticated data access and handling features normally available only to advanced SQL database tools like SQL-Server.
Python CodeBase Tools provide the following capabilities:
* Creates, Opens, Reads, Writes and Indexes DBF tables in a format compatible with Visual FoxPro type DBF tables.
* Supports record, table, and table header locking compatible with that of Visual FoxPro, allowing hundreds of users to access the same tables simultaneously either with Python-Codebase-Tools or Visual FoxPro or both.
* Optionally supports tables, fields, and index elements larger than the limits in Visual FoxPro (Maximum table size is 8GB for the standard mode required for VFP compatibility.  A "Large Table" mode (not VFP compatible) is software selectable in both 32-bit and 64-bit versions that supports tables in multiple terrabyte sizes.)
* Recognizes a wide variety of fields including:
  * Character (Assumes Code Page 1252 -Windows European Language characters)
  * Character-binary (Makes no assumptions regarding language representation so suitable for Unicode data)
  * Memo (Large text up to 2GB)(Assumes Code Page 1252 -Windows European Language characters)
  * Memo-binary (Large text up to 2GB) (Makes no assumptions regarding language representation so suitable for Unicode data)
  * Number (decimal values up to 17 digits)
  * Integer (4-byte signed integers)
  * Boolean (Logical, True/False)
  * Currency (Fixed 4-decimal point exact representation of money values)
  * Date (Year, Month, Day)
  * Datetime (Year, Month, Day, Hour, Minute, Second)
  * Float (like Number)
  * Double (like Number)
  * General (treated like Memo-binary, OLE component content is not recognized)
* Allows creation of standard VFP-style CDX indexes (with multiple "tags" for different orderings), IDX indexes (with a single ordering), and the auto updating of indexes when records are changed including updating all distinct index tags in a CDX index.
* Provides for temporary indexes which remain synchronized with the data table while in use, and are then deleted automatically when closed.
* Supports copying DBF table data to Excel spreadsheet format with the ExcelTools modules (and the non-open-source commercial product LibXL, a purchased license for which is required.)
* Supports VFP-style CURSORTOXML and XMLTOCURSOR functions for rapid conversion of DBF data to generic XML and back to DBF.
* Provides hundreds of powerful functions that emulate the capabilities of Visual FoxPro.  VFP, though now a discontinued Microsoft product, was renowned for being able to create applications 3-4 times faster than comparable .NET tools due to its unmatched data handling capabilities.
* Copies DBF tables to CSV and System Data (fixed field length) text files as well as importing those formats directly into a DBF table.
For all functionality see the HTML documentation of this module.

This version of Python-CodeBase-Tools was updated as of May, 2026.
# Support for Python and OS Versions
While the CodeBase(tm) product was originally provided in both Windows and Linux versions, at this time only the compiled Windows version is available (both 32-bit and 64-bit) in the [**CodeBase-for-DBF**](https://github.com/MPSystemsServices/CodeBase-for-DBF) GitHub repository.  Consequently, the Python CodeBase_Tools package is available only for Windows.  The maintainers may have time in the future to extend the product to Linux, depending on community interest.

The .PY files in the Python-CodeBase-Tools package are designed to be cross version, functional on all Windows versions of Python from 2.7 and up.  Python versions 3.1 through 3.5 are not supported, however, as they are effectively obsolete.  Versions 3.6, 3.7, 3.8 and 3.9 are supported with version-specific PYD files that are included in the distribution directory.  For versions 3.10 and above, the PYD file which implements all the functionality was compiled from an upgraded version of the .C source that allows creation of a cross-version PYD file using the Python "Limited API" functions.  PYD files supporting 64-bit Python have been compiled using the cross-version library as well, so 32-bit and 64-bit versions 3.10, 3.11, 3.12, 3.13, 3.14 and beyond are supported.  No further support for versions of Python prior to 2.7 will be provided.  This package has been tested with traditional CPython.  It may work with other forms of Python, but we haven't tried them.

The Excel copy features are currently only available for 32-bit Python.

Note that you do NOT need to download any files from the CodeBase-for-DBF repository, as the required compiled DLL files are included in this repository.  Of course, you are free to customize the .C source and recompile the CodeBase-for-DBF DLLS should you need to do so.
# Python-CodeBase-Tools Licensing
This package is copyright M-P Systems Services, Inc., and is released to Open Source under the GNU Lesser GPL V.3.0 license, a copy of which is found in this repository.  The CodeBase-for-DBF module, is covered by this same license.  The CodeBase package, including the core library c4dll.dll and c4dll64.dll is copyright Sequiter, Inc., and is licensed under the GNU Lesser GPL v.3.0.  There is an old demo version of libxl.dll included for testing purposes for the DBF to Excel module.  However, that is commercial software and no license information is included in the `LibXLLicenseInfo.TXT` file.  To use this module without the size limits and demo messages injected into the output, you'll need to acquire your own libxl license from libxl.com.
# Integration with Visual FoxPro Applications
DBF tables can be opened both by Visual FoxPro applications and Python applications using this package simultaneously.  Record and table locking and buffering work correctly to support multi-user applications where both Python and VFP applications are accessing the tables.  For more information on how to integrate Python components with existing Visual FoxPro applications, see our white paper at https://mpss-pdx.com/white-paper.  See notes above that the "Large Model" of CodeBase-Tools is NOT compatible with Visual FoxPro for simultaneous access.
# OS Compatibility
The Python .pyd files which wrap the CodeBase(tm) c4dll.dll and c4dll64.dll modules were compiled for Windows.  The 32-bit modules will run properly on either 32-bit or 64-bit Windows version 7 or later.  The 64-bit modules will run only on 64-bit Windows.  The COM functionality in ExcelComTools and DBFXLStools2 is specific to Microsoft Windows. As of October, 2018, the LibXL product, upon which the ExcelTools module is based (also required for the DBFXLStools2 module) is a Windows-specific component.  A 64-bit version of ExcelTools is not provided in this release, but may be added in the future.  The Python implementation you are running will determine which 32-bit or 64-bit version of this package you will be running. 
# Installation and Usage
For details see the [Installation Instructions](InstallationAndUsage.md) information.
# Other Optional components
The .PYD files (like CodeBasePYWrapper3X.pyd) provide native Python API access directly to the functions in c4dll.dll or c4dll64.dll without need for ctypes to access the raw .DLL C functions.  There is a significant speed and simplicity benefit from moving more program functionality directly into a compiled C program using the Python C API library, enabling direct access to the contained functions from the Python interpreter.  However, before we had developed the .PYD file(s) we developed "wrapper" DLLs that could be accessed directly by Python using ctypes.  We have not obsoleted these components, and they are found in the Comp_Wrappers and the Comp_Wrappers64 directories.  Full documentation is not provided for these components as we don't really see a need for them, but if you wish to make use of them, send us a GitHub message for more information.

