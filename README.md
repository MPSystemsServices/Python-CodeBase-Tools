# Python-CodeBase-Tools
Python binding for the CodeBase-for-DBF data table engine, plus tools for spreadsheet access and other functions
# Software Description
CodeBase Tools is a Python binding that encapsulates the low-level functioning of the CodeBase(tm) product.  This is a formerly proprietary product released to Open Source as of September, 2018.
Python CodeBase Tools provide the following capabilities:
* Creates, Opens, Reads, Writes and Indexes DBF tables in a format compatible with Visual FoxPro type DBF tables.
* Supports record, table, and table header locking compatible with that of Visual FoxPro, allowing hundreds of users to access the same tables simultaneously
* Optionally supports tables, fields, and index elements larger than the limits in Visual FoxPro (Maximum table size is 8GB)
* Recognizes a wide variety of fields including:
  * Character (Assumes Code Page 1252 -- Windows European Language characters)
  * Character-binary (Makes no assumptions regarding language representation so suitable for Unicode data)
  * Memo (Large text up to 2GB)(Assumes Code Page 1252 -- Windows European Language characters)
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
* Allows creation of standard VFP-style CDX indexes (with multiple "tags" for different orderings), IDX indexes (with a single ordering), and the auto updating of indexes when records are changed.
* Provides for temporary indexes which remain synchronized with the data table while in use, and are then deleted automatically when closed.
* Supports copying DBF table data to Excel spreadsheet format with the ExcelTools modules (and the non-open-source commercial product LibXL, a purchased license for which is required.)
* Supports VFP-style CURSORTOXML and XMLTOCURSOR functions for rapid conversion of DBF data to generic XML
* Copies DBF tables to CSV and System Data (fixed field length) text files as well as importing those formats directly in a DBF table.
For all functionality see the HTML documentation of this module.
# Support for Python Versions
The .PY files in Python-CodeBase-Tools package are designed to be cross platform, functional on all versions of Python from 2.7 and up.  However the compiled library .pyd file is specific to the Python version you are using.  Currently, Python 2.7, 3.6, and 3.7 are supported.  Support for earlier Python versions can be made available if interest is strong enough.
# Python-CodeBase-Tools Licensing
This package is copyright M-P Systems Services, Inc., and is released to Open Source under the GNU Lesser GPL V.3.0 license, a copy of which is found in this repository.  The CodeBase-for-DBF module, is covered by this same license.  The CodeBase package, including the core library c4dll.dll is copyright Sequiter, Inc., and is licensed under the GNU Lesser GPL v.3.0.
=======
This package is copyright M-P Systems Services, Inc., and is released to Open Source under the GNU Lesser GPL V.3.0 license, a copy of which is found in this repository.  The CodeBase-for-DBF module, is covered by this same license.
>>>>>>> 7aa1ff82fc56fb54a6ff54b3b495b221faac1bde
=======
# Integration with Visual FoxPro Applications
DBF tables can be opened both by Visual FoxPro applications and Python applications using CodeBaseTools simultaneously.  Record and table locking and buffering work correctly to support multi-user applications where both Python and VFP applications are accessing the tables.  For more information on how to integrate Python components with existing Visual FoxPro applications, see our white paper at http://www.mpss-pdx.com/Documents/Heuer_WhitePaper_PythonAsAWayForward.pdf
# OS Compatibility
The Python .pyd files which wrap the CodeBase(tm) c4dll.dll module were compiled for Windows.  They are 32-bit modules but will run properly on either 32-bit or 64-bit Windows version 7 or later.  The COM functionality in ExcelComTools and DBFXLStools2 is specific to Microsoft Windows. As of October, 2018, the LibXL product, upon which the ExcelTools module is based (also required for the DBFXLStools2 module) is a Windows-specific component.  
# Python-CodeBase-Tools Licensing
This package is copyright M-P Systems Services, Inc., and is released to Open Source under the GNU Lesser GPL V.3.0 license, a copy of which is found in this repository.  The CodeBase-for-DBF module, is covered by this same license.  The CodeBase package, including the core library c4dll.dll is copyright Sequiter, Inc., and is licensed under the GNU Lesser GPL v.3.0.
>>>>>>> 0d83a103a1eff23f1a5cb2c27b2e2aa39b7d0c78
