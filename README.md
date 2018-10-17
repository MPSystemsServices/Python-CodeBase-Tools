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
The .PY files in Python-CodeBase-Tools to function are designed to be cross platform.  However the compiled library .pyd file is specific to the Python version you are using.  Currently, Python 2.7, 3.6, and 3.7 are supported.  Support for earlier Python versions can be made available if interest is strong enough.
# Python-CodeBase-Tools Licensing
This package is copyright M-P Systems Services, Inc., and is released to Open Source under the GNU Lesser GPL V.3.0 license, a copy of which is found in this repository.  The CodeBase-for-DBF module, is covered by this same license.
