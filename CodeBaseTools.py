# Copyright (C) 2010-2018 by M-P Systems Services, Inc.,
# PMB 136, 1631 NE Broadway, Portland, OR 97232.
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, version 3 of the License.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
# or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General
# Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
""" This module is a wrapper and Pythonizer for the CodeBasePYWrapper.pyd shared library written
in 'C' using the Python 'C' API, which provides relatively easy and Pythonic access to Visual FoxPro DBF tables.  The
underlying engine is the CodeBase library, which, as of September 28, 2018, is released by its owner Sequiter, Inc.,
into Open Source, and is available for free use under the GNU Lesser GPL V3.0 license.

Basic documenation for each function is provided in the docstrings below.

NOTE FOR THIS VERSION 2.00 -- October, 2014:
    This version is a wrapper for a compiled core python module named CodeBasePYWrapper.pyd which is a revised and
improved version of the original CodeBaseWrapper.dll.  The core module is now built with the Python C-API which
allows much faster transfer of data from disk to memory.  The original dll was accessed using the cTypes modules
and the Win32 extensions.  cTypes is not required for this enhanced version.

Python Version information:
This module has been developed and tested with Python 2.7.6 and 2.7.10.

Tested also with Python 3.6.  Earlier 3x versions haven't been tested and may not work.

***********************************************
String handling for Python 2/3 compatibility
***********************************************
Effective with the April, 2018 changes for Python 2 and 3 compatibility, all text values passed to the
Codebase Tools engine will be coerced to str types.  That means plain strings for 2.x and Unicode
strings for 3.x.  The engine will test for these types and throw an error if the type is wrong.
This applies to path names (no "pathlike" objects), alias names, field names, tag names, and logical expressions
for selecting and filtering.

Field return values will be converted per the table below, if conversion is requested.  Unconverted field
contents will be returned as type bytes in all versions, if stored in the tables as text.  Field setting values may
be supplied as any valid type (including text representations of values like numbers and dates) and will be converted
by the engine.

**************************************
FoxPro and Visual FoxPro Compatibility
**************************************
All field types used in FoxPro 2.6 DOS and Windows are supported natively by this tool.  Also the indexes
generated are compatible with .CDX compound indexes with one or more index tags.

Visual FoxPro 5.0 through 9.0 SP2 tables are supported with the following exceptions:

1. Tables contained in a database container can be read, written to and reindexed, but only the short field names
   will be usable.  Field names that are 10 characters long or less can be accessed by their normal names.
2. Tables contained in a database container cannot have their structure changed, nor have indexes removed, added,
   or changed, as those changes will not be recorded in the database container.
3. VFP supports multiple Windows Code Pages, including 1252, Western Europe.  This module only supports
   tables marked for Code Page 1252, however, NO translations between code pages are performed by this module.
4. VFP may or may not support Unicode strings, including multibyte languages like Cantonese, in its C and M type
   fields.  See the docs below on string and Unicode handling in this module.
5. Most, but not all field types are supported, and some type codes are different as given in the following table:

VFP Type Code  | Type Name        | CodeBase Type Code     |  Python 2 Type    |  Python 3 Type***
  B              Double               B                         float               float
  C              Character            C                         str                 str (Unicode)
  C              Character(binary)    Z                         bytearray           bytes
                                                                (opt) unicode       (opt)str (Unicode)
  D              Date                 D                         datetime.date       datetime.date
  F              Float                F (same as Numeric)       float               float
  G              General              G (see **note below)      bytes               bytes
  I              Integer              I                         int                 int
  I              Integer(auto incr)   [not supported
                                       - see *note below]
  L              Logical              L                         bool                bool
  M              Memo                 M                         str                 str (Unicode)
  M              Memo(binary)         X                         bytes               bytes
  N              Number               N                         float               float
  Q              Varbinary            [not supported]
  T              Datetime             T                         datetime.datetime   datetime.datetime
  V              Varchar              [not supported]
  V              Varchar(binary)      [not supported]
  W              Blob                 [not supported]
  Y              Currency             Y                         Decimal             Decimal

*Note -- The last Codebase version added support for an auto incrementing integer field, but these fields are NOT
recognized by Visual FoxPro, and VFP auto incrementing integer fields are not recognized by Codebase and can
cause Codebase to crash.  This module therefore does not support auto incrementing integer fields directly, but a
utility function getNewKey() is provided as a source for consecutive integer values.

**Note -- CodeBase and this module support G (General) type fields for reading and writing, but unlike in Visual
Foxpro, their contents cannot be filled with some kind of APPEND GENERAL sort of function.  VFP tables which have
had a G field populated by the APPEND GENERAL command from an OLE document can be copied to other tables by this
module, however, as these fields are read as simple strings of generic binary data (possibly containing null bytes),
exactly as X Memo(binary) fields, there is no interpretation done with them.

***Note -- All "strings" in Python 3.x are Unicode.  Sequences of bytes that are not intended to represent text of any
kind are stored in "bytes" variables, not "strings".  Accordingly, for Py 3.x the default Python type target for a
standard char or memo field is a string, and that of char or memo binary fields is bytes.  But you may want to store
Unicode data in binary fields to ensure no CodePage conversions are attempted by FoxPro.  As a result options are
provided to write to and read from binary fields in several different Unicode codings.

***********************************************************************
String handling and internationalization... Depends on Python version
***********************************************************************
Field Names:
    ASCII is default, but extended ASCII with Code Page 1252 is supported for displaying table structure,
    but fields having names with characters above ASCII 127 cannot be accessed for reading and writing.
Tag Names:
    ASCII is default, but extended ASCII tag names work properly

Alias Names:
    Must be standard ASCII strings.  If specified with cp1252 high order bit characters, the Alias Name will be
    coerced to 7-bit ASCII.

Table Names:
    Names supported by the operating system will work (but see note above for Alias Names, as auto-assigned
    Alias Names are also forced to 7-bit ASCII strings.)

Data Fields:
------------
Python 2.7 (earlier versions may work but are not officially supported)
----------
Char and Memo:
    Default data storage is 8-bit strings as found in Python str objects.  Embedded nulls will be treated as end of
    string markers and will not be stored.  Unicode strings will be coerced to Code Page 1252 before saving.
    Applies to scatter, gather, insert, replace, and curval type functions.

    If bytearray types (with or without embedded nulls) are supplied for these fields, they will be copied into
    the target field as pure binary data.  In the case of char fields, any unused space to the right of the data
    will be padded with blanks.  Memo fields will contain only the data content.  If you write binary data from
    bytearray objects into these standard text fields, be sure that you handle the content accordingly when
    reading it back.  Also beware of FoxPro Code Page translations to which these fields are subject by default.

    The following special coding will be allowed optionally in these functions for string and unicode object
    data sources.:
    ' ', 'X', or None = The default storage and retrieval
    'W' = Windows Double Byte Unicode encoding for the specified Code Page (default Code Page in the U.S.A. is
          usually cp1252) on the local machine.  Source data may be string or bytearray, and conversion will be
          performed.  Unicode will be coerced to the default Code Page.
    '8' = UTF-8 encoding. Source may be string, bytearray, or Unicode.

Char(binary) and Memo(binary):
    Default data storage is a non-text sequence of 8-bit bytes with no textual meaning.  The raw contents of strings
    will be stored, but with strings terminated when a NULL is encountered.  Bytearrays will be copied character
    for character, regardless of the presence of null bytes.

    The following special coding will be allowed optionally in these functions with string or unicode type
    data source objects:
    ' ', 'X', or None = The default storage and retrieval
    'D' = Windows Double Byte Unicode encoding for the specified Code Page (default Code Page in the U.S.A. is
          usually cp1252) on the local machine.
    'U' = UTF-8 encoding.
    'C' = Custom Code Page code.  See Python codecs documentation for allowable values.  Examples include:
          "cp1252" - Western Europe, "cp1251" - Cyrillic (Russian, etc.), "gb2312" - Simplified Chinese.

Python 3.x (3.4 and higher.  Earlier versions not supported.)
------------------------------------------------------------
Char and Memo:
    All 3.x Python strings are Unicode.  This module will attempt to code Python strings for storage in the
    field type specified.  If the Unicode string value cannot be converted into a form suitable for the field type,
    an error will be triggered. Applies to scatter, gather, insert, replace, and curval type functions.
    By default conversion of the strings to and from 8-bit Code Page 1252 will be attempted.  This will handle
    most special characters for Western European languages.

    Version 3.x also has a bytes and bytearray type (differences are very subtle, but the both contain
    sequences of bytes that don't necessarily have avy textual meaning.  Both of these binary data types may be
    stored in both plain and binary char/memo fields, but you will need to be careful with their data
    reteieval.  In all such cases, the contents of the bytes and bytearray objects will be stored byte for
    byte into the target field.

    The following special coding will be allowed optionally in these functions:
    ' ', 'X', or None = The default storage and retrieval (cp1252, 8-bit coding or plain ASCII)
    'W' = Windows Double Byte Unicode encoding for the specified Code Page (default Code Page in the U.S.A. is
          usually cp1252) on the local machine.
    '8' = UTF-8 encoding.

Char(binary) and Memo(binary):
    Default data storage is a non-text sequence of 8-bit bytes with no textual meaning.
     The following special coding will be allowed optionally in these functions:
    ' ', 'X', or None = The default storage and retrieval.  For Bytes and ByteArray data, this will be a simple
    copy of the binary data.  Char fields will be filled out with nulls.  String (Unicode) data will be stored
    as Code Page 1252 and padded out with blanks. (For pure binary data, memo (binary) should be preferred.)
    'D' = Windows Double Byte Unicode encoding for the specified Code Page (default Code Page in the U.S.A. is
          usually cp1252) on the local machine.
    'U' = UTF-8 encoding.
    'C' = Custom Code Page code.  See Python codecs documentation for allowable values.  Examples include:
          "cp1252" - Western Europe, "cp1251" - Cyrillic (Russian, etc.), "gb2312" - Simplified Chinese.  There
          are many others.

All Python versions
-------------------
Note that with all the non-default codings, you are expected to know what codings are contained the fields
you are reading and writing.  Data which cannot be transformed by the Python codecs manager will trigger
a Python error and will need to be trapped by a try... except... block.

Note that if you specify a custom coding and are sharing the data with Visual FoxPro, you are responsible for
ensuring the codings match on both sides.  VFP will dynamically change field data in Char and Memo
fields (but not their binary versiona) if the table and local machine Code Pages do not match.  This could
result in badly corrupted data, as this module does no automatic Code Page conversions.

Also, be careful with functions that read or write data from and to multiple fields.  In such cases, the custom
coding spcifiers shown above will be applied to every field of the appropriate type.  If that is not what you
want, process fields individually using replace() or curvalstrex().
"""
import decimal
import os
import csv
from time import time, localtime, strftime, sleep
from locale import atof, atoi
from datetime import date, datetime
import copy
import sys
import string
import shutil
import random
from xml.dom import minidom
import CodeBasePYWrapper as cbp
import MPSSBaseTools as mTools
import collections

if sys.version_info[0] <= 2:
    _ver3x = False
    xLongType = long
else:
    _ver3x = True
    xLongType = int

__author__ = "Jim Heuer"
__version__ = "2.01"
__copyright__ = "M-P System Services, Inc., Portland, OR, 2008-2016"
__versiondate__ = "February 19, 2015"
__license__ = """Provided to all users on an as-is basis.  No warranties, express or implied are
made for this code.  Users are asked to supply information on required changes, perceived bugs,
and other comments about the code to the author.  This license and author information must be included
in any derived versions of this code modified by others.  The author reserves the right to modify
this code at any time and may or may not post updated versions to a website for retrieval by others.
Use of this code is completely free.  HOWEVER, this code will NOT WORK without a copy
of Sequiter Software's CodeBase software.  It also requires a compiled version of the CodeBasePYWrapper.c
program, built into a DLL (copied to a PYD) with Microsoft Visual Studio C, using the header files supplied in the
CodeBase install package.

This software is working successfully in the author's systems with CodeBase 6.5 updated with all the latest
CodeBase updates and enhancements available to registered users.  Other versions may or may not work with this
software, and the author takes no responsibility and makes no guarantees for compatibility with this or any other
versions of CodeBase.

No support is provided to users of this code by the author.  However, the author will respond to questions on a
time available basis.  The author may be contacted at jsheuer at mpss-pdx.com

This software and its companion CodeBasePYWrapper.c program are provided free of charge.  The terms of use
for this software are the GNU Lesser GPL, Version 3.0, a copy of which should have been provided with this file.
"""

# 10/16/2014 JSH. Start conversion from Version 1.0 of this code
# 10/30/2014 JSH. Final testing completed.
# 12/02/2014 JSH. Added cbToolsX() which returns a unique stand-alone version of the _cbTools() object and
#                  supports the use of multiple sessions to keep table opens and closes isolated from those
#                  of other modules using another instance.
# 01/06/2015 JSH. Added extra table open try to getNewKey().
# 02/19/2015 JSH. Major changes to use() method.  No longer allows same table to be opened multiple times
#                 with the same alias.

gbIsRandomized = False
"""
Set True when the randomizer functions are seeded, as that only needs to happen once.
"""
gnNextKeyTableNameLength = 0
gcLastErrorMessage = ""


class VFPFIELD(object):
    def __init__(self):
        self.cName = ""
        self.cType = ""
        self.nWidth = 0
        self.nDecimals = 0
        self.bNulls = False


class VFPINDEXTAG(object):
    def __init__(self):
        self.cTagName = ""
        self.cTagExpr = ""
        self.cTagFilt = ""
        self.nDirection = 0
        self.nUnique = 0


class _cbTools(object):

    recDict = {}
    recArray = list()
    _tablenames = list()
    
    def __init__(self, bLargeMode=False):
        """ Initializer routine. """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        self.bUseDecimals = False  # Set this to True to return Number field values as decimal.Decimal() type values.
        self.tally = 0 # Stores the count of records most recently processed by several methods.
        self.xCalcTypes = {"SUM": 1, "AVG": 2, "MAX": 3, "MIN": 4}
        # if not _ver3x:
        #     self.workIdent = string.maketrans('', '')
        # else:
        #     self.workIdent = str.maketrans('', '')
        self.workIdent = None
        self.cNonPrintables = "\x01\x02\x03\x04\x05\x06\x07\x08\x09\x0A\x0B\x0C\x0D\x0E\x0F\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F"
        self.xNonPrintables = dict()
        for c in self.cNonPrintables:
            self.xNonPrintables[ord(c)] = None
        self.cbt = cbp
        self.nDataSession = self.cbt.initdatasession(bLargeMode)
        self.oCSV = csv
        self.xLastFieldList = []  # Used by scattertorecord() to avoid having to repeat afields()
        self.cLastRecordAlias = ""  # ditto
        self.xLastRecordTemplate = None
        self.name = "CodeBaseTools"
        self.cPreferredEncoding = "cp1252"  # Western European languages.  Like "Latin-1".

    def __getitem__(self,xfile):
        return self.TableObj(xfile)

    def __del__(self):
        if self.nDataSession >= 0:
            if self.cbt is not None:
                try:
                    self.cbt.closedatasession(self.nDataSession)
                except:
                    pass  # Nothing to be done about it.
            # print "Closed Datasession", self.nDataSession
        self.nDataSession = -1
        self.cbt = None

    def cb_shutdown(self):
        """
        Manually does what the __del__() does above.
        """
        if self.nDataSession >= 0:
            self.cbt.closedatasession(self.nDataSession)
            self.nDataSession = -1
        return True

    def TableObj(self, *args, **kwargs):
        """
        Create database object from cbTools connection object.  Ignores deleted record.
        :tablename: can include full path
        optional :reset: will cause file to be re-opened if already open
        """
        return TableObj( self, *args, **kwargs)

    def setcustomencoding(self, lpcEncoding=""):
        """
        If you want string data to be retrieved and saved in an encoding different from the three basic options
        (ASCII, UTF-8, and Code Page 1252), you can specify one the many standard Python encodings (see the Python
        documentation for the latest complete list.)  For this custom configuration to be applied, the conversion
        code of 'C' must be specified and the target field must be a binary version (either of char or memo).
        :param lpcEncoding: Python encoding name like "cp1252" or "cp1251".
        :return: True.  Note that there is no error checking on the value of the lpcEncoding parameter!  If
        you specify a non-existent name, this module will generate a coding error when processing binary data
        with the 'C' conversion option.
        """
        self.cbt.setcustomencoding(lpcEncoding)
        self.cPreferredEncoding = lpcEncoding
        return True

    def setdateformat(self, lpcFormat=""):
        """
        Codebase, like the USA versions of Visual FoxPro assumes U.S. standard date
        expressions of the form MM/DD/YY or MM/DD/YYYY, in contrast to other parts of
        the world that use DD/MM/YY or DD/MM/YYYY or other combinations.  This default
        assumptions about date text strings can be changed.  This is especially useful
        when importing text data via the appendfrom() function.  If non-numeric
        characters are found in the text, and it is intended to be stored in a date or
        datetime field, then an attempt will be made to parse the text based on the
        standard date expression, either MM/DD/YYYY or an alternate set prior to the
        append process.
        Parameters:
         - lpcFormat - This must be a character string with at minimum the characters
           'YY' or 'YYYY', and 'MM', and 'DD' in some combination with punctuation or
           it must be one of the Visual FoxPro standard SET DATE TO values such as:
           AMERICAN = MM/DD/YY
           ANSI = YY.MM.DD
           BRITISH = DD/MM/YY
           FRENCH = DD/MM/YY
           GERMAN = DD.MM.YY
           ITALIAN = DD-MM-YY
           JAPAN = YY/MM/DD
           TAIWAN = YY/MM/DD
           USA = MM-DD-YY
           MDY = MM/DD/YY
           DMY = DD/MM/YY
           YMD = YY/MM/DD
           Pass NULL or the empty string to restore the MDY/AMERICAN default.
        Returns 1 on OK, 0 on any kind of error, mainly unrecognized formats.
        The total length of the lcFormat string must be no greater than 16 bytes.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lnReturn = self.cbt.setdateformat(lpcFormat)
        if lnReturn == 0:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnReturn

    def getdateformat(self):
        """
        Obtains the current setting of date format from the CodeBase engine.  Returns a text string indicting
        the type of format. See setdateformat() for format codes.  Will always return a string value.
        """
        return self.cbt.getdateformat()

    def newfield(self, cName, cType, nWidth, nDecimals, bNulls):
        lxRet = VFPFIELD()
        lxRet.cName = cName
        lxRet.cType = cType
        lxRet.nWidth = nWidth
        lxRet.nDecimals = nDecimals
        lxRet.bNulls = bNulls
        return lxRet

    def isexclusive(self):
        """
        Returns the exclusive open status on the currently selected table.  Returns True if exclusive, False if
        NOT exclusive.  Returns None on error.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        bRet = self.cbt.isexclusive()
        if bRet is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber  = self.cbt.geterrornumber()
        return bRet

    def islargemode(self):
        """
        Returns the status of Large Mode operations.  If this instance of the Tools is using large mode
        then, this will return True.  Otherwise False.  If True, this tools instance may NOT be used to
        open DBF tables that might also be opened concurrently by Visual FoxPro applications.
        """
        return self.cbt.getlargemode()

    def use(self, tableName, alias="", readOnly=False, exclusive=False, noBuffering=False):
        """
        Basic method for starting work with a DBF table. This must be called before any work with
        the table can be performed. After the table is opened with use(), it is the "Currently
        Selected Table" and can be accessed by the other methods of this class that operate on the
        currently selected table. If you call use() again for another table, that table becomes the
        currently selected table. You can change the currently selected table later by calling the
        select() method and passing to it the alias name assigned by the original use() method call.
         When finished with the table call close() and pass the alias name assigned when the table
        was opened. The method closedatabases() will also close the table (and close all other open
        tables too). (Changed the mechanism by which the alias is handled 02/19/2015. JSH.)

        Parameters:
        :param: tableName:  path name to the dbf table you want to open.  Must have the .dbf extension
                unless the table you wish to open has some other extension -- .DBF is not
                enforced, in other words.  Up to you as to how much info you pass in the
                path name, but if the file can't be found the method reports an error.

        :param: alias: an optional character string with an alphanumeric string starting with a
                letter or underscore which will be used as the alias for subsequent table and field
                manipulations.  The presence of this value determines how tables are opened when
                multiple tables might be opened with the same base file name but in different
                directories.  Two cases:
                1) An alias is specified.
                If a table with that alias is already open (either this same table or another one in
                a different directory but with the same base name), that table is closed, and then
                this table open happens and the alias you specified is assigned.  Otherwise the
                table is simply opened and this alias name is assigned.
                2) NO alias is specified.
                A default alias is determined, which is equal to the base table name (other than
                path and extension) of the table with embedded spaces and punctuation removed, if any.
                If this same table is already open with that alias, it is closed and re-opened with your
                required properties.
                If a different table is already open with that alias, it is left open, and a new
                unique alias is created as 'TMP' followed by a string of 12 random characters.  The
                table specified in this function call is then opened with that new alias.

                NOTE: If you do NOT specify an alias, the alias name is NOT guaranteed to be the table
                base name.  If you will need to manipulate the tables and you don't specify an alias
                name, you MUST store the alias name value provided by the alias() method immediately
                after the table is opened, and use that alias name for select() and similar functions
                which take an alias name parameter.

                The same table is allowed to be opened multiple times, each with its own alias.  Remember
                in that case that closing one of the instances will NOT close the other(s).

        :param: readOnly: boolean value which defaults to False.  If set to True, then the table is
                opened in a mode which prevents any record updates or record appends.
                Reads from tables opened readOnly are faster, and readOnly is preferred when
                writing is not actually required.

        :param: exclusive: boolean value which defaults to False.  If set to True, then the table is
                opened in exclusive mode.  This means that while the table is open, no other
                process may open or access it.  Exclusive mode is required for safely
                applying indexes in situations where other processes may attempt access.  When
                exclusive is True, the setting for readOnly is ignored and defaults to False.

        :param: noBuffering: Boolean value defaults to False.  Set to True to force the CodeBase engine
                to read/write directly from/to disk without any in-memory buffering of data.
                This may be useful for tables with lots of multi-user access.  However you
                may still need to call flush() if data corruption issues are encountered.

        Returns:
            True on success, False on any error
        """
        lbReturn = True
        if tableName == "" or tableName == "?":
            import easygui  # on demand.  should only be used in interactive mode
            tableName = easygui.fileopenbox("Select file to open", "Open DBF file",
                                            " *.dbf")  # the initial space IS required because of a bug in easygui
        if (tableName == "") or (tableName is None):
            self.cErrorMessage = "NO table name supplied"
            self.nErrorNumber = -49348
            return False
        # if "LTLACCESSENABLE" in tableName:
        #     import traceback
        #     stack_str = ''.join(traceback.format_stack())
        #     MPSSCommon.MPSSBaseTools.MPMsgBox("OPENING LAE with Alias: >" + alias + "< stack:" + stack_str)
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        if not exclusive:
            nReadOnly = (1 if readOnly else 0)
            nNoBuffering = (1 if noBuffering else 0)
            lbReturn = self.cbt.use(tableName, alias, nReadOnly, nNoBuffering)
        else:
            lbReturn = self.cbt.useexcl(tableName, alias)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage() + " OPEN FAILED"
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def maketempindex(self, lcExpr, tagFilter="", descending=0):
        """
        Similar to indexon except that this creates a temporary, single use index that only lives as long
        as the table is open.  This index has only one order embedded in it.  This method call applies the
        index to the currently selected table and makes the new temp index the current order selection.  You
        can traverse the file in the temp index order but seek() is not currently supported.  However, the
        locate() method will exploit available temporary indexes if they allow the locate to be optimizable.

        You can change to different orders in the regular index file (CDX) by using setorderto() while you keep
        this temp index open.  Open temp indexes are updated when records are changed, added or deleted in the
        table.

        The method returns a temp index code which you should store in case you need to set the ordering back to
        this temp index with selecttempindex(), which uses the index code as its parameter.

        The temp index is associated with one alias -- the currently selected one when the index was created.  If
        the same table is open with a different alias, it cannot be accessed by that different alias with this
        index order.

        Multiple temporary indexes can be opened for a table at the same time, and you can switch back and forth
        among them as required.  Closing the table with the closetable() method, closes all temporary indexes
        associated with the specified alias and deletes the files they were stored in.

        lcExpr MUST be a valid VFP index expression, and tagFilter, if non-blank, must be a valid VFP logical
        filter expression. See indexon() for details.  Set descending to 1 to have the index constructed in
        reverse order (highest at the top).  The 'unique' option is not available for temp indexes.

        Returns a valid index code >= 0 or -1 on error.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        if not mTools.isstr(tagFilter):
            raise TypeError("tagFilter parm MUST be a string or be omitted")
        lnResult = self.cbt.tempindex(lcExpr, tagFilter, descending)
        if lnResult == -1:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnResult

    def selecttempindex(self, lnCode):
        """
        Once a temporary index has been created with maketempindex, if you have needed to change the index order
        for any reason, you can restore the temporary ordering of a previously created temp index. Store the return
        value from maketempindex() and use it as the lnCode parm value in this method.  Subsequent goto() and skip()
        method calls will use the temp index.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lbReturn = self.cbt.selecttempindex(lnCode)
        if (not lbReturn):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber  = self.cbt.geterrornumber()
        return lbReturn

    def closetempindex(self, lnCode):
        """
        Temporary indexes are automatically destroyed when the table they relate to is closed.  If you want
        to close the temporary index before the table, pass the index code to this method.  You may wish to do
        this if you are going to be adding a large number of records to the table, and don't want the possibly
        useless overhead of updating the temporary table(s) on every append.  Returns True on OK, False on failure.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lbReturn = self.cbt.closetempindex(lnCode)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def updatestructure(self, cSourceTable, cTargetTable):
        """
        This function examines the field structure and index tags of the source table referenced by cSourceTable
        and forces the table named in cTargetTable (not currently open) to have the same structure and index tags.  The
        current content of the target table is retained so long as all of its fields are retained.  Data in fields
        which are not contained in the cSourceTable table will be lost.
        :param cSourceTable: Alias of pattern table, the source of the structure/tags, not currently open.
        :param cTargetTable: DBF table to be modified with the structure and index tags of the pattern (source) table.
        :return: True on success, False on failure or any error.

        NOTE: Do NOT, repeat DO NOT, use this method to modify the structure or indexes of a VFP table contained
        in a VFP database container.  If you do, the table may become un-readable!
        """
        bReturn = self.use(cSourceTable, readOnly=True)
        cErr = ""
        nErr = 0
        cTempFile = ""
        cTempAlias = ""
        # mTools = MPSSCommon.MPSSBaseTools
        cPath, cFile = os.path.split(cTargetTable)
        cTempFile = os.path.join(cPath, "X" + mTools.strRand(8) + ".DBF")

        if not bReturn:
            cErr = self.cErrorMessage
            nErr = self.nErrorNumber
        else:
            cSrcAlias = self.alias()
            xStru = self.afields()

            bReturn = self.createtable(cTempFile, xStru)
            if bReturn:
                cTempAlias = self.alias()
                self.closetable(cTempAlias)
                self.use(cTempFile, alias=cTempAlias, exclusive=True)
                self.closetable(cSrcAlias)
                nACount = self.appendfrom(cTempAlias, cSource=cTargetTable, cType="DBF")
                if nACount == -1:
                    bReturn = False
                    cErr = self.cErrorMessage
                    nErr = self.nErrorNumber
                if bReturn:
                    bReturn = self.use(cSourceTable, alias=cSrcAlias, readOnly=True)
                    if bReturn:
                        bReturn = self.copyindexto(cSrcAlias, cTempAlias)
                    if not bReturn:
                        cErr = self.cErrorMessage
                        nErr = self.nErrorNumber
                    self.closetable(cSrcAlias)
            else:
                cErr = self.cErrorMessage
                nErr = self.nErrorNumber

        cTargetMemo = mTools.FORCEEXT(cTargetTable, "FPT")
        cTargetIndex = mTools.FORCEEXT(cTargetTable, "CDX")
        cTempMemo = mTools.FORCEEXT(cTempFile, "FPT")
        cTempIndex = mTools.FORCEEXT(cTempFile, "CDX")

        if bReturn:
            self.closetable(cTempAlias)
            mTools.DELETEFILE(cTargetTable)
            mTools.DELETEFILE(cTargetMemo)
            mTools.DELETEFILE(cTargetIndex)
            try:
                if os.path.exists(cTempFile):
                    shutil.copy2(cTempFile, cTargetTable)
                if os.path.exists(cTempMemo):
                    shutil.copy2(cTempMemo, cTargetMemo)
                if os.path.exists(cTempIndex):
                    shutil.copy2(cTempIndex, cTargetIndex)
            except:
                bReturn = False
                cErr = "Copy to replacement table failed"
                nErr = -30598
        mTools.DELETEFILE(cTempFile)  # These don't throw an error if the file isn't there.
        mTools.DELETEFILE(cTempMemo)
        mTools.DELETEFILE(cTempIndex)

        self.cErrorMessage = cErr
        self.nErrorNumber = nErr
        return bReturn

    def indexon(self, lcTag, lcExpr, tagFilter="", descending=0, unique=0):
        """
        Method that creates an index tag for the currently selected open DBF table.
        If no DBF is currently selected, or some other problem occurs, it returns False
        otherwise returns True.  This method creates and modifies the VFP type CDX index
        files which may contain multiple physical indexes, each identified by a Tag name.
        It is expected that there will be one and only one CDX index for the table.  In VFP
        and CodeBase, this one-to-one limit is not enforced, but with one CDX file having the
        same base file name as the table, you are assured that the index will be updated with
        its key values whenever the underlying table field values change.

        If no CDX file exists for the table, one is created and the requested tag added to it.

        If the tag exists, it is replaced with your new tag without any warning.

        Parameters:
            - lcTag = a character string of up to 10 alphanumeric values starting with a letter
            - lcExpr = a character string containing an index expression with one or more fields
                defining the index 'key'.
            - tagFilter = optional character string with an expression limiting the number of
                records indexed
            - unique = one of several values determining how duplicate key values will be handled when the table
                is indexed or new records are added:
                0 = No special handling, duplicate key values are recorded in the index normally.
                15 = VFP Compatible "Candidate" index.  Duplicate key values will result in an error condition if
                     encountered in the table.  Both VFP and CodeBase recognize this index type.
                20 = CodeBase compatible setting.  Duplicate key values will result in an error condition if
                     encountered in the table by CodeBase BUT, when VFP accesses or indexes the table it will behave
                     like code 25.
                25 = VFP type UNIQUE index.  In this case when the table is indexed, only the first occurrence of a
                     duplicate is stored in the index.  If the table is traversed or browsed in this index order,
                     it is possible that not all records will be visible or be displayed.  Behavior is the same
                     for VFP and CodeBase.

        Returns:
            True on success, False on any error (see properties cErrorMessage and nErrorNumber for
                more information)

        NOTE: As of Oct. 26, 2013, see method setstrictaliasmode(), which may be required to support this method.

        IMPORTANT NOTE!!! Do NOT attempt to use this method to add an index to a table contained in a VFP Data Dictionary
        Container.  Such tables can ONLY have indexes changed using VFP itself.  However, those indexes CAN be
        reindexed successfully by the reindex() method.
        """
        lbReturn = True
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        if (unique != 0) and (unique != 20) and (unique != 25) and (unique != -340) and (unique != 15) and (unique != -360):
            unique = 0  # Make sure it's a valid value.
        lbReturn = self.cbt.indexon(lcTag, lcExpr, tagFilter, descending, unique)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def getErrorMessage(self):
        cMsg = self.cbt.geterrormessage()
        return cMsg

    def select(self, lcAlias):
        """
        Method that sets the currently selected table to the one with the alias passed as the
        parameter. If the string value you pass is not associated with an open table, a False is
        returned. If the string is associated with an open table, that table becomes the currently
        selected table, and a True is returned. Note that alias names are NOT case sensitive in VFP
        or in this module.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lbReturn = self.cbt.select(lcAlias)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage() + " SELECT FAILED"
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def dbf(self, alias=""):
        """
        Returns the fully qualified path name of the currently selected table or the table
            referenced by the optional parameter alias.  Returns the empty string if there is no
            currently open table or if the alias doesn't exist.
        """
        lcReturn = ""
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lcReturn = self.cbt.dbf(alias)
        if lcReturn == "":
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lcReturn

    def isreadonly(self, alias=""):
        """
        Returns True if the specified open data table was opened in read only mode, otherwise false.
        :param alias: Specify the target table alias or leave empty for the currently selected table.
        :return: True if opened readonly.  If in read-write or exclusive mode, returns False.  Returns None
        on error - typically table not opened.
        """
        lbReturn = False
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.isreadonly(alias)
        if lbReturn is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def used(self, alias):
        """
        Pass a text string with an alias name.  If that table is open, the function will return True,
        otherwise false.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        if alias != "":
            lbReturn = self.cbt.used(alias)
            if not lbReturn:
                self.cErrorMessage = self.cbt.geterrormessage()
                self.nErrorNumber = self.cbt.geterrornumber()
        else:
            self.cErrorMessage = "Alias parameter was blank"
            self.nErrorNumber = -59595
            lbReturn = False
        return lbReturn

    def alias(self):
        """
        Method that returns the alias name of the currently selected table. If there is no
        currently selected table, returns an empty string "". Call this immediately after a table
        use() or createtable() if you need to capture the alias, especially if you may have opened
        or created a table with a temporary or externally defined name, and don't want to make
        assumptions about what default name might have been assigned to the table and have not set
        the alias name yourself.
        """
        lcStr = self.cbt.alias()
        return lcStr

    def goto(self, gomode, gonum=1):
        """
        Primary method for navigating through the currently selected table. Functionality is
        similar to the GOTO command in Visual FoxPro and xBase languages. Unlike SQL-based
        databases, VFP tables can be accessed navigationally by moving around amongst the records in
        the table using variations of this method. Returns True when navigation is successful,
        otherwise False. If False, inspect the cErrorMessage and nErrorNumber properties for more
        information. The effect of this method is to move the current record pointer. The current
        record will be the source of field information when the scatter() method is executed and the
        target of updates when the gathermemvar() method is executed. If an order based on an index
        Tag has been set by setorderto(), all gomode values except RECORD honor the index order.

        Parameters:
            - gomode = character string containing one of the following values:
                RECORD - In this case you must supply the second parameter gonum with an integer
                        value specifying the record number to move to. If the record number
                        specified doesn't exist, the method returns False
                SKIP - Moves the current record point forward or back a specified number of records
                        (in index order if one has been set). If gonum is not supplied, then moves
                        forward by one record.  If gonum > 1, moves forward by gonum records if
                        possible. If EOF() is reached before gonum records, returns False.  If gonum
                        < 0, moves backwards by gonum records if possible. If BOF() is reached
                        before gonum records, returns False.
                NEXT - Identical to SKIP with gonum = 1.
                PREV - Identical to SKIP with gonum = -1.
                TOP  - Moves to the top record by the current ordering (or record 1 if there is no
                        index Tag ordering in effect).  If there are no records in the table, returns
                        False.
                BOTTOM - Move to the the last record by the current ordering (or the highest record
                        number if there is no index Tag ordering in effect).  If there are no records
                        in the table, returns False.

        Returns:
            True on Success, False on Failure for any reason. Inspect the cErrorMessage or nErrorNumber
            for reasons.

        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.goto(gomode, gonum)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def calcfor(self, stat="SUM", fldexpr="", forexpr=""):
        """
        Calculates a statistic from a numeric expression in fldexpr based on the records in the currently selected table
        that satisfy the logical expression contained in forexpr.  Returns a float value result of the calculation.
        Returns None on any kind of error (for example an invalid expression string).  See locate() for examples of
        valid contents for the forexpr parameter.  This obeys the setting of DELETED.

        Values allowed in stat are:
        "SUM" - totals the values in the fldexpr
        "AVG" - produces a simple average
        "MAX" - what it says
        "MIN" - what it says.

        Other considerations:
        - Must NEVER be used in the midst of a locate/continue/locateclear sequence. If it is,
          it will return None error.
        - Does not alter the current record pointer position, so can be called on the current table in the middle
          of a scan().
        - Since this executes entirely in the 'C' engine, it will be much faster than scanning through a table
          with Python code calling lower level CBTools functions.
        - If the fldexpr contains only the name of an Integer type field, the result will always be 0.0 UNLESS
          you append a "+0" to the expression (due to an oddity in the way CodeBase handles Integer fields).  So,
          if S_NUM is an Integer, and you need the sum of these values, pass "S_NUM+0" as the value of fldexpr.  This
          does not apply to Number type fields.  It has not been tested with VFP Double, Float, or Currency type fields.
        """
        lnStat = self.xCalcTypes.get(stat.upper(), 0)
        if not lnStat:
            self.cErrorMessage = "Bad Statistic Type"
            self.nErrorNumber = -34934
            return None

        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lnResult = self.cbt.calcstats(lnStat, fldexpr, forexpr)
        if lnResult is None:
            # Error condition
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnResult

    def deleted(self):
        """
        Determines if the current record of the currently selected table has been marked for deletion.
        Returns True if the record has been marked for deletion.  Returns False if it hasn't.  Returns
        None if there has been an error.

        Inspect cErrorMessage and nErrorNumber if None is returned for more information.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        bResult = self.cbt.deleted()
        if bResult is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return bResult

    def recno(self):
        """
        Returns the record number of the current record of the currently
        selected table. Returns the record number on success, -1 on any kind of
        failure.

        When a table is opened the record pointer is automatically moved to
        record 1, if it exists, and there is no selected index tag. If you need
        to work on a record, then move to some other records, and later go back
        to the original record, it is much faster to save the recno() value and
        return to that record using the goto("RECORD", nTheRecord) method.

        It is possible for the recno() value returned to be one greater than
        the number of records as returned by the reccount() method. In that
        case, the record pointer is at an EndOfFile condition, and eof() will
        return true.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lnReturn = self.cbt.recno()
        if lnReturn < 0:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnReturn

    def countfor(self, expr):
        """
        Counts the number of records in the currently selected table that satisfy the logical expression contained
        in expr.  Returns a value from 0 to the number of records in the table on success.  Returns -1 on any kind
        of error (for example an invalid expression string).  See locate() for examples of valid contents for the
        expr parameter.  This obeys the setting of DELETED.
        Other considerations:

        - Must NEVER be used in the midst of a locate/continue/locateclear sequence. If it is,
          it will return -1 error.
        - Does not alter the current record pointer position, so can be called on the current table in the middle
          of a scan().
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        if not expr.strip():
            raise ValueError("Expression may NOT be empty or blank")
        lnReturn = self.cbt.count(expr)
        if lnReturn == -1:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnReturn

    def locate(self, expr):
        """
        Comparable to the VFP LOCATE command.  This operates ONLY on the currently selected table and
        performs a search of the table for the first record which satisfies the logical expression provided
        in the parameter expr.  Parm expr must contain a valid VFP/DBASE type expression and may include
        .AND., .OR., and other logical expression elements.  Field names in the currently selected table
        may be included as well as VFP/DBASE functions supported by CodeBase.  If any of the logical expressions
        contained in expr match index expressions in index tags on the current table, CodeBase will optimize
        the search to take advantage of the tags.

        Returns True on success (found a record), False on failure for any reason.  Inspect nErrorNumber if there
        is a possibility of an error condition.  It will have the value 0, if it is a simple no-find situation.

        Examples of valid expr values:

        UPPER(LAST_NAME) = "JONES" .AND. UPPER(FIRST_NAME) = "PETER" .AND. SSN = "123456789"
        STR(RecordKey) + TTOC(dateField, 1) + COMPANY = "   123456720120629053329UNITED COPPER"

        In the first example if there is an index on the expression UPPER(LAST_NAME) the search will be partially
        optimizable and will be substantially faster than a brute force search through the table.  In the second
        example, for it to be optimizable, an index tag would need to exist with the expression exactly the
        same as the expression to the left of the = sign.

        Note that logical expression operators like .AND. and .OR. are required to have the periods.  VFP is
        tolerant of having those operators given simply as AND or OR, but the CodeBase DLL is not.

        See also the locateclear() and locatecontinue() methods which may be used on conjunction with locate()

        While you are debugging, or where the expr value may be created by a user, the error number should always
        be checked when False is returned, as False may be the result of an illegal expression.

        WARNING!  The locate(), locatecontinue(), locateclear() sequence may NOT be nested.  Only ONE locate
        process may be in effect at any one time!!!

        NOTE: calling locate() clears the previously called locate(), so locate() may be called repeatedly if just one
        record is needed each time.  At the end of calling a bunch of locates, you should then clear the locate sub-
        system with a locateclear() since some other lookup functions may use the locate() logic internally, and will
        fail if there is a pending un-cleared locate active.  To be safe, call locateclear() after any locate() if
        you won't be needing its results or locatecontinue() any more.

        ALSO NOTE: Locate will not work consistently with the current index set to a descending index.  Use some
                   form of scan with an expression instead, which can traverse an index in reverse.
        """
        lbReturn = True
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        if not expr.strip():
            lbReturn = False
            self.cErrorMessage = "No expression provided for LOCATE"
            self.nErrorNumber = -24593
        else:
            lbReturn = self.cbt.locate(expr)
            if lbReturn is None:
                self.cErrorMessage = self.cbt.geterrormessage()
                self.nErrorNumber = self.cbt.geterrornumber()
                lbReturn = False  # This is what users are expecting from the old version.
        return lbReturn

    def locatecontinue(self):
        """
        A successful VFP LOCATE command may be followed by any number of CONTINUE commands to continue the
        search for records which meet the criteria specified by the initial LOCATE.  locatecontinue() provides
        that functionality after the locate() method.  If it returns True, another record has been found
        that satisfies the expr value as provided to the initial locate() method.  Otherwise, it returns False.
        On False, nErrorNumber will be non-zero if an error has occurred.  nErrorNumber will be 0 if there
        simply are no more matching records.

        locatecontinue will traverse the table in the currently set order or in record number order if there
        is no set order.  locatecontinue() called without a corresponding initial locate() will produce
        either an error condition or an undefined result.

        When locatecontinue() is called the currently selected table MUST be the same as was selected at the time
        the preceding locate() was called or the results will be undefined.

        It is STRONGLY recommended that after the last locatecontinue() method call in the series, the
        locateclear() method be called to release memory and allow normal processing, seek(), etc., of the
        currently selected table.

        See also locate(), locateclear()
        """

        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lbReturn = self.cbt.locatecontinue()
        if (not lbReturn) and (self.cbt.geterrornumber() != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()

        return lbReturn

    def locateclear(self):
        """
        There is no VFP correlate to this method.  Due to memory management in CodeBase, when a locate/continue
        sequence is complete, it is best to call locateclear() to release memory and make sure that the table involved
        in the locate sequence (currently selected table) be made available for normal seeking, skipping, etc.

        Returns True on success, False on failure, but the return value may normally be discarded.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lbReturn = self.cbt.locateclear()
        if (not lbReturn) and (self.cbt.geterrornumber > 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def deletetag(self, tagname=""):
        """
        Specify a tag name in the CDX index of the currently selected table to be removed from the CDX file.  If the
        last tag is removed from a CDX index file, the CDX index file is deleted and the DBF file is marked as having
        no index files.  DO NOT use this method to remove an index tag from a VFP table contained in a VFP Database
        Container.  Doing so will result in data corruption.  This can only be applied to so-called VFP Free Tables.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lbReturn = self.cbt.deletetag(tagname)
        if (not lbReturn) and (self.cbt.geterrornumber() != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def setorderto(self, tagname):
        """
        If the currently selected table has one or more index tags available, this method will set
        the current ordering (which is recognized by the various modes of the goto() method) to the
        tag specified by tagname. Returns True on success, False on failure for any reason. Inspect
        the cErrorMessage or nErrorNumber for reasons.  Pass "" as tagname to set to native recno order.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        bResult = self.cbt.setorderto(tagname)
        if (not bResult) and (self.cbt.geterrornumber != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return bResult

    def ataginfo(self):
        """
        Retrieves index tag information for the currently selected table.  If there are no index tags, returns
        an empty list.  Otherwise returns a list consisting of one VFPINDEXTAG object for each list element.

        The VFPINDEXTAG objects each have five properties defining the characteristics of the index tag:
            - cTagName - The name of the sort order tag name.  Use this string value to set the current order
                         to this index using the setorderto() method.
            - cTagExpr - The expression defining the tag.  At minimum this will be the name of one field or a function
                         like DELETED() that relates to the contents of each record.  It can be a much more complex
                         expression like "UPPER(first_name)+UPPER(last_name)+STR(salary)".
            - cTagFilt - If the index tag is to only have visibility to a subset of records, this filter expression can
                         contain a logical expression containing references to one or more fields.  In that case, if you
                         set order to the tag and execute skip() or copytoarray(), only records for which this expression
                         evaluates to True will be recognized.
            - nDirection - 0 if ordering is standard ascending order.  10 if ordering is descending with largest values
                           at the top.
            - nUnique - value of 0 if duplicate index key values are allowed.  Value of 20 if an attempt to add a record
                        that has a key identical to an existing key value will trigger an error.

        If there is no currently selected table or some other error occurs, returns None.

        Sets the value the property tally to the number of tags.  0 is a valid value (no index tags defined), and -1
        indicates an error.
        """
        lxListReturn = list()
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        self.tally = 0

        lxInfoList = self.cbt.ataginfo()
        if lxInfoList is not None:
            self.tally = len(lxInfoList)
            for xI in lxInfoList:
                vTag = VFPINDEXTAG()
                vTag.cTagName = xI["cTagName"]
                vTag.cTagExpr = xI["cTagExpr"]
                vTag.cTagFilt = xI["cTagFilt"]
                vTag.nDirection = xI["nDirection"]
                vTag.nUnique = xI["nUnique"]
                lxListReturn.append(vTag)
                vTag = None
            del lxInfoList
        else:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
            lxListReturn = None
        return lxListReturn

    def order(self):
        """
        Retrieves the index tag previously set by setorderto() from the currently selected table.  Returns a
        string containing the tag name of the index tag set or an empty string if one isn't set or there is no
        currently active table.  If an error occurred (likely no table selected), the cErrorMessage and
        nErrorNumber properties will be set.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        cReturn = self.cbt.order()
        if (not cReturn) and (self.cbt.geterrornumber != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return cReturn

    def deletetagall(self):
        """
        Deletes all index tags for the currently selected table. The CDX index file is deleted and the
        DBF file is marked as having no index files.  Returns True on success, False on failure.  Returns True
        if there are no tags set up for the table, even though there is nothing to do.  DO NOT use this method to
        remove index tags from a VFP table contained in a VFP Database Container.  Doing so will result in data
        corruption.  This can only be applied to so-called VFP Free Tables.
        """
        bReturn = True
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        self.tally = 0

        lxTagList = self.cbt.ataginfo()  # Returns a list() of dict()s
        if lxTagList is not None:
            if len(lxTagList) > 0:
                for xT in lxTagList:
                    bResult = self.cbt.deletetag(xT["cTagName"])
                    if not bResult:
                        bReturn = False
                        break
                    else:
                        self.tally += 1
        else:
            bReturn = False
        if not bReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return bReturn

    def dtt2expression(self, dateTimeValue):
        """
        Pass a date/time value and this will return a text string appropriate for a constant datetime value
        in an expression for a location or calculate function.  In VFP a constant date/time value is expressed
        with elaborate notation like {^2013-02-15 12:35:00}, but Codebase doesn't recognize this.  Instead,
        it wants a function like "DATETIME(2013, 2, 15, 12, 35, 0)".  This function produces that latter string value.
        """
        cReturn = "DATETIME(%d, %d, %d, %d, %d, %d)" % (dateTimeValue.year, dateTimeValue.month, dateTimeValue.day,
                                                        dateTimeValue.hour, dateTimeValue.minute, dateTimeValue.second)
        return cReturn

    def dtd2expression(self, dateTimeValue):
        """
        Similar to dtt2expression() except it uses the STOD() function to match date type fields. NOTE that the
        Python values returned from VFP DATE and DATETIME fields are always Python datetime objects, BUT in VFP
        they are different entities and can't be compared directly.  So if a field is a type DATE, you must create
        the matching expression with the STOD() function using this method.  Also note that STOD() is NOT a VFP
        function but a dBase function recognized by CodeBase in expressions.
        """
        cReturn = 'STOD("%4d%02d%02d")' % (dateTimeValue.year, dateTimeValue.month, dateTimeValue.day)
        return cReturn

    def dt2seek(self, dateValue):
        """
        Pass a date value, and this returns the proper string format required for the seek() method below.
        """
        lcRet = ""  # default if something goes wrong.
        try:
            lcRet = dateValue.strftime("%Y%m%d")
        except:
            pass # nothing to be done
        return lcRet

    def dtt2seek(self, dateTimeValue):
        """
        Pass a datetime value, and this returns the proper string format required for the seek() method.
        """
        lcRet = ""
        try:
            lcRet = dateTimeValue.strftime("%Y%m%d%H:%M:%S:000")
        except:
            pass
        return lcRet

    def getdatasession(self):
        """
        When multiple instances of this class exist, each will be given its own data session.
        Other instances may set the "current" session to a different value.  If you find your
        table is not reported 'open', when you think it should be try storing the session number
        of your instance and restore to it with setdatasession().  If you change a current data session
        be sure to put it back when exiting the function.
        :return: Integer session number
        """
        return self.cbt.getcurrentsession()

    def setdatasession(self, nSession):
        """
        set the data session back to what it was earlier.  the session number MUST be one that came
        from getdatasession() or the results will be undefined.
        :param nSession:
        :return: The session number switched to or -1 on error
        """
        return self.cbt.switchdatasession(nSession)

    def curvalstr(self, lcFieldName):
        """
        Retrieves the contents of one field specified by the parm lcFieldName.
        If you want a value from a table other than the currently selected one,
        use the alias.fieldname notation in the name of the string you pass in
        the lcFieldName parameter.

        Returns a text string representation of the field value. No conversions
        are performed. An empty string is a valid return value, but if you
        expect something else, check the error messages. An empty string is also
        returned if the lcFieldName is not found in the current table or if your
        specified alias doesn't exist.

        Designed for speed with minimal checking. Works OK with char, memo,
        number, currency, and date type fields. Does not produce useful results from
        integer, double, datetime, or float type fields.  Note that the contents of char fields
        are padded with blanks to the full size of the fields.
            - Returns None on error, and sets errormessage and errornumber
            - Returns None when field allows null and has a null value, but does NOT set errors.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lcReturn = self.cbt.scatterfieldex(lcFieldName, False)
        if self.cbt.geterrornumber():
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lcReturn

    def curvalstrEX(self, lcFieldName, lbStrip=False):
        """
        Retrieves the contents of one field specified by the parm lcFieldName.
        If you want a value from a table other than the currently selected one,
        use the alias.fieldname notation in the name of the string you pass in
        the lcFieldName parameter.

        Returns a text string representation of the field value. No conversions
        are performed. An empty string is a valid return value, but if you
        expect something else, check the error messages. An empty string is also
        returned if the lcFieldName is not found in the current table or if your
        specified alias doesn't exist.

        Designed for speed with minimal checking. Works OK with char, memo,
        number, currency, and date type fields. Does not produce useful results from
        integer, double, or float type fields.  May truncate the contents of Unicode
        strings contained in char(binary) and memo(binary) fields.

        Enhancements over curvalstr:
            - Allows specification of stripping trailing blanks.

        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lcReturn = self.cbt.scatterfieldex(lcFieldName, lbStrip)
        if lcReturn is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber  = self.cbt.geterrornumber()
        return lcReturn

    def seek(self, targmatch, alias="", indextag=""):
        """
        This method is closely analogous to the VFP SEEK() function and uses an
        index Tag to find a record matching a specified lookup string (the
        targmatch value). If a matching record is found, returns True, otherwise
        False. False can also be returned for a number of other failure causes,
        and the nErrorNumber should be checked for a non-zero value to identify
        error situations. Unlike setorderto(), gathermemvar(), alias() and many
        other methods, this one doesn't have to work on the currently selected
        data table. This allows you to do a seek() into another table in the
        middle of traversing the currently selected table. You can retrieve
        field values from the found record in that other table using curval()
        which allows you to specify the table from which to return the field
        value.

        Parameters:
            - targmatch = A text string containing the value to find by the specified index tag.
                For example if there is an index Tag with an expression of "UPPER(xCode
                + yCode)" where xCode and yCode are each character fields of length
                10, and you wanted to find a record with an xCode value of "abc" and a
                yCode value = "yior" you would construct the targmatch as "ABC       YIOR      ".
                Note that in DBF files character fields are automatically
                padded to their full length with space characters, and a concantentation
                of field values will include those spaces. Also note that the targmatch
                values are all upper case because the field values in the example tag
                expression are wrapped in a UPPER() function.

              Some further rules for Tag expressions:
                  - If the field indexed is a date field, the index expression must be a
                        character string of the form "YYYYMMDD".
                  - If the field indexed is any type of number, the index expression must
                        be a character string containing a text representation of the value
                        being searched for. For example, if the field is an integer and
                        you are searching for 29502, the targmatch string should be "29502". For
                        number, double, and currency type fields, be sure to include the decimal
                        point in your text string.
                  - If the field indexed is a datetime value, the index expression
                        required is not exactly what you would expect.
                        The required text string should be in the form "YYYYMMDDHH:MM:SS:TTT"
                        where YYYY is the year, MM the month, DD the calendar day, HH the time
                        hour (24 hour clock time), MM minutes, SS seconds, and TTT thousandths
                        of seconds. BUT NOTE: the colons ARE REQUIRED. Also, since DBF tables
                        can't store datetime values accurate to the one thousandth of a second,
                        the TTT value should always be "000".
                  - If the tag expression is a character string, the targmatch value may
                        be equal to or less than the length of the tag expression. If it is less
                        than the length of the tag expression it will "find" the first record
                        which matches the characters of the tag expression. If the targmatch
                        string length is larger than the actual tag expression length, the
                        result will always be False.

            - alias = Allows you to specify the alias name of any currently open
                        table. If passed as an empty string or omitted the currently selected
                        table is assumed. Otherwise the table specified by the alias name is the
                        subject of the seek. The currently selected table is NOT changed.

            - indextag = Optional specification of the index tag to use for the
                        seek. If not specified, then seeks using the tag specified in the latest
                        setorderto() method for the table. In that case, if no current order has
                        been set for the table, the method returns False.


        Returns:
            True for success (record is found).  False for any other condition.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.seek(targmatch, alias, indextag)
        if lbReturn is None:
            lbReturn = False
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def skip(self, gonum=1):
        """
        Identical to the goto() method with the "SKIP" gomode value, except it runs a stripped down
        function for speed with only very basic error checking and reporting.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.skip(gonum)
        if (not lbReturn) and (self.cbt.geterrornumber() != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def refreshbuffers(self):
        """
        This is a more drastic update of the information on the currently selected table than refreshrecord().
        In this case, the entire cached data for the currently selected table is purged, forcing a re-read from
        the disk of any part of the table required for subsequent processing, including the table header.

        This function incurs substantial disk overhead the next time the table is accessed, so should be used with
        care in loops.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.refreshbuffers()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def refreshrecord(self):
        """
        While VFP properly handles the refresh of its internal buffers when another application has changed a table,
        CodeBase does NOT.  If you are working with a table that is potentially subject to changes by an external
        application, and if you have been keeping this table open to save the time of re-opening prior to each
        use, then before accessing the data, you should select the table and then call this method to be sure
        the most recent version of the current record data is in the CodeBase internal buffers.

        Note that this ONLY operates on the data in the current record.  Other records may remain buffered and
        potentially out of date relative to changes made by other processes.

        Returns True on success, False on failure.  On failure records the error message.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.refreshrecord()
        if (not lbReturn) and (self.cbt.geterrornumber != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def flush(self):
        """
        Call this function immediately after a gatherfromarray() to force the data changes to be written
        immediately to the disk.  This will ensure that other users of the data will see it immediately.  Otherwise
        CodeBase will buffer writes and could delay these writes until the buffer is full.

        This action applies only to the currently selected data table.

        Returns True on success, False on Error, in which case you need to check cErrorMessage and nErrorNumber.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.flush()
        if (not lbReturn) and (self.cbt.geterrornumber != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def recall(self):
        """
        The opposite of delete().  This works on the current record of the currently selected table.  If
        that record is marked for deletion, then removes the mark.  If not marked for deletion, then does
        nothing.

        Returns True on success, False on Error, in which case you need to check cErrorMessage and nErrorNumber.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.recall()
        if (not lbReturn) and (self.cbt.geterrornumber != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def recallall(self):
        """
        Performs a recall() operation on every record in the currently selected table currently marked for
        deletion.  When this is done successfully, there will be no records in the table marked for deletion.

        Returns True on success, False on error, in which case you need to check cErrorMessage and nErrorNumber.

        Sets the value of the tally property to the number of records recalled.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        self.tally = 0
        lbReturn = self.cbt.recallall()
        if (not lbReturn) and (self.cbt.geterrornumber != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        else:
            self.tally = self.cbt.gettally()
        return lbReturn

    def flushall(self):
        """
        Similar to flush() but applies to all open tables.

        Will attempt to lock each record being updated.  Also checks to make sure that there are no duplicate
        values for indexes that require unique keys.  If these tests fail, the flushall() can fail.

        Returns True on success, False on Error, in which case you need to check cErrorMessage and nErrorNumber.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.flushall()
        if (not lbReturn) and (self.cbt.geterrornumber != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def fcount(self):
        """
        Returns the number of fields in the currently selected table or -1 on any kind of error.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lnReturn = self.cbt.fcount()
        if lnReturn < 0:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnReturn

    def reccount(self):
        """
        Returns the physical number of records contained in the currently
        selected table. This will include records that have been marked for
        deletion. Use countrecords() to get the number of undeleted records.

        Returns -1 on error, in which case see cErrorMessage and nErrorNumber.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lnReturn = self.cbt.reccount()
        if lnReturn < 0:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnReturn

    def tagcount(self):
        """
        Returns the number of index tags stored in the CDX index file for the
        currently selected table. Returns 0 if there aren't any.

        Returns -1 on error, in which case see cErrorMessage and nErrorNumber.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lnReturn = self.cbt.tagcount()
        if lnReturn < 0:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnReturn

    def eof(self):
        """
        Interrogates the record pointer status of the currently selected table.
        Returns True if the record pointer is at the EOF position or on error. Otherwise
        returns False. If False, the record pointer is not at the end. Otherwise EOF or
        an error has occurred and you should take further action.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.eof()
        if lbReturn is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
            lbReturn = True  # Perversely.  If there's an error we can't iterate, so we have to say "Yes, EOF is True".
        return lbReturn

    def fieldinfo(self, fieldname=""):
        """
        Gets the field structure info for just one field without the bother of afields()
        :param fieldname: may include the alias
        :return: tuple - (type, length, decimals, allow nulls, is binary) or None on error.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lxReturn = self.cbt.fieldinfo(fieldname)
        if lxReturn is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lxReturn

    def afields(self, lpcAlias=""):
        """
        Retrieves field structure information for the currently selected table or the specified table designated
        by alias name.

        Returns a list of VFPFIELD objects, each containing information on one field, the list is in order as the
        fields physically appear in the table.  See the doc for createtable() for how the output of this method
        can be used to create a duplicate of the table.

        The properties of the VFPFIELD object are detailed in the createtable() documentation.

        If there is no currently selected table, or the alias isn't recognized, or some other error occurs, returns None.
        """

        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lxRawList = self.cbt.afields(lpcAlias)
        if lxRawList is None:
            lxListReturn = None
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        else:
            lxListReturn = list()
            for xF in lxRawList:
                xFld = VFPFIELD()
                cTest = xF["cName"]
                # print(repr(cTest), type(cTest), ord(cTest[5:6]))
                xFld.cName = cTest
                xFld.cType = xF["cType"]
                xFld.nWidth = xF["nWidth"]
                xFld.nDecimals = xF["nDecimals"]
                xFld.bNulls = xF["bNulls"]
                lxListReturn.append(xFld)
        return lxListReturn

    def afielddict(self, lpcAlias=""):
        """
        Similar to afields in that it retrieves field data, but it returns a dict() keyed by the field name (always
        upper case), containing the object pointer for the field info structure.  Makes it easy to get specs
        on individual fields without looping through the list returned by afields()
        :param lpcAlias:
        :return: see above. Returns None on error
        """
        xFlds = self.afields(lpcAlias=lpcAlias)
        if xFlds is not None:
            xRet = dict()
            for xF in xFlds:
                xRet[xF.cName] = xF
            return xRet
        else:
            return None

    def getfieldinfo(self, lpcFieldName=""):
        """
        Returns info on the width, decimals, type, nulls for a field in the currently selected table.

        Return value is None, if the field doesn't exist in the current table or there is no currently selected table.

        Returns tuple of (cType, nLength, nDecimals, bNulls) if field is found.
        Added 07/30/2015. JSH. [Consider in future implementing directly in the cbtools C engine.
        """

        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lxRawList = self.cbt.afields("")
        lxReturn = None
        if lxRawList is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber  = self.cbt.geterrornumber()
        else:
            lpcFieldName = lpcFieldName.upper()
            for xF in lxRawList:
                if xF["cName"] == lpcFieldName:
                    lxReturn = (xF["cType"], xF["nWidth"], xF["nDecimals"], xF["bNulls"])
                    break
        return lxReturn

    def afieldtypes(self):
        """
        Utility function which returns a dictionary containing field type information for the currently selected table.

        The keys to the dictionary are the field names.  This provides an easy way to tailor any data preparation for
        storage in a specific field, based on the target field type.

        Returns the dictionary on success, None on failure.  Much faster than getting all afields() info, but NOTE:
        the ordering of the dict() is unrelated to the sequence of fields in the table.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lxRetDict = self.cbt.afieldtypes()
        if lxRetDict is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lxRetDict

    def scatterblank(self, lcAlias="", lbBinaryAsUnicode=True):
        """
        Returns a dictionary for an empty record in the currently selected table or the table specified by
        alias.  Contains one item keyed by the uppercase of the field name and having a value representing an
        empty field.  For example, if the field is an INTEGER, the value will be 0.  If the field is a LOGICAL,
        the value will be .F. (VFP False).  If the field is CHARACTER, the value will be an empty string.  For
        DATE and DATETIME fields, the value will be Python None.

        Added 03/20/2012. JSH.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lxRec = self.cbt.scatterblank(lcAlias, lbBinaryAsUnicode)
        if lxRec is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lxRec

    def scatter(self, alias="", converttypes=True, stripblanks=False, fieldList=None, coding="XX"):
        """
        The most generally used way to retrieve the contents of a data record
        from the current record of the currently selected table or a table whose
        alias you specify. Returns information on the contents of all fields in
        the table.

        Parameters:
            - alias = specify the alias of a currently open table if you want a
                table other than the currently selected table. Does not change the
                currently selected table unless there is no currently selected table.
            - converttypes = Pass a value of False to get all values as text
                representations of their field values. For example, with convertypes
                True, date values are returned as Python dates. With convertypes False,
                date values are returned as strings of the form "YYYYMMDD".
            - stripblanks = If set to True (default is False), this tells the method
                to strip off all the trailing blanks on all character fields. Note that
                in VFP and all DBF tables, character fields are padded with blanks to
                the full width of the field.
            - fieldList = if not None, it should be a string of upper case field names
                to be included in the dictionary, separated by commas.
            - coding = XX (default) or codes for char and binary fields, see class docstring.

        Returns:
            On success returns a dictionary consisting of field name keys and field
            values. On failure returns the value of None and the cErrorMessage and
            nErrorNumber values will be set for more information. You can modify the
            values for individual fields in this dictionary and then store the
            changes back to the table by passing the dictionary as the first
            parameter to the gatherfromarray() method.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        if fieldList is None:
            fieldList = ""

        xRetDict = self.cbt.scatter(alias, converttypes, stripblanks, fieldList, False, coding)
        if xRetDict is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return xRetDict

    def scattertolist(self, alias="", converttypes=True, stripblanks=False, fieldList=None, coding="XX"):
        """
        Like scatter except that instead of returning a dict(), it returns a list of tuples which
        can be sent via a COM connection to a COM client with the full record.

        Parameters:
            - alias = specify the alias of a currently open table if you want a
                table other than the currently selected table. Does not change the
                currently selected table unless there is no currently selected table.
            - converttypes = Pass a value of False to get all values as text
                representations of their field values. For example, with convertypes
                True, date values are returned as Python dates. With convertypes False,
                date values are returned as strings of the form "YYYYMMDD".
            - stripblanks = If set to True (default is False), this tells the method
                to strip off all the trailing blanks on all character fields. Note that
                in VFP and all DBF tables, character fields are padded with blanks to
                the full width of the field.
            - fieldList = if not None, it should be a string of upper case field names
                to be included in the dictionary, separated by commas.
            - coding = default string text coding or values per class docstring.

        Returns:
            On success returns a list consisting of field names and field
            values as a tuple for each field. On failure returns the value of None and the cErrorMessage and
            nErrorNumber values will be set for more information. You can modify the
            values for individual fields in this dictionary and then store the
            changes back to the table by passing the dictionary as the first
            parameter to the gatherfromarray() method.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        if fieldList is None:
            fieldList = ""
        xReturn = None
        xRetDict = self.cbt.scatter(alias, converttypes, stripblanks, fieldList, False, coding)
        if xRetDict is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        else:
            xReturn = list()
            for cKey, xVal in xRetDict.items():
                xReturn.append((cKey, xVal))
        return xReturn

    def scattertorecord(self, alias="", converttypes=True, stripblanks=False, fieldList=None, coding="XX"):
        """
        Works just like scattertoarray() except that it returns a named tuple with each element having a name
        equal to the uppercased field name, thus: myResult.MYFIELD01
        :return: The named tuple can either be accessed by index number or by the name of the item in the tuple.
        """
        if fieldList:
            xFlds = fieldList.split(",")
            if xFlds != self.xLastFieldList:
                self.xLastFieldList = xFlds
                xWorkRec = collections.namedtuple("CBRecord", xFlds)
                self.xLastRecordTemplate = xWorkRec
            else:
                xWorkRec = self.xLastRecordTemplate
        else:
            if not alias:
                alias = self.alias()
            if alias == self.cLastRecordAlias:
                xWorkRec = self.xLastRecordTemplate
            else:
                self.cLastRecordAlias = alias
                xNames = [xx.cName for xx in self.afields(lpcAlias=self.cLastRecordAlias)]
                xWorkRec = collections.namedtuple("CBRecord", xNames)
                self.xLastRecordTemplate = xWorkRec
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        if fieldList is None:
            fieldList = ""
        bIsList = True
        xRetList = self.cbt.scatter(alias, converttypes, stripblanks, fieldList, bIsList, coding)
        if xRetList is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        else:
            xRetList = xWorkRec._make(xRetList)  # This attaches the names to the elements.
        return xRetList

    def scattertoarray(self, alias="", converttypes=True, stripblanks=False, fieldList=None, coding="XX"):
        """
        Retrieves the contents of a data record from the current record of the
        currently selected table or a table whose alias you specify. Returns
        information on the contents of all fields in the table.

        Parameters:
            - alias = specify the alias of a currently open table if you want a
                    table other than the currently selected table. Does not change the
                    currently selected table unless there is no currently selected table.
            - converttypes = Pass a value of False to get all values as text
                    representations of their field values. For example, with convertypes
                    True, date values are returned as Python dates. With convertypes False,
                    date values are returned as strings of the form "YYYYMMDD".
            - stripblanks = If set to True (default is False), this tells the method
                    to strip off all the trailing blanks on all character fields. Note that
                    in VFP and all DBF tables, character fields are padded with blanks to
                    the full width of the field.
            - fieldList = Comma delimited list of field names to include in the output.
            - coding = XX for system text conversion defaults or a codec or Unicode value from
                    docstring of this class.

        Returns:
            On success returns a simple list of field values in the order in which the fields are found
            in the table OR in the order in which the fields were listed in the fieldList parameter.  Returns
            None on error.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        if fieldList is None:
            fieldList = ""

        bIsList = True
        xRetList = self.cbt.scatter(alias, converttypes, stripblanks, fieldList, bIsList, coding)
        if xRetList is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return xRetList

    def curval(self, lcFieldName, convertType=False, coding="XX"):
        """
        This is a generic CURVAL() (see VFP Function CURVAL()) that returns the
        value of a specified field found either in the currently selected table
        (just pass the field name) or in any open table (prefix the field name
        with the table alias and a '.' in the lcFieldName parameter. If
        convertType is False or omitted, returns a tuple consisting of the
        string representation of the field value followed by the type of the
        value. See scattertoarray() for more details on the formatting of date
        and datetime values and type codes. If convertType is True, then returns
        a single scalar value of the Python type corresponding to the field
        value type.  Char and Memo fields will always return type string with
        coding as set by the coding parm.  If convertType is False all other field types
        will be in plain ASCII stored in string variables, except for binary char and
        memo fields, which will be returned as bytearray (ver 2.x) or bytes (ver 3.x).

        Returns (None, None) or None on error. May also return None on fields
        containing empty dates. If you are expecting a date field and receive
        (None, None) or None, check the cErrorNumber for a non-0 value
        indicating an error.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lxReturn = self.cbt.curval(lcFieldName, convertType, coding)
        # if types are converted, strings are also stripped of trailing blanks.
        if (lxReturn is None) and (self.cbt.geterrornumber() != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        else:
            if not convertType:
                if (lxReturn[0] is None) and (self.cbt.geterrornumber() != 0):
                    self.cErrorMessage = self.cbt.geterrormessage()
                    self.nErrorNumber = self.cbt.geterrornumber()
        return lxReturn

    def curvallogical(self, lcFieldName):
        """
        Retrieves the contents of one field specified by the parm lcFieldName.
        This field MUST be of type LOGICAL. If you want a value from a table
        other than the currently selected one, use the alias.fieldname notation
        in the name of the string you pass in the lcFieldName parameter.

        Returns a Python value of true or false. A value of None is returned on
        error. If you have any doubt about the None value returned, then check
        the error messages. An non-blank error message may indicate an invalid
        field name or a non-existant alias name.

        Designed for speed with minimal checking. Works ONLY with logical type
        fields. Does NOT produce useful results from ANYTHING else.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.scatterfieldlogical(lcFieldName)
        if lbReturn is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def curvaldatetime(self, lcFieldName):
        """
        Retrieves the contents of one field specified by the parm lcFieldName.
        This field MUST be of type DATETIME or of type DATE.

        If you want a value from a table other than the currently selected one,
        use the alias.fieldname notation in the name of the string you pass in
        the lcFieldName parameter.

        Returns a Python datetime.datetime() value. A value of None is returned if the
        field value is empty or on error. If you have any doubt about the None
        value returned, then check the error messages. A non-blank error
        message or an error number other than 0 may indicate an invalid field
        name or a non-existant alias name. If the field is of type DATE, the
        return will still be a python datetime type but the values for hour,
        minute, and second will all be 0.

        Designed for speed with minimal checking. Works ONLY DATETIME and DATE
        type fields. Does NOT produce useful results from ANYTHING else.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        ltReturn = self.cbt.scatterfielddatetime(lcFieldName)
        if (ltReturn is None) and (self.cbt.geterrornumber() != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return ltReturn

    def curvaldate(self, lcFieldName):
        """
        Retrieves the contents of one field specified by the parm lcFieldName.
        This field MUST be of type DATETIME or of type DATE.

        If you want a value from a table other than the currently selected one,
        use the alias.fieldname notation in the name of the string you pass in
        the lcFieldName parameter.

        Returns a Python datetime.date() value. A value of None is returned if the
        field value is empty or on error. If you have any doubt about the None
        value returned, then check the error messages. A non-blank error
        message or an error number other than 0 may indicate an invalid field
        name or a non-existant alias name.

        Designed for speed with minimal checking. Works ONLY for DATETIME and DATE
        type fields. Does NOT produce useful results from ANYTHING else.  For DATETIME
        fields still only returns the date portion of the value.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        ldReturn = self.cbt.scatterfielddate(lcFieldName)
        if (ldReturn is None) and (self.cbt.geterrornumber() != 0):
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return ldReturn

    def curvalfloat(self, lcFieldName):
        """
        Retrieves the contents of one field specified by the parm lcFieldName as a Python float type variable.
        If you want a value from a table other than the currently selected one,
        use the alias.fieldname notation in the name of the string you pass
        in the lcFieldName parameter. This method ONLY works with fields that
        are of type NUMBER, FLOAT, DOUBLE or CURRENCY.  (It MAY work with INTEGER, but
        will return a float, so may not be terribly useful.)

        Returns the number as a float or None on error. Error can be triggered by a bad
        field name, a non-numeric field, or no table open, etc.

        Designed for speed with minimal checking.

        NOTE:!!!!!  In DBF tables, decimal places are defined for NUMBER fields, and their values are
        explicit... in a field of type N(10,2) the value 3.33 is stored exactly as 3.33, not
        as some rounding of a binary representation of that value.  The same is true of CURRENCY fields.
        This method, however, returns the value as a Python float, which IS a binary representation which may NOT
        be exactly the value desired.  For CURRENCY field types, especially, this is generally NOT what is desired.
        To preserve the decimals exactly, use curvalstr() and convert the result into the required
        python type that preserves the decimal exactness you require OR use curvaldecimal().
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lnReturn = self.cbt.scatterfielddouble(lcFieldName)
        if lnReturn is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnReturn

    def curvaldecimal(self, lcFieldName):
        """
        Retrieves the contents of one field specified by the parm lcFieldName as a Python decimal.Decimal()
        type value, which contains an exact representation of any fractional portion.
        If you want a value from a table other than the currently selected one,
        use the alias.fieldname notation in the name of the string you pass
        in the lcFieldName parameter. This method ONLY works with fields that
        are of type NUMBER, FLOAT, or CURRENCY.  Note that in Visual FoxPro, a FLOAT type is exactly
        the same as NUMBER.  Conversely, a DOUBLE is stored in binary form, and may not be able to
        represent fractional data precisely.  An attempt to retrieve a DOUBLE field with this function
        will return an ERROR condition (None).

        Returns the number as a decimal.Decimal() object or None on error. Error can be triggered by a bad
        field name, a non-numeric field, or no table open, etc.

        Designed for speed with minimal checking.

        NOTE:!!!!!  In DBF tables, decimal places are defined, and their values are
        explicit... in a field of type N(10,2) the value 3.33 is stored exactly as 3.33, not
        as some rounding of a binary representation of that value.  This method, however,
        preserves the decimal exactness you may require.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lxReturn = self.cbt.scatterfielddecimal(lcFieldName)
        if lxReturn is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber  = self.cbt.geterrornumber()
        return lxReturn

    def curvallong(self, lcFieldName):
        """
        Retrieves the contents of one field specified by the parm lcFieldName.
        If you want a value from a table other than the currently selected one,
        use the alias.fieldname notation in the name of the string you pass
        in the lcFieldName parameter. This method ONLY works with fields that
        are of type INTEGER, DATE, or NUMBER. In the case of DATE, you'll get
        the Julian day number back. In the case of NUMBER, you'll get the whole
        number part of the value back as a long.

        Returns the number or None on error. Error can be triggered by a bad
        field name or no table open, etc.

        Designed for speed with minimal checking.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lnReturn = self.cbt.scatterfieldlong(lcFieldName)
        if lnReturn is None:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnReturn

    def scatterraw(self, alias="", converttypes=True, stripblanks=False, lcDelimiter="<~!~>", coding="XX"):
        """
        IMPORTANT: This function is here for backwards compatibility with older versions of this
        module.  It is deprecated for future new development.  Use scatter() or other retrieval
        mechanisms instead that are now just as fast or faster.

        Retrieves the contents of a data record from the current record of the
        currently selected table or a table whose alias you specify. Returns
        information on the contents of all fields in the table. Applies NO
        formatting or translation of the result data. Use for speed when you are
        confident you know what is in the table and what order the fields are
        found in.

        Parameters:
            - alias = specify the alias of a currently open table if you want a
                    table other than the currently selected table. Does not change the
                    currently selected table unless there is no currently selected table.

            - delimiter = string of one or more characters (up to 32) which will be
                    used as the delimiter between the values in the output string.

        Returns:
            On success returns a simple character string containing a simple text
            representation of the value of each table field in order. Returns no
            information about field names. Values are delimited in the string by the
            delimiter string, which defaults to "<~!~>", but can be overridden...
            Returns None on error.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lxRetList = self.cbt.scatter(alias, False, False, "", True, "XX")
        if lxRetList is not None:
            lcReturn = lcDelimiter.join(lxRetList)
        else:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
            lcReturn = None
        return lcReturn

    def gatherfromarray(self, lxValues, lxFields=None, lcDelimiter="<~!~>", lbIgnoreMissing=True, coding="X"):
        """
        ------------------------------------------------------------------------------------------------------------
        NOT Recommended with Python 3x!  Note field types not supported!  Limited to text coded in Extended ASCII
        using CodePage 1252. DEPRECATED.

        Use insertdict or gatherdict INSTEAD.

        This function is retained for backwards compatibility and for its modest speed advantage when used within
        its limitations.
        ------------------------------------------------------------------------------------------------------------
        Writes values into the current record of the currently selected table.  If the current record is null or the
        record pointer is at EOF, returns False.  Also returns False if there is no currently selected table.

        VFP's most powerful way to put values into a table record in the "GATHER FROM" command.  This method emulates
        some of its power to be able update a selection of multiple fields all at once without the complication of SQL
        syntax.  The function operates in two modes: Field Specific Mode and All Fields Mode.

        Field Specific Mode:
        In this case you must specify which fields are being updated (by name) and provide a value for each.  This may
        be done in one of three ways:

        Dictionary: Supply the first parameter as a dictionary containing the field values keyed by the field names.
        The field values may be any Python type compatible with the type of the associated field.

        For example, either a 1, True, or "T" are all accepted for the value of a LOGICAL field.  Any other type will
        raise an exception of BadFieldValueType.  As another example, for DATETIME or DATE fields you may provide
        either a datetime.date type value or a datetime.datetime value.  But be aware that integers and numbers will
        be accepted for DATE values. In this case, the number will be truncated to an integer and interpreted as the
        Julian Day value (in VFP the number of days since APPROXIMATELY Jan 1, 4800 BCE).  Specifically, the date
        (Current Era or A.D.) February 1, 2011, has a VFP Julian Day value of 2455594.  DATE and DATETIME values can
        also be supplied by tuples of integer Year, Month, and Day values (for DATE) or integer Year, Month, Day, Hour,
        Minute, Second values (for DATETIME).

        For character fields, any value type will be accepted and converted to a character representation.  See
        comments under the Strings input type for more on type conversions.

        Strings are allowed for all field types, but for correct evaluation they must follow the rules specified in the
        Strings section below.

        Lists: This requires a list of field names as the second parameter and a list of field values as the first
        parameter.  The index of the field names and values must correspond.  As with the dictionary input, the field
        value types must be compatible with the type of the target field.

        Strings: Supply the second parameter as a string consisting of an ordered list of field names delimited by a
        non-alpha string.  By default the delimiter is assumed to be <~!~>, but you may specify another string value
        in the lcDelimiter parameter.  The first parameter must be another string containing a concantenated and
        delimited string representation of all values referenced in the field list.  This value string must be ordered
        identically to the field name string.  No type conversions or checking is done in this mode.  For this reason,
        processing is fast, but you are responsible for expressing each field value in the correct format.
        Formats are as follows for each type of DBF field:

            CHARACTER - Any string value.  Strings longer than the length of the field will be truncated, shorter
                        strings will be right padded with blanks.  Extended ASCII wirh CodePage 1252 ONLY unless
                        data passed as a dict().
            NUMBER - String representation of a floating point value in decimal format, e.g. "394039.395".  For negative
                     values, prepend the minus sign, e.g. "-25904.99".  The actual number of decimals stored is
                     determined by the field format.  Excess digits to the right of the decimal point will be rounded
                     off.  Excess digits to the left of the decimal point will trigger a NumericOverflow error.
                     Number type fields store the exact value, e.g. 2.1 is stored as 2.1 if at least one decimal of
                     precision is specified, not 2.09999999999999999.  Relative to field size, the size of NUMBER
                     fields is defined by a total field width and the number of decimal places.  The total field width
                     must be sufficient to handle the largest value intended for the field INCLUDING the decimal point
                     AND any prepended minus sign.  For example a field defined as N(7, 3) can accommodate a value of
                     999.349 or -99.349 but NOT -999.349, as the latter value requires a field of size N(8, 3).
            FLOAT - Similar to NUMBER, but the precision of the input number will be preserved in the value stored.
                    However, float values, like all IEEE float and double numeric values are stored as approximations
                    of the fractional component, e.g. 2.1 double may be stored as 2.0999999999999999.
            INTEGER - String representation of any 4-byte integer value, e.g. "195023985" or "-234095".  Values with
                      embedded decimals will be stored with decimal fraction truncated.
            DATE - String representation of a date as 8 digits in the format YYYYMMDD.
            DATETIME - String representation of a datetime value (thousandths of a second are not supported).  The
                       string must consist of 14 digits in the format YYYYMMDDHHMMSS.
            LOGICAL - The single character "T" for True and "F" for False.
            MEMO - Any string value.  The maximum length of the string is 2 Gigabytes, but in practice it is much less
                   since the maximum size of the FPT file which contains all the memo values is itself 2 Gigabytes.  Non
                   ASCII values may be stored in Visual FoxPro memo fields, but some other versions of DBF tables don't
                   support this.
            DOUBLE - Like Number.
            CURRENCY - Like Number in that it stores an exact representation of the value but with a fixed 4 decimal
                       places.
            GENERAL - Not supported except when data are supplied in a dictionary.
            CHAR BINARY - Not supported except when data are supplied in a dictionary.
            MEMO BINARY - Not supported except when data are supplied in a dictionary.

            NOTE: Empty DATE and DATETIME values ---
            Visual FoxPro and the DBF format allow the concept of an empty date and empty datetime value.  This is
            NOT the same as a NULL value in these fields.  A NULL is an undetermined value.  An empty date is a value
            of type date that doesn't contain an actual date.  Ditto for datetime.  In VFP an expression like
            somedate={ / / } (that's VFP date notation) is completely legal.  Date and datetime fields in blank
            records start out with the value of empty date and empty datetime respectively.  You can think of the empty
            date and datetime values as analogs of "" and 0 in the string and number worlds.

            Since Python doesn't support empty date or datetime values, this module handles this using the value None.
            To supply an empty date or datetime value in this method, supply the value of None.  If you are supplying
            a string representation of a date or datetime, supply a string of 8 blanks for the empty date value or
            "        000000" (8 blanks followed by 6 zeros) as the empty datetime value.  In methods like scatter()
            and copytoarray() empty date and datetime values are converted to the Python value None.  This will be
            indistinguishable from the value None which can be returned when the contents of the field is NULL.

        All Fields Mode:
        If you are confident that you know how many fields are found in the table and want to quickly update all of them
        at once very simply, you pass as the first parameter an ordered set of values corresponding to all the
        fields in the table.  This single parameter may either be a list[] containing a value for every field in order
        or a delimited string of field values as described above for the Strings type under the Field Specific Mode.  If
        the number of fields you supply is not exactly equal to the number of fields, a MismatchedFields exception is
        raised.  The list form allows non-string values to be passed as above.  Value type checking will be performed
        on non-string types.  The string form does not do value or format checking and passes the delimited string
        directly to the CodeBase engine for storage.

        The lbIgnoreMissing parameter applies if the source (lxValues) is a dictionary.  In that case, set
        lbIgnoreMissing to True if you want to store any data where there IS a match on the field name vs the key into
        the dictionary, and where there isn't a match, the problem is ignored.  If this value is left to False, the
        default, such a condition will trigger a KeyError exception, which will have to be trapped externally.

        Returns True on success, False on any kind of failure, in which case cErrorMessage and nErrorNumber will be
        set with explanatory values.

        Note that the dictionary returned by scatter() may be used as the first parameter of this method after you make
        any required changes to the field values.
        July 21, 2014 - Changed lbIgnoreMissing default to True. JSH.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = True
        lnResult = 0
        lbIsAllFields = False
        lbIsDict = False
        lbIsString = False
        lbIsList = False
        if lxFields is None and not isinstance(lxValues, dict):
            lbIsAllFields = True
        lcFldStr = ""
        lcValStr = ""
        if lbIsAllFields:  # All Fields Mode
            lbIsString = mTools.isstr(lxValues)
            if not lbIsString:
                lbIsList = isinstance(lxValues, list)
            if not lbIsString and not lbIsList:
                raise ValueError("First parameter must be a list or string in All Fields Mode")

            if lbIsString:
                lcValStr = lxValues
            else:
                lxFlds = self.afields()
                for ix, lf in enumerate(lxFlds):  # checking and converting the types
                    lxValues[ix] = self._py2cbtype(lxValues[ix], lf.cType)
                    if lxValues[ix] is None:
                        self.cErrorMessage = "Bad Field Type: " + lf.cName
                        self.nErrorNumber = -8593
                        lbReturn = False
                        break
                if lbReturn:
                    lcValStr = lcDelimiter.join(lxValues)
        else:  # Just the specified fields get updated: Field Specific Mode
            lbIsDict = isinstance(lxValues, dict)
            if not lbIsDict:
                if isinstance(lxValues, list) and isinstance(lxFields, list):
                    lbIsList = True
                elif mTools.isstr(lxValues) and mTools.isstr(lxFields):
                    lbIsString = True
            if not lbIsDict and not lbIsList and not lbIsString:
                raise ValueError("Need a dict, string, or list as the value parameter")

            if lbIsString:
                lcValStr = lxValues
                lcFldStr = lxFields  # We take what they gave us and hope they got it right.

            elif lbIsList:
                lcFldStr = lcDelimiter.join(lxFields)
                lxFldTypes = self.afieldtypes()
                for ix, lf in enumerate(lxFields):
                    lxValues[ix] = self._py2cbtype(lxValues[ix], lxFldTypes[lxFields[ix].upper()])
                    if lxValues[ix] is None:
                        self.cErrorMessage = "Bad Field Type: " + lxFields[ix]
                        self.nErrorNumber = -8593
                        lbReturn = False
                        break
                if lbReturn:
                    lcValStr = lcDelimiter.join(lxValues)

            else:  # Has to be lbIsDict
                lxFldTypes = self.afieldtypes()
                lxList = list()
                lxNames = list()
                for lx in lxValues:
                    lcName = lx.upper()
                    if lbIgnoreMissing:
                        try:
                            lcStr = self._py2cbtype(lxValues[lx], lxFldTypes[lcName])
                        except KeyError:
                            lcStr = None
                    else:
                        lcStr = self._py2cbtype(lxValues[lx], lxFldTypes[lcName])
                    if lcStr is None and not lbIgnoreMissing:
                        self.cErrorMessage = "Bad Field Type: " + lx
                        self.nErrorNumber = -8593
                        lbReturn = False
                        break
                    else:
                        if lcStr is None:
                            continue
                    lxList.append(lcStr)
                    lxNames.append(lx)
                if lbReturn:
                    lcValStr = lcDelimiter.join(lxList)
                    lcFldStr = lcDelimiter.join(lxNames)

        if lbReturn:
            lnResult = self.cbt.gathermemvar(lcFldStr, lcValStr, lcDelimiter)
            if lnResult == -1:
                lbReturn = False
                self.cErrorMessage = self.cbt.geterrormessage()
                self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

#  These next three are for internal use only.
    def _py2cbint(self, lnInt, lcTargetType):
        lnReturn = int(lnInt)
        return lnReturn

    def _py2cbnum(self, lnNum, lcTargetType):
        lnReturn = float(lnNum)
        return lnReturn

    def _py2cblog(self, lbBool, lcTargetType):
        lnReturn = (1 if lbBool else 0)
        return lnReturn

    def _removenonprint(self, xStr):
        if _ver3x:
            if isinstance(xStr, bytes) or isinstance(xStr, bytearray):
                xRet = xStr.translate(None, bytes(self.cNonPrintables, "utf-8"))
            if isinstance(xStr, str):
                xRet = xStr.translate(self.xNonPrintables)
        else:
            xRet = str(xStr).translate(None, self.cNonPrintables)
        return xRet

    def _py2cbtype(self, lxSourceValue, lcTargetType, lbForFilter=False):
        """
        Utility function that converts Python variable (object) types to the string format needed for
        the CodeBaseWrapper function cbwGATHERMEMVAR() which optionally takes a string of character representations
        of the values for all fields.  [NOT required when you are using a dict() to supply data to gathermemvar.)

        lxSourceValue may be any legal Python value including string, int, float, Decimal, date, datetime, or bool.
        Other types will be coerced to a string value one way or another.

        Returns the string result of the conversion or "" when it throws up its hands and can't do anything, which is
        mainly on bad datetime values.

        ----------------------------
        NOTE!!!
        ----------------------------
        Unicode types in Python 2.x are not supported. Bytes and bytearray types likewise are NOT supported either
        in Python 2.x or 3.x.  Strings in Python 3.x which contain non-ASCII Unicode values likewise are NOT SUPPORTED.

        When working with these more modern types of data use gatherdict() and insertdict(), which perform
        sophisticated type and coding conversion consistent with the Python object type and the target field type.
        """
        if mTools.isstr(lxSourceValue) and (lcTargetType != "D"):
            if not _ver3x:
                if isinstance(lxSourceValue, str):
                    lxSourceValue = lxSourceValue.decode('cp1252', 'replace')
                    # The VFP type tables we have created assume
                    # Windows Codepage 1252, which is an 8-bit Western European language code page.  Unicode strings
                    # are not supported without special coding.
                    lxSourceValue = lxSourceValue.encode("cp1252", "replace")
                elif isinstance(lxSourceValue, unicode):
                    lxSourceValue = lxSourceValue.encode("cp1252", "replace")
            if _ver3x:
                lxSourceValue = lxSourceValue.encode("cp1252", "replace")
                lxSourceValue = lxSourceValue.decode("cp1252", "replace")
                # back to a string with the proper encodings
            if (lcTargetType == 'C') or lbForFilter:
                lxSourceValue = self._removenonprint(lxSourceValue)
                # if a character field, get rid of non-printing characters.
            if lbForFilter:
                return '"' + lxSourceValue + '"'
            else:
                return lxSourceValue

        lxReturn = None
        if (lcTargetType == "C") or (lcTargetType == "M"):
            if isinstance(lxSourceValue, date):
                lxReturn = lxSourceValue.strftime("%Y-%m-%d")
            elif isinstance(lxSourceValue, datetime):
                lxReturn = lxSourceValue.strftime("%c")
            else:
                lxReturn = repr(lxSourceValue)  # and hope for the best.
            if lbForFilter:
                lxReturn = '"' + lxReturn + '"'
            if (lcTargetType == 'C') or lbForFilter:
                # Clear out any non-printing characters, they don't belong here...
                lxReturn = self._removenonprint(lxReturn)

        elif lcTargetType == "T":
            lcTypestr = repr(type(lxSourceValue))
            if isinstance(lxSourceValue, datetime):
                lxReturn = (lxSourceValue.strftime("%Y%m%d%H%M%S") if not lbForFilter else lxSourceValue.strftime("DATETIME(%Y,%m,%d,%H,%M,%S)"))
            elif isinstance(lxSourceValue, date):
                lxReturn = (lxSourceValue.strftime("%Y%m%d%H%M%S") if not lbForFilter else lxSourceValue.strftime("DATETIME(%Y,%m,%d,%H,%M,%S)"))
            elif isinstance(lxSourceValue, tuple) and (len(lxSourceValue) >= 6) and isinstance(lxSourceValue[0], int):
                ltDateTime = datetime(lxSourceValue[0], lxSourceValue[1], lxSourceValue[2], lxSourceValue[3], lxSourceValue[4], lxSourceValue[5])
                lxReturn = (ltDateTime.strftime("%Y%m%d%H%M%S") if not lbForFilter else ltDateTime.strftime("DATETIME(%Y,%m,%d,%H,%M,%S)"))
            elif "NoneType" in lcTypestr:
                lxReturn = ("        000000" if not lbForFilter else "{^ - -  : : }")
            elif "<type 'time'>" in lcTypestr:  # This is a special datetime type received from COM clients
                lbGoodValue = True
                try:
                    lxDT = localtime(int(lxSourceValue))
                except:
                    lbGoodValue = False
                    lxDT = None
                if not lbGoodValue:
                    lxReturn = ("        000000" if not lbForFilter else "{^ - -  : : }")
                else:
                    lxDateTime = datetime(lxDT.tm_year, lxDT.tm_mon, lxDT.tm_mday, lxDT.tm_hour, lxDT.tm_min, lxDT.tm_sec)
                    lxReturn = lxReturn = (lxDateTime.strftime("%Y%m%d%H%M%S") if not lbForFilter else lxDateTime.strftime("DATETIME(%Y,%m,%d,%H,%M,%S)"))
            else:
                lxReturn = ""

        elif lcTargetType == "D":  # We have logic here to try to interpret a date string, if in USA
            # format MM/DD/YYYY or MM-DD-YYYY
            if isinstance(lxSourceValue, datetime):
                lxReturn = (lxSourceValue.strftime("%Y%m%d") if not lbForFilter else lxSourceValue.strftime("CTOD(%m/%d/%Y)"))

            elif isinstance(lxSourceValue, date):
                lxReturn = (lxSourceValue.strftime("%Y%m%d") if not lbForFilter else lxSourceValue.strftime("CTOD(%m/%d/%Y)"))

            elif isinstance(lxSourceValue, tuple) and (len(lxSourceValue) >= 3) and isinstance(lxSourceValue[0], int):
                # Probably a time tuple...
                ltDate = date(lxSourceValue[0], lxSourceValue[1], lxSourceValue[2])
                lxReturn = (ltDate.strftime("%Y%m%d") if not lbForFilter else ltDate.strftime("CTOD(%m/%d/%Y)"))
            elif mTools.isstr(lxSourceValue):
                if lxSourceValue.count("/") == 2:  # Probably a "MM/DD/YYYY" type date as a string
                    laStrs = lxSourceValue.split("/")

                    try:
                        lnMonth = atoi(laStrs[0])
                        lnDay = atoi(laStrs[1])
                        lnYear = atoi(laStrs[2])
                        ldDate = date(lnYear, lnMonth, lnDay)
                        lxReturn = (ldDate.strftime("%Y%m%d") if not lbForFilter else ldDate.strftime("CTOD(%m/%d/%Y)"))
                    except:
                        lxReturn = ""
                elif lxSourceValue.count("-") == 2:
                    laStrs = lxSourceValue.split("-")
                    try:
                        lnTest = atoi(laStrs[0])
                    except:
                        lnTest = -1
                    if lnTest > 0:
                        # A valid number, so test for YYYY-MM-DD or MM-DD-YYYY form: - Added YYYY-MM-DD format
                        # 02/16/2012. JSH.
                        if lnTest > 1000:
                            # year is first...
                            try:
                                lnMonth = atoi(laStrs[1])
                                lnYear = atoi(laStrs[0])
                                lnDay = atoi(laStrs[2])
                                ldDate = date(lnYear, lnMonth, lnDay)
                                lxReturn = (ldDate.strftime("%Y%m%d") if not lbForFilter else ldDate.strftime("CTOD(%m/%d/%Y)"))
                            except:
                                lxReturn = ""
                        else:
                            try:
                                lnMonth = atoi(laStrs[0])
                                lnDay = atoi(laStrs[1])
                                lnYear = atoi(laStrs[2])
                                ldDate = date(lnYear, lnMonth, lnDay)
                                lxReturn = (ldDate.strftime("%Y%m%d") if not lbForFilter else ldDate.strftime("CTOD(%m/%d/%Y)"))
                            except:
                                lxReturn = ""
                    else:
                        lxReturn = ""  # unidentified string.
                else:
                    lxReturn = lxSourceValue  # hope for the best.
            elif "NoneType" in repr(type(lxSourceValue)):
                lxReturn = ("        " if not lbForFilter else "{ / / }")
            else:
                lxReturn = ""

        elif (lcTargetType == "N") or (lcTargetType == "I") or (lcTargetType == "F") or (lcTargetType == "B"):
            if isinstance(lxSourceValue, float) or isinstance(lxSourceValue, int) or isinstance(lxSourceValue, xLongType):
                lxReturn = repr(lxSourceValue)
            elif isinstance(lxSourceValue, decimal.Decimal):
                lcStr = repr(lxSourceValue)
                lxList = lcStr.split("'")  # repr produces a value like "Decimal('123.4234234234')"
                lxReturn = lxList[1]  # Just the number representation.
            else:
                lxReturn = repr(lxSourceValue)

        elif lcTargetType == "Y":
            if isinstance(lxSourceValue, decimal.Decimal):
                lcStr = repr(lxSourceValue.quantize(decimal.Decimal("1.0000")))
                lxList = lcStr.split("'")
                lxReturn = lxList[1]

            elif isinstance(lxSourceValue, float):
                lnDec = decimal.Decimal(repr(lxSourceValue))
                lcStr = repr(lnDec.quantize(decimal.Decimal("1.0000")))
                lxList = lcStr.split("'")
                lxReturn = lxList[1]
            elif isinstance(lxSourceValue, int) or isinstance(lxSourceValue, xLongType):
                lnDec = decimal.Decimal(lxSourceValue)
                lcStr = repr(lnDec.quantize(decimal.Decimal("1.0000")))
                lxList = lcStr.split("'")
                lxReturn = lxList[1]
            else:
                lxReturn = repr(lxSourceValue)  # They need to get this right -- their problem.

        elif lcTargetType == "L":
            if isinstance(lxSourceValue, bool):
                lxReturn = (("T" if lxSourceValue else "F") if not lbForFilter else (".T." if lxSourceValue else ".F."))
            elif isinstance(lxSourceValue, int):
                lxReturn = (("T" if lxSourceValue == 1 else "F") if not lbForFilter else (".T." if lxSourceValue == 1 else ".F."))
            else:
                lxReturn = repr(lxSourceValue)  # Hope for the base.

        elif lcTargetType == "G":
            lxReturn = repr(lxSourceValue)  # Generally not supported

        elif lcTargetType == "Z":
            lxReturn = repr(lxSourceValue)  # Rare and not widely supported

        elif lcTargetType == "X":
            lxReturn = repr(lxSourceValue)  # ditto.

        else:
            lxReturn = repr(lxSourceValue)  # Shouldn't get here.

        return lxReturn

    def convert_vfp_type(self, lcSourceValue, lcSourceType, lbStripBlanks):
        """
        ----------------------------------------------------------------------------------------------------------
        Since the development of the CodeBaseTools Python API module that quickly and easily returns native Python
        objects for field values, there is limited use for this feature.  However, scatter and other functions
        do offer a "no data conversion" for slightly enhanced speed, and those unconverted values can be
        translated by this function when you are ready to use them in your program.

        Note that for char (binary) and memo (binary) and any text fields containing Unicode data, you MUST
        allow the CodeBaseTools module to do the conversions internally rather than using this function.
        ----------------------------------------------------------------------------------------------------------
        Converts the field value information returned from the CodeBase engine into the appropriate native Python value.

        If you elect to have converttypes=False in copytoarray() or scatter(), you may want to use this method at some
        point to convert the raw values returned to native Python values.  But, to do so you'll need to determine the
        source field value type.  That can be extracted using the afields() method and inspecting the cType property
        of the appropriate field.

        Returns the Python value (int, string, datetime, date, bool, etc.) or None on error.  Will also return None on
        empty DATE and DATETIME field values.
        """
        lxReturn = ""
        if lcSourceType == "C":
            lxReturn = (lcSourceValue.rstrip(" ") if lbStripBlanks else lcSourceValue)

        elif lcSourceType in "NBF":
            lcSourceValue = lcSourceValue.rstrip(" ")
            if lcSourceValue == "":
                lcSourceValue = "0.0"
            try:
                lxReturn = float(lcSourceValue)
            except ValueError:
                lxReturn = None

        elif lcSourceType == "L":
            lxReturn = (True if (lcSourceValue == 'T') else False)

        elif lcSourceType == "T":
            if lcSourceValue == "        000000":
                lxReturn = None
            else:
                lxReturn = datetime(atoi(lcSourceValue[0:4]), atoi(lcSourceValue[4:6]), atoi(lcSourceValue[6:8]),
                                    atoi(lcSourceValue[8:10]), atoi(lcSourceValue[10:12]), atoi(lcSourceValue[12:14]))

        elif lcSourceType == "D":
            if (lcSourceValue == "        ") or (lcSourceValue == ""):  # This is the VFP Empty Date
                # Can have a completely empty string if the last field in the table structure.
                lxReturn = None
            else:
                lxReturn = date(atoi(lcSourceValue[0:4]), atoi(lcSourceValue[4:6]), atoi(lcSourceValue[6:8]))

        elif lcSourceType == "I":
            if (lcSourceValue == ""): lcSourceValue = "0"
            try:
                lxReturn = int(lcSourceValue)
            except ValueError:
                lxReturn = None
            lxReturn = atoi(lcSourceValue)

        elif lcSourceType == "M":
            lxReturn = lcSourceValue

        elif lcSourceType == "Y":  # Currency Value, so have to return an exact decimal value.
            if lcSourceValue == "":
                lcSourceValue = "0.0000"
            try:
                lxReturn = decimal.Decimal(lcSourceValue)
            except decimal.InvalidOperation:
                lxReturn = None

        elif lcSourceType == "G":
            lxReturn = lcSourceValue

        elif lcSourceType == "Z":
            lxReturn = lcSourceValue

        elif lcSourceType == "X":
            lxReturn = lcSourceValue

        else:
            lxReturn = lcSourceValue

        return lxReturn

    def gatherdict(self, cAlias="", dData=None, codes="XX"):
        """
        Fast alternative to gatherfromarray() which updates the current record of the specified table
        with the data contained in the dictionary dData.  Will accept the optional alias name of a currently
        open table if you wish to store the data to a table that is not currently selected.  BE CAREFUL to make sure
        that the target table current record is the one you want to update.  Attempting to update a record
        in any table where there is no active record (either no records in the table or the record pointer at EOF)
        results in an error condition.

        Returns True on success, False on any kind of error.  Consult the error properties for reasons.

        Note that some types of errors will trigger Python exceptions.  These include passing illegal values
        to a field and passing field names that are too long or not of type string.

        Codes specify how string, bytes and bytearray objects are coded into field data.  XX is "default".  For
        details, see the docs for curval().  You can specify a conversion code for normal char and memo fields amd
        another code for conversions of char(binary) and memo(binary).  For type 'C', you can also set alternative
        codecs by setting the
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        if dData is None:
            raise ValueError("dData parameter must be a Dictionary")

        lbReturn = self.cbt.gatherdict(cAlias, dData, 0, codes)  # last 0 says "update don't append".
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def copytoarray(self, alias="", maxcount=1000, fieldtomatch=None, matchvalue=None, matchtype="=",
                    converttypes=True, stripblanks=False, fieldlist="", coding="XX"):
        """
        Copies values of all fields and all (or selected) records of the specified table into a list of dictionaries.
        If the optional parameter alias is left empty, the method works on the currently selected table.  Otherwise
        returns the contents of the open table with the alias name specified.

        The list will be in order as set for the requested table.  If no order is set, the ordering will be by record
        number.  Each element of the list object returned will be a dictionary consisting of one element for each field
        keyed by the field name.  On error will return an empty list.

        The values in the dictionary will be converted from VFP table field types to native Python types by default,
        but you can override this behavior by specifying the named parameter converttypes=False.  By default,
        character fields will be returned with their native VFP table trailing blanks intact.  If you want them removed
        automatically, specify the named parameter stripblanks=True.

        By default the resulting list contains one element for every non-deleted record in the table up to a limit of
        1000 records.  To limit the number of records to a different number to be returned, specify a non-zero value
        for the named parameter maxcount.  To return all records, regardless of the number, set maxcount=0.

        You can filter the records to be returned by supplying the name of a field in the named parameter fieldtomatch.
        This field name is not case sensitive.  In this case, you must also specify a value for matchvalue.  matchvalue
        may be assigned an appropriate Python variable type (object) congruent with the type of data table field you
        specified in the fieldtomatch variable.  Valid Python types associated with DBF Field types are:
            Field type I (Integer) - int or long
            Field type C (Character) - string
            Field type D (Date) - datetime.date (None for empty date)
            Field type T (Datetime) - datetime.datetime (None for empty datetime)
            Field type N (Number) - int, long, float or decimal.Decimal
            Field type Y (Currency) - decimal.Decimal
            Field type M (Memo) - string
            Field type L (Logical) - bool (True or False)

        If you specify a filter field and matchvalue, the field value will be compared to the matchvalue based on the
        matchtype.  If matchtype is its default value of '=' (equality, yes, it is a single equal sign), the values
        must match for a record to be included.  (Note that for string comparisons, if the matchvalue is shorter than
        the length of the field value, a comparison will only be performed up to the length of the match value.  If
        you want to have an exact string match, character for character you should supply a matchtype of '=='
        (double quotes)).  Valid values for the matchtype include:
            =
            >=
            <=
            == (Exactly equal. See note on string comparisons above. Identical to = for non-string values.)
            <> (NOT equal)
            $ (Is Contained In.  Equivalent to the Python 'in' operator.  The field value string must be entirely
               contained in the matchvalue string.)

        The reason for the filter mechanism is to allow very fast record filtering in the C component before Python
        ever sees the record values.  This can potentially eliminate significant numbers of records you just don't want
        to see.  Then you can perform tests on the resulting array to refine your search more precisely.  If you are
        using the client-server version of CodeBase, the filter test will be performed on the server before the data
        is actually retrieved across the network.

        Returns the list object which will be empty on error (len() of 0), in which case cErrorMessage and nErrorNumber
        will be set with explanatory values.  If the list is empty and you included values for fieldtomatch and
        matchvalue and if nErrorNumber = 0, your filter criteria didn't match any records.

        ValueError will be raised if non-matching value types are passed as matchvalue or if a non-existing field name
        is passed.

        Sets the value of the tally property to the number of records added to the array.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        self.tally = 0
        lcOldAlias = self.alias()
        if alias != "":
            self.select(alias)
        xRecArray = list()

        laFldArray = self.afields()
        if laFldArray is None:
            self.cErrorMessage = "No table is selected for the copy"
            self.nErrorNumber = 69384
            return xRecArray  # empty list.

        lnCount = len(laFldArray)
        lbTestMatch = False
        lnTally = 0

        if fieldtomatch is not None:
            fieldtomatch = fieldtomatch.upper()
            lcType = ""
            lcFiltExpr = ""
            for jj in range(0, lnCount):
                if fieldtomatch == laFldArray[jj].cName:
                    lcType = laFldArray[jj].cType
                    break

            if lcType == "":
                # Bad field name...
                raise ValueError("Field name: " + fieldtomatch + " doesn't exist")

            lcFiltExpr = self._makefilterstring(fieldtomatch, matchvalue, matchtype, lcType)
            lbTestMatch = True
            lnTestFilt = self.cbt.preparefilter(lcFiltExpr)
            if lnTestFilt == 0:
                self.cErrorMessage = "Invalid filter for this table"
                self.nErrorNumber = -8349
                return self.recArray

        lnRecCnt = 0
        if maxcount == 0:
            try:
                maxcount = sys.maxint  # the most records that a VFP table can hold anyway.
            except:
                maxcount = sys.maxsize
        lnResult = self.cbt.goto("TOP", 0)
        lbRecordOK = True
        while not self.cbt.eof():
            lbRecordOK = True
            if lbTestMatch:
                lnTestFilt = self.cbt.testfilter()
                if lnTestFilt == 0:
                    lbRecordOK = False
                elif lnTestFilt == -1:
                    raise ValueError("Could not test filter")
            if lbRecordOK:
                # converttypes=True, stripblanks=False
                xDict = self.scatter(converttypes=converttypes, stripblanks=stripblanks, fieldList=fieldlist,
                                     coding=coding)
                lnRecCnt += 1
                xRecArray.append(xDict)
                lnTally += 1
                if lnRecCnt > maxcount:
                    break
            self.cbt.skip(1)

        self.cbt.select(lcOldAlias)
        if lnTally > 0:
            self.tally = lnTally
        self.cbt.clearfilter()
        fieldtomatch = None
        return xRecArray

    def _makefilterstring(self, lpcFld, lpcMatch, lpcRelate, lpcFldType):
        """
        Internal method to make a filter string that CodeBase can use to evaluate a match for a current
        record.
        """
        if ((lpcFldType == "D") or (lpcFldType == "T")) and (lpcMatch is None):
            lcReturn = "(EMPTY(" + lpcFld + ")" + ("=.T.)" if lpcRelate == "=" else "=.F.)")
        else:
            lcReturn = lpcFld + lpcRelate + self._py2cbtype(lpcMatch, lpcFldType, lbForFilter = True)
        return lcReturn

    def replace(self, lpcFldName, lpxValue, coding="XX"):
        """
        Provides a quick and convenient way to replace the value of a field in the current record in any open table
        without worrying about selecting the current table.  Further, it takes any native Python type and properly
        stores it in the field, assuming that the Python type is compatible with the data field type.  See the
        documentation for field and value types for the matchvalue in copytoarray().

        To specify a field in any open table, prepend the field name with the table alias separated by a period.  For
        example for the last_name field of the employee1 table, set the value of lpcFldName to "employee1.last_name".

        To Replace() a date field value with the VFP empty date, supply a value of 0, not None.

        Returns True on success, False on failure, and stores error info into cErrorMessage and nErrorNumber.

        The 'coding' parm added April 21, 2018.  See the class docstring for codes that define how string values will
        be converted for different text encodings. "XX" is the default coding.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        # if isinstance(lpxValue, unicode):
        #    lpxValue = lpxValue.encode('cp1252', 'ignore') ## The VFP type tables we have created assume
        #    ## Windows Codepage 1252, which is an 8-bit Western European language code page.  Unicode strings
        #    ## are not supported without special coding.
        #    ## Added 11/10/2014. JSH.
        lbReturn = self.cbt.replace(lpcFldName, lpxValue, coding)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def replacelong(self, lpcFieldName, lpnValue=0):
        """
        Very fast storage of a long value into a long field. Very little error checking.
        :param lpcFieldName: Field name with optional alias specification.
        :param lpnValue: Long value.  NO conversions are performed!!
        :return: 1 on OK, else 0 or -1 on failure.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.replacelong(lpcFieldName, lpnValue)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def replacedatetimeN(self, lpcFieldName, lpnYear=0, lpnMonth=0, lpnDay=0, lpnHour=0, lpnMinute=0, lpnSecond=0):
        """
        Very fast storage of a long value into a long field. Very little error checking.
        :param lpcFieldName: Field name with optional alias specification.
        :param lpnYear: The year number (3 or 4 digits).  Recognizes years back to 200 AD.
        :param lpnMonth: Month number 1 to 12.
        :param lpnDay: Day of month number.
        :param lpnHour: Hours part of the time from 0 to 23
        :param lpnMinute: Minutes past the hour, from 0 to 59
        :param lpnSecond: Seconds past the minute, from 0 to 59
        :return: 1 on OK, else 0 or -1 on failure.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.replacedatetimen(lpcFieldName, lpnYear, lpnMonth, lpnDay, lpnHour, lpnMinute, lpnSecond)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def closetable(self, alias=""):
        """
        Closes the currently open table or the table with the alias specified by the optional alias parameter.  Flushes
        all unsaved data to the disk and releases all the file handles.

        The table cannot be accessed again without another use() method, and references to this alias will fail.

        Returns True on success (table was open and is now closed) or False on failure (didn't find anything to close).
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.closetable(alias)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def fielddicttolist(self, lxDict, lcAlias="", bPartial=False, cDelim="<~!~>"):
        """
        Transforms a record value dictionary as provided by the scatter() method to a list in the format required
        by the insertintotable() method, together with a field name string as required by that method as well.

        The first parameter is the dictionary.  The second parameter is the alias name of the target table which
        you'll be updating from the values in the list.  Note this parameter is required.  Returns a tuple consisting
        of the list of values as the first element and a delimited string list of field names in the correct
        order as the second element.
        bPartial:  Indicates the lxDict has only SOME of the fields.  Make sure ALL other fields are set.

        Note that the lxDict dictionary must have an entry for every field in the table, but it is NOT limited
        just to those values.  This allows you to build a dictionary with other information besides the required
        fields for whatever purpose in subsequent processing, and just use the fields required for the table when
        processing through this method.

        NOTE: This method is useful for supplying a list and a string of field names for the older version of
        insertintotable().  In this version of CodeBaseTools for Python, a dictionary of field values keyed by
        field name may be passed directly to insertintotable(), which uses the function insertintotableex() to
        pass the dictionary directly to the CodeBase engine.
        """
        lcOldSel = self.alias()
        lxResult = True
        if lcAlias:
            lxResult = self.select(lcAlias)
        lxListRet = list()
        if bPartial:
            blankDict = self.scatterblank()  # make sure all fields have something
            blankDict.update(lxDict)
            lxDict = blankDict
        laFldRec = list()
        if lxResult:
            lxFields = self.afields()
            for lx in lxFields:
                lxListRet.append(lxDict[lx.cName])
                laFldRec.append(str(lx.cName))
        if lcOldSel:
            self.select(lcOldSel)
        lcFldRet = cDelim.join(laFldRec)
        return lxListRet, lcFldRet

    def closedatabases(self):
        """
        Closes all open tables, flushes all pending writes to the disk, but keeps the CodeBase engine running.  Call this
        method when you are done with a bunch of work, but need to keep the object alive for additional work.  Any calls
        that require a currently selected table or an open table will fail after this has been called and until new calls
        to use().
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.closedatabases()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def createtable(self, lcTableName, lxFieldInfo):
        """
        Method that creates a DBF type table, overwriting any pre-existing table of the same name.  Pass
        the name of the table and an array of field info structures or a string containing field definitions.

        The first parameter is a string value containing a file name specification.  The name should have a '.DBF'
        extension for compatibility with other programs that might access it; however, the .DBF extension is not
        strictly required so long as you know what the file is -- these routines will open it again as a DBF file
        if you pass its full name and extension to the use() method.

        The second parameter to pass takes one of two different types of values:
            - String - containing field specification information
            - List - with each element containing an instance of a VFPFIELD object having information on one field.

        The string version is the easiest to program.  It is a newline-delimited list of field specification strings.
        Each field specification string is a comma delimited set of field property values.  The contents are:
        <field name>,<field type>,<field width>,<field decimals>,<field allows nulls>

        For example:
        "FIRSTFIELD,C,30,0,FALSE\n" would create a character field named FIRSTFIELD with a width of 30 bytes and would
        not allow nulls.  For a brief list of field types see the doc for copytoarray().  For complete documentation on
        defining fields in a DBF, see the CodeBase documentation.

        One field will be created in the table for each specification string, in the order listed.

        Alternatively, you may supply a list of VFPFIELD objects each containing one field spec.  The properties of
        each object must be set to define the field.  For example, for a VFPFIELD object vfx you might set:
            vfx.cName="FIRSTFIELD"
            vfx.cType="C"
            vfx.nWidth=30
            vfx.nDecimals=0
            vfx.bNulls=False

        The most common use of this list of VFPFIELD type input is to obtain a list generated by the afields() method
        from a currently open table.  This allows you to easily create an exact empty copy of an existing table with just
        two lines of code, or by append to or deleting from the afields() result list, you can alter the table structure
        to fit your immediate needs.  Note that the elements of the VFPFIELD object class ARE type-specific.  For example
        the cName value MUST be a string, the nWidth value MUST be an integer.

        General rules:
            1) Field names must be 10 characters or less and start with a letter or underscore
            2) No two field names may be the same
            3) Maximum number of fields is 254
            4) Maximum size (width) of a character field is 255 bytes
            5) Maximum size of a record is 65,500 bytes
            6) If you include one or more type 'M' fields (memo or large text) fields, a memo file will also
               be created with the same base name as the table and an extension of .FPT

        Limitations:
            1) This version of CodeBase cannot read or update VFP tables with auto-incrementing Integer fields.
            2) Tables contained in a Visual FoxPro database container can be opened, but features derived from
               the database container like long field names, triggers, validations, captions, etc., will not be
               available.
            3) Tables will always be created in the Visual FoxPro FREE mode, meaning that they are not contained
               in a database container.

        After the createtable() method finishes successfully, the newly created table will be empty and will be
        the currently selected table, opened in read/write mode, ready for appending records.  The alias of this
        table will be the base file name using the default alias naming rules when a specific alias is not provided.
        Be sure that is no table open with that alias already, or that table will be automatically closed and
        possibly overwritten!  If you want the new table to have a custom alias, close it immediately using
        the alias() value and re-open it with your custom alias specified explicitly.

        Returns True on success, False on failure, in which case inspect properties cErrorMessage and nErrorNumber
        for an explanation.  Some errors such as invalid field names or other field characteristics will trigger
        a Python ValueError exception.  When a list is supplied, the CodeBasePYWrapper pyd module generates the
        Python exceptions, if required.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        bMadeFields = False
        if mTools.isstr(lxFieldInfo):
            # they are passing a string, so we will convert it to an array of VFPFIELD objects
            xFields = list()
            bMadeFields = True
            xItems = lxFieldInfo.split("\n")
            for xI in xItems:
                xI = xI.replace("\r", "")
                xI = xI.replace("\n", "")  # Just in case.
                if xI == "":
                    break
                xParts = xI.split(",")
                if len(xParts) != 5:
                    raise ValueError("Field definition string %s has too few elements: %d" % (xI, len(xParts)))
                xV = VFPFIELD()
                xV.cName = xParts[0].strip()
                if len(xV.cName) > 10:
                    raise ValueError("Field name is too big, max 10 characters")
                xV.cType = xParts[1]
                if not (xV.cType in "C,I,N,Y,B,F,L,M,D,T,G"):
                    raise ValueError("Invalid field type value")
                xV.nWidth = int(xParts[2])
                if xV.nWidth <= 0:
                    raise ValueError("Invalid field width value")
                elif xV.nWidth > 255:
                    raise ValueError("Field Width Value exceeds maximum")
                xV.nDecimals = int(xParts[3])
                if xV.nDecimals < 0 or xV.nDecimals > 17:
                    raise ValueError("Decimals value invalid")
                xV.bNulls = (True if xParts[4].strip() == "TRUE" else False)
                xFields.append(xV)
        elif isinstance(lxFieldInfo, list):  # we pass the list directly since that's what the engine expects.
            xFields = lxFieldInfo
        else:
            raise ValueError("Structure information must be string or list")

        lbReturn = self.cbt.createtable(lcTableName, xFields)
        if bMadeFields:
            del xFields
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()

        return lbReturn

    def appendblank(self):
        """
        Adds a new, completely blank record at the end of the currently selected table.  Once the record is added
        you can call gatherfromarray() or replace() to fill data into the fields of the new record.

        Returns True on success, False on error, in which case check cErrorMessage and nErrorNumber for more information.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lbReturn = self.cbt.appendblank()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def dispstru(self, lbToText=False):
        """
        Writes text to the stdout via a series of print commands which contains information about the structure of the
        currently selected table.  This is the equivalent of the VFP command DISPLAY STRUCTURE.

        Pass a True to lbToText and the function will return a string with this information.

        Returns True for OK and False on failure (usually because there is no currently selected table if lbToText is False.
        If lbToText is True, then returns the string on Success and an empty string on failure.
        """
        lxFlds = self.afields()
        if lxFlds is None:
            return False if not lbToText else ""
        lnCnt = len(lxFlds)
        if lnCnt == 0:
            return False if not lbToText else ""
        lcDBF = self.dbf()
        lnReccount = self.reccount()
        chgTime = os.stat(lcDBF)
        chgTimex = localtime(chgTime.st_mtime)
        lcChgTime = strftime("%Y-%m-%d %H:%M:%S", chgTimex)
        xLines = list()
        lcStr = """
Structure for table:    %s
Number of data records: %d
Date of last update:    %s
Field  Field Name     Type     Width    Dec  Nulls
"""
        xLines.append(lcStr % (lcDBF, lnReccount, lcChgTime))
        lcFldStr = "%5d  %-13s     %1s%10d%7d   %-5s"
        lnTotWidth = 0

        for jj, lx in enumerate(lxFlds):
            xLines.append(lcFldStr % (jj + 1,
                                      lx.cName, lx.cType, lx.nWidth, lx.nDecimals, ("Yes" if lx.bNulls else "No")))
            lnTotWidth = lnTotWidth + lx.nWidth
        xLines.append("** Total ** %24d" % (lnTotWidth+2))
        # The total width is the sum of the field widths + 1 byte for the deleted flag and 1 byte for the null flags.
        lcStr = "\r\n".join(xLines)
        if not lbToText:
            print(lcStr)
            return True
        else:
            return lcStr

    def copyindexto(self, cSourceAlias, cTargetAlias):
        """
        Takes two open tables and copies the indexes from the source table to the target table.  Removes the indexes
        from the target table at the start.

        It is recommended that the target alias table be opened in Exclusive mode unless you are sure that there
        cannot be any conflicts in opening the table.
        :param cSourceAlias:
        :param cTargetAlias:
        :return:  True on success, False on failure of any kind -- including if the target table doesn't have all
        the fields found in the source indexes.

        NOTE: The target table must be a VFP "FREE" table NOT part of a database container.  Performing this method
        on a target table in a database container will result in data corruption.  Source tables, however, may
        be part of a VFP database container without problems, however PRIMARY KEY index tags may NOT copy correctly,
        but CANDIDATE and UNIQUE indexes will.
        """
        bReturn = self.cbt.copytags(cSourceAlias, cTargetAlias)
        if not bReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return bReturn

    def copyto(self, cAlias="", cOutput="", cType="", cTestExpr="", bHeader=False, cFieldList="", bStripBlanks=False):
        """
        An output routine similar to the VFP COPY TO.  It takes any currently open table
        and writes out a text or DBF file in one of four possible formats with the table data
        which meets an optional selection criteria.
        Parameters:
        - cAlias -- The alias name of the open table with contents to be written out.
                    If blank, the currently selected table will be used.
        - cOutput -- [REQUIRED] The fully qualified path name of the output file.  Be sure to include
                  the file name extension, as this function makes no assumptions about the output file name.
        - cType -- [REQUIRED] One of four string values: "SDF", "CSV", "TAB", "DBF", indicating:
            SDF = System Data Format -- Fixed width fields in records delimited by CR/LF
            CSV = Comma Separated Value -- Plain text with fields separated by commas
                  unless they are text, in which case, they are set off with double quote
                  signs thus: "joe's Diner","Hambergers To Go",35094,"Another Name".
                  If the text field has an embedded quote, the quote is transformed into a
                  double-double quote: "Joe's ""Best"" Hambergers","Pete's Grill".
                  Embedded line breaks within character/text fields are NOT supported.
                  If encountered, they WILL BE REMOVED from the data prior to output.
            TAB - Tab separated values.  Plain text with fields separated by a single
                  Tab (ASCII Character 9) character.  Character and number data is given
                  without any quotations other than what might exist in the data.  In
                  this implementation embedded line breaks in character fields and em-
                  bedded TAB characters in text fields are NOT supported and WILL BE
                  REMOVED prior to output of the data.
            DBF - Standard VFP-type DBF table.  Data structure will be identical to the
                  source unless a cFieldList is supplied.  No indexes will be copied.  Be SURE
                  to supply the file name extension as none will be assumed.
        - bHeader -- If True, then CSV and TAB outputs will include a first line consisting
                     of the names of each field, separated by the appropriate separator,
                     either Comma or TAB respectively.  Otherwise no field name data will
                     be included.  Default is False.
        - cFieldList -- A list of field names to include in the output.  If omitted, ALL
                     fields will be included (except as indicated in NOTES).
        - cTestExpr -- Text string containing a VFP Logical expression which must be True
                     for the record to be output.  If empty or null, ALL records are output.
        - bStripBlanks -- If True, trailing blanks will be stripped at the end of records.
                     Defaults to False.
        NOTES:
        - For all types except DBF, fields which cannot be represented by simple text
          strings without line breaks will be skipped.  These include types MEMO, GENERAL,
          CHARACTER(BINARY), and BLOB.
        - In text formats Date fields will be output as MM/DD/YYYY or as specified by the current setting
          of cbwSETDATEFORMAT().
        - In text formats Date-Time fields will be represented as CCYYMMDDhh:mm:ss, which is re-importable
          back into CodeBase DBF tables via cbwAPPENDFROM().
        - The Deleted status of records will be respected based on the setdeleted() setting.
        - Current Order will be respected.
        - The TYPE of the file will be determined by the cType parameter regardless of the
          name of the output file!!
        - The table named in the cOutput parm should NOT be open.  If it exists, it will be overwritten

        - Field lengths for the SDF type output are, in general determined by the length
          of the data fields themselves.  Rules for length and output are as follows by
          type:
            - C - Width of field.
            - D - 8 characters in the form YYYYMMDD
            - T - 16 characters in the form YYYYMMDDhh:mm:ss
            - L - 1 character ('F' or 'T')
            - N - Determined by size and decimals of field.
            - F (Float) - like N
            - B (Double) - 16 bytes, left-filled
            - I (Integer) - 11 bytes, left-filled
            - Y (Currency) - 16 bytes, 4 decimals, left filled.

        RETURNS:
            Returns number of records output or -1 on error.  Zero records returned may be a valid result,
            if there are no records in the source table or none match the cTestExpr evaluation.

            If the function was successful, the new table will exist as specified but it will NOT be open.  If you
            need to work with it, you will need to open it normally with the use() method after this method terminates.

        WARNING: If the output type is DBF and the base file name of the target table is the same as the source
        table, then the alias of the source table MUST be a non-default string (not the same as the base table name).
        The CreateTable() function, which is at the core of this function when outputting a DBF, creates the table
        and opens it with the default alias name.  If that alias name is already in use, the original table is closed,
        and the results will likely NOT be what you expect.

        KNOWN ISSUE: If you do a copyto() on a currently open table, and wish to repeat the process on that table
        with a different target result, it may be necessary to close the source table and re-open it.  In some
        instances the repeat copy will create the target table but NOT put any records into it.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lnReturn = 0
        lbResult = True
        if not cOutput:
            raise ValueError("cOutput may not be empty")

        if not cType:
            raise ValueError("cType may not be empty")

        if cAlias is None:
            cAlias = ""

        if cFieldList is None:
            cFieldList = ""

        if cTestExpr is None:
            cTestExpr = ""

        lnHdr = (1 if bHeader else 0)
        lnStrip = (1 if bStripBlanks else 0)
        lnReturn = self.cbt.copyto(cAlias, cOutput, cType, lnHdr, cFieldList, cTestExpr, lnStrip)

        if lnReturn < 0:
            self.cErrorMessage = self.cbt.geterrormessage() + " COPY TO FAILED"
            self.nErrorNumber = self.cbt.geterrornumber()
        return lnReturn

    def appendfrom(self, cAlias="", cSource="", cExpr="", cType="", cDialect="", cDelimiter="", cQuoteChar="",
                   cFieldList="", bFieldsFromRow1=True):
        """
        Works much like the APPEND FROM command in Visual FoxPro.  You specify a currently open alias as
        the target, or pass "" or None, if you want to use the currently selected table as the target.  The
        engine will attempt to append the data from the file cSource into that table.  It will select for the
        append process only records which meet the cExpr logical expression if it is not "" or None.  NOTE: as
        of this version, the cExpr filter is ONLY applied for DBF type source files.

        Four types of files can be the source: DBF tables, System Data Format (fixed position fields in a
        text file using CR/LF separators for the records, CSV (comma separated values), and TAB/TXT (tab character
        delimited values).  CSV and TAB/TXT file reading makes use of the Python csv library, which means that it
        may be slightly slower to execute than 'C' native imports of DBF or SDF type files, but also means that there
        is a great deal of flexibility in configuring the csv module to handle variations in CSV formatting.

        Four optional parameters provide some configuration flexibility for use of this module:

        cDialect = supports the dialect options supported by csv.  Typically these are "excel" and "excel-tab", and
        provide for reading csv or tab delimited (.TXT) files created by Excel.  Leave empty for native appendfrom()
        processing or to override with the remaining parms.

        cDelimiter = define a character which separates fields in the input text string.  This defaults to "," (comma)
        for native appendfrom() operations and 'excel' dialect.  For tab-delimited files, use "\t" (tab char), which
        is also the default for 'excel-tab' dialect.

        cQuoteChar = the character used in the file to set off text strings, if any.  For native appendfrom() behavior,
        this defaults to the double quote charcter: '"'.  Typically tab-delimited files don't use quote characters,
        but some custom formats may.  By default, the csv handling of double quote characters embedded within text
        fields is applied (it recognizes "" embedded in the string as a quote character).  For custom handling of
        such situations, see customizing csv for your needs below.

        bFieldsFromRow1 = Indicates whether the reader should look for field names in the first line of the input
        file.  VFP does NOT support this feature, but csv DOES, and the default for appendfrom() behavior with
        CSV or TAB/TXT input is for this to be True.  If you do NOT use this feature, then you have two options:
        1) Supply cFieldList with a comma delimited text string containing one field name for each field in the
        input table or 2) Be ABSOLUTELY CERTAIN that the target DBF table has exactly the fields you want in the
        exact order as found in the input.

        Customizing csv for your needs: When this object class is created, the self.oCSV property is initialized
        to an instance of csv.  That instance, configured to provide default behavior will be used for the
        appendfrom() function.  However, if you configure your own csv instance and set custom dialect information
        for it, you can use that by setting the oCSV attribute to a reference to your custom instance of csv.  See
        the Python documentation for its csv.Dialect class for how to customize the format.  Alternatively, if
        you know you'll always be wanting a particular Dialect configuration, then specify the properties of
        Dialect on the self.oCSV object:
            - delimiter
            - doublequote
            - escapechar
            - lineterminator
            - quotechar
            - quoting
            - skipinitialspace

        If you set any one of these values to a non-None value, you'll need to set them all to meet your specific
        needs, and no default Dialect information will be applied, and the parameters cDialect, cDelimiter,
        cQuoteChar will ALL be ignored.

        NOTE: The default behavior of appendfrom() with CSV and TAB/TXT type files will correctly append files
        created by the copyto() method of this class.

        Will use the file name extension of the source file to determine the type.  If that value is not "DBF" for
        DBF table, SDF or DAT for System Data Format, or CSV for Comma Separated Values type, then you will need
        to specify the type in the cType parameter as one of these 4 possible strings.  Since TAB is typically
        not used as an extension for a tab-delimited file, you should expect to use the string 'TAB' for the
        cType parameter and use .TXT or other appropriate extension for the file name.

        Returns the number of records appended or -1 on error.  Zero records is a potentially valid return if
        for some reason the source table/file is empty or if cExpr matches no records..

        For SDF appends, the expected field widths for type C and type N fields are determined by the
        table field sizes themselves.  However other field types like D, T, and B have special rules.  See
        the documentation for the copyto() method for size rules for output to SDF from these fields.  This
        appendfrom() method uses the same width assumptions for append from an SDF type file.  If these special rules
        are not consistent with an SDF file supplied from another source, create a DBF table with all type C
        fields and set the width of each to the corresponding width of each field in the source file.  After appending
        from the SDF file, you can then copy from the resulting DBF table to one with the desired field types, making
        conversions as necessary.
        """
        lnReturn = 0
        lbResult = True
        if not cSource:
            raise ValueError("cSource may not be empty")

        self.cErrorMessage = ""
        self.nErrorNumber = 0
        cTestName = cSource.upper()
        cFile, cExt = os.path.splitext(cTestName)
        cExt = cExt.replace(".", "")
        if not cType:
            cType = cExt
        if cType in ",DBF,SDF,":
            lnReturn = self.cbt.appendfrom(cAlias, cSource, cExpr, cType)
            if lnReturn < 0:
                self.cErrorMessage = self.cbt.geterrormessage() + " APPEND FAILED"
                self.nErrorNumber = self.cbt.geterrormessage()
        else:
            try:
                if _ver3x:
                    fFile = open(cSource, mode="r", encoding=self.cPreferredEncoding, errors="replace", newline="")
                else:
                    fFile = open(cSource, "rb")
            except:
                lnReturn = -1
                fFile = None
                self.cErrorMessage = "Unable to open %s for reading" % (cSource,)
                self.nErrorNumber = -3549
            if lnReturn >= 0:
                bCustomCSV = False
                if self.oCSV.Dialect.delimiter is not None or \
                   self.oCSV.Dialect.doublequote is not None or \
                   self.oCSV.Dialect.escapechar is not None or \
                   self.oCSV.Dialect.lineterminator is not None or \
                   self.oCSV.Dialect.quotechar is not None or \
                   self.oCSV.Dialect.quoting is not None or \
                   self.oCSV.Dialect.skipinitialspace is not None:
                    bCustomCSV = True
                if not cAlias:
                    cAlias = self.alias()
                # Set up the reader configuration variables.
                if bFieldsFromRow1:
                    cFields = None
                else:
                    if cFieldList:
                        cFields = cFieldList
                    else:  # If they don't provide a field list, then we have to assume we are loading all fields.
                        xFlds = self.afields(cAlias)
                        xFList = list()
                        for xF in xFlds:
                            xFList.append(xF.cName)
                        cFields = ",".join(xFList)
                        cFields = cFields.upper()
                if cQuoteChar:
                    lcQuoteChar = cQuoteChar
                else:
                    lcQuoteChar = '"'
                if cDialect:
                    lcDialect = cDialect
                else:
                    lcDialect = ""

                if cFields:
                    xFields = cFields.split(",")
                else:
                    xFields = None
                # Note: The fieldnames parm below must be a list or none, NOT a comma delimited string.
                if bCustomCSV:
                    xReader = self.oCSV.DictReader(fFile, fieldnames=xFields, restkey=None, restval=None)
                elif cType == "TAB":
                    if not cDelimiter:
                        lcDelimiter = "\t"
                    else:
                        lcDelimiter = cDelimiter
                    xReader = self.oCSV.DictReader(fFile, fieldnames=xFields, restkey="UNKNOWN", restval="",
                                                   dialect=None, delimiter=lcDelimiter, quotechar=lcQuoteChar)
                elif cType == "CSV":
                    if not cDelimiter:
                        lcDelimiter = ","
                    else:
                        lcDelimiter = cDelimiter
                    xReader = self.oCSV.DictReader(fFile, fieldnames=xFields, restkey="UNKNOWN", restval="",
                                                   dialect=None, delimiter=lcDelimiter, quotechar=lcQuoteChar)
                else:
                    fFile.close()
                    lnReturn = -1
                    raise ValueError("Unrecognized source file type")
                if lnReturn == 0:  # OK to start reading and appending...
                    for xRow in xReader:
                        bResult = self.insertdict(xRow, lcAlias=cAlias)
                        if not bResult:
                            lnReturn = -1  # ERROR MESSAGES
                            break  # Already set the error flags
                        else:
                            lnReturn += 1
                fFile.close()

        return lnReturn

    def insertintotable(self, lxValues, lcFields="", lcAlias="", lcDelimiter="<~!~>", coding="XX"):
        """
        Performs in one function the tasks of appending a new record to a table and filling the fields with new data.
        This approach requires that values for every field must be supplied in the lxValues parameter.  The rules for
        this field are identical to those for gatherfromarray() in the All Fields Mode.

        Unlike gatherfromarray() this method allows you to specify an alias name if you don't want to insert the
        data into the currently selected table.  Of course the table specified by alias must first have been opened
        by the use() method.

        Call this function repeatedly in a tight loop to add lots of records to the table.  Do NOT call flush() or
        flushall() until you are done.

        The lxValues field is where you pass field data to the method.  There are three forms this value may take.
        Depending on which form you supply, you may need to supply the lcFields parameter as well:

        1) Dict() should be keyed by field name with native python values for each field as the item value.  Do not
           pass any value for lcFields or an empty string.  lcDelimiter is ignored.  This is a fast, native Python
           field update.  In this form, you do not need to populate the dict() with every field in the table,
           just the ones you need to update.  However, be sure to pre-populate any primary keys that must be
           unique!

        2) String should contain a delimited text representation of the contents of each field.  Either all fields must be
           supplied in the correct order as in the table, or supply a list of field names in the appropriate order
           in the lcFields parameter.  The lcDelimiter value should be used for both.  See the gatherfromarray()
           doc string for details on how strings should be formatted for logical, date and other non-text fields.

        3) List() containing one element for each field you want to update.  In this case you MUST supply an
           element of the list for every field in the table in the correct order.

        Returns True on Success, False on error.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lbReturn = True
        if isinstance(lxValues, dict):
            lbReturn = self.cbt.gatherdict(lcAlias, lxValues, 1, coding)
            # last 1 says "Append the record with this data"
            if not lbReturn:
                self.cErrorMessage = self.cbt.geterrormessage()
                self.nErrorNumber = self.cbt.geterrornumber()
        else:
            if mTools.isstr(lxValues):
                lcValStr = lxValues
            elif isinstance(lxValues, list):
                lxFlds = self.afields(lcAlias)
                for ix, lf in enumerate(lxFlds):  # checking and converting the types
                    lxValues[ix] = self._py2cbtype(lxValues[ix], lf.cType)
                    if lxValues[ix] is None:
                        self.cErrorMessage = "Bad Field Type: " + lf.cName
                        self.nErrorNumber = -8593
                        lbReturn = False
                        # print("bad type:", lf.cName, lf.cType)
                        break
                lcValStr = lcDelimiter.join(lxValues)
            else:
                raise ValueError("lxValues must be dict(), list() or string()")

            if lbReturn:
                lbReturn = self.cbt.insertintotable(lcAlias, lcValStr, lcFields, lcDelimiter)
                if not lbReturn:
                    self.cErrorMessage = self.cbt.geterrormessage()
                    self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def insertdict(self, dDict, lcAlias="", coding="XX"):
        """
        Simple version of insertintotable() that just takes a dict() of field names and values and inserts
        the data into the specified lcAlias table or the current table if it's empty.  Returns True on OK,
        False on any error.  Synonymous with insertintotable() with a dict() parm, but simpler calling
        signature.

        For coding specifies txt string conversion rules for text and binary data.  See class docstring for
        details of alloable codes.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0

        lbReturn = self.cbt.gatherdict(lcAlias, dDict, 1, coding)
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def setdeleted(self, lbHow=None):
        """
        This function sets the DELETED status for all table access.  If DELETED ON or TRUE is set by passing a
        value of True to the method() then most table navigation methods will automatically skip over deleted
        records.  When this module is initialized, the status is set to DELETED OFF in which case all records
        are visible regardless of deleted status.  To turn DELETED OFF, pass a False to the method.  This may be done
        at any time and will affect all future table navigation until changed.

        Again, note that this follows the FoxPro convention where the DELETED ON status tells the data engine to
        ignore deleted records (Delete testing is ON).  When DELETED OFF has been set in VFP, all records, including
        those on which the deleted flag has been set to TRUE are visible.

        For most normal operations, you will want to issue a setdeleted(True) command to ensure that you don't
        have to test constantly for the deleted() status of every record.

        The DELETED status affects all skip() functions and goto() functions except for those where you goto()
        a specific record number.  In that case you will still need to test the deleted() status.

        Also, the copytoarray() method adheres to the DELETED status.

        If you pass None as the parameter lbHow, then the function returns the current state of the DELETED flag:
        True for DELETED is ON, False for DELETED is OFF.  Otherwise, always returns True.
        """
        if lbHow is None:
            lnParm = -1
        else:
            lnParm = (1 if lbHow else 0)

        return self.cbt.setdeleted(lnParm)

    def delete(self):
        """
        Marks the current record of the currently selected table for deletion.  Does not move the record pointer
        even if DELETED is set ON.

        Returns True on success, False on failure, in which case cErrorMessage and nErrorNumber will provide more
        information.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.delete()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def flock(self, retries=0, interval=.1):
        """
        Locks the currently selected table.  This allows you to safely update multiple records or collect statistics
        about the contents of a table without risk of other processes changing the contents.  Also if you cannot place
        the flock() and can afford the time, pause for a few milliseconds and then try again.  You can do this
        explicitly by specifying a number of retries in the retries parm and an interval specified in fractions of a
        second (default is 0.1 seconds).
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.flock()
        rtcnt = 0
        while (not lbReturn) and (rtcnt < retries):
            rtcnt +=1
            sleep(interval)
            lbReturn = self.cbt.rlock()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def rlock(self, retries=0, interval=.1):
        """
        Locks the currently selected record of the currently selected table.  This allows you to safely update
        the contents of the record either by a replace() or a gatherfromarray().

        This method plus its corresponding one unlock() gives you the added measure of certainty relative to
        the record updating process.  If activity in the target table is light, with few other users making updates,
        it is not necessary to call rlock() explicitly.  After your field value updates, CodeBase waits until you move off
        the record or call flush() to lock the record, write the data to the file, and then automatically unlock the
        record.  Typically, this works without a hitch.  But in situations where there is a lot of save activity it
        can happen that it is impossible to obtain the lock, and an error is reported from the record move or flush.

        To avoid that possibility, the sequence would be:
            - navigate to the record with seek(), skip() or goto()
            - scatter() the field values so you can work with them in your program
            - when you are ready to save changes back to the record, call rlock()
            - if rlock() returns True, you can call replace() or gatherfromarray() to make the changes
            - if rlock() and the update was successful, call flush()
            - call unlock()

        Returns True on Locked OK, False if unable to get lock for any reason.

        Note: Leave the rlock() in place as short a time as possible.  Also if you cannot place the rlock() and can
        afford the time, pause for a few milliseconds and then try again.  You can do this explicitly by specifying
        a number of retries in the retries parm and an interval specified in fractions of a second (default is
        0.1 seconds).
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.rlock()
        rtcnt = 0
        while (not lbReturn) and (rtcnt < retries):
            rtcnt +=1
            sleep(interval)
            lbReturn = self.cbt.rlock()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def unlock(self):
        """
        Unlocks the current record of the currently selected table.  See rlock() for more information.  Returns True
        on success, False on failure.  Also is used to remove the table lock applied by flock().
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.unlock()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def pack(self):
        """
        Pack performs the VFP PACK function which copies all records NOT marked
        for deletion to a temporary table, then empties out the table, and
        copies the surviving records back into it. This method performs that
        task on the currently selected table.

        The VFP rule that this can only be performed on a table opened
        exclusively is NOT enforced in this module, but if it is expected that
        other users might access the table, then a use() with the exclusive
        option should be done before applying a pack().

        One of the side effects of a pack() is that memo files (the .FPT file)
        storing memo field data, are restructured, which usually produces a
        smaller file.

        Returns True on success, False on failure of any kind. A typical error
        would be that there is no currently selected table.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.pack()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def reindex(self):
        """
        Call this method to refresh the indexes on the currently selected table.

        Normally indexes in CDX index files (the Visual FoxPro standard) are updated automatically by this module
        as they are by VFP when changes are made to the table.  However, after a while the index pointers can become
        inefficiently arranged, and a reindex can possibly improve performance on large (over 1mm records) tables.

        Returns True on success, False on failure of any kind.

        Note: If it is possible that other users will be accessing the table you wish to reindex, it is best to
        use() the table with the exclusive=True parameter so as to prevent index corruption and data retrieval
        errors by the other users.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.reindex()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def scan(self, indexTag="", forExpr="", noData=False, fieldList=None, getList=False, stripblanks=False,
             bNoTop=False, bDescending=False, coding="XX"):
        """
        Functions as an iterator that can be used for successively returning rows from the currently selected table
        either as dictionaries, similar to scatter() which returns a dictionary for the current record, OR
        as a list of values as from scattertoarray() if you pass getList as True.  Note that the items in the
        returned list() will be either in the order the fields appear in the table or, if fieldList was passed,
        in the order found in the fieldList string.

        Call as:
            for xx in xcbTools.scan(indexTag="MyTag", forExpr="FIRSTNAME='John'")
                print xx["FIRSTNAME"], xx["LASTNAME"]

        Returns None on error and when done.  Note that if forExpr is provided (non blank) the iteration
        will stop when there are no more records matching the forExpr.  If the indexTag value (or the currently
        selected index tag) has an expression exactly matching the left hand side of the forExpr, then the iteration
        may stop before an actual EOF() condition is reached.  Otherwise when the iteration is completed, the table
        record pointer will be positioned at EOF().

        noData parameter:  This returns a record number rather than the scattered contents of the entire record.
        Use this if you just want a few fields and don't want to waste the time for parsing the full record.

        Support for temporary indexes:  If you want to scan in order by a temporary index, create the index
        and then when scanning pass "" or omit the indexTag parameter.  Your temporary index ordering will apply
        to the scan.  Specifying a different tag in the scan will substitute that ordering even if one or more
        temporary indexes exist.

        The  parameter fieldList is an optional comma-delimited list of field names.  If it is non-blank and if
        noData is False, then the returned dict() or list() ONLY has the fields listed in the fieldList.  This allows fast
        in-dll retrieval of field values even when you only want a few of the many fields in the table.

        By default, this function starts the scan at the beginning of the table, based on the currently set or
        specified index tag (or by recno, if no tag is active).  Pass bNoTop=True to override this behavior, and to
        start from where the record pointer is currently positioned.  If bDescending is True, then by default goes
        to the bottom of the table.  bNoTop is respected by not going to the bottom first in this case.

        If bDescending is True, then the scan moves backward through the table in the specified order.

        NOTE: Nested scans are NOT supported.  seek() is supported into another table in the middle of the scan.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        if indexTag:
            if not self.setorderto(indexTag):
                # Bad news.  Out of here
                return
        if fieldList is None:
            fieldList = ""
        else:
            if (fieldList != "") and (noData is True):
                raise ValueError("fieldList must be None if noData is True")
        # We have a good indexTag value set...
        lbFilterActive = False
        nSkipper = 1
        if bDescending:
            nSkipper = -1

        if forExpr:  # corrected suggested by Jerry. 06/20/2012. JSH.
            if self.cbt.preparefilter(forExpr) == 0:
                self.cErrorMessage = self.cbt.geterrormessage()
                self.nErrorNumber = self.cbt.geterrornumber()
                return  # Nothing we can do, illegal filter
            lbFilterActive = True
        if not bNoTop:
            if not bDescending:
                self.goto("TOP")
            else:
                self.goto("BOTTOM")
        lcCurrentTable = self.alias()
        oCBT = self.cbt
        if lbFilterActive:
            if noData:
                while not oCBT.eof() and not oCBT.bof():
                    if oCBT.testfilter():
                        lnRecno = oCBT.recno()
                        if lnRecno == -1:
                            break
                        yield lnRecno
                    if oCBT.alias() != lcCurrentTable:
                        oCBT.select(lcCurrentTable)
                    oCBT.skip(nSkipper)
            else:
                while not oCBT.eof() and not oCBT.bof():
                    if oCBT.testfilter():
                        lxDict = oCBT.scatter("", True, stripblanks, fieldList, getList, coding)
                        if lxDict is None:
                            if self.cbt.geterrornumber() != 0:
                                self.cErrorMessage = self.cbt.geterrormessage()
                                self.nErrorNumber = self.cbt.geterrornumber()
                                raise ValueError("Scan failed with err: " + str(self.cErrorMessage))
                            break
                        yield lxDict
                    if oCBT.alias() != lcCurrentTable:
                        oCBT.select(lcCurrentTable)
                    oCBT.skip(nSkipper)
        else:
            if noData:
                while not oCBT.eof() and not oCBT.bof():
                    lnRecno = oCBT.recno()
                    if lnRecno == -1:
                        break
                    yield lnRecno
                    if oCBT.alias() != lcCurrentTable:
                        oCBT.select(lcCurrentTable)
                    oCBT.skip(nSkipper)
            else:
                while not oCBT.eof() and not oCBT.bof():
                    lxDict = oCBT.scatter("", True, stripblanks, fieldList, getList, coding)
                    if lxDict is None:
                        if self.cbt.geterrornumber() != 0:
                            self.cErrorMessage = self.cbt.geterrormessage()
                            self.nErrorNumber = self.cbt.geterrornumber()
                            raise ValueError("Scan failed with err: " + str(self.cErrorMessage))
                        break
                    yield lxDict
                    if oCBT.alias() != lcCurrentTable:
                        bTest = oCBT.select(lcCurrentTable)
                    oCBT.skip(nSkipper)
        return

    def zap(self):
        """
        Removes all records from the currently selected table.  This is more drastic than deleting all records
        because the table is actually reconstructed from scratch with no records.  All existing records will be
        lost.  USE WITH CARE.  There is NO WARNING MESSAGE.

        Use on a table that has not been opened in exclusive mode is extremely risky.

        Returns True on success, False on failure.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lbReturn = self.cbt.zap()
        if not lbReturn:
            self.cErrorMessage = self.cbt.geterrormessage()
            self.nErrorNumber = self.cbt.geterrornumber()
        return lbReturn

    def xmltocursor(self, cAlias="", cXML="", cFileName=""):
        """
        In VFP the XMLTOCURSOR() function is the converse of the CURSORTOXML() function.  It makes possible the loading
        of XML created by CURSORTOXML() back into a cursor.  This replicates a portion of the functionality of the
        VFP function.  In this case, the xml data will be stored back into a table (appending the records) which
        has some or all of the fields specified in the xml data.  It will be up to you to make sure that if the
        table has primary keys, which must be set externally to unique values, that those will be handled somehow.
        No capability is provided to create a table into which to load the XML (as is provided for in VFP), so you
        must first create a table to receive the XML records, and have field names which correspond.

        The primary tag name for the XML data is typically <VFPData> but may be something else.  Tag names for
        individual records may be any value.  Tag names for fields must match the spelling (but not the case) of fields
        in the target table, or the data element will be ignored.

        :param cAlias: Alias of an open table.  If blank, the currently selected table will be used.
        :param cXML: Text string with the XML content to be loaded.  Leave "" if loading from a file.
        :param cFileName: File name from which to load the XML.  Leave "" if cXML is passed.
        :return:  Number of records loaded.  0 indicates an error.
        """
        nRecs = 0
        cOldAlias = self.alias()
        cTopName = ""
        bOK = True
        cSource = cXML
        bGoodRead = True
        oTopNode = None
        xTypes = None
        cOldError = ""
        nOldErrNum = 0
        if not cSource:
            try:
                xFile = open(cFileName, "rb")
                cSource = xFile.read()
                xFile.close()
            except:
                bGoodRead = False
        if not bGoodRead:
            self.cErrorMessage = "Unable to read XML File"
            self.nErrorNumber = -9235
            return 0
        if cSource:
            if cAlias:
                bGoodRead = self.select(cAlias)
            if bGoodRead:
                xTypes = self.afieldtypes()

                cTest = cSource[0:5].upper()
                if cTest != "<?XML":
                    self.cErrorMessage = "not XML data"
                else:
                    try:
                        oDom = minidom.parseString(cSource)
                        oNodes = oDom.childNodes
                        oTopNode = oNodes[0]
                    except:
                        bGoodRead = False # we don't really care why.
                        self.cErrorMessage = "Mal-Formed XML: can't parse XML"
                    if bGoodRead:
                        # have at least one top level node...
                        for oRec in oTopNode.childNodes:
                            if oRec.nodeType == oRec.ELEMENT_NODE:
                                xData = dict()
                                xFlds = oRec.childNodes
                                for xFld in xFlds:
                                    if xFld.nodeType == oRec.ELEMENT_NODE:
                                        cName = xFld.tagName.upper()
                                        oValue = xFld.firstChild
                                        if oValue is not None:
                                            cValue = oValue.nodeValue
                                        else:
                                            cValue = ""
                                        xData[cName.upper()] = self.getValueFromXMLString(cValue, xTypes.get(cName, ""))

                                if len(xData) > 0:
                                    # we have some data in the dict...
                                    bOK = self.insertintotable(xData)
                                    if bOK:
                                        nRecs += 1
                                    else:
                                        nRecs = 0
                                        break

        cOldError = self.cErrorMessage
        nOldErrNum = self.nErrorNumber
        self.select(cOldAlias) # This clears the error, so if we have one we have to store it
        if nOldErrNum != 0:
            self.cErrorMessage = cOldError
            self.nErrorNumber = nOldErrNum
        return nRecs

    def cursortoxml(self, cAlias="", cForExpr="", cFields=None, bStripBlanks=False, cFileName=""):
        """
        In VFP the CURSORTOXML() function is a powerful means of copying all or part of a cursor or table into
        XML form in several different ways.  This method provides a subset of that functionality, creating a file
        that is identical to the default VFP output where each record is has a tag, and each field in that record has
        its tag and the corresponding value contained in the tag.  The top-level tag name is always <VFPData>.  All
        field and cursor record tags are lower case.  Outputs fields in each record in the order in which they
        appear in the table or in the order in the cFields list. See also xmltocursor().

        :param cAlias: Pass the alias name if not the currently selected table.
        :param cForExpr: Pass a valid logical expression to select records.  If omitted, all records will be output
        :param cFields: Pass a comma delimited list of field names to limit the output to selected fields.
        :param bStripBlanks: By default, trailing blanks on text fields are left in place.  Set this true to trim them
        :param cFileName: Pass a file name to output the result directly into that text file.  If omitted, then
            the method will return the resulting XML as a string
        :return: The XML text as a string unless the cFileName parm is not empty.  If cFileName is "", if an error
            occurs, the function will return "".  If cFileName is a file name, returns True on success, False on
            failure, and sets cErrorMessage.

        NOTE: The XML text is stored in memory until output at the end.  Very large tables may produce memory
        issues as a result, especially in 32-bit versions of Python.
        """
        cOldAlias = self.alias()
        bOK = True
        cXML = ""
        bFileOutput = (cFileName != "")
        cPreserve = ("" if bStripBlanks else ' xml:space="preserve"')
        if cAlias:
            bOK = self.select(cAlias)
        if bOK:
            cXML = '<?xml version = "1.0" encoding="Windows-1252" standalone="yes"?>\r\n<VFPData%s>\r\n' % (cPreserve,)
            cRecName = ""
            for jj in range(0, 8):
                cRecName += random.choice("abcdefghijklmnopqrstuvwxyz")
            if cFields is not None:
                xFields = cFields.split(",")
            else:
                xFields = list()
                xFld = self.afields()
                for xF in xFld:
                    xFields.append(xF.cName.upper()) # official order of the fields in the table.
            xTypes = self.afieldtypes()
            cRecTagStart = "\t<%s>\r\n" % (cRecName,)
            cRecTagEnd = "\r\n\t</%s>" % (cRecName,)
            xRecXML = list()
            for xRec in self.scan(forExpr=cForExpr, fieldList=cFields):
                xTemp = list()
                for cF in xFields:  # List of field names in the required order
                    cType = xTypes[cF]
                    xValue = xRec[cF]
                    cF = cF.lower()
                    xTemp.append('\t\t<%s>%s</%s>' % (cF, self.makeXMLValueString(xValue, cType), cF))
                xRecXML.append(cRecTagStart + "\r\n".join(xTemp) + cRecTagEnd)
            cXML = cXML + "\r\n".join(xRecXML) + "\r\n</VFPData>\r\n"

        self.select(cOldAlias)
        if bFileOutput:
            try:
                xFile = open(cFileName, "wb")
                xFile.write(cXML)
                xFile.close()
            except:
                bOK = False
                self.cErrorMessage = "Writing XML to file %s failed" % (cFileName,)

            xRet = bOK
        else:
            if bOK:
                xRet = cXML
            else:
                xRet = ""
        return xRet

    def getValueFromXMLString(self, cValue, cType):
        """
        Utility function that takes an XML value in standard VFP output format, and returns the Python native type
        corresponding to the field type specified in the cType parameter
        :param cValue:
        :param cType:
        :return: Native Python value or None on error.
        """
        xRet = None
        if cType:
            if cValue is not None:
                if isinstance(cValue, unicode):
                    cValue = cValue.encode("cp1252", "ignore")
                cValue = self._removenonprint(cValue)
                if cType == "C":
                    xRet = str(cValue) # whatever it is.
                elif cType in "NFYB":
                    try:
                        xRet = float(cValue)
                    except:
                        xRet = 0.0
                elif cType == "I":
                    try:
                        nTemp = float(cValue)
                        xRet = int(nTemp)
                    except:
                        xRet = 0
                elif cType == "D":
                    # We presume here that the text is of the form YYYY-MM-DD.  We'll allow for "." and "/" and "_"
                    # as the delimiters, and that is all.
                    try:
                        cTest = cValue.replace("_", "-")
                        cTest = cTest.replace("/", "-")
                        cTest = cTest.replace(".", "-")
                        xParts = cTest.split("-")
                        nYear = int(xParts[0])
                        nMonth = int(xParts[1])
                        nDay = int(xParts[2])
                        xRet = date(nYear, nMonth, nDay)
                    except:
                        xRet = None

                elif cType == "L":
                    cTest = cValue[0:1].upper()
                    if not cTest:
                        xRet = False
                    else:
                        if cTest in "TY1":
                            xRet = True
                        else:
                            xRet = False
                elif cType == "T":
                    # The presumption is that the date/time value is in standard VFP/XML format YYYY-MM-DDTHH:MM:SS
                    # but with date delimiters of /, _ and . permitted.
                    try:
                        cTest = cValue.replace("_", "-")
                        cTest = cTest.upper().replace("/", "-")
                        cTest = cTest.replace(".", "-")
                        xParts1 = cTest.split("T")
                        xTime = xParts1[1].split(":")
                        nHr = int(xTime[0])
                        nMin = int(xTime[1])
                        nSec = int(xTime[2])
                        xDate = xParts1[0].split("-")
                        nYear = int(xDate[0])
                        nMonth = int(xDate[1])
                        nDay = int(xDate[2])
                        xRet = datetime(nYear, nMonth, nDay, nHr, nMin, nSec)
                    except:
                        xRet = None
                elif cType == "M":
                    xRet = str(cValue)
        return xRet

    def makeXMLValueString(self, xValue, cType):
        """
        Utility function that transforms a field type into a text value string for cursortoxml() output.
        :param xValue: Any legal python value
        :param cType: The official type of the database table from which it comes.
        :return: Text string of the value to put into the XML block
        """
        cRet = ""
        if xValue and cType:
            if cType == "L":
                if xValue:
                    cRet = "T"
                else:
                    cRet = "F"
            elif cType == "C":
                cRet = str(xValue) # just in case it isn't a string value.
            elif cType in "NIFBY":
                cRet = str(xValue)
            elif cType == "D":
                try:
                    cRet = xValue.strftime("%Y-%m-%d")
                except:
                    cRet = "    -  -  "
            elif cType == "T":
                try:
                    cRet = xValue.strftime("%Y-%m-%dT%H:%M:%S")
                except:
                    cRet = "    -  -  T  :  :  "
            elif cType == "M":
                if xValue:
                    cRet = "<![CDATA[" + str(xValue) + "]]>"
                else:
                    cRet = ""
            else:
                cRet = str(xValue)

        return cRet

    def getNewKey( self, dataDir, filename=None, readOnly=False, stayOpen=False, noFlush=False):
        """
        Copy of GetNewKey from VFP.  Returns the next sequential key number for the passed filename or a negative
        value on error.  M-P System Services, Inc., specific code.

        dataDir: cDataDirectory from configurator.  If NEXTKEY.DBF is not open and this is blank, triggers an
                 exception.  This is where the NEXTKEY.DBF table should be found if the method has to open it.
        filename: name of file to get next key.  Defaults to *current*, but if no current and blank, triggers
                  an exception.
        readOnly: if set, return *previous* key number, write nothing - (for testing) - NOT USED!
        stayOpen: if set, does not close NextKey.dbf when done.  For looping and testing.  For maximum speed when
                  using this function to add many records in a loop, open the NEXTKEY.DBF table prior to starting the
                  loop.  This function will then not try to open it every time it runs.
        noFlush: By default the method flushes the contents of the NEXTKEY.DBF table to disk using the OS flush
                 to disk command to make sure that any other users see the updates immediately. However, if that
                 won't be needed for operational requirements, this function can be speeded up by more than a
                 factor of 10 by passing noFlush=True, to bypass this flushing action.  In that case, you may still
                 wish to issue a flush() function at the end of your looping update process.

        Placed here because it is dbf file related but could be useful for many applications.

        NOTE: CodeBase claims to support auto-generated key fields, but that functionality doesn't appear to be
        compatible with the corresponding feature in Visual FoxPro.  This function helps get around that limitation
        by using a NextKey.dbf table to store the "next key" value of each table in your system.  The NextKey.dbf
        table is expected to have two fields:
        Field  Field Name      Type                Width    Dec   Index   Collate Nulls    Next    Step
        1  TABLE_NAME      Character              15            Asc   Machine    No
        2  NEXTKEYVAL      Integer                 4                             No

        and to have an index called TABLE_NAME on UPPER(TABLE_NAME).  The dataDir parameter tells the function
        where to find the NextKey.dbf table.  This function uses record locking to ensure that no keys are duplicated
        within the same table.

        It is the responsibility of the programmer to know which field is the key field in the table and to do
        the replace of the value generated by this function into the appropriate field in the table.
        """
        global gbIsRandomized
        global gnNextKeyTableNameLength
        nReturn = 0
        error = self.nErrorNumber = 0
        errormessage = self.cErrorMessage = ""
        curAliasName = self.cbt.alias() # save current selected table
        if (filename is None) or (filename == ""):
            cTest = self.cbt.dbf(curAliasName)
            cPath, cFName = os.path.split(cTest)
            cBase, cExt = os.path.splitext(cFName)
            filename = cBase.upper()
            if not filename:
                raise ValueError("No table specified for the New Key")
        oCBT = self.cbt
        if not oCBT.used("NEXTKEY"):
            if not dataDir:
                raise ValueError("Must specify a data directory location for NextKey table")
            cNKfilename = os.path.join(dataDir, 'NextKey.dbf')
            bOK = oCBT.use(cNKfilename, "NEXTKEY", False, True)
            if not bOK:
                sleep(0.2)  # Added this extra try. 01/06/2015. JSH.
                bOK = oCBT.use(cNKfilename, "NEXTKEY", False, True)
        else:
            bOK = oCBT.select("NEXTKEY")
            stayOpen = True # we won't close it if already open.
        if bOK:
            if not gbIsRandomized:
                random.seed(mTools.getrandomseed())
                # This makes sure we have a truly random set of values
                # without any dependence on the system clock for generating the seed. added 02/21/2015. JSH.
                gbIsRandomized = True
            # Added code to make delay between lock attempts random so we don't have multiple instances in lock step
            # with each other.  02/21/2015. JSH.
            if gnNextKeyTableNameLength == 0:
                gnNextKeyTableNameLength = len(self.curvalstr("TABLE_NAME"))
            cFilename = ("%-" + str(gnNextKeyTableNameLength) + "s") % (filename.upper().strip(),)
            cTag = 'TABLE_NAME'
            # if cTag not in [x.cTagName for x in self.ataginfo()]:
            #    # needs this index.  Didn't find it.  Create it.
            #    oCBT.indexon(cTag, cTag)
            if oCBT.seek(cFilename, "NEXTKEY", cTag):
                lbLocked = oCBT.rlock()
                rtcnt = 0
                while (not lbLocked) and (rtcnt < 20):  # Allows roughly 2 full seconds
                    rtcnt += 1
                    nSleepTime = 0.055 + (random.random() / 6.0)
                    sleep(nSleepTime)
                    lbLocked = oCBT.rlock()
                if lbLocked:
                    oCBT.refreshrecord()
                    nReturn = oCBT.scatterfieldlong("NEXTKEYVAL")
                    bOK = oCBT.replacelong("NEXTKEYVAL", nReturn + 1)
                    if not bOK:
                        nReturn = -1
                        self.cErrorMessage = "Unable to update NextKey table with updated value"
                        self.nErrorNumber = -4
                    if (noFlush is None) or (not noFlush):
                        oCBT.flush()
                    oCBT.unlock()
                else: # Added 10/06/2012. JSH.
                    nReturn = -1
                    self.cErrorMessage = "Unable to LOCK NextKey: " + oCBT.geterrormessage()
                    self.nErrorNumber = -3
            else:
                # add new counter
                nReturn = 11000
                if not readOnly:
                    xRec = dict()
                    xRec["TABLE_NAME"] = filename
                    xRec["NEXTKEYVAL"] = nReturn + 1
                    self.insertintotable(xRec)
            if not stayOpen:
                oCBT.closetable("NEXTKEY")
        else:
            self.cErrorMessage = "Unable to access NEXT KEY table because %s" % oCBT.geterrormessage()
            self.nErrorNumber = -2
            nReturn = -3
        if self.nErrorNumber:
            error = self.nErrorNumber
            errormessage = self.cErrorMessage
        self.select(curAliasName)  # reset current file.  If stayOpen and no current file, NextKey will be current
        if error:
            self.nErrorNumber = error  # select() might reset it
            self.cErrorMessage = errormessage
        return nReturn

    def setstrictaliasmode(self, bMode=False):
        """
        Obsolete, but here because some old code may make this call
        """
        return True

    def openTableList(self, aTables, cPath=""):
        """
        Pass this a list of values each consisting EITHER of a string containing a comma or space delimited
        list of table names OR a list of tuples consisting of (tablename, readonly, alias).  Then the method
        will attempt to open all tables, set each to the specified readonly status, and assign the
        specified alias.  All three elements of each tuple must be passed, but they may be passed as None
        or "".  In that case the default readonly will be False, and the normal alias will be assigned.  The
        same defaults will apply if the aTables value is a string of table names.

        If one of the tables is to be opened exclusively, the readonly element of the tuple must be passed as
        the text string "EXCL".

        Returns a list of the aliases, in the same order as the tables originally passed.  Returns None on
        error, in which case you can inspect the cErrorMessage property for the reason why.  If None is returned,
        NONE of the tables will be opened.  Any opened partway through the process will be closed before the return.

        If cPath is non-empty, it will be prepended to each of the table names to make the full path name of
        the table to be opened.  Otherwise you'll need to put the fully qualified path name of each table
        in the list of tuples.
        """
        if not aTables:
            raise ValueError("aTables MUST be a list of names or tuples with names, aliases, and RO flags.")
        if mTools.isstr(aTables):
            xTables = aTables.replace(",", " ").split()
            laList = list()
            for ix in xTables:
                laList.append((ix, False, ""))
        else:
            laList = copy.copy(aTables)
        lxReturn = list()
        for ix in laList:
            lbRO = False
            lbExcl = False
            lcName = ix[0]
            lcAlias = ix[2]
            lxFlag = ix[1]
            if lxFlag is None:
                lbRO = False
            elif mTools.isstr(lxFlag):
                if lxFlag.upper() == "EXCL":
                    lbExcl = True
            elif isinstance(lxFlag, bool):
                lbRO = lxFlag
            if cPath:
                lcName = os.path.join(cPath, lcName)
            if not lcAlias:
                lcDir, lcFName = os.path.split(lcName)
                lcBase, lcExt = os.path.splitext(lcFName)
                lcAlias = lcBase.upper()  # The default as UPPER()
            lbOK = self.use(lcName, alias=lcAlias, readOnly=lbRO, exclusive=lbExcl)
            if lbOK:
                lxReturn.append(self.alias())
            else:
                cTempMessage = self.cErrorMessage
                nTempNumber = self.nErrorNumber
                for xt in lxReturn:
                    self.closetable(xt)  # Clears the error messasges.
                lxReturn = None
                self.cErrorMessage = cTempMessage + ": " + lcName
                self.nErrorNumber = nTempNumber
                break
        if not (lxReturn is None):
            if len(lxReturn) == 0:
                self.cErrorMessage = "No table names found"
                lxReturn = None
        else:
            self.cErrorMessage = "openTableList Failed because: %s, %d" % (self.cErrorMessage, self.nErrorNumber)
        return lxReturn

_cbToolsobj = _cbTools()    # always in use
_cbToolsLargeobj = None     # only instantiated if needed


def cbTools(bIsLarge=False):
    global _cbToolsLargeobj
    if bIsLarge:
        if _cbToolsLargeobj is None:
            _cbToolsLargeobj = _cbTools(True)
        return _cbToolsLargeobj
    else:
        return _cbToolsobj


def cbToolsX(bIsLarge=False):
    """
    This version of the cbTools caller creates a brand new instance every time.  This is preferred for new work
    because it provides for a unique session number for the instance to reduce interference between instances
    in opening and closing tables.  Should work fine with TableObj type routines as well as legacy methods.
    Added Dec. 2, 2014. JSH.
    :param bIsLarge:
    :return: The object reference to a new instance of cbTools()
    NOTE: While Codebase does assign each instance a unique data session number, the opening and closing of
    tables is not really isolated, despite what their docs claim. Oct. 7, 2015. JSH.
    """
    if bIsLarge:
        return _cbTools(True)
    else:
        return _cbTools(False)


class TableObj(object):
    """Treat vfp table access as if separate objects for each table.
    'reset' keyword will close/reopen file to flush file buffer  (still requires full path in tablename)
    opens with 'setdeleted' True, so deleted records are ignored by default
    """

    def __init__(self, vfp, tablename, alias="", readOnly=False, exclusive=False, reset=False):
        """
        Create database object from cbTools connection object.  Ignores deleted record.
        :tablename: can include full path
        optional :reset: will cause file to be re-opened if already open
        """
        self._fields = dict()
        self.__dict__.update(  # required because of def __setattr__
            open=True,
            error=0,
            errormessage='',
            name='',
            _readOnly=readOnly,
            _exclusive=exclusive,
            nErrorNumber=0,
            cErrorMessage='',
        )
        self._vfp = vfp
        self.name = alias or os.path.splitext( os.path.split(tablename)[1])[0]
        self.name = self.name.upper()
        if not vfp.select(self.name) or (reset and (vfp.closetable(self.name) or True)):
            found = vfp.use(tablename, alias, readOnly, exclusive)
            if found:
                vfp.setdeleted(True)    # set to ignore deleted records by default
            else:
                self.error = vfp.nErrorNumber
                self.errormessage = vfp.cErrorMessage
                self.open = False
                self._vfp = None    # to return an unusable instance
        if self.open:
            self._fields = vfp.afieldtypes()
            if not self.name in self._vfp._tablenames:
                self._vfp._tablenames.append( self.name)

    def __getattr__(self, fieldName):
        """
        allows getting VFP fields AND methods via i.field_or_method_name
        does not change current active file when only accessing field names.
        DOES change current active file when using CodeBaseTools methods
        current selected table is NOT reset on error or method call
        """
        if not self.open:
            raise AttributeError( "Closed file")
        self.cErrorMessage = self._vfp.cErrorMessage    # otherwise accessing any other function might
        self.nErrorNumber = self._vfp.nErrorNumber      # clear these
        curname = self._saveName()
        # in case the field was in the format FILE.FIELD - not fully working
        type = self._fields.get( fieldName.upper().split('.')[-1], '')
        if type:
            result = self._vfp.curval( fieldName, True)
            self._vfp.select(curname)  # empty string means no change
        else:
            try:
                result = getattr( self._vfp, fieldName)
            except AttributeError:
                raise AttributeError('%s is not a valid vfp command or field in "%s"' % (fieldName,self.__dict__['name']))
#        if curname: self._vfp.select(curname)  #cannot reset in the middle of a method
        return result

    def __setattr__(self, fieldName, value):
        """
        allows setting VFP fields with i.field_name = new_value
        does not change current active file unless an error occurs
        """
        if fieldName.startswith("_") or fieldName in vars(self).keys():
            # internal attributes
            super(TableObj,self).__setattr__(fieldName, value)
            return
        if self.open:
            curname = self._saveName()
            # in case the field was in the format FILE.FIELD - not fully working
            type = self._fields.get( fieldName.upper().split('.')[-1], '')
            if type:
                if not self._vfp.replace( fieldName, value):
                    raise AttributeError( 'Cannot set %s to %s' % (fieldName, value))
            else:
                raise AttributeError( '%s is not a valid field in "%s"' % (fieldName,self.__dict__['name']))
            self._vfp.select(curname)

    def __getitem__(self, fieldName):
        """
        allows retrieving VFP fields with i[field_name]
        """
        try:
            it = self.__getattr__(fieldName)
        except AttributeError as msg:
            raise IndexError(msg)
        return it

    def __setitem__(self, fieldName,value):
        """
        allows setting VFP fields with i[field_name] = new_value
        """
        try:
            self.__setattr__(fieldName,value)
        except AttributeError as msg:
            raise IndexError(msg)

    def __repr__(self):
        return 'vfp.%s("%s")' % (self.__class__.__name__, self.name if self.open else "<closed>")

    def __str__(self):
        if self.open:
            ret = 'Table "%s" with %s record at %s"' % (
                self.name,
                self._vfp.reccount(),
                str(os.path.split(self._vfp.dbf())[0])
                )
        else:
            ret = "Closed table. (was %s)" % self.name
        return ret

    def __iter__(self):
        """
        simple iterator to loop through records.  Works with indexes
        Sample 1:  for r in fileInstance:
                     print r.doc_code
        Sample 2:  [r.doc_code for r in fileInstance if r.doc_reqd]
        :return: generator
        """
        self._vfp.goto("TOP")

        def nextr():
            if not self.eof():
                yield self   # first one is already set
            while self.next():
                yield self
        return nextr()

    def _saveName(self):
        """
        If the "current" VFP table isn't the the one we want, switch to the one we want, then return the...
        :return:  string name of the previous "current" VFP table
        """
        curname = self._vfp.alias()
        curname = curname.upper()
        if self.name != curname:
            self._vfp.select(self.name)
        else:
            curname = "" # no need to change back
        return curname

    def close(self):
        """closes table AND renders instance inert.
        does not change current active file
        """
        ret = 0
        self.errormessage = ""
        if self.open:
            curname = self._vfp.alias()
            ret = self._vfp.closetable( self.name)
            self.open = False
            self._vfp.select(curname)
            self._vfp = None
        else:
            self.errormessage = "Cannot close a closed file"
        return ret

    def newId(self):
        """return next unique ID.
        does not change current active file
        NextKey file should be already opened
        """
        res = False
        self.errormessage = ""
        if self.open:
            table = self._vfp['NextKey']
            if table.open:
                # it SHOULD already be open.  Otherwise *try* to open.
                self._vfp.select(self.name)
                res = self._vfp.getNewKey('',stayOpen=True)
                self.errormessage = self._vfp.cErrorMessage
            else:
                self.errormessage = 'NewId() error: NextKey must be already open.'
        else:
            self.errormessage = "Cannot read a closed file"
        return res

    def next(self):
        """next record.
        does not change current active file
        """
        res = False
        self.errormessage = ""
        if self.open:
            curname = self._saveName()
            self._vfp.skip(1)
            res = not self._vfp.eof()
            self._vfp.select(curname)
        else:
            self.errormessage = "Cannot read a closed file"
        return res

    def reset(self):
        """close and re-open file.
        does not change current active file
        """
        res = False
        self.errormessage = ""
        if self.open:
            tablename = self._vfp.dbf()
            curname = self._saveName()
            self._vfp.closetable( self.name)
            self.open = False
            found = self._vfp.use( tablename, self.name, self._readOnly, self._exclusive)
            if found:
                self._vfp.setdeleted(True)
                res = True
                self.open = True
            else:
                self.error = self._vfp.nErrorNumber
                self.errormessage = self._vfp.cErrorMessage
                self._vfp = None
            if curname and self.name != curname:
                self._vfp.select(curname)
        else:
            self.errormessage = "Cannot reset a closed file"
        return res


#####################################################################################
def copydatatable(cTableName="", cSourceDir="", cTargetDir="", bByZap=False, oCBT=None):
    """ Function which copies all files for one DBF type table from one directory to another.
     Takes 5 parameters:
     - Name of the table, no path, extension optional unless NOT .DBF
     - Name of the source directory - fully qualified
     - Name of the target directory - fully qualified
     - By Zap, which causes table content to be removed and replaced
     - oCBT which is a handle to a CodeBaseTools object, if an external one is to be used.

     Returns True on success, False on failure.  If there is an existing table in the
     target directory, it and its related files are renamed to a random name prior
     to the copy process and then copied to an ARCHIVE subdirectory.  If failure,
     the partially copied files are removed and the originals renamed to their starting names.
     On returning False, writes the error into the gcLastErrorMessage global.
     In that case call getlasterrormessage() to retrieve the description of the error.
     As of 12/29/2013, this method will attempt to copy by zapping the contents of the
     target table and appending in the source, if the bByZap parameter is True.  If failing
     that or if bByZap is False, then copies using the operating system copy functions.
     Caller should provide a CodeBaseTools object as oCBT or this program will create its own.
     Note that if the source table is in a VFP database container, and if the function
     uses the OS to copy the file, the resulting copy will have a reference to the database
     container in it, which may not be what you intended.
    """
    global gcLastErrorMessage
    gcLastErrorMessage = ""
    cRoot, cExt = os.path.splitext(cTableName)
    if cExt and cExt.upper() != ".DBF":
        cTable = cTableName
    else:
        cTable = cRoot + ".DBF"
        cExt = ".DBF"
    cSourceDir = cSourceDir.strip()
    cTargetDir = cTargetDir.strip()

    cSrc = os.path.join(cSourceDir, cTable)
    cTarg = os.path.join(cTargetDir, cTable)

    if not os.path.exists(cSrc):
        gcLastErrorMessage = "Source Table does not exist"
        return False
    cSrcCdx = mTools.FORCEEXT(cSrc, "CDX")
    cTargCdx = mTools.FORCEEXT(cTarg, "CDX")
    if not os.path.exists(cSrcCdx):
        cSrcCdx = ""
        cTargCdx = ""
    cSrcFpt = mTools.FORCEEXT(cSrc, "FPT")
    cTargFpt = mTools.FORCEEXT(cTarg, "FPT")
    if not os.path.exists(cSrcFpt):
        cSrcFpt = ""
        cTargFpt = ""

    bRenameRequired = False
    cTempDbf = ""
    cTempCdx = ""
    cTempFpt = ""
    bRenameOK = False
    if os.path.exists(cTarg):
        # We have to rename it to something else...
        # Or copy it to some separate table IF it is to be zapped in place.
        cTimeStamp = mTools.TTOC(datetime.now(), 1)
        bRenameRequired = True
        cTemp = cRoot + "_" + cTimeStamp + "_" + mTools.strRand(4, isPassword=1)
        cTempDbf = cTemp + cExt
        cTempDbf = os.path.join(cTargetDir, cTempDbf)
        if cTargCdx:
            cTempCdx = mTools.FORCEEXT(cTempDbf, "CDX")
        if cTargFpt:
            cTempFpt = mTools.FORCEEXT(cTempDbf, "FPT")

        bRenameOK = True
        xProc = os.rename
        if bByZap:
            xProc = shutil.copy2
        try:
            xProc(cTarg, cTempDbf)
        except:
            bRenameOK = False
        if bRenameOK and cTargCdx:
            try:
                xProc(cTargCdx, cTempCdx)
            except:
                bRenameOK = False
            if not bRenameOK:
                try:
                    xProc(cTempDbf, cTarg)
                except:
                    bRenameOK = False  # Duhhh...
        if bRenameOK and cTargFpt:
            try:
                xProc(cTargFpt, cTempFpt)
            except:
                bRenameOK = False
            if not bRenameOK:
                try:
                    xProc(cTempDbf, cTarg)
                    if cTargCdx:
                        xProc(cTempCdx, cTargCdx)
                except:
                    bRenameOK = False
    if bRenameRequired and not bRenameOK:
        gcLastErrorMessage = "Existing Target Table could not be renamed or archived for backup"
        return False

    bCopyOK = True
    if bByZap:
        if oCBT is None:
            oCBT = cbTools()
        oCBT.setdeleted(True)
        bCopyOK = oCBT.use(cSrc, "THESOURCE", readOnly=True)
        if bCopyOK:
            bCopyOK = oCBT.use(cTarg, "THETARG", exclusive=True)
        if bCopyOK:
            bCopyOK = oCBT.select("THETARG")
        if bCopyOK:
            bCopyOK = oCBT.zap()
        if bCopyOK:
            oCBT.select("THESOURCE")
            for xRec in oCBT.scan(noData=True):
                cData = oCBT.scatterraw(alias="THESOURCE", converttypes=False, stripblanks=False, lcDelimiter="~!~")
                bCopyOK = oCBT.insertintotable(cData, lcAlias="THETARG", lcDelimiter="~!~")
                if not bCopyOK:
                    break
        if not bCopyOK:
            gcLastErrorMessage = "Codebase ERR: " + oCBT.cErrorMessage
        oCBT.closetable("THETARG")
        oCBT.closetable("THESOURCE")
        oCBT = None
    else:
        try:
            shutil.copy2(cSrc, cTarg)
        except:
            bCopyOK = False
            # lcType, lcValue, lxTrace = sys.exc_info()
            # gcLastErrorMessage = "Err: " + str(lcType) + " " + str(lcValue)
            gcLastErrorMessage = "shutil() copy Failure from: " + cSrc + " To " + cTarg

        if bCopyOK and cSrcCdx:
            try:
                shutil.copy2(cSrcCdx, cTargCdx)
            except:
                bCopyOK = False
                # lcType, lcValue, lxTrace = sys.exc_info()
                # gcLastErrorMessage = "Err: " + str(lcType) + " " + str(lcValue)
                gcLastErrorMessage = "shutil() copy Failure from: " + cSrcCdx + " to " + cTargCdx
        if bCopyOK and cSrcFpt:
            try:
                shutil.copy2(cSrcFpt, cTargFpt)
            except:
                bCopyOK = False
                # lcType, lcValue, lxTrace = sys.exc_info()
                # gcLastErrorMessage = "Err: " + str(lcType) + " " + str(lcValue)
                gcLastErrorMessage = "shutil() copy Failure from: " + cSrcFpt + " to " + cTargFpt

    if not bCopyOK:
        mTools.DELETEFILE(cTarg)
        mTools.DELETEFILE(cTargCdx)
        mTools.DELETEFILE(cTargFpt)  # These are resilient and don't fail if the file isn't there.
        if bRenameRequired:
            try:
                os.rename(cTempDbf, cTarg)
                if cTargCdx:
                    os.rename(cTempCdx, cTargCdx)
                if cTargFpt:
                    os.rename(cTempFpt, cTargFpt)
            except:
                bCopyOK = False  # of course
        gcLastErrorMessage = "Copy Failed, Target Table restored: " + gcLastErrorMessage
    else:
        # we try archiving the replaced tables, if any, to an ARCHIVE directory under the specified target...
        if bRenameRequired:
            cArchiveDir = os.path.join(cTargetDir, "ARCHIVE")
            bOkToArchive = True
            if not os.path.exists(cArchiveDir):
                try:
                    os.mkdir(cArchiveDir)
                except:
                    bOkToArchive = False
            cPath, cBaseTempDbf = os.path.split(cTempDbf)
            cBackupDbf = os.path.join(cArchiveDir, cBaseTempDbf)
            try:
                shutil.move(cTempDbf, cBackupDbf)
                if cTargCdx:
                    cBackupCdx = mTools.FORCEEXT(cBackupDbf, "CDX")
                    shutil.move(cTempCdx, cBackupCdx)
                if cTargFpt:
                    cBackupFpt = mTools.FORCEEXT(cBackupDbf, "FPT")
                    shutil.move(cTempFpt, cBackupFpt)
            except:
                gcLastErrorMessage = "Moving Backup Files to Archive Failed"
    return bCopyOK


def cbtWork():
    tmpdir = 'c:\\MPSSPythonScripts\\TestDataOutput'
    from MPSSCommon.MPSSBaseTools import confirmPath, DELETEFILE
    if not confirmPath(tmpdir):
        tmpdir = ''
    # Create Table Example #1
    xFlds = list()
    print("starting cbtWork")
    xFlds.append("firstname,C,25,0,FALSE")
    xFlds.append("lastname,C,25,0,FALSE")
    xFlds.append("city,C,35,0,FALSE")
    xFlds.append("state,C,2,0,FALSE")
    xFlds.append("zipcode,C,10,0,FALSE")
    xFlds.append("hire_date,D,8,0,FALSE")
    xFlds.append("login_date,T,8,0,TRUE")
    xFlds.append("num_kids,I,4,0,FALSE")
    xFlds.append("salary,N,10,2,FALSE")     # Note Number type.  Exact decimals in the table but will be
    #  converted in Python to a double.
    xFlds.append("wageperhr,Y,10,3,FALSE")  # Note currency type with exact decimals
    xFlds.append("ssn,C,10,0,FALSE")
    xFlds.append("married,L,1,0,FALSE")
    xFlds.append("dob,D,8,0,TRUE")
    xFlds.append("comments,M,10,0,FALSE")
    xFlds.append("coverage,B,12,4,FALSE")  # This is a DOUBLE field, which functions as a floating point
    lcFld = "\n".join(xFlds)
    # The init() of the cbTools class initializes the engine for you.
    vfp = cbTools()
    lnStart = time()

    testname = os.path.join(tmpdir, "employee1.dbf")
    DELETEFILE(testname)
    print(testname)
    bResult = vfp.createtable(testname, lcFld)
    lnEnd = time()
    print("Create Time: %d" % (lnEnd - lnStart))
    print("Create Work Result: %s" % str(bResult))
    print("ERR: ", vfp.cErrorMessage)
    cAlias = vfp.alias()
    print("Alias>%s" % cAlias)
    xFlds = vfp.afields()
    print("Fields>%s" % str(xFlds))
    if not bResult:
        print("ERROR", vfp.cErrorMessage)
        exit(1)
    print("Alias Name: %s" % vfp.alias())
    print(vfp.refreshbuffers(), "buffers refreshed")
    vfp.closetable("employee1")
    bexcl = vfp.use(testname, alias="EMPLOYEE1", exclusive=True)
    print("bxl", bexcl)
    # Index new table
    lnStart = time()
    bResult = vfp.indexon("fullname", "UPPER(lastname+firstname)")
    bResult = vfp.indexon("geo", "state+city")
    print("index result", bResult)
    print("indx err:", vfp.cErrorMessage)
    lnEnd = time()
    print("Index time: ", lnEnd - lnStart)

    # Append record from two lists (Value and Field Name)
    xVal = list()
    xFld = list()
    xFld.append("firstname")  # Note that field names may be specified in lower case, but will be UPPER in table.
    xVal.append("George")
    xFld.append("lastname")
    xVal.append("Weibladson")
    xFld.append("city")
    xVal.append("Chicago")
    xFld.append("state")
    xVal.append("IL")
    xFld.append("zipcode")
    xVal.append("60607")
    xFld.append("hire_date")
    xVal.append(date(2009, 6, 12))
    xFld.append("login_date")
    xVal.append(datetime(2011, 1, 24, 9, 30, 00))
    xFld.append("salary")
    xVal.append(decimal.Decimal("35535.50"))
    xFld.append("num_kids")
    xVal.append(2)
    xFld.append("wageperhr")
    xVal.append(23.45)
    xFld.append("ssn")
    xVal.append("123456789")
    xFld.append("married")
    xVal.append(True)
    xFld.append("dob")
    xVal.append(date(1955, 10, 3))
    xFld.append("comments")
    xVal.append("hello to all")
    xFld.append("coverage")
    xVal.append(2345.19003)

    lnStart = time()
    bResult = vfp.appendblank()
    print("Append Result: ", bResult)
    print("append err ", vfp.cErrorMessage)
    bResult = vfp.gatherfromarray(xVal, xFld)
    print("Gather Result: ", bResult)
    lnEnd = time()
    print("Tot Append Time: ", lnEnd - lnStart)
    print("Gather Error: ", vfp.cErrorMessage)

    # Append record from string of values (all fields updated) -- 10X faster
    lnStart = time()
    bResult = vfp.appendblank()
    lcVals = "Pete<~!~>Smith<~!~>Boston<~!~>MA<~!~>03493<~!~>19980224<~!~>20110103140203<~!~>2<~!~>66039.39<~!~>45.37<~!~>987734959<~!~>FALSE<~!~>19591203<~!~>Testing Again<~!~>0.5343"
    bResult = vfp.gatherfromarray(lcVals, None, "<~!~>")
    lnEnd = time()
    print("String Gather Time: ", lnEnd - lnStart)
    print("Gather From String Result: ", bResult)
    # Append record from a dictionary
    lxD = dict()
    lnStart = time()
    lxD["firstname"] = "William"
    lxD["lastname"] = "Jones"
    lxD["city"] = "Los Angeles"
    lxD["state"] = "CA"
    lxD["zipcode"] = "90272"
    lxD["hire_date"] = date(2005, 7, 29)
    lxD["login_date"] = datetime(2011, 2, 12, 14, 34, 21)
    lxD["salary"] = 42000.00
    lxD["num_kids"] = 2
    lxD["wageperhr"] = 22.35
    lxD["ssn"] = "292929292"
    lxD["married"] = False
    lxD["dob"] = None
    lxD["comments"] = ""
    lxD["coverage"] = 0.35239
    bResult = vfp.appendblank()
    bResult = vfp.gatherfromarray(lxD)
    lnEnd = time()
    print("gather from dict result: ", bResult)
    print("gather from dict time: ", lnEnd - lnStart)

    # Go back to record 1 and get several field values
    bResult = vfp.goto("RECORD", 1)
    print("record 1?", vfp.recno(), "count", vfp.reccount())
    lxDict = vfp.scatter()
    if lxDict is None:
        print(vfp.cErrorMessage, 1)
    print(lxDict["FIRSTNAME"])  # Note upper case field names are REQUIRED.
    print(lxDict["LASTNAME"])
    print(lxDict["LOGIN_DATE"])  # Note conversion to Python datetime value
    print(lxDict["SALARY"])  # Converted to a double
    print(lxDict["WAGEPERHR"])  # Converted to currency value with exact decimals
    print("NAME", vfp.curvalstr("EMPLOYEE1.FIRSTNAME"), "ENDNAME")
    print("NUMKIDS", vfp.curvallong("EMPLOYEE1.NUM_KIDS"))
    #
    bResult = vfp.goto("RECORD", 2)
    vfp.replace("employee1.firstname", "Charles")
    vfp.replace("employee1.wageperhr", 37.50)
    vfp.replace("employee1.dob", date(1945, 6, 13))
    vfp.replace("employee1.married", True)
    lxDict = vfp.scatter()
    if lxDict is None:
        print(vfp.cErrorMessage, 2)
    print(lxDict["FIRSTNAME"])
    print(lxDict["DOB"])
    print(lxDict["MARRIED"])
    print("READY TO SCAtteR")
    bResult = vfp.goto("RECORD", 1)
    print("lbResult", bResult)
    lxDict = vfp.scatter(converttypes=False, fieldList="firstname,LASTNAME,SALARY")
    print(lxDict)
    print("that was dict")
    print("READY TO SCATTER RECORD")
    lxRec = vfp.scattertorecord()
    print(lxRec)
    print("FIRSTNAME:", lxRec.FIRSTNAME)
    print("LASTNAME:", lxRec.LASTNAME)
    cXML = vfp.cursortoxml(cFileName="c:\\temp\\myxml.xml")
    print(vfp.dbf())
    vfp.select("employee1")
    vfp.xmltocursor(cFileName="c:\\temp\\myxml.xml")
    xArray = vfp.copytoarray()
    for xA in xArray:
        print(xA)
    vfp.closedatabases()
    return True


def cbt_test():
    cDest = r"c:\temp\smsrc5.dbf"
    vfp = cbTools()
    vfp.use("e:\\loadbuilder2\\appdbfs\\73rdstrt\\shipmstr.dbf", alias="SHIPMSTR")
    nStart = time()
    nResult = vfp.copyto(cAlias="SHIPMSTR", cOutput=cDest, cType="DBF", cTestExpr='', bHeader=True, bStripBlanks=True)
    nEnd = time()
    vfp.closetable("SHIPMSTR")
    print("ELAPSED TIME: ", (nEnd - nStart))
    if nResult < 0:
        print("THE ERROR:", vfp.cErrorMessage)
        print("THE ERR NUMBER:", vfp.nErrorNumber)
    print("NRESULT: ", nResult)
    # return True
    tmpdir = 'c:\\MPSSPythonScripts\\TestDataOutput'
    from MPSSCommon.MPSSBaseTools import confirmPath
    if not confirmPath(tmpdir):
        tmpdir = ''
    #  Create Table Example #1
    lcFld = "firstname,C,25,0,FALSE\n"
    lcFld += "lastname,C,25,0,FALSE\n"
    lcFld += "city,C,35,0,FALSE\n"
    lcFld += "state,C,2,0,FALSE\n"
    lcFld += "zipcode,C,10,0,FALSE\n"
    lcFld += "hire_date,D,8,0,FALSE\n"
    lcFld += "login_date,T,8,0,TRUE\n"
    lcFld += "num_kids,I,4,0,FALSE\n"
    lcFld += "salary,N,10,2,FALSE\n"   # Note Number type.  Exact decimals in the table but will be
    # converted in Python to a double.
    lcFld += "wageperhr,Y,10,3,FALSE\n"  # Note currency type with exact decimals
    lcFld += "ssn,C,10,0,FALSE\n"
    lcFld += "married,L,1,0,FALSE\n"
    lcFld += "dob,D,8,0,TRUE\n"
    lcFld += "comments,M,10,0,FALSE\n"
    lcFld += "coverage,B,12,4,FALSE\n"  # This is a DOUBLE field, which functions a floating point
    # The init() of the cbTools class initializes the engine for you.
    vfp = cbTools()
    lnStart = time()
    testname = os.path.join(tmpdir, "employee1.dbf")
    print(testname)
    bResult = vfp.createtable(testname, lcFld)
    lnEnd = time()
    print("Create Time: ", lnEnd - lnStart)
    print("Create Result: ", bResult)
    if not bResult:
        print("ERROR", vfp.cErrorMessage)
        exit(1)
    print("Alias Name X: ", vfp.alias())
    zFlds = vfp.afields()
    print(zFlds)
    print("zflds above")

    # Append record from a dictionary
    lxD = dict()
    lnStart = time()
    lxD["firstname"] = "William"
    lxD["lastname"] = "Jones"
    lxD["city"] = "Los Angeles"
    lxD["state"] = "CA"
    lxD["zipcode"] = "90272"
    lxD["hire_date"] = date(2005, 7, 29)
    lxD["login_date"] = datetime(2011, 2, 12, 14, 34, 21)
    lxD["salary"] = 42000.00
    lxD["num_kids"] = 2
    lxD["wageperhr"] = 22.35
    lxD["ssn"] = "292929292"
    lxD["married"] = False
    lxD["dob"] = None
    lxD["comments"] = ""
    lxD["coverage"] = 0.35239
    bResult = vfp.appendblank()
    bResult = vfp.gatherfromarray(lxD)
    lnEnd = time()
    print("gather from dict result: ", bResult)
    print("gather from dict time: ", lnEnd - lnStart)
    vfp.closetable(vfp.alias())
    vfp = None
#
#     ## Go back to record 1 and get several field values
#     lbResult = vfp.goto("RECORD", 1)
#     lxDict = vfp.scatter()
#     print lxDict["FIRSTNAME"] ## Note upper case field names are REQUIRED.
#     print lxDict["LASTNAME"]
#     print lxDict["LOGIN_DATE"] ## Note conversion to Python datetime value
#     print lxDict["SALARY"] ## Converted to a double
#     print lxDict["WAGEPERHR"] ## Converted to currency value with exact decimals
#     print "NAME", vfp.curvalstr("EMPLOYEE1.FIRSTNAME"), "ENDNAME"
#     print "NUMKIDS", vfp.curvallong("EMPLOYEE1.NUM_KIDS")
#
#     lbResult = vfp.goto("RECORD", 2)
#     vfp.replace("employee1.firstname", "Charles")
#     vfp.replace("employee1.wageperhr", 37.50)
#     vfp.replace("employee1.dob", date(1945, 6, 13))
#     vfp.replace("employee1.married", True)
#     lxDict = vfp.scatter()
#     print lxDict["FIRSTNAME"]
#     print lxDict["DOB"]
#     print lxDict["MARRIED"]
#
#     lxaData = vfp.copytoarray(fieldtomatch="dob", matchvalue=None, matchtype="<>")
#     print lxaData[0]["FIRSTNAME"], lxaData[0]["LASTNAME"]
#     print lxaData[1]["FIRSTNAME"], lxaData[1]["LASTNAME"]
#
#     lcTableName = vfp.dbf()
#     print "Table Name: ", lcTableName
#     lcTableName = vfp.dbf("NOSUCHTABL")
#     print "Bad Table Name: >" + lcTableName + "<"
#     print "Error Message: " + vfp.cErrorMessage
#     print "Error Number: ", vfp.nErrorNumber
#
#     ## Create a new table as a copy of another by capturing a list of the first
#     ## table's fields and then pass that list to the createtable function.
#     vfp.closedatabases()
#     lnStart = time()
#     lbResult = vfp.use(testname, "emp1")
#     lxTempList, lxTempFields = vfp.fielddicttolist(lxDict, "emp1")
#     print lxTempList
#     print lxTempFields
#     for jj in range(0, 20):
#         vfp.insertintotable(lxTempList, lxTempFields, lcAlias="emp1")
#     print vfp.alias(), vfp.dbf()
#     print vfp.reccount()
#
#     vfp.goto("RECORD", 19)
#     vfp.delete()
#     vfp.pack()
#     print "After pack: ", vfp.reccount()
#     lnCnt = 0
#     lxFlds = vfp.afields()
#
#     vfp.closedatabases()
#     lnStartC = time()
#     assert True == vfp.createtable(os.path.join(tmpdir,"employee2.dbf"), lxFlds)
#     print "### test TableObj iterator ###"
#     cnt=0
#     for this in vfp['employee2']:
#         cnt += 1
#     assert 0 == cnt , "Iterator not working"
#     assert True == vfp.appendblank()
#     cnt=0
#     for this in vfp['employee2']:
#         cnt += 1
#     assert 1 == cnt , "Iterator not working"
#     assert True == vfp.closetable("employee2")
#     assert True == vfp.use(os.path.join(tmpdir,"employee2.dbf"))
#     xflds = vfp.scatter(converttypes=False)
#     print xflds
#     lnEndC = time()
#     print "Create Time 2: ", lnEndC - lnStartC
#     vfp.dispstru()
#     lnEnd = time()
#     print "Open and Copy Time: ", lnEnd - lnStart
#     if (lbResult == False):
#         print vfp.cErrorMessage, vfp.nErrorNumber
#
#     vfp.zap()
#
# #    print "testing zippairs"
# #    lbResult = vfp.use("e:\\loadbuilder2\\appdbfs\\geo.dbf", "geo")
# #    print "use", lbResult
# #    lbResult = vfp.setdeleted(True)
# #    lbFound = vfp.seek("62305", "geo", "postalcode")
# #    print vfp.cErrorMessage
# #    print lbFound
# #    lbResult = vfp.select("geo")
# #    print lbResult
# #    lbResult = vfp.recno()
# #    print lbResult, lbFound
#
#     ## test table from TableObj
#     print 'TableObj from vfp',
#     table = vfp.TableObj( testname)
#     if not table.open: print 'Open error:',table.errormessage
#     print table.name
#     print 'Fields: ',table._fields
#     del table
#     print 'TableObj direct',
#     table = TableObj( vfp, testname)
#     print 'kids =',table.num_kids,  # test getattr
#     print ', salary =',table['salary'],  # test getitem
#     table.num_kids = 1  # test setattr
#     table['salary'] = 20000.5  # test setitem
#     print
#     fields = 'firstname hire_date married num_kids salary login_date coverage wageperhr'.split()
#     fsize = 10
#     ff = '%%-%ss' % fsize
#     print '  ',
#     for n in fields:
#         print ff%n,
#     print
#     while not vfp.eof():
#         print;print '%2d' % vfp.recno(),
#         for n in fields:
#             print ff % str(table[n])[:fsize],
#         table.next()
#     print
#     lastAlias = vfp.alias()
#     print 'open file=',vfp.alias()
#
#     ## test getNewKey()
#     print 'testing getNewKey()'
#     testLocation = 'e:/loadbuilder2/appdbfs'
#     assert vfp.getNewKey(testLocation,'bob',readOnly=True) == 11000  # new file
#     assert vfp.alias() == lastAlias  #doesn't change current file
#     vfp.getNewKey('c:/','xyz', readOnly=True)
#     assert (vfp.cErrorMessage[:14] == "Unable to open"), "ErrorMessage = '%s'. curFile = '%s'" % (
#         vfp.cErrorMessage, vfp.alias()
#         )
#     assert vfp.alias() == lastAlias  #doesn't change current file on error either
#     if vfp.getNewKey(testLocation,'bob',readOnly=True, stayOpen=True) != 11000:
#         # delete item 'bob'
#         vfp.select('NextKey')
#         x = vfp.curval('table_name',True) == 'BOB' and vfp.delete()
#     assert vfp.getNewKey(testLocation,readOnly=True) == 11000  # `testname` is new file also
#     assert vfp.getNewKey(testLocation,'SHIPMSTR',readOnly=True) > int(1e8)  # bigger than 100 million
#     assert vfp.alias() == lastAlias  # make sure we haven't changed the active file
#     assert vfp.getNewKey(testLocation,'bob', stayOpen=True) == 11000  # new file
#     assert vfp.getNewKey(testLocation,'bob', stayOpen=True) == 11001  # added one
#     assert vfp.getNewKey(testLocation,'bob', stayOpen=True) == 11002  # added one
#     assert vfp.select('NextKey') is True
#     assert vfp.curval('table_name',True) == 'BOB' and vfp.delete()
#     assert vfp.getNewKey(testLocation,'bob',readOnly=True) == 11000  # new file
#     assert vfp.getNewKey(testLocation,'bob',readOnly=True) == 11000  # still new file
#     # test rlock() changes
#     vfp.select(lastAlias)
#     assert vfp.rlock(1) is True
#     assert vfp.unlock() is True
#
#     vfp.closedatabases()
#     lnTest = vfp.use(r"e:\loadbuilder2\appdbfs\geo.dbf", "GEO")
#     print "GEO: ", lnTest
#     print vfp.alias()
#     lbFound = vfp.locate("ST_PROV = 'OR' .AND. UPPER(CITY) = 'PORTLAND'")
#     print lbFound
#     print vfp.cErrorMessage
#     lcTest = vfp.curval("POSTALCODE")
#     print lcTest
#     lbFound = vfp.locatecontinue()
#     print lbFound
#     lcTest = vfp.curval("POSTALCODE")
#     print lcTest
#     vfp.locateclear()
#
#     vfp.setorderto("POSTALCODE")
#     print "IN CHICAGO"
#     lbFound = vfp.locate("ST_PROV = 'IL' .AND. CITY = 'CHICAGO'")
#     while lbFound:
#         print vfp.curvalstr("POSTALCODE")
#         lbFound = vfp.locatecontinue()
#     vfp.locateclear()
#
#     ## Shut down the DBF engine.
#     vfp.closedatabases()
#     vfp.cb_shutdown()
#     return lbResult


if __name__ == "__main__":
    print("***** Testing CodeBaseTools.py components")
    cbt_test()
    print("************ cbt test done **************")
    lbResult = cbtWork()
    print(lbResult)
    print("***** Testing Complete")
    if 'stop' in sys.argv:
        raw_input('press <enter>')

    
        
