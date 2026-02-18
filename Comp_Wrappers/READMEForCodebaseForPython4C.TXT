CodeBase Extension Module for Python
(c) Copyright Sequiter Software Inc., 2000.  All rights reserved.

There is no technical support available for this product.
However, if you discover bugs or have suggestions, you may send mail to
general@sequiter.com.

The CodeBase Extension Module is built as a shared library called codebase
(i.e. codebase.dll on Win32).
To build it for Win32 with Visual C++ 6, use the following command line:
cl -WX -LD Python4.c -Ic:\Python152\include -I[codebasesource] -ocodebase.dll python15.lib c4dll.lib
To install the extension module, place it in the "DLLs" directory of your
Python installation.

In your Python program, import the CodeBase functions with this statement:
from codebase import *

The CodeBase Extension Module for Python contains the following functions
which correspond directly to the CodeBase Basic functions of the same name.
For documentation, see the CodeBase Visual Basic Reference Guide.
   code4accessMode
   code4autoOpen
   code4close
   code4codePage
   code4collate
   code4collateUnicode
   code4compatibility
   code4createTemp
   code4dateFormat
   code4errDefaultUnique
   code4indexExtension
   code4init
   code4initUndo
   code4readLock
   code4readOnly
   code4safety
   code4singleOpen
   d4append
   d4appendStart  (second parameter is optional - default: 0)
   d4blank
   d4bof
   d4bottom
   d4close
   d4create  (fourth parameter is optional - default: no index)
   (see below for parameter info)
   d4delete
   d4deleted
   d4eof
   d4field
   d4fieldJ
   d4flush
   d4go
   d4memoCompress
   d4numFields
   d4open
   d4pack
   d4position
   d4recall
   d4recCount
   d4recNo
   d4recWidth
   d4refreshRecord
   d4reindex
   d4remove
   d4seek
   d4seekDouble
   d4seekNext
   d4seekNextDouble
   d4skip  (second parameter is optional - default: 1)
   d4tag
   d4tagSelect  (second parameter is optional - default: null)
   d4top
   d4zap  (second and third parameters are optional - default: 1 and (last record))
   f4assign
   f4assignChar
   f4assignDouble
   f4assignInt
   f4assignLong
   f4blank
   f4char
   f4decimals
   f4double
   f4int
   f4len
   f4long
   f4name
   f4str
   f4true
   f4type
   i4create  (see below for parameter info)
   i4open

The field definition parameter for d4create is a list of tuples. Each list
item defines one field and is in this format:
name, type [, length, decimals, nulls]
name is the name of the field.
type is a 1-character string which identifies the field type (e.g. "D" =
Date, "C" = Character).
length is the length of the field in bytes (optional)
decimals if the number of decimals in a Numeric or Float field (optional)
nulls is a true/false flag which determines if the field can contain null
(optional) (not supported yet)

The tag definition parameter for d4create and i4create is a list of tuples,
each in this format: name, expression [, filter, unique, descending]
name is the name of the tag
expression is the dBASE expression that defines how the tag will be sorted
filter is the logical dBASE expression that defines which records will be
included in the tag (optional)
unique an integer which is the unique setting for the tag (optional)
descending is the descending setting for the tag (optional)

Example:
   from codebase import *

   cb = code4init()
   safety(cb,0)

   fields = [("NAME","C","20"),("BIRTHDAY","D"),("INCOME","N",9,2)]
   tags = [("NAME","NAME"),("INCOME","INCOME","INCOME < 100000")]
   data = d4create(cb,"Employee",fields,tags)
   fName = d4field(data,"NAME")
   fBDay = d4field(data,"BIRTHDAY")
   fIncome = d4field(data,"INCOME")

   d4appendStart(data)
   f4assign(fName,"John Duncan")
   f4assign(fBDay,"19690225")
   f4assignLong(fIncome,56000)
   d4append(data)

   d4close(data)
   code4initUndo(cb)


The CodeBase Extension Module also contains the following specialized higher
level functions.

Close( data )
Close the table.
data is the table handle returned by Open. data is no longer valid and
should be discarded.
Returns 0 on failure, 1 on success. On failure, the table is still closed.

Deletes( data, column, keys )
Delete one or more records.
data is the table handle returned by Open.
column is the name of the column to search.
keys is one value or an array of values to search for.
The column and keys parameters determine which records should be deleted.
Each key in keys is sought in the column. Each record that matches is
deleted.
Returns a 2-part tuple. The first part is an integer that indicates the
number of records that were successfully deleted. The second part is an
array of integers which indicate which elements in the keys array failed to
be deleted.

Drop( data )
Close the table and delete it from disk.
data is the table handle returned by Open. data is no longer valid and
should be discarded.
Returns 0 on failure, 1 on success. On failure, the table is still closed
and deleted.

Find( data, keys )
Find records and return them in an array of dictionaries.
data is the table handle returned by Open.
keys is a dictionary that defines criteria for the find. In each entry of
the dictionary, the key is the column name and the value is the value to
search in that column. For example:
   keys = { "F_NAME":"Tom", "L_NAME":"Black" }
This would find and return all records where F_NAME is "Tom" and L_NAME is
"Black".
Returns an array of dictionaries. Each dictionary corresponds a record that
matches the find criteria. In each dictionary, the key is the column name
and the value is the value in the column for that record.

Open( fileName, columns )
Open a table, or create it if it does not exist.
fileName is the name of the file to open or create. It can have an absolute
path, relateive path, or no path. If no extension is given, the .dbf
extension will be used.
columns is the same as the field info array passed to d4create.
This function will also generate an index for each column.

Updates( data, column, keysAndValues )
Modify existing records.
data is the table handle returned by Open.
column is the name of the column to search.
keysAndValues is a dictionary that defines the value to find in the column
and the new values for the found records. In each entry of the dictionary,
the key is the value to find in the column, and the value is a dictionary
which defines the new column values for the record. For example:
   Updates( data, "CITY", {
      "Chicago":{
         "HOCKEYTEAM":"Black Hawks",
         "POPULATION":2783726
      }, {
      "New Orleans":{
         "HOCKEYTEAM":"",
         "POPULATION":496938
      } }
   )
This example would find all records where CITY is "Chicago" and "New
Orleans", and replace the values in the HOCKTEAM and POPULATION fields.
Any columns that are not listed in the dictionary retain their previous
value.
Returns a 2-part tuple. The first part is an integer that indicates the
number of records that were successfully modified. The second part is an
array of integers which indicate which entries in the keysAndValues
dictionary failed to be deleted.

Writes( data, newRecords )
Add records to the table.
data is the table handle returned by Open.
newRecords is a dictionary or an array of dictionaries. Each dictionary
defines a new record. In each dictionary, the key is the column name and the
value is the value in the column for the new record.
Any columns that are not listed in the dictionary will be blank in the new
record.
