/* python4.c   (c)Copyright Sequiter Software Inc., 1988-2001.  All rights reserved. */
#include <Python.h>
#include <d4all.h>
#include <time.h>
#include <stdio.h>
#include <stdlib.h>
#include <math.h>

/* To build Win32 DLL:

cl -WX -LD Python4.c -Ic:\Python152\include -I[codebasesource] -oc:\Python152\DLLs\codebase.dll c:\Python152\libs\python15.lib c4dll.lib
*/

/* To build Linux .so:

g++ -I/usr/src/redhat/SOURCES/python/Python-1.5.2/Include -I[codebasesource] -L[codebasesource] -shared -Wl,-soname,codebasemodule.so -o codebasemodule.so Python4.c -lcb
*/

static PyObject *CBException(CODE4 *c4)
{
   /* Generate a CodeBase exception. */
   int rc = c4getErrorCode(c4);
   c4setErrorCode(c4,0);
   return PyErr_Format(PyExc_Exception,"Error %d: %s",rc,error4text(c4,rc));
}

static PyObject *cb_code4accessMode(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int newVal = 10005;
   int oldVal;

   if (!PyArg_ParseTuple(args,"l|i",&cb,&newVal))
      return NULL;

   oldVal = c4getAccessMode(cb);
   if (newVal != 10005)
      c4setAccessMode(cb,(char)newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4autoOpen(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int newVal = 10005;
   int oldVal;

   if (!PyArg_ParseTuple(args,"l|i",&cb,&newVal))
      return NULL;

   oldVal = c4getAutoOpen(cb);
   if (newVal != 10005)
      c4setAutoOpen(cb,newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4codePage(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int newVal = 10005;

   if (!PyArg_ParseTuple(args,"li",&cb,&newVal))
      return NULL;

   c4setCodePage(cb,newVal);
   return Py_BuildValue("i",c4getErrorCode(cb));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_code4collate(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   short newVal = -1;

   if (!PyArg_ParseTuple(args,"l|h",&cb,&newVal))
      return NULL;

   return Py_BuildValue("h",code4collate(cb,newVal));
}

static PyObject *cb_code4collateUnicode(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   short newVal = -1;

   if (!PyArg_ParseTuple(args,"l|h",&cb,&newVal))
      return NULL;

   return Py_BuildValue("h",code4collateUnicode(cb,newVal));
}

static PyObject *cb_code4compatibility(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   short newVal = 10005;
   short oldVal;

   if (!PyArg_ParseTuple(args,"l|h",&cb,&newVal))
      return NULL;

   oldVal = c4getCompatibility(cb);
   if (newVal != 10005)
      c4setCompatibility(cb,newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4createTemp(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int newVal = 10005;
   int oldVal;

   if (!PyArg_ParseTuple(args,"l|i",&cb,&newVal))
      return NULL;

   oldVal = c4getCreateTemp(cb);
   if (newVal != 10005)
      c4setCreateTemp(cb,newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4dateFormat(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   const char *newVal = 0;
   char oldVal[64];

   if (!PyArg_ParseTuple(args,"l|s",&cb,&newVal))
      return NULL;

   strcpy(oldVal,code4dateFormat(cb));
   if (newVal)
      code4dateFormatSet(cb,newVal);

   return Py_BuildValue("s",oldVal);
}

static PyObject *cb_code4errDefaultUnique(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   short newVal = 10005;
   int oldVal;

   if (!PyArg_ParseTuple(args,"l|h",&cb,&newVal))
      return NULL;

   oldVal = c4getErrDefaultUnique(cb);
   if (newVal != 10005)
      c4setErrDefaultUnique(cb,newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4indexExtension(PyObject *self,PyObject *args)
{
   CODE4 *cb;

   if (!PyArg_ParseTuple(args,"l",&cb))
      return NULL;

   return Py_BuildValue("s",code4indexExtension(cb));
}

static PyObject *cb_code4lockAttempts(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int newVal = 10005;
   int oldVal;

   if (!PyArg_ParseTuple(args,"l|i",&cb,&newVal))
      return NULL;

   oldVal = c4getLockAttempts(cb);
   if (newVal != 10005)
      c4setLockAttempts(cb,newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4readLock(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int newVal = 10005;
   int oldVal;

   if (!PyArg_ParseTuple(args,"l|i",&cb,&newVal))
      return NULL;

   oldVal = c4getReadLock(cb);
   if (newVal != 10005)
      c4setReadLock(cb,(char)newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4readOnly(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int newVal = 10005;
   int oldVal;

   if (!PyArg_ParseTuple(args,"l|i",&cb,&newVal))
      return NULL;

   oldVal = c4getReadOnly(cb);
   if (newVal != 10005)
      c4setReadOnly(cb,newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4safety(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int newVal = 10005;
   int oldVal;

   if (!PyArg_ParseTuple(args,"l|i",&cb,&newVal))
     return NULL;

   oldVal = c4getSafety(cb);
   if (newVal != 10005)
      c4setSafety(cb,(char)newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4singleOpen(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int newVal = 10005;
   int oldVal;

   if (!PyArg_ParseTuple(args,"l|i",&cb,&newVal))
      return NULL;

   oldVal = c4getSingleOpen(cb);
   if (newVal != 10005)
      c4setSingleOpen(cb,(short)newVal);

   return Py_BuildValue("i",oldVal);
}

static PyObject *cb_code4init(PyObject *self,PyObject *args)
{
   CODE4 *cb;

   if (!PyArg_ParseTuple(args,""))
      return NULL;

   cb = code4initAllocLow( "c4sock.dll" ) ;
   if (!cb)
      return PyErr_Format(PyExc_Exception,"CodeBase initialization failed.");

   c4setErrOff(cb,1);
   return Py_BuildValue("l", (long)cb);
}

static PyObject *cb_code4initUndo(PyObject *self,PyObject *args)
{
   int rc;
   CODE4 *cb;

   if (!PyArg_ParseTuple(args,"l",&cb))
      return NULL;

   rc = code4initUndo(cb);

   return Py_BuildValue("i",rc);
}

static PyObject *cb_code4close(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   int rc;

   if (!PyArg_ParseTuple(args,"l",&cb))
      return NULL;

   rc = code4close(cb);
   if (rc < 0)
      return CBException(cb);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4append(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4append(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("h",rc);
}

static PyObject *cb_d4appendStart(PyObject *self,PyObject *args)
{
   DATA4 *data;
   short useMemoEntries = 0;
   int rc;

   if (!PyArg_ParseTuple(args,"l|h",&data,&useMemoEntries))
      return NULL;

   rc = d4appendStart(data,useMemoEntries);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4blank(PyObject *self,PyObject *args)
{
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   d4blank(data);

   return Py_BuildValue("i",c4getErrorCode(data->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_d4bof(PyObject *self,PyObject *args)
{
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   return Py_BuildValue("i",d4bof(data));
}

static PyObject *cb_d4bottom(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;
   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4bottom(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4close(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4close(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4create(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   DATA4 *data;
   const char *fileName;
   PyObject *pyFields;
   FIELD4INFO *fields = 0;
   int numFields;
   PyObject *pyTags = 0;
   TAG4INFO *tags = 0;
   int numTags;
   int i;

   // collect d4create parameters; tags is optional.
   if (!PyArg_ParseTuple(args,"lsO|O:d4create",&cb,&fileName,&pyFields,&pyTags))
      return NULL;

   if (!PyList_Check(pyFields))
      return PyErr_Format(PyExc_SyntaxError,"Field info must be a list");

   numFields = PyList_Size(pyFields);
   fields = (FIELD4INFO *)u4allocFree(cb,(numFields + 1) * sizeof(FIELD4INFO));
   memset(fields,0,(numFields + 1) * sizeof(FIELD4INFO));
   for (i = 0;i < numFields;i++)
   {
      PyObject *pyField;
      pyField = PyList_GetItem(pyFields,i);
      if (!PyArg_ParseTuple(pyField,"sc|hhh:d4create(get fields)",
         &(fields[i].name),&(fields[i].type),&(fields[i].len),
         &(fields[i].dec),&(fields[i].nulls)))
      {
         u4free(fields);
         return NULL;
      }
   }

   if (pyTags)
   {
      if (!PyList_Check(pyTags))
         return PyErr_Format(PyExc_SyntaxError,"Tag info must be a list");

      numTags = PyList_Size(pyTags);
      tags = (TAG4INFO *)u4allocFree(cb,(numTags + 1) * sizeof(TAG4INFO));
      memset(tags,0,(numTags + 1) * sizeof(TAG4INFO));
      for (i = 0;i < numTags;i++)
      {
         PyObject *pyTag;
         pyTag = PyList_GetItem(pyTags,i);
         if (!PyArg_ParseTuple(pyTag,"ss|shh:d4create(get tags)",
            &(tags[i].name),&(tags[i].expression),&(tags[i].filter),
            &(tags[i].unique),&(tags[i].descending)))
         {
            u4free(fields);
            u4free(tags);
            return NULL;
         }
      }
   }

   data = d4create(cb,fileName,fields,tags);
   u4free(fields);
   u4free(tags);
   if (!data)
      return CBException(cb);

   return Py_BuildValue("l", (long)data);
}

static PyObject *cb_d4delete(PyObject *self,PyObject *args)
{
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   d4delete(data);
   return Py_BuildValue("i",c4getErrorCode(data->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_d4deleted(PyObject *self,PyObject *args)
{
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   return Py_BuildValue("i",d4deleted(data));
}

static PyObject *cb_d4eof(PyObject *self,PyObject *args)
{
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   return Py_BuildValue("i",d4eof(data));
}

static PyObject *cb_d4field(PyObject *self,PyObject *args)
{
   DATA4 *data;
   const char *fieldName;
   FIELD4 *field;

   if (!PyArg_ParseTuple(args,"ls",&data,&fieldName))
      return NULL;

   field = d4field(data,fieldName);
   if (!field)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("l",(long)field);
}

static PyObject *cb_d4fieldJ(PyObject *self,PyObject *args)
{
   DATA4 *data;
   long fieldNum;
   FIELD4 *field;

   if (!PyArg_ParseTuple(args,"ll",&data,&fieldNum))
      return NULL;

   field = d4fieldJ(data,(short)fieldNum);
   if (!field)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("l",(long)field);
}

static PyObject *cb_d4flush(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;
   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4flush(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4go(PyObject *self,PyObject *args)
{
   DATA4 *data;
   long recno;
   int rc;

   if (!PyArg_ParseTuple(args,"ll",&data,&recno))
      return NULL;

   rc = d4go(data,recno);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4numFields(PyObject *self,PyObject *args)
{
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   return Py_BuildValue("l",(long)d4numFields(data));
}

static PyObject *cb_d4open(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   DATA4 *data;
   const char *fileName;

   if (!PyArg_ParseTuple(args,"ls",&cb,&fileName))
      return NULL;

   data = d4open(cb,fileName);
   if (!data)
      return CBException(cb);

   return Py_BuildValue("l", (long)data);
}

static PyObject *cb_d4pack(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4pack(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4position(PyObject *self,PyObject *args)
{
   DATA4 *data;
   double newPos = -22;
   double oldPos;
   int rc;

   if (!PyArg_ParseTuple(args,"l|d",&data,&newPos))
      return NULL;

   oldPos = d4position(data);

   if (newPos >= 0)
   {
      rc = d4positionSet(data,newPos);
      if (rc < 0)
         return CBException(data->codeBase);
   }

   return Py_BuildValue("i",rc);
}

static PyObject *cb_d4memoCompress(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4memoCompress(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4recWidth(PyObject *self,PyObject *args)
{
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   return Py_BuildValue("l",(long)d4recWidth(data));
}

static PyObject *cb_d4refreshRecord(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4refreshRecord(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4seek(PyObject *self,PyObject *args)
{
   DATA4 *data;
   const char *seekStr;
   int rc;
   if (!PyArg_ParseTuple(args,"ls",&data,&seekStr))
      return NULL;

   rc = d4seek(data,seekStr);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4seekDouble(PyObject *self,PyObject *args)
{
   DATA4 *data;
   double seekVal;
   int rc;
   if (!PyArg_ParseTuple(args,"ld",&data,&seekVal))
      return NULL;

   rc = d4seekDouble(data,seekVal);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4seekNext(PyObject *self,PyObject *args)
{
   DATA4 *data;
   const char *seekStr;
   int rc;
   if (!PyArg_ParseTuple(args,"ls",&data,&seekStr))
      return NULL;

   rc = d4seekNext(data,seekStr);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4seekNextDouble(PyObject *self,PyObject *args)
{
   DATA4 *data;
   double seekVal;
   int rc;
   if (!PyArg_ParseTuple(args,"ld",&data,&seekVal))
      return NULL;

   rc = d4seekNextDouble(data,seekVal);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4recall(PyObject *self,PyObject *args)
{
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   d4recall(data);
   return Py_BuildValue("i",c4getErrorCode(data->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_d4recCount(PyObject *self,PyObject *args)
{
   DATA4 *data;
   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   return Py_BuildValue("l",d4recCount(data));
}

static PyObject *cb_d4recNo(PyObject *self,PyObject *args)
{
   DATA4 *data;
   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   return Py_BuildValue("l",d4recNo(data));
}

static PyObject *cb_d4reindex(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;
   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4reindex(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4remove(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;
   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4remove(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4skip(PyObject *self,PyObject *args)
{
   DATA4 *data;
   long nSkip = 1L;
   int rc;
   if (!PyArg_ParseTuple(args,"l|l",&data,&nSkip))
      return NULL;

   rc = d4skip(data,nSkip);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4tag(PyObject *self,PyObject *args)
{
   DATA4 *data;
   const char *tagName;
   TAG4 *tag;

   if (!PyArg_ParseTuple(args,"ls",&data,&tagName))
      return NULL;

   tag = d4tag(data,tagName);
   if (!tag)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("l",(long)tag);
}

static PyObject *cb_d4tagSelect(PyObject *self,PyObject *args)
{
   DATA4 *data;
   TAG4 *tag = 0;

   if (!PyArg_ParseTuple(args,"l|l",&data,&tag))
      return NULL;

   d4tagSelect(data,tag);
   return Py_BuildValue("i",c4getErrorCode(data->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_d4top(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;
   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4top(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4unlock(PyObject *self,PyObject *args)
{
   DATA4 *data;
   int rc;
   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = d4unlock(data);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_d4zap(PyObject *self,PyObject *args)
{
   DATA4 *data;
   long recStart = 1,recEnd = -10005;
   int rc;

   if (!PyArg_ParseTuple(args,"l|ll",&data,&recStart,&recEnd))
      return NULL;

   if (recEnd == -10005)
      recEnd = d4recCount(data);

   rc = d4zap(data,recStart,recEnd);
   if (rc < 0)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("i",rc);
}

static PyObject *cb_f4assign(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   const char *newValue;
   int len;
   if (!PyArg_ParseTuple(args,"ls#",&field,&newValue,&len))
      return NULL;

   f4memoAssignN(field,newValue,len);
   return Py_BuildValue("i",c4getErrorCode(f4data(field)->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_f4blank(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   f4blank(field);
   return Py_BuildValue("i",c4getErrorCode(f4data(field)->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_f4assignChar(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   int newValue;
   if (!PyArg_ParseTuple(args,"lc",&field,&newValue))
      return NULL;

   f4assignChar(field,newValue);
   return Py_BuildValue("i",c4getErrorCode(f4data(field)->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_f4assignDouble(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   double newValue;
   if (!PyArg_ParseTuple(args,"ld",&field,&newValue))
      return NULL;

   f4assignDouble(field,newValue);
   return Py_BuildValue("i",c4getErrorCode(f4data(field)->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_f4assignInt(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   int newValue;
   if (!PyArg_ParseTuple(args,"li",&field,&newValue))
      return NULL;

   f4assignInt(field,newValue);
   return Py_BuildValue("i",c4getErrorCode(f4data(field)->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_f4assignLong(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   long newValue;
   if (!PyArg_ParseTuple(args,"ll",&field,&newValue))
      return NULL;

   f4assignLong(field,newValue);
   return Py_BuildValue("i",c4getErrorCode(f4data(field)->codeBase));  // SHOULD RETURN NONE BUT THAT CRASHES PYTHON
   //return Py_None;
}

static PyObject *cb_f4char(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   if (f4null(field))
      return Py_None;
   else
      return Py_BuildValue("c",f4char(field));
}

static PyObject *cb_f4decimals(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   return Py_BuildValue("i",f4decimals(field));
}

static PyObject *cb_f4double(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   if (f4null(field))
      return Py_None;
   else
      return Py_BuildValue("d",f4double(field));
}

static PyObject *cb_f4int(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   if (f4null(field))
      return Py_None;
   else
      return Py_BuildValue("i",f4int(field));
}

static PyObject *cb_f4len(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   return Py_BuildValue("l",f4memoLen(field));
}

static PyObject *cb_f4long(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   if (f4null(field))
      return Py_None;
   else
      return Py_BuildValue("l",f4long(field));
}

static PyObject *cb_f4name(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   return Py_BuildValue("s",f4name(field));
}

static PyObject *cb_f4str(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   if (f4null(field))
      return Py_None;
   else
      return Py_BuildValue("s",f4memoStr(field));
}

static PyObject *cb_f4true(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   if (f4null(field))
      return Py_None;
   else
      return Py_BuildValue("i",f4true(field));
}

static PyObject *cb_f4type(PyObject *self,PyObject *args)
{
   FIELD4 *field;
   if (!PyArg_ParseTuple(args,"l",&field))
      return NULL;

   return Py_BuildValue("c",f4type(field));
}

static PyObject *cb_i4create(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   DATA4 *data;
   INDEX4 *index;
   const char *fileName;
   int fileNameLen;
   PyObject *pyTags = 0;
   TAG4INFO *tags = 0;
   int numTags;
   int i;

   if (!PyArg_ParseTuple(args,"ls#O:i4create",&data,&fileName,&fileNameLen,&pyTags))
      return NULL;

   cb = data->codeBase;

   if (fileNameLen == 0)
      fileName = 0;

   if (!PyList_Check(pyTags))
      return PyErr_Format(PyExc_SyntaxError,"Tag info must be a list");

   numTags = PyList_Size(pyTags);
   tags = (TAG4INFO *)u4allocFree(cb,(numTags + 1) * sizeof(TAG4INFO));
   memset(tags,0,(numTags + 1) * sizeof(TAG4INFO));
   for (i = 0;i < numTags;i++)
   {
      PyObject *pyTag;
      pyTag = PyList_GetItem(pyTags,i);
      if (!PyArg_ParseTuple(pyTag,"ss|shh:i4create(get tags)",
         &(tags[i].name),&(tags[i].expression),&(tags[i].filter),
         &(tags[i].unique),&(tags[i].descending)))
      {
         u4free(tags);
         return NULL;
      }
   }

   index = i4create(data,fileName,tags);
   u4free(tags);
   if (!data)
      return CBException(cb);
   else
      return Py_BuildValue("i", c4getErrorCode(cb));
}

static PyObject *cb_i4open(PyObject *self,PyObject *args)
{
   DATA4 *data;
   INDEX4 *index;
   const char *fileName;

   if (!PyArg_ParseTuple(args,"ls",&data,&fileName))
      return NULL;

   index = i4open(data,fileName);
   if (!index)
      return CBException(data->codeBase);
   else
      return Py_BuildValue("l", c4getErrorCode(data->codeBase));
}

/******* FUNCTIONS FOR ND *******/

static PyObject *assignFields(DATA4* data,PyObject *record)
{
   // record is a Dictionary that contains entries for field values
   int pos;
   PyObject *fieldName;
   PyObject *fieldValue;
   FIELD4 *field;

   pos = 0;
   while (PyDict_Next(record,&pos,&fieldName,&fieldValue))
   {
      // Exit with an exception if the field name portion
      // of the dictionary entry is not a string.
      if (!PyString_Check(fieldName))
         return PyErr_Format(PyExc_TypeError,"Field name must be a String");

      field = d4field(data,PyString_AsString(fieldName));
      if (!field)
      {
         data->codeBase->errorCode = 0;
         return PyErr_Format(PyExc_Exception,"Field %s not found",PyString_AsString(fieldName));
      }

      switch (f4type(field))
      {
         case r4currency:
            if (PyFloat_Check(fieldValue))
               f4assignDouble(field,PyFloat_AsDouble(fieldValue));
            else if( PyString_Check(fieldValue) )
               f4assignCurrency(field, PyString_AsString(fieldValue)) ;
            else if( PyLong_Check(fieldValue))
               f4assignDouble(field,PyLong_AsDouble(fieldValue)); 
            else if( PyInt_Check(fieldValue) )
               f4assignDouble(field,(double)PyInt_AsLong(fieldValue)); 
            else
               return PyErr_Format(PyExc_TypeError,"Value for field %s not valid",PyString_AsString(fieldName));
            break;   
         case r4float:
            if (PyFloat_Check(fieldValue))
               f4assignDouble(field,PyFloat_AsDouble(fieldValue));
            else if( PyString_Check(fieldValue) )
               f4assign(field, PyString_AsString(fieldValue)) ;
            else if( PyLong_Check(fieldValue))
               f4assignLong(field,PyLong_AsLong(fieldValue)); 
            else if( PyInt_Check(fieldValue) )
               f4assignLong(field,PyInt_AsLong(fieldValue)); 
            else
               return PyErr_Format(PyExc_TypeError,"Value for field %s not valid",PyString_AsString(fieldName));
            break;            
         case r4num:
            if (PyFloat_Check(fieldValue))
               f4assignDouble(field,PyFloat_AsDouble(fieldValue));
            else if( PyString_Check(fieldValue) )
               f4assign(field, PyString_AsString(fieldValue)) ;
            else if( PyLong_Check(fieldValue))
               f4assignLong(field,PyLong_AsLong(fieldValue)) ; 
            else if( PyInt_Check(fieldValue) )
               f4assignLong(field,PyInt_AsLong(fieldValue)) ; 
            else
               return PyErr_Format(PyExc_TypeError,"Value for field %s not valid",PyString_AsString(fieldName));
            break;                     
         case r4dateTime:
            if (PyFloat_Check(fieldValue))
               f4assignDouble(field,PyFloat_AsDouble(fieldValue));
            else if( PyString_Check(fieldValue) )
               f4assignDateTime(field, PyString_AsString(fieldValue)) ;
            else
               return PyErr_Format(PyExc_TypeError,"Value for field %s not valid",PyString_AsString(fieldName));
            break;                     
         case r4int:
            if( PyString_Check(fieldValue) )
               f4assignLong(field, c4atoi(PyString_AsString(fieldValue), PyString_Size(fieldValue)) ) ;
            else if( PyLong_Check(fieldValue))
               f4assignLong(field,PyLong_AsLong(fieldValue)); 
            else if( PyInt_Check(fieldValue) )
               f4assignLong(field,PyInt_AsLong(fieldValue)); 
            else
               return PyErr_Format(PyExc_TypeError,"Value for field %s not valid",PyString_AsString(fieldName));
            break;            
         case 'B':
            if (code4indexFormat(data->codeBase) == r4cdx)
            {
               if (PyFloat_Check(fieldValue))
                  f4assignDouble(field,PyFloat_AsDouble(fieldValue));
               else if( PyString_Check(fieldValue) )
                  f4assignDouble(field, c4atod(PyString_AsString(fieldValue), PyString_Size(fieldValue)) ) ;
               else if( PyLong_Check(fieldValue))
                  f4assignLong(field,PyLong_AsLong(fieldValue)); 
               else if( PyInt_Check(fieldValue) )
                  f4assignLong(field,PyInt_AsLong(fieldValue)); 
               else
                  return PyErr_Format(PyExc_TypeError,"Value for field %s not valid",PyString_AsString(fieldName));
            }
            else
            {
               if( PyString_Check(fieldValue) )
                  f4assign(field, PyString_AsString(fieldValue)) ;
               else
                  return PyErr_Format(PyExc_TypeError,"Value for field %s not valid",PyString_AsString(fieldName));
            }
            break;             
         case r4str:
            if( PyString_Check(fieldValue) )
               f4assign(field, PyString_AsString(fieldValue)) ;
            else
               return PyErr_Format(PyExc_TypeError,"Value for field %s not valid",PyString_AsString(fieldName));         
         default:
            if( PyString_Check(fieldValue) )
               f4memoAssign(field, PyString_AsString(fieldValue)) ;
            else
               return PyErr_Format(PyExc_TypeError,"Value for field %s not valid",PyString_AsString(fieldName));                  
            break;
      }
   }
   return NULL;
}

int doSeek(DATA4 *data,PyObject *pySeek,short next,short partial, FIELD4* field)
{
   int rc;

   if (PyString_Check(pySeek))
   {
      char *seekStr;
      char *newSeek;
      short seekLen;
      short allocFlag = 0 ;

      seekStr = PyString_AsString(pySeek);
      seekLen = PyString_Size(pySeek);

      // if partial = true, just d4seek
      // if partial = false, make a new seek string which is padded with blanks
      // if string but seeking on numerical data then just d4seek      
      if (partial)
         newSeek = seekStr;
      else 
      {
         EXPR4 *e;
         int keyLen;
         if( field )
         {         
            switch (f4type(field))
            {
               case r4currency:
               case r4float:
               case r4num:
               case r4dateTime:
               case r4int:
               case r4memo:
               case 'B':
               case r4gen :
                  newSeek = seekStr ;
                  break;
               case r4str:
               default:
                  e = expr4parse(data,t4expr(d4tagSelected(data)));
                  keyLen = expr4len(e);
                  expr4free(e);

                  newSeek = (char *)u4allocFree(data->codeBase,keyLen);
                  allocFlag = 1 ;

                  memset(newSeek,' ',keyLen);
                  memcpy(newSeek,seekStr,seekLen);
                  seekLen = keyLen;
                  break ;
            }
         }
         else
         {         
            e = expr4parse(data,t4expr(d4tagSelected(data)));
            keyLen = expr4len(e);
            expr4free(e);
            
            newSeek = (char *)u4allocFree(data->codeBase,keyLen);
            allocFlag = 1 ;

            memset(newSeek,' ',keyLen);
            memcpy(newSeek,seekStr,seekLen);
            seekLen = keyLen;
         }
      }

      if (next)
         rc = d4seekNextN(data,newSeek,seekLen);
      else
         rc = d4seekN(data,newSeek,seekLen);
         
      
      if (allocFlag)
         u4free(newSeek);
   }
   else
   {
      double seekVal;
      if (PyFloat_Check(pySeek))
         seekVal = PyFloat_AsDouble(pySeek);
      else if (PyInt_Check(pySeek))
         seekVal = (double)PyInt_AsLong(pySeek);
      else if (PyLong_Check(pySeek))
         seekVal = (double)PyLong_AsLong(pySeek);
      else
      {
         PyErr_Format(PyExc_TypeError,"Value for seek not valid");
         return -1;
      }
                 
      if (next)
         rc = d4seekNextDouble(data,seekVal);
      else
         rc = d4seekDouble(data,seekVal);
         
   }

   return rc;
}

static PyObject *getRecord(DATA4 *data)
{
   // used by Read() and Find()
   // read the current record and return all fields as a dictionary
   PyObject *recordDict;
   short j;
   int numFields;
   FIELD4 *field;
   char fieldName[11];
   PyObject *fieldValue;
   char *strValue;

   numFields = d4numFields(data);
   recordDict = PyDict_New();
   for (j = 1;j <= numFields;j++)
   {
      field = d4fieldJ(data,j);
      switch (f4type(field))
      {
         case r4currency:
         case r4float:
         case r4num:
            fieldValue = Py_BuildValue("d",f4double(field));
            break;
         case r4dateTime:
            fieldValue = Py_BuildValue("s",f4dateTime(field));
            break;
         case r4int:
            fieldValue = Py_BuildValue("i",f4int(field));
            break;
         case r4str:
            strValue = f4str(field);
            c4trimN(strValue,f4len(field)+1);
            fieldValue = Py_BuildValue("s",strValue);
            break;
         case 'B':
            if (code4indexFormat(data->codeBase) == r4cdx)
            {
               fieldValue = Py_BuildValue("d",f4double(field));
               break;
            }
         default:
            fieldValue = Py_BuildValue("s#",f4memoPtr(field),f4memoLen(field));
      }
      strncpy(fieldName,f4name(field),10);
      fieldName[10] = 0;
      PyDict_SetItemString(recordDict,fieldName,fieldValue);
      Py_DECREF(fieldValue);
   }

   return recordDict;
}

PyObject *LockHandle(CODE4 *cb)
{
   // Use CODE4.s4cr2 as a flag to indicate if the process is using the table
   int lockCount;
   for (lockCount = 70 ; lockCount ; lockCount--)
   {
      cb->s4cr2++;
      if (cb->s4cr2 == 1)
         break;

      // Someone else has it locked. Keep trying for seven seconds.
      cb->s4cr2--;
      u4delayHundredth(10); // retry 10 times per second
   }
   if (lockCount == 0)
      return PyErr_Format(PyExc_IOError,"Timeout waiting for other process.");

   return 0;
}

void UnlockHandle(CODE4 *cb)
{
   cb->s4cr2--;
}

int spacecat(char *queryStr,DATA4 *data,const char *fieldName,const char *fieldValue)
{
   // contatenate spaces to queryStr to make the difference between the
   // field value and the field.
   // e.g. if the field is 10 bytes and fieldValue is "abcd", append 6 spaces
   // Returns the number of spaces appended.

   FIELD4 *field = d4field(data,fieldName);
   size_t valLen;
   unsigned long fLen;
   unsigned long count;

   if (!field)
   {
      data->codeBase->errorCode = 0;
      return -1;
   }

   valLen = strlen(fieldValue);
   fLen = f4len(field);
   count = 0;
   while (fLen > valLen)
   {
      strcat(queryStr," ");
      count++;
      fLen--;
   }

   return count;
}

static PyObject *nd_Close(PyObject *self,PyObject *args)
{
   int rc;
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   rc = code4initUndo(data->codeBase);
   return Py_BuildValue("i", rc);
}

static PyObject *nd_Delete(PyObject *self,PyObject *args)
{
   return cb_d4delete(self,args);
}

static PyObject *nd_Deletes(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   DATA4 *data;
   const char *fieldName;
   TAG4 *tag;
   FIELD4 *field ;
   PyObject *keys;
   short doPack = 1;
   int rc;
   long numSuccessful = 0;
   PyObject *listFailed = PyList_New(0), *pyLockRC;

   if (!PyArg_ParseTuple(args,"lsO|h:Deletes",&data,&fieldName,&keys,&doPack))
      return NULL;

   cb = data->codeBase;

   tag = d4tag(data,fieldName);
   if (!tag)
      return CBException(cb);      
   d4tagSelect(data,tag);
   field = d4field( data, fieldName ) ;
   if( !field )
      return CBException(cb);      

   pyLockRC = LockHandle(cb);
   if (pyLockRC)
      return pyLockRC;

   if (!PyList_Check(keys))
   {
      UnlockHandle(cb);
      return PyErr_Format(PyExc_TypeError,"Parameter must be a list.");
   }
   else
   {
      PyObject *key;
      int i,size;

      size = PyList_Size(keys);
      if (size == 0)  // delete ALL records
      {
         d4zap(data, 1L, d4recCount(data));
         doPack = 0;  // no need to pack after zap
      }
      else
      {
         for (i = 0;i < size;i++)
         {
            key = PyList_GetItem(keys,i);
            rc = doSeek(data, key, 0, 0, field);
            if (rc != r4success)
            {            
               c4setErrorCode(data->codeBase, 0);
               PyList_Append(listFailed,Py_BuildValue("i",i));
            }
            else
            {
               do
               {
                  d4delete(data);
                  rc = d4write(data, -1L);
                  if (rc == r4success)
                     rc = d4unlock(data);
                  if (rc == r4success)
                     numSuccessful++;
                  else
                  {
                     PyErr_Clear();
                     PyList_Append(listFailed,Py_BuildValue("i",i));
                     break;
                  }
                  rc = doSeek(data, key, 0, 0, field);
                  c4setErrorCode(data->codeBase, 0);
               } while (rc == r4success);
            }  // if rc
         }  // for i
      }  // if size
   }  // if PyList_Check
   
   if (doPack)
      d4pack(data);

   UnlockHandle(cb);
   return Py_BuildValue("iO",numSuccessful,listFailed);
}

static PyObject *nd_Drop(PyObject *self,PyObject *args)
{
   PyObject *ret = cb_d4remove(self,args);
   PyErr_Clear();
   return Py_BuildValue("i",ret?1:0);
}

void strcatWithQuoteTranspose(char *dst,const char *src)
{
   // copy expression string literal from src to dst, but when encountering
   // a double quote character, put it in its own.
   // e.g. if src points to   Call me "Cam"
   //      dst should be      "Call me " + '"' + "Cam" + '"'

   int i;
   size_t len;

   strcat(dst,"\"");
   len = strlen(dst);

   for (i = 0;src[i];i++)
   {
      if (src[i] == '\"')
      {
         strcat(dst,"\" + \'\"\' + \"");  // append this:   " + '"' + "
         len = strlen(dst);
      }
      else
      {
         dst[len++] = src[i];
         dst[len] = 0;
      }
   }
   strcat(dst,"\"");
}

static PyObject *nd_Find(PyObject *self, PyObject *args)
{
   // find all records and return in an array of dictionaries
   // function takes (data,tagName,seekValue,partial)
   // data is the DATA4 pointer
   // tagName is the name of the tag in which to seek
   // seekValue is the value for which to seek
   //    seekValue can be string, float, int or long
   // partial is a true/false flag that indicates if the seek
   // should find exact or partial matches.
   CODE4 *cb;
   DATA4 *data;
   PyObject *keys;
   short partial = 0;
   PyObject *sort = 0;
   PyObject *foundRecords = PyList_New(0), *pyLockRC;
   char queryStr[1024] = "", sortStr[1024] = "";

   int pos = 0, rc;
   PyObject *pyfieldName;
   PyObject *pyfieldValue;
   const char *fieldName;
   FIELD4 *field;

   RELATE4 *rel;

   if (!PyArg_ParseTuple(args, "lO|Oh", &data, &keys, &sort, &partial))
      return NULL;

   cb = data->codeBase;

   if (!PyDict_Check(keys))
      return PyErr_Format(PyExc_TypeError, "Parameter must be a dictionary.");

   if (sort)
   {
      if (!PyList_Check(sort))
      {
         // could be "partial" as third parameter
         if (PyInt_Check(sort))
         {
            partial = (short)PyInt_AsLong(sort);
            sort = 0;
         }
         else
            return PyErr_Format(PyExc_TypeError, "Third parameter must be a string array or integer.");
      }
   }

   pyLockRC = LockHandle(cb);
   if (pyLockRC)
      return pyLockRC;

   while (PyDict_Next(keys, &pos, &pyfieldName, &pyfieldValue))
   {
      if (!PyString_Check(pyfieldName))
      {
         UnlockHandle(cb);
         return PyErr_Format(PyExc_TypeError, "Field name must be a string.");
      }
      fieldName = PyString_AsString(pyfieldName);
      field = d4field(data, fieldName);
      if (!field)
      {
         UnlockHandle(cb);
         return CBException(data->codeBase);
      }
      if (strlen(queryStr) > 0)
         strcat(queryStr, " .AND. ");

      /* LY 2001/07/24 : changed from !f4type() == r4date */
      if (partial && (f4type(field) == r4str || f4type(field) == r4memo))  // inexact matches should be case insensitive
         strcat(queryStr,"UPPER(");
      strcat(queryStr, fieldName);
      if (partial && (f4type(field) == r4str || f4type(field) == r4memo))  // inexact matches should be case insensitive
         strcat(queryStr, ")");
      strcat(queryStr, " = ");

      if (PyString_Check(pyfieldValue))
      {
         const char *strFieldValue = PyString_AsString(pyfieldValue);
         // CJ July 26/01 -Changed to a switch value as the f4type function will not have to be called as often
         // also need to put in a check for when a string is passed but were finding a value in a numerical field.
         switch ( f4type(field) )
         {
            case r4date:
               strcat(queryStr, "STOD('");
               strcat(queryStr, strFieldValue);
               strcat(queryStr, "')");
               break;
            case r4log:
               if ( c4stricmp( strFieldValue, "T" ) || c4stricmp( strFieldValue, "Y" ) )
                  strcat( queryStr, ".T." ) ;
               else
                  strcat( queryStr, ".F." ) ;
               break;
            case r4num:
            case r4float:
            case r4int:
               strcat(queryStr, strFieldValue);
               break;
            default:
               if (partial)
                  strcat(queryStr, "UPPER");
               strcat(queryStr,"(");
               strcatWithQuoteTranspose(queryStr, strFieldValue);
               if (!partial)
               {
                  strcat(queryStr," + '");
                  if (spacecat(queryStr, data, fieldName, strFieldValue) < 0)
                  {
                     UnlockHandle(cb);
                     return PyErr_Format(PyExc_ValueError, "Unrecognized field name: s.", fieldName);
                  }
                  strcat(queryStr, "'");
               }
               strcat(queryStr, ")");
               break;      
         }        
      }
      else if (PyFloat_Check(pyfieldValue))
         sprintf(queryStr + strlen(queryStr),"%f", PyFloat_AsDouble(pyfieldValue));
      else if (PyInt_Check(pyfieldValue))
         sprintf(queryStr + strlen(queryStr),"%ld", PyInt_AsLong(pyfieldValue));
      else if (PyLong_Check(pyfieldValue))
         sprintf(queryStr + strlen(queryStr),"%d", PyLong_AsLong(pyfieldValue));
      else
      {
         // If we get here, it means that the field value
         // passed in is not compatible with an xBASE field
         // type, or at least can't be easily converted.
         UnlockHandle(cb);
         return PyErr_Format(PyExc_TypeError, "Value for field %s not valid", fieldName);
      }

      strcat( queryStr, " .AND. .NOT. DELETED()" );  // don't find deleted records
   }

   // now create sort expression
   if (sort)
   {
      const int sortListSize = PyList_Size(sort) / 2;  // Each sort specification is in a FIELD/DIRECTION pair.

      if (!PyList_Check(sort))
      {
         UnlockHandle(cb);
         return PyErr_Format(PyExc_TypeError, "Sort parameter name must be a list.");
      }

      for (pos = 0;pos < sortListSize;pos++)
      {
         pyfieldName = PyList_GetItem(sort, (pos*2)+0);
         pyfieldValue = PyList_GetItem(sort, (pos*2)+1);
         if (!PyString_Check(pyfieldName) || !PyString_Check(pyfieldValue))
         {
            UnlockHandle(cb);
            return PyErr_Format(PyExc_TypeError, "Sort list elements must be strings.");
         }

         if (sortStr[0] != 0)
            strcat(sortStr, " + ");

         rc = c4stricmp(PyString_AsString(pyfieldValue), "desc");
         if (rc == 0)
         {
            strcat(sortStr, "DESCEND( ");
            strcat(sortStr, PyString_AsString(pyfieldName));
            strcat(sortStr, " )");
            continue;
         }

         rc = c4stricmp(PyString_AsString(pyfieldValue), "asc");
         if (rc == 0)
         {
            strcat(sortStr, "ASCEND( ");
            strcat(sortStr, PyString_AsString(pyfieldName));
            strcat(sortStr, " )");
            continue;
         }

         // If we get here, the value was neither ASC or DESC.
         UnlockHandle(cb);
         return PyErr_Format(PyExc_ValueError, "Sort direction must be ASC or DESC.");
      }
   }
      
   rel = relate4init(data);
   relate4querySet(rel, queryStr);
   if (sortStr[0] != 0)
      relate4sortSet(rel, sortStr);
   for (rc = relate4top(rel);rc == r4success;rc = relate4skip(rel,1L))
   {
      PyObject *recordDict = getRecord(data);
      PyList_Append(foundRecords, recordDict);
      Py_DECREF(recordDict);
   }
   relate4free(rel,0);

   UnlockHandle(cb);

   if (rc < 0)
      return CBException(data->codeBase);

   return foundRecords;
}

/*static PyObject *nd_Find(PyObject *self,PyObject *args)
{
   return nd_FindLow(self,args,0);
}*/

/*static PyObject *nd_FindNext(PyObject *self,PyObject *args)
{
   return nd_FindLow(self,args,1);
}*/

void revertData4(DATA4* data,long recNo)
{
   if (recNo > d4recCount(data))
      d4goEof(data);
   else
      d4go(data,recNo);
}

static PyObject *nd_Open(PyObject *self, PyObject *args)
{
   CODE4 *cb;
   DATA4 *data;
   const char *fileName;
   PyObject *pyFields = 0;
   PyObject *pyAIStart = 0 ;
   FIELD4INFO *fields = 0;
   int num;
   TAG4INFO *tags = 0;
   int i;
   char indexName[LEN4PATH];
   double aiStart = 1.0;

   if (!PyArg_ParseTuple(args, "s|OO:Open", &fileName, &pyFields, &pyAIStart))
      return NULL;

   cb = code4initAllocLow( "c4sock.dll" ) ;
   if (!cb)
      return PyErr_Format(PyExc_Exception,"CodeBase initialization failed.");
   cb->s4cr2 = 0;
   c4setErrOff(cb,1);
   c4setAutoOpen(cb,0);  // allow for i4close when updating by opening/creating index separately
   
   u4namePiece(indexName, sizeof(indexName), fileName, 1, 0);

   data = d4open(cb,fileName);
   if (data)
   {
      i4open(data,indexName);
      cb->errorCode = 0;  // assume i4open succeeds
      return Py_BuildValue("l", (long)data);
   }
    
   // At this point, the table could not be opened. Try creating.

   if (!pyFields  // no field definition provided, so can't create
         || (cb->errorCode/10 != -6))  // table exists, but can't be opened
   {
      return CBException(cb);
   }

   cb->errorCode = 0;

   if (!PyList_Check(pyFields))
      return PyErr_Format(PyExc_SyntaxError,"Field info must be a list");

   num = PyList_Size(pyFields);
   fields = (FIELD4INFO *)u4allocFree(cb,(num + 1) * sizeof(FIELD4INFO));
   tags = (TAG4INFO *)u4allocFree(cb,(num + 1) * sizeof(TAG4INFO));
   memset(fields,0,(num + 1) * sizeof(FIELD4INFO));
   memset(tags,0,(num + 1) * sizeof(TAG4INFO));
   for (i = 0;i < num;i++)
   {
      PyObject *pyField;
      pyField = PyList_GetItem(pyFields, i);
      if (!PyArg_ParseTuple(pyField,"sc|hhh:d4create(get fields)",
         &(fields[i].name),&(fields[i].type),&(fields[i].len),
         &(fields[i].dec),&(fields[i].nulls)))
      {
         u4free(fields);
         return NULL;
      }
      tags[i].name = fields[i].name;
      tags[i].expression = tags[i].name;
      tags[i].filter = ".NOT. DELETED()";
   }
   
   if( pyAIStart != 0 )
   {
      if( PyString_Check(pyAIStart) )
      {
         char *aiStartString = PyString_AsString(pyAIStart) ;
         aiStart = c4atod( aiStartString, strlen(aiStartString) ) ;
      }
      else if( PyFloat_Check( pyAIStart ) )
      {
         aiStart = PyFloat_AsDouble( pyAIStart ) ;
      }
      else if( PyLong_Check( pyAIStart ) )
      {
         aiStart = (double)PyLong_AsLong( pyAIStart ) ;
      }
      else if( PyInt_Check( pyAIStart) )
      {
         aiStart = (double)PyInt_AsLong( pyAIStart ) ;
      }
   }
   code4autoIncrementStart(cb, aiStart);
   data = d4create(cb,fileName,fields,0);
   if (c4getErrorCode(cb) == e4data)
   {
      // might be because trying to create VFP fields
      c4setCompatibility(cb,30);
      c4setErrorCode(cb,0);
      data = d4create(cb,fileName,fields,0);
   }
   if (data)
   {
      i4create(data,indexName,tags);
      cb->errorCode = 0;
   }
   u4free(fields);
   u4free(tags);
   if (!data)
   {
      code4initUndo(cb);
      return CBException(cb);
   }

   return Py_BuildValue("l", (long)data);
}

static PyObject *nd_Read(PyObject *self,PyObject *args)
{
   /* return current record */
   DATA4 *data;

   if (!PyArg_ParseTuple(args,"l",&data))
      return NULL;

   return getRecord(data);
}

static PyObject *nd_Updates(PyObject *self,PyObject *args)
{
   CODE4 *cb;
   DATA4 *data;
   const char *fieldName;
   TAG4 *tag;
   PyObject *stuff;
   int rc;
   long numSuccessful = 0;
   PyObject *listFailed = PyList_New(0);
   PyObject *key;  // value to search for in the field
   PyObject *fields;  // new field info
   PyObject *pyLockRC;
   int pos = 0;

   if (!PyArg_ParseTuple(args,"lsO:Updates",&data,&fieldName,&stuff))
      return NULL;

   cb = data->codeBase;

   tag = d4tag(data,fieldName);
   if (!tag)
      return CBException(cb);
   d4tagSelect(data,tag);

   if (!PyDict_Check(stuff))
      return PyErr_Format(PyExc_TypeError,"Parameters must be a dictionary.");

   pyLockRC = LockHandle(cb);
   if (pyLockRC)
      return pyLockRC;

   while (PyDict_Next(stuff,&pos,&key,&fields))
   {
      rc = doSeek(data,key, 0, 0, 0);
      if (rc != r4success)
      {
         c4setErrorCode(cb, 0);
         PyList_Append(listFailed,Py_BuildValue("O",key));
      }
      else
      {
         do
         {
            assignFields(data,fields);
            if (PyErr_Occurred())
            {
               PyErr_Clear();
               PyList_Append(listFailed,Py_BuildValue("O",key));
            }
            else
            {
               rc = d4write(data, -1L);
               if (rc == r4success)
                  rc = d4unlock(data);
               if (rc == r4success)
                  numSuccessful++;
               else
                  PyList_Append(listFailed,Py_BuildValue("O",key));
            }
            rc = doSeek(data, key, 1, 0, 0);
            c4setErrorCode(cb, 0);
         } while (rc == r4success);
      }
   }

   UnlockHandle(cb);
   return Py_BuildValue("iO",numSuccessful,listFailed);
}

static PyObject *nd_Write(PyObject *self,PyObject *args)
{
   /* update current record */
   DATA4 *data;
   PyObject *record;
   int rc = 0;
   PyObject *exception;

   // get the DATA4 pointer and the new record
   if (!PyArg_ParseTuple(args,"lO",&data,&record))
      return NULL;

   // exit with an exception if the record is not passed as a dictionary
   if (!PyDict_Check(record))
      return PyErr_Format(PyExc_TypeError,"Record must be a Dictionary");

   // assign each field in the Dictionary
   exception = assignFields(data,record);
   if (PyErr_Occurred())
      return exception;

   rc = d4write(data, -1 ) ;
   if (rc != 0)
   {
      d4refreshRecord( data ) ;
      return CBException(data->codeBase);
   }
   else
      return Py_BuildValue("i",0);
}

static PyObject *nd_Writes(PyObject *self,PyObject *args)
{
   /* Write sequential (append) */
   CODE4 *cb;
   DATA4 *data;
   PyObject *records; // list of dictionaries which hold field names and values
   int rc = 0;
   short closeIndex = 0;

   int i;  // used in for loop
   int numRecs = 0;  // number of entries in dictionary
   short arrayFlag = 1 ;
   int numSuccessful = 0;  // number of successful appends
   PyObject *failedItems;  // list of list indexes for records that could not be appended
   PyObject *record;  // dictionary of field names and values which defines one record
   PyObject *retVal;  // return value
   PyObject *pyLockRC;

   // get the DATA4 pointer and the new record
   if (!PyArg_ParseTuple(args,"lO",&data,&records))
      return NULL;

   cb = data->codeBase;

   // exit with an exception if the record is not passed as a dictionary or array
   if (!PyList_Check(records))
   {
      if( !PyDict_Check(records))
      {
         return PyErr_Format(PyExc_TypeError,"Record must be an array of Dictionaries.");
      }
      arrayFlag = 0 ;
      numRecs = 1 ;
      record = records ;
   }

   pyLockRC = LockHandle(cb);
   if (pyLockRC)
      return pyLockRC;

   failedItems = PyList_New(0);  // initialize

   // loop through all records in the list and append each
   if( arrayFlag )
   	numRecs = PyList_Size(records);

   // Calculate the threshold based on current number of records and
   // number of fields (i.e. tags). If the number of records to be appended
   // exceeds that threshold, then it is likely faster to append with the
   // index closed, then open the index and reindex.
   // The formula for getting this threshold is:
   //     r * (i + 8)
   // t = -----------
   //         400
   // where r is the number of records initially in the table
   // and i is the number of index tags.
   if (numRecs > (d4recCount(data) * (d4numFields(data) + 8) / 400))
   {
      closeIndex = 1;
      i4close(d4index(data,0));
   }

   for (i = 0;i < numRecs;i++)
   {
      if( arrayFlag )
      {
         record = PyList_GetItem(records,i);

         // fail on this record if this list element is not a dictionary
         if (!PyDict_Check(record))
         {
            PyErr_Clear();
            PyList_Append(failedItems,Py_BuildValue("i",i));
            continue;
         }
      }

      d4appendStart(data,0);
      d4blank(data);

      // assign each field in the Dictionary
      assignFields(data,record);
      if (PyErr_Occurred())
      {
         PyErr_Clear();
         PyList_Append(failedItems,Py_BuildValue("i",i));
         continue;
      }

      rc = d4append(data);
      if (rc == r4success)
         numSuccessful++;
      else
      {
         data->codeBase->errorCode = 0;
         PyList_Append(failedItems,Py_BuildValue("i",i));
      }
   }

   if (closeIndex)
   {
      d4flush(data);
      i4open(data,0);
      d4reindex(data);
   }

   retVal = Py_BuildValue("iO",numSuccessful,failedItems);
   Py_DECREF(failedItems);
   UnlockHandle(cb);
   return retVal;
}

/******* METHOD TABLE *******/

/* The PyMethodDef array is the Method Table (http://www.python.org/doc/current/ext/methodTable.html).
   The Method Table must have an entry for every function
   that is to be called by the Python interpretor. */
static PyMethodDef cbMethods[] =
{
   { "code4accessMode",cb_code4accessMode,METH_VARARGS },
   { "code4autoOpen",cb_code4autoOpen,METH_VARARGS },
   { "code4close",cb_code4close,METH_VARARGS },
   { "code4codePage",cb_code4codePage,METH_VARARGS },
   { "code4collate",cb_code4collate,METH_VARARGS },
   { "code4collateUnicode",cb_code4collateUnicode,METH_VARARGS },
   { "code4compatibility",cb_code4compatibility,METH_VARARGS },
   { "code4createTemp",cb_code4createTemp,METH_VARARGS },
   { "code4dateFormat",cb_code4dateFormat,METH_VARARGS },
   { "code4errDefaultUnique",cb_code4errDefaultUnique,METH_VARARGS },
   { "code4indexExtension",cb_code4indexExtension,METH_VARARGS },
   { "code4init",cb_code4init,METH_VARARGS },
   { "code4initUndo",cb_code4initUndo,METH_VARARGS },
   { "code4lockAttempts",cb_code4lockAttempts,METH_VARARGS },
   { "code4readLock",cb_code4readLock,METH_VARARGS },
   { "code4readOnly",cb_code4readOnly,METH_VARARGS },
   { "code4safety",cb_code4safety,METH_VARARGS },
   { "code4singleOpen",cb_code4singleOpen,METH_VARARGS },
   { "d4append",cb_d4append,METH_VARARGS },
   { "d4appendStart",cb_d4appendStart,METH_VARARGS },
   { "d4blank",cb_d4blank,METH_VARARGS },
   { "d4bof",cb_d4bof,METH_VARARGS },
   { "d4bottom",cb_d4bottom,METH_VARARGS },
   { "d4close",cb_d4close,METH_VARARGS },
   { "d4create",cb_d4create,METH_VARARGS },
   { "d4delete",cb_d4delete,METH_VARARGS },
   { "d4deleted",cb_d4deleted,METH_VARARGS },
   { "d4eof",cb_d4eof,METH_VARARGS },
   { "d4field",cb_d4field,METH_VARARGS },
   { "d4fieldJ",cb_d4fieldJ,METH_VARARGS },
   { "d4flush",cb_d4flush,METH_VARARGS },
   { "d4go",cb_d4go,METH_VARARGS },
   { "d4memoCompress",cb_d4memoCompress,METH_VARARGS },
   { "d4numFields",cb_d4numFields,METH_VARARGS },
   { "d4open",cb_d4open,METH_VARARGS },
   { "d4pack",cb_d4pack,METH_VARARGS },
   { "d4position",cb_d4position,METH_VARARGS },
   { "d4recall",cb_d4recall,METH_VARARGS },
   { "d4recCount",cb_d4recCount,METH_VARARGS },
   { "d4recNo",cb_d4recNo,METH_VARARGS },
   { "d4recWidth",cb_d4recWidth,METH_VARARGS },
   { "d4refreshRecord",cb_d4refreshRecord,METH_VARARGS },
   { "d4reindex",cb_d4reindex,METH_VARARGS },
   { "d4remove",cb_d4remove,METH_VARARGS },
   { "d4seek",cb_d4seek,METH_VARARGS },
   { "d4seekDouble",cb_d4seekDouble,METH_VARARGS },
   { "d4seekNext",cb_d4seekNext,METH_VARARGS },
   { "d4seekNextDouble",cb_d4seekNextDouble,METH_VARARGS },
   { "d4skip",cb_d4skip,METH_VARARGS },
   { "d4tag",cb_d4tag,METH_VARARGS },
   { "d4tagSelect",cb_d4tagSelect,METH_VARARGS },
   { "d4top",cb_d4top,METH_VARARGS },
   { "d4unlock",cb_d4unlock,METH_VARARGS },
   { "d4zap",cb_d4zap,METH_VARARGS },
   { "f4assign",cb_f4assign,METH_VARARGS },
   { "f4assignChar",cb_f4assignChar,METH_VARARGS },
   { "f4assignDouble",cb_f4assignDouble,METH_VARARGS },
   { "f4assignInt",cb_f4assignInt,METH_VARARGS },
   { "f4assignLong",cb_f4assignLong,METH_VARARGS },
   { "f4blank",cb_f4blank,METH_VARARGS },
   { "f4char",cb_f4char,METH_VARARGS },
   { "f4decimals",cb_f4decimals,METH_VARARGS },
   { "f4double",cb_f4double,METH_VARARGS },
   { "f4int",cb_f4int,METH_VARARGS },
   { "f4len",cb_f4len,METH_VARARGS },
   { "f4long",cb_f4long,METH_VARARGS },
   { "f4name",cb_f4name,METH_VARARGS },
   { "f4str",cb_f4str,METH_VARARGS },
   { "f4true",cb_f4true,METH_VARARGS },
   { "f4type",cb_f4type,METH_VARARGS },
   { "i4create",cb_i4create,METH_VARARGS },
   { "i4open",cb_i4open,METH_VARARGS },

   { "Close",nd_Close,METH_VARARGS },
   { "Delete",nd_Delete,METH_VARARGS },
   { "Deletes",nd_Deletes,METH_VARARGS },
   { "Drop",nd_Drop,METH_VARARGS },
   { "Find",nd_Find,METH_VARARGS },
//   { "FindNext",nd_FindNext,METH_VARARGS },
   { "Open",nd_Open,METH_VARARGS },
   { "Read",nd_Read,METH_VARARGS },
   { "Updates",nd_Updates,METH_VARARGS },
   { "Write",nd_Write,METH_VARARGS },
   { "Writes",nd_Writes,METH_VARARGS },
   { 0,0 }
};

#ifdef S4WIN32
void _declspec(dllexport) initpython4c(void)
#else
extern "C" initpython4c(void)
#endif
{
   Py_InitModule("python4c", cbMethods);
}
