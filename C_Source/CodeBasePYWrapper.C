/* --------------------------------------- November 12, 2008 - 12:40pm -------------------------
 * C language program for new modules for PlanTools(tm) and BidTools(tm) accessed from Python.
 * --------------------------------------------------------------------------------------------- */
 /* *****************************************************************************
Copyright (C) 2010-2018 by M-P Systems Services, Inc.,
PMB 136, 1631 NE Broadway, Portland, OR 97232.

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Lesser General Public License as published by
the Free Software Foundation, version 3 of the License.
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General
Public License for more details.

You should have received a copy of the GNU Lesser General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
******************************************************************************** */
/* 09/30/2012 - Added support for large table mode supported by CodeBase allowing tables much larger than 2GB.  When
   this mode is active, this module should NOT attempt to open and access tables simultaneously with Visual FoxPro applications.
   The result of trying to do that could be drastic data corruption.  Tables opened in this mode will still be openable by VFP
   until they grow in size beyond that supported by VFP after which point VFP will generate an error when attempting to open them. */
   
/* October, 2014, rebuilt this module to use Python functions and to be compiled to a .PYD file for direct use by
   Python modules. Jim Heuer, M-P Systems Services, Inc. */
   
/* Jan. 9, 2015 - Added cbwCOPYTAGS(). JSH. */
#define TRUE 1
#define FALSE 0
#define DOEXPORT __declspec(dllexport)
#define BUFFER_LEN 8920000
#define MAXDATASESSIONS 32
#define ERRORMSGSIZE 1000
#define SEEKSTRINGSIZE 1000
#define DATEFMATSIZE 20
#define QUERYALIASSIZE 100
#define MAXTEMPINDEXES 200
#define MAXINDEXTAGS 512
#define MAXEXCLCOUNT 300
#define MAXFIELDCOUNT 513
#define MAXDELIMCOUNT 12
#define MAXALIASNAMESIZE 100
#define MAXOPENTABLECOUNT 500

#include <Python.h>
#include <datetime.h>
#include <d4all.h>
#include <time.h>
#include <stdio.h>
#include <stdlib.h>
#include <math.h>

#if PY_MAJOR_VERSION >= 3
#define IS_PY3K
#endif

#if PY_MAJOR_VERSION < 3
#define IS_PY2K
#endif


long glInitDone = FALSE;
long gnMaxTempTags = 20;
long gnMaxTempIndexes = MAXTEMPINDEXES;
long gnOpenTempIndexes = 0;
long gbLargeMode = FALSE;
CODE4 codeBase; /* Currently active code base structure. Copied from the selected item in the gaCodeBases[] array. */
CODE4 gaCodeBases[MAXDATASESSIONS];
long  gnCodeBaseFlags[MAXDATASESSIONS];
long  gnCodeBasesCount = MAXDATASESSIONS;
long  gnCodeBaseItem = -1;

DATA4 *gpCurrentTable = NULL;
TAG4  *gpCurrentIndexTag = NULL; /* Future use */
EXPR4 *gpCurrentFilterExpr = NULL;
FIELD4 *gpLastField = NULL;
RELATE4 *gpCurrentQuery = NULL;
char gcQueryAlias[QUERYALIASSIZE]; /* Stores the table alias to which the gpCurrentQuery applies.  If cbwCONTINUE() is
                          called and a different table is currently selected, an error is reported. */

long gnDeletedFlag = FALSE; /* If TRUE, then will attempt to skip over deleted records. */
char* cTextBuffer = NULL; /* Allocate a bunch of space to this when starting up. */
char gcErrorMessage[ERRORMSGSIZE]; /* Contains the most recent error message.  Guaranteed to be less than 1000 characters in length. */
long gnLastErrorNumber = 0;
char gcStrReturn[BUFFER_LEN]; /* Where we put stuff for static string return */
long gnProcessTally = 0; /* Keeps track of the TALLY of last records processed like VFP _TALLY. */
char gcDelimiter[MAXDELIMCOUNT];
char gcLastSeekString[SEEKSTRINGSIZE]; /* We store the last SEEK string value for future needs. */
char gcDateFormatCode[DATEFMATSIZE];
unsigned char gaOtherLetters1252[7] = {138, 140, 142, 154, 156, 158, 159}; // Extended ASCII, cp1252 isolated characters.
long gnOtherLetterCount = 7;
char gcCustomCodePage[50];

typedef struct VFPINDEXTAGtd {
	char cTagName[130];
	char cTagExpr[256];
	char cTagFilt[256];
	long nDirection;
	long nUnique;
} VFPINDEXTAG;
VFPINDEXTAG gaIndexTags[MAXINDEXTAGS]; /* The max we will allow, and well over the expected typical numbers */

typedef struct TEMPINDEXtd {
	INDEX4 *pIndex;
	char cFileName[255];
	char cTableAlias[100];
	VFPINDEXTAG aTags[20];
	long bOpen;
	long nTagCount;
	DATA4 *pTable;
} TEMPINDEX;
TEMPINDEX gaTempIndexes[MAXTEMPINDEXES];

typedef struct CODEBASESTATUStd {
	DATA4 *pCurrentTable;
	TAG4  *pCurrentIndexTag; /* Future use */
	EXPR4 *pCurrentFilterExpr;
	FIELD4 *pLastField;
	RELATE4 *pCurrentQuery;
	long bLargeMode;
	char cQueryAlias[QUERYALIASSIZE]; /* Stores the table alias to which the gpCurrentQuery applies.  If cbwCONTINUE() is
	called and a different table is currently selected, an error is reported. */
	
	long nDeletedFlag; /* If TRUE, then will attempt to skip over deleted records. */
	char cErrorMessage[ERRORMSGSIZE]; /* Contains the most recent error message.  Guaranteed to be less than 1000 characters in length. */
	long nLastErrorNumber;
	char cDelimiter[MAXDELIMCOUNT];
	char cLastSeekString[SEEKSTRINGSIZE]; /* We store the last SEEK string value for future needs. */
	char cDateFormatCode[DATEFMATSIZE];	
	TEMPINDEX aTempIndexes[MAXTEMPINDEXES];
	VFPINDEXTAG aIndexTags[MAXINDEXTAGS];
	char *aExclList[MAXEXCLCOUNT];
} CODEBASESTATUS;
CODEBASESTATUS gaCodeBaseEnvironments[MAXDATASESSIONS];

typedef struct VFPFIELDtd {
	char cName[132];
	char cType[4];
	long nWidth;
	long nDecimals;
	long bNulls;	
} VFPFIELD;
VFPFIELD gaFields[MAXFIELDCOUNT];

typedef struct FieldPlusTd {
	VFPFIELD xField;
	long nFldType;
	long bMatched;
	long nSource; // Index into the Source array. Used in gaMatchFields
	long nTarget; // Index into the Target array.  Used in gaMatchFields
	long bIdentical; // Set to TRUE if the type, width, and decimals are all identical.
	FIELD4 *pCBfield;
} FieldPlus;
FieldPlus gaSrcFields[MAXFIELDCOUNT];
long gnSrcCount;
long gnTrgCount;
FieldPlus gaTrgFields[MAXFIELDCOUNT];
FieldPlus gaMatchFields[MAXFIELDCOUNT];

FIELD4INFO gaFieldSpecs[MAXFIELDCOUNT];
// FIELD4INFO members:
// char *name;
// short type;
// unsigned short len;
// unsigned short dec;
// unsigned short nulls/ 

char* gaExclList[MAXEXCLCOUNT]; // A list of alias values of currently open tables with the Exclusive property set to TRUE.
// We do this since the CodeBase engine doesn't give us a good way to test whether a table was opened in exclusive
// mode or not. Added 10/25/2013. JSH.  Alias strings are ALWAYS lower CASE.
long gnExclMax = MAXEXCLCOUNT;
long gnExclLimit = 20;
long gnExclCnt = 0;

char *gaFieldList[MAXFIELDCOUNT];

//typedef struct AliasTrackTD {
//	char cAlias[MAXALIASNAMESIZE];
//	char cFullName[400];
//	DATA4 *TableID;
//} AliasTrack;
//
//AliasTrack gaAliasTrack[MAXOPENTABLECOUNT];


long cbxCLOSETABLE(char *);

/* ******************************************************************************** */
/* Generic converter from plain ASCII string literal to a Python string object.     */
static PyObject* stringToPy(const char *cStr)
{
	static PyObject *lxStr = NULL;	
#ifdef IS_PY3K
	lxStr = PyUnicode_FromString(cStr);
#endif
#ifdef IS_PY2K
	lxStr = PyString_FromString(cStr);
#endif	
return(lxStr);
}

/* ******************************************************************************** */
/* Returns a Python string value from a CodeBase field name, applying the decoding  */
/* necessary for both PY versions.  Note that in VFP, a field name can have extended*/
/* characters with a Code Page of 1252.                                             */
static PyObject* cbNameToPy(const char *cName)
{
	static PyObject *lxName;
	size_t lnTest = 0;

#ifdef IS_PY3K					
	lnTest = strlen(cName);	
	lxName = PyUnicode_DecodeLatin1((const char*) cName, lnTest, "replace");
	if (lxName == (void*) NULL)
		{
		strcpy(gcErrorMessage, "Illegal character found in field name, not defined in Code Page 1252.");
		gnLastErrorNumber = -10011;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		// error out now to avoid General protection fault on 0 pointer.	
		return NULL;
		}
#endif
#ifdef IS_PY2K					
	lxName = PyString_FromString((unsigned char*) cName);
	if (lxName == (void*) NULL)
		{
		strcpy(gcErrorMessage, "Illegal character found in field name, not convertable to string.");
		gnLastErrorNumber = -10011;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		// error out now to avoid General protection fault on 0 pointer.	
		return NULL;
		}						
#endif
	return(lxName);
}

#ifdef IS_PY3K
/* ******************************************************************************** */
/* Unicode utf-8 converter to extended ascii with the standard codepage 1252.
   Note that CodeBase ONLY supports CodePage 1252, Western European languages.      */
const char *Unicode2Char(PyObject* cFromUnicode)
{
	return(PyBytes_AsString(PyUnicode_AsEncodedString(cFromUnicode, "cp1252", "replace")));
}

/* ******************************************************************************** */
/* Converts a plain C string of char in utf-8 format as passed to functions back to
   Extended Ascii using Code Page 1252.                                             */
const char *Char2ExtendedAscii(char* cFromChar)
{
	PyObject* cIsUnicode;
	cIsUnicode = Py_BuildValue("s", cFromChar);
	return(Unicode2Char(cIsUnicode));
}

const unsigned char *testStringTypes(PyObject* pxValue)
{
if (!PyUnicode_Check(pxValue)) return("Parameter must be Unicode String.");
else return("");
}
#endif   


#ifdef IS_PY2K
/* ******************************************************************************** */
/* Unicode utf-8 converter to extended ascii with the standard codepage 1252.
   Note that CodeBase ONLY supports CodePage 1252, Western European languages.      */
const char *Unicode2Char(PyObject* cFromUnicode)
{
	PyObject *xResult;
	char *cStr;
	if (PyUnicode_Check(cFromUnicode))
		{
		xResult = PyUnicode_AsEncodedObject(cFromUnicode, "cp1252", "replace");
		cStr = PyString_AsString(xResult);
		}
	else // string type
		{
		cStr = PyString_AsString(cFromUnicode);
		}
	
	return(cStr);
}

/* ******************************************************************************** */
/* Converts a plain C string of char in utf-8 format as passed to functions back to
   Extended Ascii using Code Page 1252.                                             */
const char *Char2ExtendedAscii(char* cFromChar)
{
	PyObject* cIsUnicode;
	cIsUnicode = Py_BuildValue("s", cFromChar);
	return(Unicode2Char(cIsUnicode));
}

const unsigned char *testStringTypes(PyObject* pxValue)
{
if (!PyString_Check(pxValue) && !PyUnicode_Check(pxValue)) return("Parameter must be String or Unicode");
else return("");
}
#endif

/* ****************************************************************************** */
/* Utility function for changing all letters in a string to lower case            */
char* StrToLower(char* lpcStr)
{
	char* pOrig = lpcStr;
	for (lpcStr; *lpcStr != (char) 0; lpcStr++)
		{
		*lpcStr = tolower(*lpcStr);	
		}
	return(pOrig);
}

///* ****************************************************************************** */
///* Stores the alias name and a table identifier to the AliasTrack table.          */
///* Returns 1 on OK, 0 on failure -- typically no available spaces in the table.   */
//long StoreAliasToTracker(DATA4* lppTable, char* lpcAlias, char* lpcFullName)
//{
//	register long jj;
//	long lnReturn = 0;
//	for (jj = 0; jj < MAXOPENTABLECOUNT; jj++)
//		{
//		if (gaAliasTrack[jj].TableID == NULL)
//			{
//			gaAliasTrack[jj].TableID = lppTable;
//			strncpy(gaAliasTrack[jj].cAlias, lpcAlias, MAXALIASNAMESIZE);
//			strncpy(gaAliasTrack[jj].cFullName, lpcFullName, 399);
//			StrToLower(gaAliasTrack[jj].cFullName);
//			lnReturn = 1;
//			break;	
//			}	
//		}
//	printf("Stored Alias %s  %p\n", lpcAlias, lppTable);
//	return lnReturn;
//}
//
///* ****************************************************************************** */
///* Returns the table ID pointer for a currently open table based on the full path */
///* of the table.  Returns the DATA4* pointer or NULL if not found.  If the table  */
///* appears more than once due to multiple opens of the same file, only the first  */
///* DATA4* value found is returned.                                                */
//DATA4* GetTableFromName(char* lpcFileName)
//{
//	register long jj;
//	DATA4* lpTable = NULL;
//	char lcTestName[400];
//	
//	strcpy(lcTestName, lpcFileName);
//	StrToLower(lcTestName);
//	
//	for (jj = 0; jj < MAXOPENTABLECOUNT; jj++)
//		{
//		if (gaAliasTrack[jj].TableID)
//			{
//			if (strcmp(gaAliasTrack[jj].cFullName, lcTestName) == 0)
//				{
//				lpTable = gaAliasTrack[jj].TableID;
//				break;	
//				}	
//			}	
//		}
//	printf("Got table for %s = %p \n", lpcFileName, lpTable);
//	return lpTable;		
//}
//
///* ****************************************************************************** */
///* Returns the table ID pointer for a specified alias name.                       */
///* Returns NULL if there is no such alias in the tracking table.                  */
//DATA4* GetTableFromAlias(char* lpcAlias)
//{
//	register long jj;
//	DATA4* lpTable = NULL;
//	
//	for (jj = 0; jj < MAXOPENTABLECOUNT; jj++)
//		{
//		if (gaAliasTrack[jj].TableID)
//			{
//			if (strcmp(gaAliasTrack[jj].cAlias, lpcAlias) == 0)
//				{
//				lpTable = gaAliasTrack[jj].TableID;
//				break;	
//				}	
//			}	
//		}
//	printf("Got table for %s = %p \n", lpcAlias, lpTable);
//	return lpTable;		
//}
//
///* ****************************************************************************** */
///* Removes a specified alias from the alias tracker table.                        */
///* Returns the Table ID pointer on success, returns NULL on not found.            */
//DATA4* ClearAliasFromTracker(char* lpcAlias)
//{
//	register long jj;
//	DATA4* lpTable = NULL;
//	for (jj = 0; jj < MAXOPENTABLECOUNT; jj++)
//		{
//		if (gaAliasTrack[jj].TableID)
//			{
//			if (strcmp(gaAliasTrack[jj].cAlias, lpcAlias) == 0)
//				{
//				lpTable = gaAliasTrack[jj].TableID;
//				gaAliasTrack[jj].TableID = (DATA4*) NULL;
//				gaAliasTrack[jj].cAlias[0] = (char) 0;
//				break;	
//				}
//			}	
//		}
//	printf("Cleared Alias %s Table %p \n", lpcAlias, lpTable);
//	return lpTable;
//} 
//
///* ****************************************************************************** */
///* Returns the alias associated with a table identifier.                          */
///* Returns a pointer to the static string or NULL on failure.                     */
//char* GetTableAlias(DATA4* lppTable)
//{
//	static char cTheAlias[MAXALIASNAMESIZE + 1];
//	register long jj;
//	long lnFound = FALSE;
//	
//	memset(cTheAlias, 0, (size_t) (MAXALIASNAMESIZE + 1) * sizeof(char));
//	for (jj = 0; jj < MAXOPENTABLECOUNT; jj++)
//		{
//		if (gaAliasTrack[jj].TableID == lppTable)
//			{
//			strcpy(cTheAlias, gaAliasTrack[jj].cAlias);
//			lnFound = TRUE;
//			break;				
//			}
//		}	
//	if (lnFound) return cTheAlias;
//	else return NULL;
//}

/* ****************************************************************************** */
/* Utility function for changing all letters in a string to UPPER case            */
char* StrToUpper(char* lpcStr)
{
	char* pOrig = lpcStr;
	for (lpcStr; *lpcStr != (char) 0; lpcStr++)
		{
		*lpcStr = toupper(*lpcStr);	
		}	
	return(pOrig);
}

/* ****************************************************************************** */
/* Extension of the strtok() library function that looks for a delimiter of       */
/* one or more adjacent characters, for example '@#' or '~~' to separate parts    */
/* of a string into their components.  Works like strtok().  Call it with a non-  */
/* null char* value as the string to parse and a char* string with the null-      */
/* terminated delimiter string.  Returns a pointer to a static string buffer with */
/* the first token in the source string.  Call it again with a first parameter of */
/* NULL, and it returns the next, and the next, etc.  When it returns NULL we're  */
/* done.                                                                          */
/* NOTE: The difference between strtok() and strtokX() is that in strtok() any one*/
/* of the values in the lpcDelim string can act as a delimiter.  The first one    */
/* found breaks the string.  In strtokX() ALL the characters in their exact order */
/* must be found for a delimiter to be "found".  The delimiter characters are     */
/* discarded in both versions.                                                    */
char* strtokX(char* lpcSource, char* lpcDelim)
{
	static char* lpStr; /* Pointer to allocated memory where we'll hold the working source string */
	static char* lpPtr; /* Pointer we'll move through the source string as we tokenize it */
	char* lpTest;
	char* lpReturn;
	
	if (lpcSource != NULL)
		{
		if (lpStr != NULL) free(lpStr);
		lpStr = malloc((size_t) strlen(lpcSource) + 1);
		strcpy(lpStr, lpcSource);
		lpPtr = lpStr;
		}
	else
		{
		if (lpPtr == NULL) return(NULL); /* Nothing to do */	
		}
	if (*lpPtr == (char) 0)
		{
		/* It's now pointing to the null at the end of the source string.  We're done. */
		lpReturn = lpPtr;
		lpPtr = NULL; /* don't try calling this again, cause we're done. */
		}
	else
		{
		lpTest = strstr(lpPtr, lpcDelim);
		if (lpTest == NULL)
			{
			lpReturn = lpPtr;
			//lpPtr += strlen(lpPtr);	
			lpPtr = NULL; // that was the last one.
			}
		else
			{
			lpReturn = lpPtr;
			*lpTest = (char) 0;
			lpPtr = lpTest + strlen(lpcDelim);	
			}
		}
	return(lpReturn);
}

/* ******************************************************************************* */
/* Returns a pointer to a static char buffer which contains just the path name,    */
/* minus the trailing back slash from any fully qualified file name.               */
static char* justFilePath(char* lpcFileName)
{
	static char lcName[500];
	register long jj;
	long lnLen;
	memset(lcName, 0, (size_t) 500 * sizeof(char));
	strcpy(lcName, lpcFileName);
	lnLen = strlen(lcName);
	for (jj = lnLen - 1; jj >= 0; jj--)
		{
		if (lcName[jj] != (char) 92)
			{
			lcName[jj] = (char) 0;
			}
		else
			{
			lcName[jj] = (char) 0;
			break; // Note we convert that first backslash we encounter too..
			}
		}
	return lcName;
}

/* ******************************************************************************* */
/* Utility function that serves the role of VFP ALLTRIM(), LTRIM(), and RTRIM()    */
/* by trimming the ends, the left, or the right of a string, removing blanks.      */
/* The cHow parameter must be a character string 'A', 'L', or 'R' which indicates  */
/* ALL, LEFT, or RIGHT trimming.  The cString is the pointer to the source string. */
/* Since the resulting string can only either be the same length or smaller, this  */
/* modification is done in place.  The return value is the cString pointer.        */
/* NOTE: ONLY space characters are removed.  Other white space characters like TAB */
/* are left in place.                                                              */
char* MPStrTrim(char cHow, char* cString)
{
	long lnLen;
	register long jj;
	char *pPtr;
	long lbAtStart = TRUE;
	long lnStartSpaceCount = 0;
	
	lnLen = strlen(cString);
	if ((cHow == 'A') || (cHow == 'L'))
		{
		pPtr = cString;
		for (jj = 0; jj <= lnLen; jj++)
			{
			if (!lbAtStart)
				{
				*(pPtr++) = *(cString + jj);
				}
			else
				{
				if (*(cString + jj) != ' ')
					{
					lbAtStart = FALSE;
					*(pPtr++) = *(cString + jj);
					}
				else
					{
					lnStartSpaceCount += 1;	
					}
				}
			}	
		}
	
	if ((cHow == 'A') || (cHow == 'R'))
		{
		lnLen = lnLen - lnStartSpaceCount; // Adjust for what we've already removed.	
		for (jj = lnLen; jj >= 0; jj--)
			{
			if (((*(cString + jj)) == ' ') || ((*(cString + jj)) == (char) NULL))
				{
				(*(cString + jj)) = '\0';
				}
			else break;				
			}
		}
		
	return cString;	
}

/* ******************************************************************************* */
/* Utility function that produces a string of random uppercase letters and numbers */
/* of a specified length.  Returns a pointer to a static buffer which will change  */
/* every time this is called, so copy it out immediately.                          */
/* Max length of 255 characters.  If lnLen is greater then 255, still returns 255. */
static char* MPStrRand(long lpnLen)
{
	static char lcStr[256];
	static char lcChars[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
	long lnLen = 0;
	double lnChrCnt = 0.0;
	long jj;
	long lnVal = 0;
	
	lnLen = lpnLen;
	if (lnLen > 255) lnLen = 255;
	lnChrCnt = (double) strlen(lcChars);

	for (jj = 0;jj < lnLen; jj++)
		{
		lnVal = floor((double) rand() * lnChrCnt / (double) RAND_MAX);
		lcStr[jj] = lcChars[lnVal];
		}
	lcStr[lnLen] = (char) 0;
	return lcStr;	
}


/* ****************************************************************************** */
/* Utility function that, like STRTRAN() in VFP replaces a specified string in    */
/* the source with an alternative string.  This is storage safe in that it does   */
/* the translation in an allocated buffer big enough to handle the potentially    */
/* larger result.  Then returns a pointer to that static buffer.  Copy the result */
/* out of the buffer immediately, since it is replaced every time this is run.    */
/* If the alternative string is empty, then removes the specified string from the */
/* source string.                                                                 */
/* Supply a non-zero value of lpnCount to limit the number of replacements to     */
/* that value -- otherwise all instances are replaced.                            */
char* MPSSStrTran(char* lpcSource, char* lpcTarget, char* lpcAlternative, long lpnCount)
{
	static char* lpBuff;
	char* lpPtr;
	char* lpStart;
	long lnFromLen, lnToLen;
	long lnDiffLen;
	long lnNewBuffLen;
	long lnOccurances = 0;
	long lnCnt = 0;
	long lnStart = TRUE;
	
	if (lpBuff != NULL) free(lpBuff);
	lpBuff = NULL;
	lnFromLen = strlen(lpcTarget);
	lnToLen = strlen(lpcAlternative);
	lnDiffLen = lnToLen - lnFromLen;
	if (lnDiffLen > 0) /* The result can be longer than the source string, so we have to make a larger buffer */
		{
		lpStart = lpcSource;
		while(TRUE)
			{
			lpPtr = strstr(lpStart, lpcTarget);
			if (lpPtr == NULL)
				break;
			lpStart = lpPtr;
			lnOccurances += 1;
			}
		lnNewBuffLen = strlen(lpcSource) + 1 + (lnOccurances * lnDiffLen); /* worst case */
		}
	else
		{
		lnNewBuffLen = strlen(lpcSource) + 1;	
		}
	lnNewBuffLen += 15; /* Just a little safety factor */
	lpBuff = malloc((size_t) lnNewBuffLen * sizeof(char));
	memset(lpBuff, 0, (size_t) lnNewBuffLen * sizeof(char));
	lpStart = lpcSource;
	while (TRUE)
		{
		lpPtr = strstr(lpStart, lpcTarget);
		if ((lpPtr == NULL) || ((lpnCount > 0) && (lnCnt >= lpnCount)))
			{
			if (lnStart) strcpy(lpBuff, lpStart);
				else strcat(lpBuff, lpStart);
			break;	
			}
		*lpPtr = (char) 0;
		if (lnStart) strcpy(lpBuff, lpStart);
			else strcat(lpBuff, lpStart);
		lnStart = FALSE;
		strcat(lpBuff, lpcAlternative);
		lpStart = lpPtr + lnFromLen;
		lnCnt += 1;
		}
	return(lpBuff);	
}

/* ****************************************************************************** */
/* Custom function that replaces stdlib isalnum for strings coded for CodePage
   1252 (Windows Western European)                                                */
long isAlnum1252(unsigned char cByte)
{
	long nRet = 0;
	register long jj;
	if (cByte >= 48 && cByte <= 57) return(1);
	if (cByte >= 65 && cByte <= 90) return(1);
	if (cByte >= 97 && cByte <= 122) return(1);
	if (cByte >= 192 && cByte <= 214) return(1);
	if (cByte >= 216 && cByte <= 246) return(1);
	if (cByte >= 249 && cByte <= 255) return(1);
	for (jj = 0; jj < gnOtherLetterCount; jj++)
		{
		if (gaOtherLetters1252[jj] == cByte) return(1);
		}
	return(nRet);
}

/* ****************************************************************************** */
/* Converts cp1252 strings to plain ascii by using the closest match for high
   order non-ASCII cgaracters to Latin alphabet characters.  This function makes
   the substitutions in-place, as the resulting string is guaranteed not to get
   longer.                                                                        */
unsigned char *Conv1252ToASCII(unsigned char *c1252, long bRemoveNonPrint)
{
	static unsigned char aMap[256];
	register long jj;
	if (aMap[1] == 0) // first time thru.
		{
		for (jj = 0; jj <= 127; jj++)
			{
			aMap[jj] = jj;
			aMap[jj + 128] = 95; // underscore
			}
		aMap[138] = 'S';
		aMap[140] = 'O';
		aMap[142] = 'Z';
		aMap[154] = 's';
		aMap[156] = 'o';
		aMap[158] = 'z';
		aMap[159] = 'Y';
		aMap[192] = aMap[193] = aMap[194] = aMap[195] = aMap[196] = aMap[197] = aMap[198] = 'A';
		aMap[199] = 'C';
		aMap[200] = aMap[201] = aMap[202] = aMap[203] = 'E';
		aMap[204] = aMap[205] = aMap[206] = aMap[207] = 'I';
		aMap[208] = 'D';
		aMap[209] = 'N';
		aMap[210] = aMap[211] = aMap[212] = aMap[213] = aMap[214] = 'O';
		aMap[217] = aMap[218] = aMap[219] = aMap[220] = 'U';
		aMap[221] = 'Y';
		aMap[222] = 'D';
		aMap[223] = 'B';
		aMap[224] = aMap[225] = aMap[226] = aMap[227] = aMap[228] = aMap[229] = aMap[230] = 'a';
		aMap[231] = 'c';
		aMap[232] = aMap[233] = aMap[234] = aMap[235] = 'e';
		aMap[236] = aMap[237] = aMap[238] = aMap[239] = 'i';
		aMap[240] = aMap[242] = aMap[243] = aMap[244] = aMap[245] = aMap[246] = 'o';
		aMap[241] = 'n';
		aMap[249] = aMap[250] = aMap[251] = aMap[252] = 'u';
		aMap[253] = 'y';		
		aMap[254] = 'p';		
		aMap[223] = 'y';
		}
	jj = 0;
	while (TRUE)
	{
		if (*(c1252 + jj) == 0) break;
		if (bRemoveNonPrint)
			{
			if (*(c1252 + jj) < 32)
				{
				*(c1252 + jj) = '_';
				continue;	
				} 
			}
		*(c1252 + jj) = aMap[*(c1252 + jj)];
		jj += 1;
	}
	return(c1252);
}


/* ****************************************************************************** */
/* Takes the base table name and makes sure that it is a legal alias name.        */
/* We'll use VFP name restrictions: Must start with a letter or underscore, may   */
/* consist of up to 254 letters, numbers, and underscore characters.  Spaces and  */
/* other punctuation are not allowed.                                             */
/* Returns a pointer to a static string buffer which is updated every time this   */
/* is run, so copy results out immediately.                                       */
static unsigned char* MakeLegalAlias(unsigned char* lpcBaseName)
{
	long lnOutCount = 0;
	long jj;
	static unsigned char lcGoodAlias[300];
	unsigned char testChar;
	unsigned char lcWorkAlias[300];
	long lnMaxSize = MAXALIASNAMESIZE;
	long lnWorkLen;
	
	memset(lcGoodAlias, 0, (size_t) 300);
	strncpy(lcWorkAlias, MPStrTrim('A', lpcBaseName), 299);
	lcWorkAlias[299] = (char) 0;
	Conv1252ToASCII(lcWorkAlias, TRUE);
	lnWorkLen = strlen(lcWorkAlias);
	if (isdigit(lcWorkAlias[0]))
		{
		lcGoodAlias[lnOutCount++] = '_';
		}
	for (jj = 0; jj < lnWorkLen; jj++)
		{
		testChar = lcWorkAlias[jj];
		if ((testChar == '_') || isalnum(testChar))
			{
			lcGoodAlias[lnOutCount++] = toupper(testChar);	
			}
		if (lnOutCount >= lnMaxSize)
			{
			lcGoodAlias[lnOutCount] = (unsigned char) 0;
			break;	
			}
		}
	//printf("NEW ALIAS >%s<\n", lcGoodAlias);
	return lcGoodAlias;	
}


/* ****************************************************************************** */
/* Initializes the various globals that keep track of stuff while a data session  */
/* is active.                                                                     */
void cbxInitGlobals(void)
{
	long jj;
	
	memset(gcErrorMessage, 0, (size_t) (ERRORMSGSIZE * sizeof(char)));
	memset(gcDelimiter, 0, (size_t) (MAXDELIMCOUNT * sizeof(char)));
	strcpy(gcDelimiter, "||");
	memset(gcLastSeekString, 0, (size_t) (SEEKSTRINGSIZE * sizeof(char)));
	memset(gcDateFormatCode, 0, (size_t) (DATEFMATSIZE * sizeof(char)));
	memset(gcQueryAlias, 0, (size_t) (QUERYALIASSIZE * sizeof(char)));
	memset(gcCustomCodePage, 0, (size_t) (50 * sizeof(char)));
	strcpy(gcCustomCodePage, "cp1252");  // default, so as not too crash if not set.
	gpCurrentTable = NULL;
	gpCurrentIndexTag = NULL;
	gpCurrentFilterExpr = NULL;
	gpLastField = NULL;
	gpCurrentQuery = NULL;
	gnLastErrorNumber = 0;

	glInitDone = FALSE;
	gnProcessTally = 0;
	gbLargeMode = FALSE;
	memset(gaFields, 0, (size_t) 256 * sizeof(VFPFIELD)); // These aren't carried over to another session.
	memset(gaIndexTags, 0, (size_t) (MAXINDEXTAGS * sizeof(VFPINDEXTAG)));
	memset(gaTempIndexes, 0, (size_t) (MAXTEMPINDEXES * sizeof(TEMPINDEX)));
	for (jj = 0; jj < MAXEXCLCOUNT; jj++)
		{
		strcpy(gaExclList[jj], "");
		}
}

/* ****************************************************************************** */
/* Returns the current Datasession number.                                        */
static PyObject *cbwGETSESSIONNUMBER(PyObject *self, PyObject *args)
{
	return Py_BuildValue("l", gnCodeBaseItem);	
}

/* ******************************************************************************  */
/* Call this first before you call anything else or the results will be BAD.       */
/* Pass a 1 in lnLargeFlag, and this will initialize the engine to use the large   */
/* table support.  Otherwise pass 0 as lnLargeFlag.  WARNING!  Tables opened with  */
/* Large Table Support can NOT be opened simultaneously by Visual FoxPro.          */
/* This returns the index key to the Datasession that has been initialized.  Store */
/* the value.  You can switch to a different data session using cbwSwitchSession(),*/
/* but if you do so you'll need to switch back to the original session to close    */
/* it down.                                                                        */
/* When done with the Datasession, call cbwCloseDataSession() to release memory.   */
//long cbwInitDataSession(long lnLargeFlag)
static PyObject *cbwINITDATASESSION(PyObject *self, PyObject *args)
{
	long lnResult;
	long lnLargeFlag;
	long lnReturn = -1;
	long jj;
	long kk;
	long lnPtr;
	CODEBASESTATUS *lpStat;

	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (!PyArg_ParseTuple(args, "l", &lnLargeFlag))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		return NULL; // Force the error
		}

	for (jj = 0; jj < MAXDATASESSIONS; jj++)
		{
		if (gnCodeBaseFlags[jj] == -1)
			{
			// We've found an open item, so we grab it;
			lnReturn = jj;
			if (gnCodeBaseItem != -1)
				{
				lnPtr = gnCodeBaseItem;
				gaCodeBases[lnPtr] = codeBase; // We store the active one before we switch to a new one.	
				
				lpStat = &(gaCodeBaseEnvironments[lnPtr]);
				lpStat->pCurrentTable = gpCurrentTable;
				lpStat->pCurrentIndexTag = gpCurrentIndexTag;
				lpStat->pCurrentFilterExpr = gpCurrentFilterExpr;
				lpStat->pLastField = gpLastField;
				lpStat->pCurrentQuery = gpCurrentQuery;
				lpStat->bLargeMode = gbLargeMode;
				strcpy(lpStat->cQueryAlias, gcQueryAlias);
				lpStat->nDeletedFlag = gnDeletedFlag;
				strcpy(lpStat->cErrorMessage, gcErrorMessage);
				lpStat->nLastErrorNumber = gnLastErrorNumber;
				strcpy(lpStat->cDelimiter, gcDelimiter);
				strcpy(lpStat->cLastSeekString, gcLastSeekString);
				strcpy(lpStat->cDateFormatCode, gcDateFormatCode);

				for (kk = 0; kk < MAXINDEXTAGS; kk++)
					{
					lpStat->aIndexTags[kk] = gaIndexTags[kk];	
					}
				for (kk = 0; kk < MAXTEMPINDEXES; kk++)
					{
					lpStat->aTempIndexes[kk] = gaTempIndexes[kk];	
					}
				for (kk = 0; kk < MAXEXCLCOUNT; kk++)
					{
					//lpStat->aExclList[kk] = gaExclList[kk];	
					strncpy(lpStat->aExclList[kk], gaExclList[kk], MAXALIASNAMESIZE);
					}
				}
			codeBase = gaCodeBases[jj];	
 			cbxInitGlobals();
 			gnCodeBaseItem = jj;
 			gnCodeBaseFlags[jj] = 1;
 			break;
			}	
		}
	if (lnReturn == -1)
		{
		// BAD PROBLEM.  NO UNUSED DATA SESSIONS!
		strcpy(gcErrorMessage, "NO AVAILABLE DATA SESSIONS");
		gnLastErrorNumber = -20000;
		return Py_BuildValue("l", -1);	
		}

	lnResult = code4init(&codeBase);
	if (lnResult >= 0)
		{
		codeBase.errOff = 1; /* No error messages displayed... just return error codes and let us handle them. */
		codeBase.safety = 0; /* allow us to overwrite things */
		code4dateFormatSet(&codeBase, "MM/DD/YYYY");
		codeBase.compatibility = 30; // Visual Foxpro 6.0 and above.
		codeBase.optimize = OPT4ALL;
		codeBase.memExpandData = 3;
		codeBase.memStartBlock = 6;
		codeBase.memStartMax = 67108864; // A lot of memory.  Can be reduced a bit if needed.
		codeBase.codePage = cp1252; // Added this 11/19/2011 to force Visual FoxPro code page.
		codeBase.lockAttempts = 2; // Tries twice to lock a record.  Up to us to try again beyond that.  Added 02/07/2012. JSH.
		codeBase.lockDelay = 2; // 2/100ths of a second. Added 02/07/2012. JSH.
		codeBase.singleOpen = FALSE; //FALSE;
		codeBase.autoOpen = TRUE;
		codeBase.ignoreCase = TRUE;
		codeBase.errDefaultUnique = r4uniqueContinue; // The default.  Shouldn't need to set this but ??
		if (lnLargeFlag == 1)
			{
			code4largeOn(&codeBase);
			gbLargeMode = TRUE;
			}
		srand(time(NULL));
		strcpy(gcDateFormatCode, "MM/DD/YY"); // The CodeBase Default
		}
	else
		{
		glInitDone = FALSE;
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		gnLastErrorNumber = codeBase.errorCode;	
		}
	return Py_BuildValue("l", lnReturn);
}

/* ****************************************************************************** */
/* Takes as its parm the Datasession index to switch to.  Returns that number on  */
/* success, otherwise -1.                                                         */
static PyObject *cbwSWITCHDATASESSION(PyObject *self, PyObject *args)
{
	long jj;
	long kk;
	long lnTargetSession;
	long lnReturn = -1;
	long lnPtr = 0;
	CODEBASESTATUS *lpStat;	
	
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
		
	if (!PyArg_ParseTuple(args, "l", &lnTargetSession))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		return NULL;
		}
	//printf("Switching Data Session to %ld \n", lnTargetSession);	
	if (gnCodeBaseItem == lnTargetSession)
		{
		// Nothing to do, session is current
		lnReturn = lnTargetSession;	
		}
	else
		{
		if (lnTargetSession > (MAXDATASESSIONS - 1))
			{
			strcpy(gcErrorMessage, "Invalid target Session Number");
			gnLastErrorNumber = -10001;
			return Py_BuildValue("l", -1);				
			}
		else
			{
			if (gnCodeBaseItem != -1)
				{
				lnPtr = gnCodeBaseItem;
				gaCodeBases[lnPtr] = codeBase; // We store the active one before we switch to a new one.	
				
				lpStat = &(gaCodeBaseEnvironments[lnPtr]);
				lpStat->pCurrentTable = gpCurrentTable;
				lpStat->pCurrentIndexTag = gpCurrentIndexTag;
				lpStat->pCurrentFilterExpr = gpCurrentFilterExpr;
				lpStat->pLastField = gpLastField;
				lpStat->pCurrentQuery = gpCurrentQuery;
				strcpy(lpStat->cQueryAlias, gcQueryAlias);
				lpStat->nDeletedFlag = gnDeletedFlag;
				strcpy(lpStat->cErrorMessage, gcErrorMessage);
				lpStat->nLastErrorNumber = gnLastErrorNumber;
				strcpy(lpStat->cDelimiter, gcDelimiter);
				strcpy(lpStat->cLastSeekString, gcLastSeekString);
				strcpy(lpStat->cDateFormatCode, gcDateFormatCode);
				lpStat->bLargeMode = gbLargeMode;

				for (kk = 0; kk < MAXINDEXTAGS; kk++)
					{
					lpStat->aIndexTags[kk] = gaIndexTags[kk];	
					}
				for (kk = 0; kk < MAXTEMPINDEXES; kk++)
					{
					lpStat->aTempIndexes[kk] = gaTempIndexes[kk];	
					}
				for (kk = 0; kk < MAXEXCLCOUNT; kk++)
					{
					//lpStat->aExclList[kk] = gaExclList[kk];	
					strncpy(lpStat->aExclList[kk], gaExclList[kk], MAXALIASNAMESIZE);
					}
				}
			codeBase = gaCodeBases[lnTargetSession];
			lpStat = &(gaCodeBaseEnvironments[lnTargetSession]);
			gpCurrentTable = lpStat->pCurrentTable;
			gpCurrentIndexTag = lpStat->pCurrentIndexTag;
			gpCurrentFilterExpr = lpStat->pCurrentFilterExpr;
			gpLastField = lpStat->pLastField;
			gbLargeMode = lpStat->bLargeMode;
			strcpy(gcQueryAlias, lpStat->cQueryAlias);
			gnDeletedFlag = lpStat->nDeletedFlag;
			strcpy(gcErrorMessage, lpStat->cErrorMessage);
			gnLastErrorNumber = lpStat->nLastErrorNumber;
			strcpy(gcDelimiter, lpStat->cDelimiter);
			strcpy(gcLastSeekString, lpStat->cLastSeekString);
			strcpy(gcDateFormatCode, lpStat->cDateFormatCode);
			
			for (kk = 0; kk < MAXINDEXTAGS; kk++)
				{
				gaIndexTags[kk] = lpStat->aIndexTags[kk];	
				}
			for (kk = 0; kk < MAXTEMPINDEXES; kk++)
				{
				gaTempIndexes[kk] = lpStat->aTempIndexes[kk];	
				}
			for (kk = 0; kk < MAXEXCLCOUNT; kk++)
				{
				//gaExclList[kk] = lpStat->aExclList[kk];	
				strncpy(gaExclList[kk], lpStat->aExclList[kk], MAXALIASNAMESIZE);
				}
			lnReturn = lnTargetSession;	
			gnCodeBaseItem = lnTargetSession;			
			}	
		}
		
	return Py_BuildValue("l", lnReturn);
}

/* ****************************************************************************** */
/* Takes as its parm the Datasession index to close.  Returns that number on      */
/* success, otherwise -1.                                                         */
static PyObject *cbwCLOSEDATASESSION(PyObject *self, PyObject *args)
{
	long jj;
	long lnTargetSession;
	long lnReturn = 99;
	long lnResult = 0;
	long kk;
	CODEBASESTATUS *lpStat;
	CODE4 lxCodeBase;
		
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (!PyArg_ParseTuple(args, "l", &lnTargetSession))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		return NULL;
		}
	if (lnTargetSession > (MAXDATASESSIONS - 1))
		{
		strcpy(gcErrorMessage, "Invalid target Session Number");
		gnLastErrorNumber = -10001;
		return Py_BuildValue("l", -1);			
		}
	if (lnTargetSession < 0) lnTargetSession = gnCodeBaseItem; // Set to "Current".	
		
	if (lnTargetSession == gnCodeBaseItem)
		{
		if (gnCodeBaseItem != -1) // Not the undefined case.
			{
			lnResult = code4initUndo(&codeBase);
			//cbxInitGlobals();
			memset(&codeBase, 0, sizeof(CODE4));
			gaCodeBases[lnTargetSession] = codeBase;
			gnCodeBaseFlags[lnTargetSession] = -1; // Set it to undefined.
			gnCodeBaseItem = -1; // The undefined case.
			if (lnResult >= 0)
				{
				lnReturn = lnTargetSession;	
				}
			else
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;					
				}
			}
		else
			{
			lnReturn = lnTargetSession;	
			}
		}
	else
		{
		if (gnCodeBaseFlags[lnTargetSession] == -1)
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, "Specified Session NOT ACTIVE");
			gnLastErrorNumber = -10003;
			}
		else
			{
			lxCodeBase = codeBase;
			codeBase = gaCodeBases[lnTargetSession];
			lnResult = code4initUndo(&codeBase); // Requires this exact variable because it has internal pointers to itself.
			memset(&codeBase, 0, sizeof(CODE4));
			gaCodeBases[lnTargetSession] = codeBase; // Storing the un-inited version.
			codeBase = lxCodeBase; // Put it back the way it was.
			gnCodeBaseFlags[lnTargetSession] = -1;
			//gnCodeBaseItem = -1;

			lpStat = &(gaCodeBaseEnvironments[lnTargetSession]); // Code below initializes this structure.
			lpStat->pCurrentTable = NULL;
			lpStat->pCurrentIndexTag = NULL;
			lpStat->pCurrentFilterExpr = NULL;
			lpStat->pLastField = NULL;
			lpStat->pCurrentQuery = NULL;
			strcpy(lpStat->cQueryAlias, "");
			lpStat->nDeletedFlag = -1;
			strcpy(lpStat->cErrorMessage, "");
			lpStat->nLastErrorNumber = 0;
			strcpy(lpStat->cDelimiter, "||");
			strcpy(lpStat->cLastSeekString, "");
			strcpy(lpStat->cDateFormatCode, "MM/DD/YYYY");
			lpStat->bLargeMode = FALSE;

			memset(lpStat->aIndexTags, 0, (size_t) (MAXINDEXTAGS * sizeof(VFPINDEXTAG)));
			memset(lpStat->aTempIndexes, 0, (size_t) (MAXTEMPINDEXES * sizeof(TEMPINDEX)));
	
			for (kk = 0; kk < MAXEXCLCOUNT; kk++)
				{
				strcpy(lpStat->aExclList[kk], "");
				}

			if (lnResult < 0)
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;		
				}
			else
				{
				lnReturn = lnTargetSession;	
				}
			}				
		}

	return Py_BuildValue("l", lnReturn);
}

/* ****************************************************************************** */
/* Codebase, like the USA versions of Visual FoxPro assumes U.S. standard date    */
/* expressions of the form MM/DD/YY or MM/DD/YYYY, in contrast to other parts of  */
/* the world that use DD/MM/YY or DD/MM/YYYY or other combinations.  This default */
/* assumptions about date text strings can be changed.  This is especially useful */
/* when importing text data via the cbwAPPENDFROM() function.  If non-numeric     */
/* characters are found in the text, and it is intended to be stored in a date or */
/* datetime field, then an attempt will be made to parse the text based on the    */
/* standard date expression, either MM/DD/YYYY or an alternate set prior to the   */
/* append process.                                                                */
/* Parameters:                                                                    */
/*  - cPattern - This must be a character string with at minimum the characters   */
/*    'YY' or 'YYYY', and 'MM', and 'DD' in some combination with punctuation or  */
/*    it must be one of the Visual FoxPro standard SET DATE TO values such as:    */
/*    AMERICAN = MM/DD/YY                                                         */
/*    ANSI = YY.MM.DD                                                             */
/*    BRITISH = DD/MM/YY                                                          */
/*    FRENCH = DD/MM/YY                                                           */
/*    GERMAN = DD.MM.YY                                                           */
/*    ITALIAN = DD-MM-YY                                                          */
/*    JAPAN = YY/MM/DD                                                            */
/*    TAIWAN = YY/MM/DD                                                           */
/*    USA = MM-DD-YY                                                              */
/*    MDY = MM/DD/YY                                                              */
/*    DMY = DD/MM/YY                                                              */
/*    YMD = YY/MM/DD                                                              */
/*    Pass NULL or the empty string to restore the MDY/AMERICAN default.          */
/* Returns 1 on OK, 0 on any kind of error, mainly unrecognized formats.          */
/* The total length of the lcFormat string must be no greater than 16 bytes.      */
static PyObject *cbwSETDATEFORMAT(PyObject *self, PyObject *args)
{
	long lnReturn = 1;
	unsigned char cDateFmat[30];
	unsigned char lcFormat[45];
	const unsigned char *cTest;
	PyObject *lxFormat;

	if (!PyArg_ParseTuple(args, "O", &lxFormat))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}
		
	cTest = testStringTypes(lxFormat);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}

	strncpy(lcFormat, Unicode2Char(lxFormat), 29);
	lcFormat[29] = (char) 0;
	Conv1252ToASCII(lcFormat, TRUE);

	memset(cDateFmat, 0, (size_t) 20 * sizeof(char));
	while(TRUE)
		{
		if ((lcFormat == NULL) || (strlen(lcFormat)	== 0))
			{
			strcpy(cDateFmat, "MM/DD/YY");
			strcpy(gcDateFormatCode, "MM/DD/YY");
			break;	
			}
		if ((strcmp(lcFormat, "AMERICAN") == 0) || (strcmp(lcFormat, "MDY") == 0))
			{
			strcpy(cDateFmat, "MM/DD/YY");
			strcpy(gcDateFormatCode, "AMERICAN");
			break;	
			}
		if (strcmp(lcFormat, "ANSI") == 0)
			{
			strcpy(cDateFmat, "YY.MM.DD");
			strcpy(gcDateFormatCode, "ANSI");
			break;	
			}
		if ((strcmp(lcFormat, "BRITISH") == 0) || (strcmp(lcFormat, "FRENCH") == 0) || (strcmp(lcFormat, "DMY") == 0))
			{
			strcpy(cDateFmat, "DD/MM/YY");
			strcpy(gcDateFormatCode, lcFormat);
			break;	
			}			
		if (strcmp(lcFormat, "GERMAN") == 0)
			{
			strcpy(cDateFmat, "DD.MM.YY");
			strcpy(gcDateFormatCode, lcFormat);
			break;	
			}
		if ((strcmp(lcFormat, "JAPAN") == 0) || (strcmp(lcFormat, "TAIWAN") == 0) || (strcmp(lcFormat, "YMD") == 0))
			{
			strcpy(cDateFmat, "YY/MM/DD");
			strcpy(gcDateFormatCode, lcFormat);
			break;	
			}			
		if (strcmp(lcFormat, "ITALIAN") == 0)
			{
			strcpy(cDateFmat, "DD-MM-YY");
			strcpy(gcDateFormatCode, "ITALIAN");
			break;	
			}
		if (strcmp(lcFormat, "USA") == 0)
			{
			strcpy(cDateFmat, "MM-DD-YY");
			strcpy(gcDateFormatCode, "USA");
			break;	
			}
		if ((strchr(lcFormat, 'C') != NULL)	&& (strchr(lcFormat, 'Y') != NULL) && (strchr(lcFormat, 'D') != NULL))
			{
			strncpy(cDateFmat, lcFormat, 16);
			cDateFmat[16] = (char) 0;
			strcpy(gcDateFormatCode, lcFormat);
			break;	
			}
		break; // not recognized						
		}
	if (cDateFmat[0] != (char) 0)
		{
		code4dateFormatSet(&codeBase, cDateFmat);
		gcErrorMessage[0] = (char) 0;
		gnLastErrorNumber = 0;
		}
	else
		{
		lnReturn = 0;
		strcpy(gcErrorMessage, "Unrecognized Date Format Code");
		gnLastErrorNumber = -4398;
		}
	return Py_BuildValue("l", lnReturn);
}

/* ****************************************************************************** */
/* Retrieves the current dateformat code value.                                   */
//char* cbwGETDATEFORMAT(void)
static PyObject *cbwGETDATEFORMAT(PyObject *self, PyObject *args)
{
//return gcDateFormatCode; // We set this whenever we change it in CodeBase, so it has the
// latest status info.
	return Py_BuildValue("s", gcDateFormatCode);
}

/* *************************************************************************** */
/* Returns the value of the last error message                                 */
//char* cbwERRORMSG(void)
static PyObject *cbwERRORMSG(PyObject *self, PyObject *args)
{
	return Py_BuildValue("s", gcErrorMessage);	
}

/* *************************************************************************** */
/* Returns the value of the last error number.                                 */
static PyObject *cbwERRORNUM(PyObject *self, PyObject *args)
{
	return Py_BuildValue("l", gnLastErrorNumber);	
}


/* ****************************************************************************** */
/* Call this to determine the current setting of "Large Table Mode".  Returns     */
/* 1 if large table mode is ON, otherwise 0.                                      */
static PyObject *cbwISLARGEMODE(PyObject *self, PyObject *args)
{
return Py_BuildValue("N", PyBool_FromLong(gbLargeMode));	
}

/* ******************************************************************************* */
/* Sets the status of the DELETED flag meaning that if TRUE, deleted records are   */
/* invisible.  If FALSE, then deleted records are visible but can be detected with */
/* the cbwDELETED() function.                                                      */
/* Operates on all tables.                                                         */
/* NOTE ADDED FEATURE 08/04/2014.  Pass -1, and it will return the current state of*/
/* the deleted flag. Otherwise pass True or False to set the deleted status.       */
static PyObject *cbwSETDELETED(PyObject *self, PyObject *args)
{
	long lbHow;
	long lnDeletedOn = 0;
	
	if (!PyArg_ParseTuple(args, "l", &lbHow))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		return NULL;
		}	
	lnDeletedOn = lbHow;
	if (lnDeletedOn == -1)
		{
		return Py_BuildValue("N", PyBool_FromLong(gnDeletedFlag));	
		}
	else
		{
		gnDeletedFlag = lbHow;
		return Py_BuildValue("N", PyBool_FromLong(TRUE));
		}	
}

/* ******************************************************************************* */
/* Internal version of cbwUSE().                                                   */
/* Attempts to open the specified table.  Applies the following logic relative to
   tables already open:
   - If there is an alias specified:
		- If the alias already applies to an open table:
			- The table is closed
			- Then the table is re-opened with the alias and properties as specified
		- If there is no table open with that alias:
			- Tries to open the table and assign the alias
   - If there is NO alias specified:
   		- Determines what the default alias would be (base table name without extension, but test 
   		  for illegal characters: must start with a letter or underscore, max length 254 characters,
   		  and otherwise all alpha-numeric or underscore characters.)
   		- Looks to see if that alias is already in use.  If it is:
   			- Compare full table name already open with new table name.  If the same, then
   			  close table and reopen, using same alias (but new properties)
   			- If table is NOT the same, then create new custom alias for table to open
   		- If the standard alias is NOT already active, then open the table with the default
   		  alias.
*/
long cbxUSE(char *lpcTable, char *lpcAlias, long lpnReadOnlyFlag, long lbNoBufferingFlag, long lpbExclusive)
{
	DATA4 *lpTable;
	long lbReturn = TRUE;
	long lnResult;
	long lnAliasLen;
	char cWorkTable[300];
	char cWorkAlias[300];
	char cTempTable[500];
	char cBaseName[300];
	register long jj;
	//printf("THIS IS THE USE \n");
	if (lpbExclusive)
		{
		lpnReadOnlyFlag = FALSE; // By definition.	
		}	
	if (lpnReadOnlyFlag)
		{
		codeBase.readOnly = TRUE;
		}
	else
		{
		codeBase.readOnly = FALSE;
		}
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (!lpbExclusive)
		{
		codeBase.accessMode = OPEN4DENY_NONE; /* Full shared mode */
		}
	else
		{
		codeBase.accessMode = OPEN4DENY_RW; /* i.e. EXCLUSIVE */	
		}
	strncpy(cWorkTable, lpcTable, 299);
	cWorkTable[299] = (char) 0;
	strncpy(cWorkAlias, lpcAlias, 299);
	cWorkAlias[299] = (char) 0; // Just in case.
	StrToUpper(cWorkTable); // Force the DLL to store the table name as upper no matter what.  OK for Windows.
	StrToUpper(cWorkAlias);
	
	if (strlen(cWorkAlias) > 0)
		{
		// An alias name has been specified.
		lpTable = code4data(&codeBase, cWorkAlias);
		if (lpTable != (DATA4*) NULL)
			{
			// A table with this alias is already open, so we have to close it first.	
			lnResult = cbxCLOSETABLE(cWorkAlias);	
			}
		}
	else
		{
		u4namePiece(cBaseName, 299, cWorkTable, FALSE, FALSE);
		cBaseName[299] = (char) 0; // Just in case.
		strcpy(cWorkAlias, MakeLegalAlias(cBaseName));
		lpTable = code4data(&codeBase, cWorkAlias);
		if (lpTable != NULL)
			{
			strncpy(cTempTable, d4fileName(lpTable), 499);
			cTempTable[499] = (char) 0;
			StrToUpper(cTempTable);
			if (strcmp(cTempTable, cWorkTable) == 0)
				{
				// This is the same table, exactly, so we simply close it and let it be reopened.
				lnResult = cbxCLOSETABLE(cWorkAlias);	
				}
			else
				{
				strcpy(cWorkAlias, "TMP");
				strcat(cWorkAlias, MPStrRand(12));	// New randomly created Alias Name
				}
			}
		// No else.  We simply open the table with the alias name we've calculated for the default.
		}
	
	lpTable = d4open(&codeBase, cWorkTable);
//	if ((lpTable == NULL) && (codeBase.errorCode == -69))
//		{
//		lpTable = GetTableFromName(cWorkTable);
//		if (lpTable != NULL)
//			{
//			//printf("Trying reopen for %p \n", lpTable);
//			//codeBase.singleOpen = FALSE;
//			lpTable = d4openClone(lpTable);	
//			//printf("Opening clone of %s %p\n", cWorkTable, lpTable);
//			}
//		}
	if (lpTable == NULL)
		{
		lbReturn = FALSE;
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		strcat(gcErrorMessage, ": ");
		strcat(gcErrorMessage, cWorkTable);
		gnLastErrorNumber = codeBase.errorCode;	
		}
	else
		{
		//printf("RAW ALIAS <%s>\n", d4alias(lpTable));
		if (strlen(cWorkAlias) == 0) strcpy(cWorkAlias, d4alias(lpTable)); // UNLIKELY
		if (lbNoBufferingFlag)
			{
			lnResult = d4optimize(lpTable, OPT4OFF);
			lnResult = d4optimizeWrite(lpTable, OPT4OFF);	
			}
		gpCurrentTable = lpTable;
		gpCurrentIndexTag = NULL;
		if (strlen(cWorkAlias) > 0)
			{
			d4aliasSet(gpCurrentTable, cWorkAlias);	
			//printf("STORING ALIAS FROM USE: %s \n", cWorkAlias);
			//StoreAliasToTracker(gpCurrentTable, cWorkAlias, cWorkTable); // TEMP
			}

		if (lpbExclusive)
			{
			for (jj = 0;(jj < (gnExclLimit + 1)) && (jj < gnExclMax) ; jj++)
				{
				if (strlen(gaExclList[jj]) == 0)
					{
					strncpy(gaExclList[jj], StrToLower(cWorkAlias), MAXALIASNAMESIZE);
					gnExclCnt += 1;
					if (jj > gnExclLimit) gnExclLimit = jj + 1;
					break;	
					}
				}
			}

		d4top(gpCurrentTable); /* Always position at the first physical record on open */	
		}	
	//codeBase.singleOpen = TRUE;
	return(lbReturn);	
}

/* ******************************************************************************* */
/* Attempts to USE the table passed by name.  Returns True on OK, False on failure.*/
/* Accepts an optional 2nd parm to specify a non-standard alias.  Pass an empty    */
/* string to use the default.                                                      */
/* SEE cbxUSE() for documentation of alias handling logic.                         */
static PyObject *cbwUSE(PyObject *self, PyObject *args)
{
	long lbReturn;
	long lpnReadOnlyFlag;
	long lpnNoBuffering;
	char lcAlias[MAXALIASNAMESIZE + 1];
	char lcTable[500];
	const unsigned char *cTest;
	PyObject *lxAlias;
	PyObject *lxTable;

	if (!PyArg_ParseTuple(args, "OOll", &lxTable, &lxAlias, &lpnReadOnlyFlag, &lpnNoBuffering))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}
		
	cTest = testStringTypes(lxTable);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTable, Unicode2Char(lxTable), 499);
	lcTable[499] = (char) 0;
	Conv1252ToASCII(lcTable, TRUE);		
		
	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE];
	Conv1252ToASCII(lcAlias, TRUE);			
		
	lbReturn = cbxUSE(lcTable, lcAlias, lpnReadOnlyFlag, lpnNoBuffering, FALSE);
	return Py_BuildValue("N", PyBool_FromLong(lbReturn));
}

/* *************************************************************************** */
/* Like the VFP ALIAS() function returns the name of the currently selected    */
/* ALIAS or an empty string if there isn't an open table selected.             */
static PyObject *cbwALIAS(PyObject *self, PyObject *args)
{
	char lcName[200];
	if (gpCurrentTable == NULL) lcName[0] = (char) 0;
	else strcpy(lcName, d4alias(gpCurrentTable));
	//else strcpy(lcName, GetTableAlias(gpCurrentTable)); // TEMP
	
	return Py_BuildValue("s", lcName);	
}

/* *************************************************************************** */
/* Like the VFP ALIAS() function returns the name of the currently selected    */
/* ALIAS.                                                                      */
/* INTERNAL VERSION.                                                           */
char* cbxALIAS(void)
{
	char lcName[200];
	if (gpCurrentTable == NULL) lcName[0] = (char) 0;
	else strcpy(lcName, d4alias(gpCurrentTable));
	//else strcpy(lcName, GetTableAlias(gpCurrentTable)); // TEMP
	
	strcpy(gcStrReturn, lcName);
	return(gcStrReturn);	
}

/* ************************************************************************** */
/* Internal version of the cbwDBF() function.                                 */
char *cbxDBF(char* lpcAlias)
{
	long lnReturn = TRUE;
	DATA4* lpTable = NULL;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gcStrReturn[0] = (char) 0;
		
	if (strlen(lpcAlias) > 0)
		{
		lpTable = code4data(&codeBase, lpcAlias);	
		//lpTable = GetTableFromAlias(lpcAlias); // TEMP
		}
	else
		{
		if (gpCurrentTable == NULL)
			{
			strcpy(gcStrReturn, "");
			strcpy(gcErrorMessage, "No Table Selected");
			gnLastErrorNumber = -99999;
			}
		else
			{
			lpTable = gpCurrentTable;
			}	
		}
	if (lpTable != NULL)
		{
		strcpy(gcStrReturn, d4fileName(lpTable));
		if (strlen(gcStrReturn) == 0)
			{
			strcpy(gcErrorMessage, "No Alias Found");
			gnLastErrorNumber = -99999;	
			}	
		}
	else
		{
		strcpy(gcErrorMessage, "Alias Name Not Found");
		gnLastErrorNumber = -99999;
		}
	return(gcStrReturn);
}

/* ************************************************************************** */
/* Function like VFP DBF() which returns the fully qualified path name of     */
/* the currently selected table or an empty string.                           */
/* Accepts an alias or an empty string.                                       */
static PyObject *cbwDBF(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	DATA4 *lpTable = NULL;
	char lcAlias[MAXALIASNAMESIZE + 1];
	PyObject *lxAlias;
	const unsigned char *cTest;

	if (!PyArg_ParseTuple(args, "O", &lxAlias))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);
			
	return(Py_BuildValue("s", cbxDBF(lcAlias)));
}

/* ******************************************************************************* */
/* Attempts to USE the table passed by name.  Returns True on OK, False on failure.*/
/* Accepts an optional 2nd parm to specify a non-standard alias.  Pass an empty    */
/* string to use the default.                                                      */
/* Like cbwUSE except that the table is opened EXCLUSIVE (never read only!)        */
/* Updated 02/19/2015 to make use of the common cbxUSE() function.                 */
static PyObject *cbwUSEEXCL(PyObject *self, PyObject *args)
{
	DATA4 *lpTable;
	PyObject *lxTable;
	PyObject *lxAlias;
	long lbReturn = TRUE;
	char lcAlias[MAXALIASNAMESIZE + 1];
	char lcTable[500];
	const unsigned char *cTest;

	if (!PyArg_ParseTuple(args, "OO", &lxTable, &lxAlias))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}
		
	cTest = testStringTypes(lxTable);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTable, Unicode2Char(lxTable), 499);
	lcTable[499] = (char) 0;
	Conv1252ToASCII(lcTable, TRUE);		
		
	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);	

	codeBase.readOnly = FALSE;
	codeBase.accessMode = OPEN4DENY_RW; /* i.e. EXCLUSIVE */
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lbReturn = cbxUSE(lcTable, lcAlias, FALSE, FALSE, TRUE);
	return Py_BuildValue("N", PyBool_FromLong(lbReturn));
}

/* ********************************************************************************* */
/* Uses the gaExclList array to get the exclusive status of the currently selected   */
/* table.  Returns True on yes, it is exclusive, False on NOT exclusive, None on err.*/
//long DOEXPORT cbwISEXCL(void)
static PyObject *cbwISEXCL(PyObject *self, PyObject *args)
{
	char cAlias[400];
	long jj;
	long nRet = 0;
	if (gpCurrentTable == NULL)
		{
		strcpy(gcErrorMessage, "No Table Selected");
		gnLastErrorNumber = -99999;
		return Py_BuildValue(""); // Returns Python None		
		}	
	
	strcpy(cAlias, d4alias(gpCurrentTable));
	StrToLower(cAlias);
	for (jj = 0;jj < gnExclLimit; jj++)
		{
		//if (gaExclList[jj] == NULL) continue;
		if (strcmp(gaExclList[jj], cAlias) == 0)
			{
			nRet = 1;
			break;	
			}	
		}
	return Py_BuildValue("N", PyBool_FromLong(nRet));	
}

/* ************************************************************************************** */
/* VFP Allows the creation of one-time use indexes which can be used for special purpose  */
/* processing but which are not continually updated as the associated table changes.      */
/* This function allows for the creation of such an index on a table with minimum fuss    */
/* and its automatic removal when the table is closed.  Like cbwINDEXON, this function    */
/* works on the currently selected table and takes as its parameters an index             */
/* expression, filter expression, descending flag and a unique flag.                      */
/* It differs from cbwINDEXON in that it returns 0 or a positive integer value as the     */
/* index identifier, or -1 on failure.                                                    */
/*                                                                                        */
/* Retain the index identifier so that if you change the index order later, you can set it*/
/* back to this index.                                                                    */
/* NOTE: Temporary indexes are always closed and deleted when the associated table is     */
/* closed!  They are physically located on disk, but have randomly selected file names    */
/* which are NOT made available outside the function.                                     */

static PyObject *cbwTEMPINDEX(PyObject *self, PyObject *args)
{
	long lnIndexCode = -1;
	char lcFileName[500];
	char lcTableName[500];
	char lcTagExpr[1000];
	char lcTagFilter[1000];
	char lcTagName[20];
	PyObject* lxTagExpr;
	PyObject* lxTagFltr;
	long lpnDescending;
	TAG4INFO laTagInfo[3];
	TAG4* lpTag = NULL;
	INDEX4 *lpIndex;
	long jj;
	long lnDescending = 0;
	const unsigned char *cTest;

	if (!PyArg_ParseTuple(args, "OOl", &lxTagExpr, &lxTagFltr, &lpnDescending))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxTagExpr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTagExpr, Unicode2Char(lxTagExpr), 998);
	lcTagExpr[998] = (char) 0;
	Conv1252ToASCII(lcTagExpr, TRUE);		
		
	cTest = testStringTypes(lxTagFltr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTagFilter, Unicode2Char(lxTagFltr), 998);
	lcTagFilter[998] = (char) 0;	
	Conv1252ToASCII(lcTagFilter, TRUE);	
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	if (gpCurrentTable != NULL)
		{
		//memset(lcFileName, 0, (size_t) 300 * sizeof(char)); // Start out with an empty name string
		lcFileName[0] = (char) 0;
		lcTableName[0] = (char) 0;
		strcpy(lcTableName, d4fileName(gpCurrentTable));

		strcpy(lcFileName, justFilePath(lcTableName));
		strcat(lcFileName, "\\X");
		strcat(lcFileName, MPStrRand(8));
		strcat(lcFileName, ".CDX");
		memset(laTagInfo, 0, (size_t) 3 * sizeof(TAG4INFO));
		strcpy(lcTagName, "X");
		strcat(lcTagName, MPStrRand(8));

		if (lpnDescending != 0) lnDescending = r4descending; // the official code for descending.
		laTagInfo[0].name = (char*) lcTagName; // Can be anything, since there will be only one.
		laTagInfo[0].expression = (char*) lcTagExpr;
		if (strlen(lcTagFilter) == 0) laTagInfo[0].filter = NULL;
		else laTagInfo[0].filter = (char*) lcTagFilter;
		laTagInfo[0].unique = 0;
		laTagInfo[0].descending = lnDescending;
		
		// Now look for an empty index in the temp index array
		for (jj = 0; jj < gnMaxTempIndexes; jj++)
			{
			if (gaTempIndexes[jj].bOpen == FALSE)
				{
				lnIndexCode = jj;
				break;
				}	
			}
		if (lnIndexCode == -1) // still, so (amazingly) all indexes are spoken for...
			{
			gnLastErrorNumber = -9341;
			strcpy(gcErrorMessage, "No unused temp index slots.  Can't create new index");	
			}		
		else
			{
			codeBase.errorCode = 0; // make sure it's reset.
			lpIndex = i4create(gpCurrentTable, (const char*) lcFileName, laTagInfo );
			if (lpIndex != NULL) // Success
				{
				gaTempIndexes[lnIndexCode].bOpen = TRUE;
				strcpy(gaTempIndexes[lnIndexCode].cFileName, lcFileName);
				strcpy(gaTempIndexes[lnIndexCode].cTableAlias, d4alias(gpCurrentTable));
				StrToLower(gaTempIndexes[lnIndexCode].cTableAlias);
				strcpy(gaTempIndexes[lnIndexCode].aTags[0].cTagName, lcTagName);
				strcpy(gaTempIndexes[lnIndexCode].aTags[0].cTagExpr, lcTagExpr);
				strcpy(gaTempIndexes[lnIndexCode].aTags[0].cTagFilt, lcTagFilter);
				gaTempIndexes[lnIndexCode].aTags[0].nDirection = lnDescending;
				gaTempIndexes[lnIndexCode].aTags[0].nUnique = 0;
				gaTempIndexes[lnIndexCode].nTagCount = 1;
				gaTempIndexes[lnIndexCode].pTable = gpCurrentTable;
				gaTempIndexes[lnIndexCode].pIndex = lpIndex;
				
				lpTag = d4tag(gpCurrentTable, lcTagName);
				if (lpTag != NULL)
					{
					d4tagSelect(gpCurrentTable, lpTag);	
					gnOpenTempIndexes += 1;
					}
				else
					{
					// This is BAD!  Can't find the index we just now created...
					// undo it if possible and return the error.
					gaTempIndexes[lnIndexCode].bOpen = FALSE;
					i4close(gaTempIndexes[lnIndexCode].pIndex);
					lnIndexCode = -1;
					strcpy(gcErrorMessage, "Temp Index Creation Failed!");
					gnLastErrorNumber = -9113;	
					}
				}
			else
				{
				lnIndexCode = -1;
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				strcat(gcErrorMessage, " - Adding TEMP INDEX");				
				}
			}
		}
	else
		{
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		lnIndexCode = -1;	
		}		
	return Py_BuildValue("l", lnIndexCode);
}

/* ************************************************************************************** */
/* If you select a different order while a temp index is in effect, you can return to it  */
/* with this function.  In theory, changes made to the table will be reflected in the temp*/
/* indexes while they are open, so they should stay up to date.                           */
/* Pass this the temp index code.  Returns True on OK, False on any failure.              */
/* Note this does NOT set the associated table as currently selected.  It just changes    */
/* the current ordering on the table associated with this temp index.                     */
static PyObject *cbwTEMPINDEXSELECT(PyObject *self, PyObject *args)
{
	long jj;
	long lnIndx;
	long lnReturn = 1;
	TAG4* lpTag = NULL;
	
	if (!PyArg_ParseTuple(args, "l", &lnIndx))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		return NULL;
		}	
	
	if ((lnIndx > gnMaxTempIndexes) || (lnIndx < 0))
		{
		lnReturn = 0;
		strcpy(gcErrorMessage, "Temp Index Code is out of range");
		gnLastErrorNumber = -9194;	
		}
	else
		{
		if ((gaTempIndexes[lnIndx].bOpen == FALSE) || (gaTempIndexes[lnIndx].pTable == NULL))
			{
			lnReturn = 0;
			strcpy(gcErrorMessage, "Temp Index has already been closed");
			gnLastErrorNumber = -9294;	
			}
		else
			{
			lpTag = d4tag(gaTempIndexes[lnIndx].pTable, gaTempIndexes[lnIndx].aTags[0].cTagName);
			if (lpTag != NULL) d4tagSelect(gaTempIndexes[lnIndx].pTable, lpTag);
			else
				{
				lnReturn = 0;
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				strcat(gcErrorMessage, " - Selecting TEMP INDEX");					
				}			
			}
		}	
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}
/* PRIVATE VERSION of cbwTEMPINDEXCLOSE() */
long cbxTEMPINDEXCLOSE(long lpnIndx)
{
	long jj;
	long lnReturn = 1;
	long lnResult;
		
	if ((lpnIndx > gnMaxTempIndexes) || (lpnIndx < 0))
		{
		lnReturn = 0;
		strcpy(gcErrorMessage, "Temp Index Code is out of range");
		gnLastErrorNumber = -9194;	
		}
	else
		{
		if (gaTempIndexes[lpnIndx].bOpen)
			{
			gaTempIndexes[lpnIndx].bOpen = FALSE;
			if (gaTempIndexes[lpnIndx].pIndex != NULL)
				{
				lnResult = i4close(gaTempIndexes[lpnIndx].pIndex);
				if (lnResult != r4success)
					{
					lnReturn = 0;	
					}
				}
			if (gaTempIndexes[lpnIndx].pTable != NULL)
				{
				d4tagSelect(gaTempIndexes[lpnIndx].pTable, NULL);
				}
			if (gnOpenTempIndexes > 0) gnOpenTempIndexes -= 1;
	
			remove(gaTempIndexes[lpnIndx].cFileName);
			memset(gaTempIndexes[lpnIndx].cFileName, 0, (size_t) 255 * sizeof(char));
			memset(gaTempIndexes[lpnIndx].cTableAlias, 0, (size_t) 100 * sizeof(char));
			gaTempIndexes[lpnIndx].nTagCount = 0;
			gaTempIndexes[lpnIndx].pIndex = NULL;
			}	
		}	
	return(lnReturn);	
}

/* ********************************************************************************** */
/* Counterpart to the cbwTEMPINDEX() function.  Hand it the index code of a temp      */
/* index, and it will close the index.  If the table to which this was associated     */
/* is still open, it will re-associate it with its CDX and set the order to NULL      */
/* meaning "no order".  Returns True on OK, False on failure.                         */
static PyObject *cbwTEMPINDEXCLOSE(PyObject *self, PyObject *args)
{
	long jj;
	long lnReturn = 1;
	long lnResult;
	long lnIndx;

	if (!PyArg_ParseTuple(args, "l", &lnIndx))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		return NULL;
		}
	lnReturn = cbxTEMPINDEXCLOSE(lnIndx);
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));
}

/* ******************************************************************************* */
/* Attempts to close and re-open the current table in the exact same mode it had   */
/* when it was opened in the first place.  The reason for needing this is that when*/
/* we add an index TAG to a table that didn't have one, we need to reopen it to    */
/* store the proper index information for the table.                               */
long cbxReopenCurrentTable(void)
{
	DATA4* lpTable;
	long lbReturn = TRUE;
	long lbResult;
	long nAccessMode = OPEN4DENY_NONE;
	long nReadOnly = 0;
	char cTableName[400];
	char cTableAlias[400];
	char cMatchAlias[400];
	long nMatchPoint = -1;
	long jj;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable == NULL)
		{
		strcpy(gcErrorMessage, "No Table Selected");
		gnLastErrorNumber = -99999;
		return(FALSE);			
		}	
	strcpy(cTableAlias, d4alias(gpCurrentTable));
	strcpy(cMatchAlias, cTableAlias);
	StrToLower(cMatchAlias);
	nReadOnly = gpCurrentTable->readOnly;
	codeBase.readOnly = nReadOnly;
	if (gnExclCnt > 0)
		{
		for (jj = 0; jj < gnExclLimit; jj++)
			{
			//if (gaExclList[jj] == NULL) continue;
			if (strcmp(gaExclList[jj], cMatchAlias) == 0)
				{
				nAccessMode = OPEN4DENY_RW;
				nMatchPoint = jj;
				break;
				}	
			}	
		}
	
	codeBase.accessMode = nAccessMode;
	strcpy(cTableName, d4fileName(gpCurrentTable));
	lbResult = d4close(gpCurrentTable);
	if (lbResult < 0)
		{
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		gnLastErrorNumber = codeBase.errorCode;
		lbReturn = FALSE;	
		}
	if (lbReturn)
		{	
		lpTable = d4open(&codeBase, cTableName);
		if (lpTable == NULL)
			{
			lbReturn = FALSE;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			strcat(gcErrorMessage, ": ");
			strcat(gcErrorMessage, cTableName);
			gnLastErrorNumber = codeBase.errorCode;	
			gpCurrentTable = NULL;
			gpCurrentIndexTag = NULL;
			}
		else
			{
			gpCurrentTable = lpTable;
			gpCurrentIndexTag = NULL;
			strcpy(cMatchAlias, d4alias(lpTable));

			if (strlen(cTableAlias) > 0)
				{
				d4aliasSet(gpCurrentTable, cTableAlias); // Forcing alias to its original value	
				}

			d4top(gpCurrentTable); /* Always position at the first physical record on open */	
			}
		}

	return(lbReturn);	
}

/* ************************************************************************************************* */
/* Replacement for VFP TAGCOUNT() which returns the number of tags in the CDX index associated       */
/* the currently selected table.  Returns the number or 0 if there aren't any.  Returns -1 on error. */
/* This function is useable internally and is called by the public cbwTAGCOUNT().                    */
long cbxTAGCOUNT(void)
{
	long lnReturn = -1;
	long lnResult;
	INDEX4* lpIndex;
	TAG4INFO* laTags;
	long lnCount = 0;
	long jj;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpIndex = d4index(gpCurrentTable, NULL);
		if (lpIndex == NULL)
			{
			lnReturn = 0;
			}
		else
			{
			laTags = i4tagInfo(lpIndex);
			if (laTags != NULL)
				{
				for (jj = 0; jj < 1000;jj++) /* Well beyond practical limit */
					{
					if ((laTags + jj)->name == NULL) break;
					lnCount += 1;
					}
				lnReturn = lnCount;	
				if (lnReturn == 0)
					{
					strcpy(gcErrorMessage, "No Tags Found in specified CDX file");
					gnLastErrorNumber = -8149;	
					}
				}
			else
				{
				strcpy(gcErrorMessage, "Memory error");
				gnLastErrorNumber = -8001;
				lnReturn = -1;	
				}
			}
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}

	return(lnReturn);	
}

/* *********************************************************************************** */
/* External version of the above function.                                             */
static PyObject *cbwTAGCOUNT(PyObject *self, PyObject *args)
{
	long lnReturn;
	lnReturn = cbxTAGCOUNT();
	return Py_BuildValue("l", lnReturn);	
}

/* *********************************************************************************** */
/* Equivalent to VFP INDEX ON command.  Operates on the currently selected table.      */
/* The index expression may contain VFP type functions recognized by the CodeBase      */
/* engine.  See the CodeBase documentation for more information.                       */
/* Returns TRUE on success, False on failure.                                             */
static PyObject *cbwINDEXON(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	int  lnResult = 0;
	long lnTagCount = 0;
	long lnTagLength = 0;
	long lnOldAccess;
	long lpnDescending;
	long lnDescending = 0;
	long lpnUnique = 0;
	char *lpcTagName;
	PyObject* lxTagName;
	PyObject* lxTagExpr;
	PyObject* lxTagFltr;
	char lcTagName[100];
	char lcTagExpr[500];
	char lcTagFltr[500];
	TAG4INFO laTagInfo[3];
	TAG4* lpTag = NULL;
	INDEX4* lpIndex;
	char lcTableName[256];
	char lcIndexName[256];
	const unsigned char *cTest;
	
	if (!PyArg_ParseTuple(args, "OOOll", &lxTagName, &lxTagExpr, &lxTagFltr, &lpnDescending, &lpnUnique))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}

	cTest = testStringTypes(lxTagExpr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTagExpr, Unicode2Char(lxTagExpr), 499);
	// In theory we are allowing our tag names and expressions to have cp1252 characters, not just ASCII.
	//Conv1252ToASCII(lcTagExpr, TRUE);		
		
	cTest = testStringTypes(lxTagFltr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTagFltr, Unicode2Char(lxTagFltr), 499);
	lcTagFltr[499] = (char) 0;

	cTest = testStringTypes(lxTagName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTagName, Unicode2Char(lxTagName), 99);
	lcTagName[99] = (char) 0;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gcStrReturn[0] = (char) 0;
	gnLastErrorNumber = 0;
	memset(&laTagInfo, 0, (size_t) 3 * sizeof(TAG4INFO));
	memset(lcTableName, 0, (size_t) 256 * sizeof(char));
	memset(lcIndexName, 0, (size_t) 256 * sizeof(char));
	
	if (gpCurrentTable != NULL)
		{
		lnTagLength = strlen(lcTagName);
		if ((lnTagLength == 0) || (lnTagLength > 10))
			{
			lnReturn = 0;
			strcpy(gcErrorMessage, "Tag Name exceeds 10 characters");
			gnLastErrorNumber = -8002;	
			}
		else
			{
			lnTagCount = cbxTAGCOUNT();
			if (lnTagCount > 0)
				{
				/* There is a CDX file active.  Now check to see if there is a tag by the requested name already */
				lpTag = d4tag(gpCurrentTable, lcTagName);
				}
			codeBase.errorCode = 0; /* Reset this, per CodeBase Tech */
			if (lpnDescending != 0)
				{
				lnDescending = r4descending;	
				}
			laTagInfo[0].name = lcTagName;
			laTagInfo[0].expression = lcTagExpr;
			laTagInfo[0].filter = lcTagFltr;
			laTagInfo[0].unique = lpnUnique;
			laTagInfo[0].descending = lnDescending;
			
			while(TRUE) /* Not a real loop, just a complex switch setup. */
				{
				if (lnTagCount == 0) /* There is no CDX file, so we have to create one and then add the Tag. */
					{
					lpIndex = i4create(gpCurrentTable, NULL, laTagInfo);					
					if (lpIndex == NULL)
						{
						gnLastErrorNumber = codeBase.errorCode;
						strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
						strcat(gcErrorMessage, " - Adding CDX");
						lnReturn = 0;							
						}
					else
						{
						lnReturn = cbxReopenCurrentTable();	
						}
					break;	
					}
					
				if ((lnTagCount > 0) && (lpTag == NULL)) /* There is an index CDX file, but not with that tag.  So we Add it. */
					{
					lpIndex = d4index(gpCurrentTable, d4alias(gpCurrentTable));
					//lpIndex = cbxD4Index(gpCurrentTable);
					if (lpIndex == NULL)
						{
						gnLastErrorNumber = codeBase.errorCode;
						strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
						strcat(gcErrorMessage, " - Finding Index CDX Failed");
						lnReturn = 0;
						break;							
						}	
					lnResult = i4tagAdd(lpIndex, laTagInfo);
					if (lnResult != 0)
						{
						if (lnResult > 0)
							{
							strcpy(gcErrorMessage, "Unable to Lock Index File - Existing CDX");
							gnLastErrorNumber = -9980;
							lnReturn = -1; /* One of the several locking and duplicate key errors.  In any event, FAILURE. */	
							}
						else
							{
							gnLastErrorNumber = codeBase.errorCode;
							strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
							strcat(gcErrorMessage, " - Adding to Existing CDX: ");
							strcat(gcErrorMessage, i4fileName(lpIndex));
							lnReturn = 0; /* We're done... ERROR */					
							}							
						}

					break;
					}
					
				if ((lnTagCount > 0) && (lpTag != NULL)) /* There is a index CDX file AND the tag exists, so we remove and recreate it. */ 
					{
					lpIndex = d4index(gpCurrentTable, d4alias(gpCurrentTable));
					if (lpIndex == NULL)
						{
						gnLastErrorNumber = codeBase.errorCode;
						strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
						strcat(gcErrorMessage, " - Finding Index CDX Failed");
						lnReturn = 0;
						break;							
						}
					lpTag = i4tag(lpIndex, lcTagName); // Looking this up in the current index file.
					if (lpTag != NULL)
						{
						lnResult = i4tagRemove(lpTag);
						if ((lnResult != 0) && (codeBase.errorCode != e4tagName))
							{
							/* can't figure out why this happens, but it can... for no reason */
							lnResult = cbxReopenCurrentTable();
							if (lnResult == TRUE)
								{
								lpIndex = d4index(gpCurrentTable, d4alias(gpCurrentTable));
								if (lpIndex != NULL)
									{
									lpTag = i4tag(lpIndex, lcTagName);
									if (lpTag != NULL) lnResult = i4tagRemove(lpTag);	
									}	
								if ((lnResult != 0) && (codeBase.errorCode != e4tagName))
									{
									gnLastErrorNumber = codeBase.errorCode;
									strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
									strcat(gcErrorMessage, " - Removing Old Tag Failed a second time");
									lnReturn = 0; /* We're done... ERROR */
									break;	
									}
								else
									{
									lnResult = 0;	
									}
								}
							else
								{
								lnResult = 99;
								strcpy(gcErrorMessage, "Removing current tag failed: ");
								strcat(gcErrorMessage, lcTagName);
								gnLastErrorNumber = -9898;
								break;
								}
							}
						}
					if (lnResult == 0)
						{
						lnResult = i4tagAdd(lpIndex, laTagInfo);
						if (lnResult != 0)
							{
							if (lnResult > 0)
								{
								strcpy(gcErrorMessage, "Unable to Lock Index File");
								gnLastErrorNumber = -9980;
								lnReturn = 0; /* One of the several locking and duplicate key errors.  In any event, FAILURE. */	
								}
							else
								{
								gnLastErrorNumber = codeBase.errorCode;
								strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
								strcat(gcErrorMessage, " - Adding Tag After Removing Old Tag");
								lnReturn = 0; /* We're done... ERROR */					
								}
							}
						else
							{
							lnReturn = cbxReopenCurrentTable();
							}
						}
					else
						{
						if (gnLastErrorNumber != 0) // Not already set above.
							{
							gnLastErrorNumber = codeBase.errorCode;
							strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
							strcat(gcErrorMessage, " - Removing Old Tag from Existing CDX");
							lnReturn = 0; /* We're done... ERROR */
							}							
						}	
					break;
					}
				break;
				}
			}
		}
	else
		{
		lnReturn = 0;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}

/* ****************************************************************************** */
/* Function that searches forward or backward for the next non-deleted record     */
/* based on the current index (uses Skip), if any.  Returns TRUE on OK, or FALSE  */
/* on nothing found.  This is an internal function and is not exported.           */
/* Pass 1 for skipping forward.  Pass -1 for skipping backward.  Pass lnSeekFlag  */
/* as TRUE to move by seekNext otherwise will use skip().                         */
/* Operates on the lpTable.                                                       */
/* Doesn't move the record pointer if the current record is NOT deleted.          */
long cbxFINDUNDELETED(DATA4 *lpTable, long lnDirection, long lnSeekFlag)
{
	long lnIsDeleted;
	long lnResult;
	lnResult = r4success;
	lnIsDeleted = d4deleted(lpTable);
	while (lnIsDeleted != FALSE)
		{
		if (lnSeekFlag)
			{
			lnResult = d4seekNext(lpTable, gcLastSeekString);
			if (lnResult == r4success)
				{
				lnIsDeleted = d4deleted(lpTable);
				}
			else
				{
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;
				break; /* We're done.  Can't go on. */
				}
			}
		else
			{
			lnResult = d4skip(lpTable, lnDirection);
			if (lnResult == r4success)
				{
				lnIsDeleted = d4deleted(lpTable);	
				}
			else
				{
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));	
				gnLastErrorNumber = codeBase.errorCode;				
				break;
				}	
			}	
		}	
	return(lnResult);
}

/* ***************************************************************************** */
/* Function that works like VFP SELECT myAlias.  Pass the Alias Name.            */
/* Returns True on success, False on "can't find alias"                          */
static PyObject *cbwSELECT(PyObject *self, PyObject *args)
{
	long lbReturn = TRUE;
	DATA4 *lpTable;
	//char *lpcAlias;
	PyObject* lpxAlias;
	char lcAlias[MAXALIASNAMESIZE + 1];
	const unsigned char *cTest;
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (!PyArg_ParseTuple(args, "O", &lpxAlias))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}
	
	cTest = testStringTypes(lpxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lpxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);
	// Alias names must be ASCII.
	
	lpTable = code4data(&codeBase, lcAlias);
	if (lpTable == NULL)
		{
		lbReturn = FALSE;
		strcpy(gcErrorMessage, "Alias not found: ");
		strcat(gcErrorMessage, lcAlias);
		gnLastErrorNumber = -4949;
		}
	else
		{
		gpCurrentTable = lpTable;	
		}
	return Py_BuildValue("N", PyBool_FromLong(lbReturn));
}

/* ***************************************************************************** */
/* Function that sets the value for gcCustomCodePage, which is used for type 'C' */
/* character coding of binary character fields.                                  */
static PyObject *cbwSETCODEPAGE(PyObject *self, PyObject *args)
{
	long lbReturn = TRUE;
	PyObject* lpxCodePage;
	const unsigned char *cTest;
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (!PyArg_ParseTuple(args, "O", &lpxCodePage))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}
	
	cTest = testStringTypes(lpxCodePage);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(gcCustomCodePage, Unicode2Char(lpxCodePage), 49);
	gcCustomCodePage[49] = (char) 0;
	Conv1252ToASCII(gcCustomCodePage, TRUE);

	return Py_BuildValue("N", PyBool_FromLong(lbReturn));
}


/* ********************************************************************************** */
/* Equivalent to USED() function which returns TRUE if there is an alias by the       */
/* specified name, else FALSE.                                                        */
static PyObject *cbwUSED(PyObject *self, PyObject *args)
{
	long lnReturn = FALSE;
	char *lpcAlias;
	DATA4 *lpTable;
	char lcAliasName[MAXALIASNAMESIZE + 1];
	PyObject* lxAlias;
	const unsigned char *cTest;
	
	if (!PyArg_ParseTuple(args, "O", &lxAlias))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		return NULL;
		}
	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAliasName, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAliasName[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAliasName, TRUE);

	lpTable = code4data(&codeBase, lcAliasName);
	if (lpTable != NULL) lnReturn = TRUE;
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}


/* ***************************************************************************** */
/* Equivalent of DELETED() which applies to the currently selected table and     */
/* current record.  Returns 1 TRUE if deleted or 0 FALSE.  -1 on ERROR.          */
/* RAW version.  Also called by cbwDELETED() python method.                      */
long cbxDELETED(void)
{
	long lnReturn = TRUE;
	long lnResult;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gnProcessTally = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4recNo(gpCurrentTable);
		if (lnResult < 1)
			{
			/* No current record to test */
			strcpy(gcErrorMessage, "No current record");
			gnLastErrorNumber = -9990;
			lnReturn = -1;	
			}
		else
			{
			lnResult = d4deleted(gpCurrentTable);
			if (lnResult == 0) lnReturn = FALSE; /* otherwise TRUE */
			}
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return(lnReturn);	
}

/* ***************************************************************************** */
/* Python version of the above.                                                  */
static PyObject *cbwDELETED(PyObject *self, PyObject *args)
{
	long lnReturn;
	lnReturn = cbxDELETED();
	if (lnReturn == -1)
		{
		return Py_BuildValue(""); // Returns None
		}
	else
		{
		return Py_BuildValue("N", PyBool_FromLong(lnReturn));
		}
}


/* ***************************************************************************** */
/* Equivalent to VFP Function RECNO() returning the number of the current record */ 
/* in the table. Returns -1 on ERROR.                                            */
/* RAW Version for internal use.  Called by cbwRECNO() Python method.            */
long cbxRECNO(void)
{
	long lnReturn = 0;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (!gpCurrentTable)
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	else
		{
		lnResult = d4recNo(gpCurrentTable);
		if (lnResult < -1)
			{
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode;
			lnReturn = -1;				
			}
		else
			{
			if (lnResult == -1)
				{
				strcpy(gcErrorMessage, "No Current Record Number");
				gnLastErrorNumber = -9995;
				lnReturn = -1;
				}	
			else
				{
				lnReturn = lnResult;	
				}
			}
		}	
	return(lnReturn);	
}

/* ***************************************************************************** */
/* Python version of the above.                                                  */
static PyObject *cbwRECNO(PyObject *self, PyObject *args)
{
	long lnReturn = 0;
	long lnResult;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (!gpCurrentTable)
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	else
		{
		lnResult = d4recNo(gpCurrentTable);
		if (lnResult > 0)
			{
			return(Py_BuildValue("l", lnResult));	
			}
		else
			{
			if (lnResult < 0)
				{
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;
				lnReturn = -1;					
				}
			else
				{
				strcpy(gcErrorMessage, "No Current Record Number");
				gnLastErrorNumber = -9995;
				lnReturn = -1;					
				}	
			}
		}	
	return(Py_BuildValue("l", lnReturn));	
}

/* ***************************************************************************** */
/* Returns BOF status for the currently selected table.                          */
/* Returns TRUE for BOF, FALSE for not or None on ERROR.                         */
static PyObject *cbwBOF(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4bof(gpCurrentTable);
		if (lnResult == 0) lnReturn = FALSE; /* any other return value indicates BOF */
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	if (lnReturn == -1)
		{
		return Py_BuildValue(""); // Returns None
		}
	else
		{
		return Py_BuildValue("N", PyBool_FromLong(lnReturn));
		}
}

/* Utility function for cbxGOTO */
long cbxSetReturn(DATA4* lpbTable, long lpnResult)
{
	long lnReturn;
	while(TRUE)
		{	
		if (lpnResult == r4success)
			{
			lnReturn = d4recNo(gpCurrentTable);
			break;
			}
		if (lpnResult == r4eof)
			{
			lnReturn = 0;
			break;	
			}
		if (lpnResult == r4bof)
			{
			lnReturn = 0;
			break;	
			}
		if (lpnResult == r4locked)
			{
			lnReturn = -1;
			gnLastErrorNumber = r4locked;
			strcpy(gcErrorMessage, "GOTO Failed, Can't Lock Record");
			break;	
			}
		if (lpnResult == r4entry)
			{
			lnReturn = -1;
			gnLastErrorNumber = r4entry;
			strcpy(gcErrorMessage, "GOTO Failed - record doesn't exist");
			break;	
			}
		if (lpnResult == r4unique)
			{
			lnReturn = -1;
			gnLastErrorNumber = r4unique;
			strcpy(gcErrorMessage, "GOTO Failed, Current Record Unique Key Error");
			break;
			}
		lnReturn = -1;
		gnLastErrorNumber = codeBase.errorCode;
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		break;
		}
	return(lnReturn);	
}

/* **************************************************************************** */
/* Functional equivalent to the VFP GOTO command.  Takes directions or records. */
/* Pass lpcHow as "TOP", "BOTTOM", "NEXT", "PREV", "SKIP", or "RECORD".         */
/* If you pass SKIP or RECORD, then the lpnCount parameter is required (else 0) */
/* and for SKIP indicates how many records to skip (negative is backwards) and  */
/* for RECORD, the value should be the record you want to goto. (Skip 0 = Skip 1)*/
/* Returns the record number you wind up on or 0 on EOF or no records.  Returns */
/* -1 on ERROR of any kind.                                                     */
/* For brevity, you can simply pass the first character of the specified words  */
/* in the lpcHow string.                                                        */
/* THE RAW VERSION FOR INTERNAL USE.  ALSO CALLED BY cbwGOTO().                  */
long cbxGOTO(char* lpcHow, long lpnCount)
{
	long lnReturn = TRUE;
	long lnResult;
	long lnCurrCount;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		while(TRUE)
			{
			if ((*lpcHow == 'T')	|| (*lpcHow == 't')) /* Top */
				{
				lnResult = d4top(gpCurrentTable);
				if ((gnDeletedFlag == TRUE) && (lnResult == r4success))
					lnResult = cbxFINDUNDELETED(gpCurrentTable, 1, FALSE);
				lnReturn = cbxSetReturn(gpCurrentTable, lnResult);	
				break;	
				}
			if ((*lpcHow == 'B')	|| (*lpcHow == 'b')) /* Bottom */
				{
				lnResult = d4bottom(gpCurrentTable);
				if ((gnDeletedFlag == TRUE) && (lnResult == r4success))
					lnResult = cbxFINDUNDELETED(gpCurrentTable, -1, FALSE);
				lnReturn = cbxSetReturn(gpCurrentTable, lnResult);	
				break;	
				}
			if ((*lpcHow == 'N')	|| (*lpcHow == 'n')) /* Next */
				{
				lnResult = d4skip(gpCurrentTable, 1);
				if ((gnDeletedFlag == TRUE) && (lnResult == r4success))
					lnResult = cbxFINDUNDELETED(gpCurrentTable, 1, FALSE);
				lnReturn = cbxSetReturn(gpCurrentTable, lnResult);	
				break;	
				}
			if ((*lpcHow == 'P')	|| (*lpcHow == 'p')) /* Previous */
				{
				lnResult = d4skip(gpCurrentTable, -1);
				if ((gnDeletedFlag == TRUE) && (lnResult == r4success))
					lnResult = cbxFINDUNDELETED(gpCurrentTable, -1, FALSE);				
				lnReturn = cbxSetReturn(gpCurrentTable, lnResult);	
				break;	
				}		
			if ((*lpcHow == 'S')	|| (*lpcHow == 's')) /* Skip */
				{
				if (lpnCount == 0) lpnCount = 1; /* Just in case, we at least skip 1 */
				lnResult = d4skip(gpCurrentTable, lpnCount);
				if ((gnDeletedFlag == TRUE) && (lnResult == r4success))
					lnResult = cbxFINDUNDELETED(gpCurrentTable, (lpnCount > 0 ? 1 : -1), FALSE);				
				lnReturn = cbxSetReturn(gpCurrentTable, lnResult);	
				break;	
				}
			if ((*lpcHow == 'R')	|| (*lpcHow == 'r')) /* Record */
				{
				lnCurrCount = d4recCount(gpCurrentTable);
				/* Note that we allow direct record access to deleted records */
				if (lnCurrCount == 0)
					{
					lnReturn = -1;
					strcpy(gcErrorMessage, "No records in table");
					gnLastErrorNumber = -9997;
					}
				else
					{
					if ((lpnCount < 1) || (lpnCount > lnCurrCount))
						{
						lnReturn = -1;
						strcpy(gcErrorMessage, "Record Number out of range");
						gnLastErrorNumber = -9996;	
						}
					else
						{
						lnResult = d4go(gpCurrentTable, lpnCount);
						lnReturn = cbxSetReturn(gpCurrentTable, lnResult);	
						}
					}	
				break;	
				}
			strcpy(gcErrorMessage, "Invalid GOTO Parameter");
			gnLastErrorNumber = -9998;
			lnReturn = -1;
			break;				
			};
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Current Area");
		gnLastErrorNumber = -9999;	
		}
	return(lnReturn);	
}

/* **************************************************************************************************** */
/* Max performance record skipping routine.                                                             */
static PyObject *cbwSKIP(PyObject *self, PyObject *args)

{
	long lnReturn = TRUE;
	long lnResult;
	long lpnCount;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (!PyArg_ParseTuple(args, "l", &lpnCount))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		return NULL;
		}	
	
	if (gpCurrentTable)
		{
		if (lpnCount == 0) lpnCount = 1; /* Just in case, we at least skip 1 forward */
		lnResult = d4skip(gpCurrentTable, lpnCount);
		if ((gnDeletedFlag) && (lnResult == r4success))
			lnResult = cbxFINDUNDELETED(gpCurrentTable, (lpnCount > 0 ? 1 : -1), FALSE);				
		if (lnResult != r4success)
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "SKIP Failed");
			gnLastErrorNumber = -13241;	
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Current Area");
		gnLastErrorNumber = -9999;	
		}
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}


/* ********************************************************************************* */
/* Public Python version of GOTO.                                                    */
static PyObject *cbwGOTO(PyObject *self, PyObject *args)
{
	long lpnCount;
	long lnReturn;
	long lbReturn = TRUE;
	char lcHow[20];
	PyObject* lxHow;
	const unsigned char *cTest;	
	
	if (!PyArg_ParseTuple(args, "Ol", &lxHow, &lpnCount))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}	

	cTest = testStringTypes(lxHow);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcHow, Unicode2Char(lxHow), 19);
	lcHow[19] = (char) 0;
	Conv1252ToASCII(lcHow, TRUE);

	if (cbxGOTO(lcHow, lpnCount) < 1)
		lbReturn = FALSE;
	
	return Py_BuildValue("N", PyBool_FromLong(lbReturn));
}

/* ******************************************************************************* */
/* Calculates a statistic for any numeric field or a field expression evaluating to
   a number, across records meeting a selection criteria expression.  This always
   applies to the currently selected table.  The table record pointer is stored and
   the pointer is put back where it was when the function was called, so it can be
   called from the middle of a scan.
   
   The lpcFieldExpr and lpcForVFPexpr expressions must refer only to fields in the 
   currently selected table.  The fields contained in lpcFieldExpr (if more than one)
   must be related in numerically valid ways, e.g. "SALES_PRICE+SALES_TAX".  The expression
   contained in the lpcForVFPexpr parameter must evaluate to a logical value.  For example:
   "IS_EMPLOY .OR. IS_RETIRED".  Basically, expressions which work work for cbwLOCATE() will
   work here.
   
   The value passed to lpnStat indicates the type of statistic as follows:
   1 = Sum - the total of all the expressions in lpcFieldExpr from records meeting lpcForVFPexpr
   2 = Average - As above, but the simple average
   3 = Maximum - As above, but only the largest value found in the records meeting lpcForVFPexpr
   4 = Minimum - As above, but only the minimum value found in the records meeting lpcForVFPexpr
   
   Returns a float value, which is the result of the requested calculation.  Returns -1.0
   on error.  Since -1.0 may also be a valid result, if you get a -1 returned, check
   the value gnLastErrorNumber to determine if this was the result of an error condition. */

/* WARNING!  Calling this function while a LOCATE/CONTINUE/LOCATECLEAR sequence is
   still active could produce unpredictable results.  Do NOT call this until the
   LOCATECLEAR has been called. If you do, the count will return -1 as an error.    */

static PyObject *cbwCALCSTATS(PyObject *self, PyObject *args)
{
	double lnReturn = 0.0;
	double lnTemp = 0.0;
	long lnOK = TRUE;
	long lnResult = 0;
	long lnStart = 1;
	long lnCount = 0;
	long lnOldRecord = 0;
	unsigned char lcFieldExpr[250];
	unsigned char lcForVFPexpr[250];
	PyObject* lxFieldExpr;
	PyObject* lxForVFPexpr;
	const unsigned char *cTest;	
	double lpnValue;
	long lpnStat;
	RELATE4 *xpQuery = NULL;
	EXPR4 *xpExpr = NULL;
	double lnVeryLargeValue = 9999999999999999.9;
	double lnVerySmallValue = -9999999999999999.9;
	double lnSumTrack = 0.0;

	if (!PyArg_ParseTuple(args, "lOO", &lpnStat, &lxFieldExpr, &lxForVFPexpr))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}

	cTest = testStringTypes(lxFieldExpr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldExpr, Unicode2Char(lxFieldExpr), 249);
	lcFieldExpr[249] = (char) 0;
		
	cTest = testStringTypes(lxForVFPexpr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcForVFPexpr, Unicode2Char(lxForVFPexpr), 249);
	lcForVFPexpr[249] = (char) 0;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentQuery != NULL)
		{
		lnReturn = -1.0;
		strcpy(gcErrorMessage, "Locate is active.  CALCSTATS is not available");
		gnLastErrorNumber = -841294;	
		}
	else
		{
		if (gpCurrentTable != NULL)
			{
			xpQuery = relate4init(gpCurrentTable);
			if (xpQuery == NULL)
				{
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
				gnLastErrorNumber = codeBase.errorCode;
				lnReturn = -1.0;						
				}
			else
				{
				lnResult = relate4querySet(xpQuery, lcForVFPexpr);
				if (lnResult != r4success)
					{
					strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
					gnLastErrorNumber = codeBase.errorCode;
					lnReturn = -1.0;					
					}
				}
				
			if (lnReturn > -1.0)
				{
				xpExpr = expr4parse(gpCurrentTable, lcFieldExpr);
				if (xpExpr == NULL)
					{
					strcpy(gcErrorMessage, "Bad Field Expression: ");
					strcat(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
					gnLastErrorNumber = codeBase.errorCode;
					lnReturn = -1.0;						
					}
				}
				
			if (lnReturn > -1.0)
					{
					lnOldRecord = cbxRECNO();	
					do
						{
						if (lnStart == 1) lnResult = relate4top(xpQuery);
						else lnResult = relate4skip(xpQuery, 1);
						lnStart = 0;
						if (lnResult == r4success)
							{
							if ((gnDeletedFlag == FALSE) || ((gnDeletedFlag == TRUE) && (cbxDELETED() == FALSE)))
								{
								lnTemp = expr4double(xpExpr);
								if (codeBase.errorCode > 0)
									{
									strcpy(gcErrorMessage, "Bad Record Test: ");
									strcat(	gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
									gnLastErrorNumber = codeBase.errorCode;
									lnReturn = -1.0;
									break;									
									}
								lnCount += 1;
								switch (lpnStat){
									case 1: // SUM
										lnReturn += lnTemp;
										break;
										
									case 2: // Average
										lnSumTrack += lnTemp;
										break;
										
									case 3: // Max.
										if (lnTemp > lnVerySmallValue) lnVerySmallValue = lnTemp;
										break;
										
									case 4: // Min.
										if (lnTemp < lnVeryLargeValue) lnVeryLargeValue = lnTemp;
										break;
										
									default:
										strcpy(gcErrorMessage, "Bad Statistic Type");
										gnLastErrorNumber = -59384;
										lnReturn = -1;
									
									}
								}
							}
						else break;
						if (gnLastErrorNumber != 0) break; // Error condition, so we get out of here.
						} while(TRUE);
					lnResult = relate4free(xpQuery, 0);
					cbxGOTO("RECORD", lnOldRecord); // Go back to where we were.
					if (lnCount == 0)
						{
						lnReturn = -1.0;
						strcpy(gcErrorMessage, "No records met your selection criteria");
						gnLastErrorNumber = -947324;	
						}
					else
					{
					if (lpnStat == 2)
						{
						lnReturn = lnSumTrack / (double) lnCount;	
						}
					else
						{
						if (lpnStat == 3)
							{
							lnReturn = lnVerySmallValue;	
							}	
						else
							{
							if (lpnStat == 4)
								{
								lnReturn = lnVeryLargeValue;	
								}
							}
						}
					}				
				}	
			}
		else
			{
			gcStrReturn[0] = (char) 0;
			strcpy(gcErrorMessage, "No Table Open in Selected Area");
			gnLastErrorNumber = -9999;
			lnReturn = -1.0;			
			}
		}

	if ((lnReturn == -1.0) && (gnLastErrorNumber != 0))
		{
		return Py_BuildValue(""); // Return None on Error.
		}
	else
		{
		return Py_BuildValue("d", lnReturn);	
		}
}

/* ******************************************************************************* */
/* Counts the number of records meeting a specified condition expression.  If the
   condition expression is " .T. ", then returns a count of all records... however
   if DELETED is on, then only returns count of non-deleted records.  This always
   applies to the currently selected table.  The table record pointer is stored and
   the pointer is put back where it was when the function was called, so it can be
   called from the middle of a scan.
   
   Returns an integer value from 0 (none match) up to the number of records in the
   table (if all match the expression).  Returns -1 on error (most likely because
   there is no currently selected table OR because the expression is not valid).    */

/* WARNING!  Calling this function while a LOCATE/CONTINUE/LOCATECLEAR sequence is
   still active could produce unpredictable results.  Do NOT call this until the
   LOCATECLEAR has been called. If you do, the count will return -1 as an error.    */
//long DOEXPORT cbwCOUNT(char *lpcVFPexpr)
static PyObject *cbwCOUNT(PyObject *self, PyObject *args)
{
	long lnReturn = 0;
	long lnResult = 0;
	long lnStart = 1;
	long lnCount = 0;
	long lnOldRecord = 0;
	unsigned char lcVFPexpr[400];
	PyObject *lxVFPexpr;
	const unsigned char *cTest;	
	RELATE4 *xpQuery = NULL;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (!PyArg_ParseTuple(args, "O", &lxVFPexpr))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}	

	cTest = testStringTypes(lxVFPexpr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcVFPexpr, Unicode2Char(lxVFPexpr), 399);
	lcVFPexpr[399] = (char) 0;

	if (gpCurrentQuery != NULL)
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "Locate is active.  COUNT is not available");
		gnLastErrorNumber = -841292;	
		}
	else
		{
		if (gpCurrentTable != NULL)
			{
			xpQuery = relate4init(gpCurrentTable);
			if (xpQuery == NULL)
				{
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
				gnLastErrorNumber = codeBase.errorCode;
				lnReturn = -1;						
				}
			else
				{
				lnResult = relate4querySet(xpQuery, lcVFPexpr);
				if (lnResult != r4success)
					{
					strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
					gnLastErrorNumber = codeBase.errorCode;
					lnReturn = -1;					
					}
				else
					{
					lnOldRecord = cbxRECNO();	
					do
						{
						if (lnStart == 1) lnResult = relate4top(xpQuery);
						else lnResult = relate4skip(xpQuery, 1);
						lnStart = 0;
						if (lnResult == r4success)
							{
							if (gnDeletedFlag == TRUE)
								{
								if (!cbxDELETED()) lnCount += 1;	
								}
							else
								{
								lnCount += 1; // always.	
								}
							}
						else break;
						} while(TRUE);
					lnResult = relate4free(xpQuery, 0);
					cbxGOTO("RECORD", lnOldRecord); // Go back to where we were.
					lnReturn = lnCount;
					}				
				}	
			}
		else
			{
			gcStrReturn[0] = (char) 0;
			strcpy(gcErrorMessage, "No Table Open in Selected Area");
			gnLastErrorNumber = -9999;
			lnReturn = -1;			
			}
		}
	return Py_BuildValue("l", lnReturn);
}

/* ******************************************************************************* */
/* Initializes a LOCATE operation in which the first record (starting from the top
   based on the current order setting) where a specified complex VFP expression
   is TRUE.  Subsequent CONTINUE functions may be called on the same table.  This
   method executes against the currently selected table.  If the table selection is
   changed before a CONTINUE is issued, then the selection must be put back for
   CONTINUE to work.  When done, a CLEARLOCATE must be called to release memory
   and enable the specified alias to be traversed without the LOCATE query being
   in effect.
   
   Returns TRUE if there is a matching record, otherwise FALSE.  Respects the 
   SET DELETED setting.  Returns None on ERROR. */
static PyObject *cbwLOCATE(PyObject *self, PyObject *args)
{
	long lnReturn = 0;
	long lnResult = 0;
	unsigned char lcVFPexpr[400];
	PyObject *lxVFPexpr;
	const unsigned char *cTest;	

	if (!PyArg_ParseTuple(args, "O", &lxVFPexpr))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);			
		return NULL;
		}
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	cTest = testStringTypes(lxVFPexpr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcVFPexpr, Unicode2Char(lxVFPexpr), 399);
	lcVFPexpr[399] = (char) 0;
	
	if (gpCurrentTable != NULL)
		{
		if (gpCurrentQuery != NULL)
			{
			relate4free(gpCurrentQuery, 0); // Doesn't close any tables
			}	
		gpCurrentQuery = relate4init(gpCurrentTable);
		if (gpCurrentQuery == NULL)
			{
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
			gnLastErrorNumber = codeBase.errorCode;
			lnReturn = -1;			
			}
		else
			{
			lnResult = relate4querySet(gpCurrentQuery, lcVFPexpr);
			if (lnResult != r4success)
				{
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
				gnLastErrorNumber = codeBase.errorCode;
				lnReturn = -1;					
				}
			else
				{
				lnResult = relate4top(gpCurrentQuery);
				if (lnResult == r4success) // Found a matching record.  Still need to check deletion status.
					{
					lnReturn = TRUE;	
					}
				else
					{
					if ((lnResult == r4eof) || (lnResult == r4terminate)) // No matching records.
						{
						lnReturn = FALSE;	
						}
					else
						{
						strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
						gnLastErrorNumber = codeBase.errorCode;
						lnReturn = -1;							
						}
					}	
				}	
			}
		}
	else
		{
		gcStrReturn[0] = (char) 0;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		lnReturn = -1;			
		}
	if ((lnReturn == TRUE) && (gnDeletedFlag == TRUE))
		{
		if (cbxDELETED())
			{
			// This "FOUND" record has been deleted, so we try for the next non-deleted one...
			lnReturn = FALSE; // be pessimistic
			while (relate4skip(gpCurrentQuery, 1) == r4success)
				{
				if (!cbxDELETED())
					{
					lnReturn = TRUE;
					break;	
					}	
				}
			}	
		}
	if (lnReturn == -1) return Py_BuildValue("");
	else return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}


/* ******************************************************************************* */
/* Equivalent to the VFP CONTINUE that gets the next matching record after a
   successful cbwLOCATE() function.  Returns FALSE if nothing else found. -1 on
   any kind of error.  EOF() or BOF() return FALSE, not error.                     */
static PyObject *cbwLOCATECONTINUE(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		if (gpCurrentQuery == NULL)
			{
			strcpy(gcErrorMessage, "No LOCATE active");
			gnLastErrorNumber = -9834;
			lnReturn = FALSE;			
			}
		else
			{
			lnResult = relate4skip(gpCurrentQuery, 1);
			if (lnResult == r4success)
				{
				lnReturn = TRUE;	
				}
			else
				{
				if ((lnResult == r4eof) || (lnResult == r4terminate) || (lnResult == r4bof)) // No matching records.
					{
					lnReturn = FALSE;	
					}
				else
					{
					strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
					gnLastErrorNumber = codeBase.errorCode;
					lnReturn = FALSE;							
					}					
				}
			}
		}
	else
		{
		gcStrReturn[0] = (char) 0;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		lnReturn = FALSE;			
		}
	if ((lnReturn == TRUE) && (gnDeletedFlag == TRUE))
		{
		if (cbxDELETED())
			{
			// This "FOUND" record has been deleted, so we try for the next non-deleted one...
			lnReturn = FALSE; // be pessimistic
			while (relate4skip(gpCurrentQuery, 1) == r4success)
				{
				if (!cbxDELETED())
					{
					lnReturn = TRUE;
					break;	
					}	
				}
			}	
		}
			
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}


/* ******************************************************************************* */
/* Function that terminates a cbwLOCATE process.  Call this after a series of
   cbwLOCATE and cbwCONTINUE commands to end the cycle and free up memory.
   Returns TRUE unless there is no current LOCATE process active, in which case
   it returns FALSE and the error values are set.                                  */
static PyObject *cbwLOCATECLEAR(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		if (gpCurrentQuery != NULL)
			{
			lnResult = relate4free(gpCurrentQuery, 0); // Doesn't close any tables
			if (lnResult != r4success) // Some kind of error???
				{
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
				gnLastErrorNumber = codeBase.errorCode;
				lnReturn = FALSE;					
				}
			else
				{
				gpCurrentQuery = NULL;	
				}
			}		
		}
	else
		{
		gcStrReturn[0] = (char) 0;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		lnReturn = FALSE;			
		}
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));
}

/* *********************************************************************************** */
/* Copies the index tags from one currently open table to another.  On the target      */
/* table removes the existing index tags before replacing them with the new tags.      */
/* Returns True on success, False on failure.                                          */
/* Note that the source and target alias values may not be blank.                      */
static PyObject *cbwCOPYTAGS(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	int lnResult = 0;
	DATA4 *lpFromTable = NULL;
	DATA4 *lpToTable = NULL;
	char lcSrcAlias[MAXALIASNAMESIZE + 1];
	char lcTrgAlias[MAXALIASNAMESIZE + 1];
	TAG4INFO* laSrcTags = NULL;
	INDEX4* lpSrcIndex;
	TAG4INFO* laTrgTags = NULL;
	TAG4INFO* lpTagPtr;
	INDEX4* lpTrgIndex;
	long lnCount = 0;
	long lnTrgCount = 0;
	long jj;
	TAG4* lpTag = NULL;
	PyObject *lxSourceAlias;
	PyObject *lxTargetAlias;
	const unsigned char *cTest;	
		
	
	if (!PyArg_ParseTuple(args, "OO", &lxSourceAlias, &lxTargetAlias))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}	
	
	cTest = testStringTypes(lxSourceAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcSrcAlias, Unicode2Char(lxSourceAlias), MAXALIASNAMESIZE);
	lcSrcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcSrcAlias, TRUE);
	//Alias names must be pure ASCII.
	
	cTest = testStringTypes(lxTargetAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTrgAlias, Unicode2Char(lxTargetAlias), MAXALIASNAMESIZE);
	lcTrgAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcTrgAlias, TRUE);	
	
	if ((strlen(lcSrcAlias) == 0) || (strlen(lcTrgAlias) == 0))
		{
		strcpy(gcErrorMessage, "Source and Target Alias may NOT be blank.");
		gnLastErrorNumber = -10001;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;			
		}	
	
	StrToUpper(lcSrcAlias);
	StrToUpper(lcTrgAlias);
	
	lpFromTable = code4data(&codeBase, lcSrcAlias);
	if (lpFromTable == NULL)
		{
		gnLastErrorNumber = codeBase.errorCode;
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		lnReturn = FALSE;				
		}		
	
	if (lnReturn)
		{
		lpToTable = code4data(&codeBase, lcTrgAlias);
		if (lpToTable == NULL)
			{
			gnLastErrorNumber = codeBase.errorCode;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			lnReturn = FALSE;				
			}
		}	

	if (lnReturn)
		{
		lpSrcIndex = d4index(lpFromTable, NULL);
		if (lpSrcIndex == NULL)
			{
			lnCount = 0;
			}
		else
			{
			laSrcTags = i4tagInfo(lpSrcIndex);
			if (laSrcTags != NULL)			
				{
				lnCount = 1; // Regardless of how many, it isn't 0.					
				}
			}
		// Next we remove the index and all tags from the target table...
		lpTrgIndex = d4index(lpToTable, NULL);
		if (lpTrgIndex == NULL)
			{
			lnTrgCount = 0;
			}
		else
			{
			laTrgTags = i4tagInfo(lpTrgIndex);
			if (laTrgTags != NULL)
				{
				for (jj = 0; jj < 1000;jj++) /* Well beyond practical limit */
					{
					lpTagPtr = laTrgTags + jj;						
					if (lpTagPtr->name == NULL) break;
						
					lpTag = d4tag(lpToTable, lpTagPtr->name);	
					if (lpTag != NULL)
						{
						lnResult = i4tagRemove(lpTag);
						if (lnResult != 0)
							{
							gnLastErrorNumber = codeBase.errorCode;
							strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));	
							lnReturn = FALSE;		
							}
						}
					}	
				}
			}
		}			
	
	if (lnReturn && (lnCount > 0))
		{
		lpTrgIndex = i4create(lpToTable, NULL, laSrcTags);
		if (lpTrgIndex == NULL)
			{
			gnLastErrorNumber = codeBase.errorCode;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));	
			lnReturn = FALSE;
			}
		}
	if (laSrcTags != NULL) u4free(laSrcTags);
	if (laTrgTags != NULL) u4free(laTrgTags);
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));
}

/* *********************************************************************************** */
/* Removes one specified index tag from the currently selected table.  Returns True on */
/* success, False on failure.                                                          */
static PyObject *cbwDELETETAG(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	int  lnResult = 0;
	PyObject *lxTagName;
	unsigned char lcTagName[100];
	const unsigned char *cTest;
	TAG4* lpTag = NULL;
	INDEX4* lpIndex;

	if (!PyArg_ParseTuple(args, "O", &lxTagName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}

	cTest = testStringTypes(lxTagName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTagName, Unicode2Char(lxTagName), 99);
	lcTagName[99] = (char) 0;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable == NULL)
		{
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		lnReturn = FALSE;
		}
	else
		{	
		if ((lcTagName == NULL) || (strlen(lcTagName) == 0))
			{
			strcpy(gcErrorMessage, "NO Tag Specified");
			gnLastErrorNumber = -9993;
			lnReturn = FALSE;
			}
		else
			{
			lpTag = d4tag(gpCurrentTable, lcTagName);	
			if (lpTag == NULL)
				{
				strcpy(gcErrorMessage, "Index Tag Does Not Exist");
				gnLastErrorNumber = -9992;
				lnReturn = FALSE;
				}
			else
				{
				lnResult = i4tagRemove(lpTag);
				if (lnResult == 0)
					{
					lnReturn = TRUE;	
					}
				else
					{
					gnLastErrorNumber = codeBase.errorCode;
					strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));	
					lnReturn = FALSE;		
					}
				}			
			}
		}
	
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}

/* ***************************************************************************** */
/* Equivalent of VFP SET ORDER TO which specifies at TAG name.  Returns 1 on OK  */
/* or -1 on Not OK.  Sets error code on why not ok.                              */
static PyObject *cbwSETORDERTO(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	TAG4 *lpTag;
	char lcTagName[100];
	PyObject* lxTagName;
	const unsigned char *cTest;
	
	if (!PyArg_ParseTuple(args, "O", &lxTagName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxTagName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTagName, Unicode2Char(lxTagName), 99);
	lcTagName[99] = (char) 0;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		if ((lcTagName == NULL) || (strlen(lcTagName) == 0))
			{
			lpTag = NULL; /* This is legal.  They just want record number order */
			}
		else
			{
			lpTag = d4tag(gpCurrentTable, lcTagName);	
			if (lpTag == NULL)
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, "Index Tag Does Not Exist");
				gnLastErrorNumber = -9992;	
				}
			}
		if (lnReturn == TRUE)
			{
			d4tagSelect(gpCurrentTable, lpTag);			
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}

/* ****************************************************************************** */
/* Equivalent of VFP ATAGINFO() which returns a list of dictionaries containing   */
/* information about each index tag associated with the currently selected table. */
/* Returns the Python List() object, which can be empty if there are no indexes   */
/* or None on error.                                                              */
static PyObject *cbwATAGINFO(PyObject *self, PyObject *args)
{
	long lnCount = 0;
	register long jj;
	long lnTest;
	INDEX4* lpIndex;
	TAG4INFO* laTags;
	TAG4INFO* laTagPtr;
	PyObject *tagList = (PyObject*) NULL;
	PyObject *infoDict = NULL;
	PyObject *xValue = NULL;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpIndex = d4index(gpCurrentTable, NULL);
		if (lpIndex == NULL)
			{
			lnCount = 0;
			tagList = PyList_New(0);
			}
		else
			{
			laTags = i4tagInfo(lpIndex);
			if (laTags != NULL)
				{
				tagList = PyList_New(0); // We don't know how long to make it.
				
				for (jj = 0; jj < 1000;jj++) /* Well beyond practical limit */
					{
					laTagPtr = laTags + jj;						
					if (laTagPtr->name == NULL) break;
					infoDict = PyDict_New();

#ifdef IS_PY3K
					lnTest = strlen(laTagPtr->name);			
					xValue = PyUnicode_DecodeLatin1((const char*) laTagPtr->name, lnTest, "replace");
					if (xValue == (void*) NULL)
						{
						strcpy(gcErrorMessage, "Illegal character found in tag name, not defined in Code Page 1252.");
						gnLastErrorNumber = -10011;
						PyErr_Format(PyExc_ValueError, gcErrorMessage);
						// error out now to avoid General protection fault on 0 pointer.	
						return NULL;
						}
#endif
#ifdef IS_PY2K					
					xValue = PyString_FromString((unsigned char*) laTagPtr->name);					
#endif
					PyDict_SetItemString(infoDict, "cTagName", xValue);
					
					Py_DECREF(xValue);
					xValue = NULL;

#ifdef IS_PY3K
					lnTest = strlen(laTagPtr->expression);			
					xValue = PyUnicode_DecodeLatin1((const char*) laTagPtr->expression, lnTest, "replace");
					if (xValue == (void*) NULL)
						{
						strcpy(gcErrorMessage, "Illegal character found in tag expression, not defined in Code Page 1252.");
						gnLastErrorNumber = -10011;
						PyErr_Format(PyExc_ValueError, gcErrorMessage);
						// error out now to avoid General protection fault on 0 pointer.	
						return NULL;
						}
#endif
#ifdef IS_PY2K					
					xValue = PyString_FromString((unsigned char*) laTagPtr->expression);					
#endif
					PyDict_SetItemString(infoDict, "cTagExpr", xValue);
					
					Py_DECREF(xValue);
					xValue = NULL;

#ifdef IS_PY3K
					lnTest = strlen(laTagPtr->filter);			
					xValue = PyUnicode_DecodeLatin1((const char*) laTagPtr->filter, lnTest, "replace");
					if (xValue == (void*) NULL)
						{
						strcpy(gcErrorMessage, "Illegal character found in tag filter, not defined in Code Page 1252.");
						gnLastErrorNumber = -10011;
						PyErr_Format(PyExc_ValueError, gcErrorMessage);
						// error out now to avoid General protection fault on 0 pointer.	
						return NULL;
						}
#endif
#ifdef IS_PY2K					
					xValue = PyString_FromString((unsigned char*) laTagPtr->filter);					
#endif
					PyDict_SetItemString(infoDict, "cTagFilt", xValue);
					
					Py_DECREF(xValue);
					xValue = NULL;
					
					xValue = Py_BuildValue("l", laTagPtr->descending);					
					PyDict_SetItemString(infoDict, "nDirection", xValue);
					
					Py_DECREF(xValue);
					xValue = NULL;
					
					xValue = Py_BuildValue("l", laTagPtr->unique);			
					PyDict_SetItemString(infoDict, "nUnique", xValue);
					
					Py_DECREF(xValue);
					xValue = NULL;
					
					PyList_Append(tagList, infoDict);
					Py_DECREF(infoDict);
					infoDict = NULL;
					lnCount += 1;
					}
				u4free(laTags);
				}
			else
				{
				strcpy(gcErrorMessage, "Memory error");
				gnLastErrorNumber = -8001;
				lnCount = -1;	
				}				
			}
		}
	else
		{
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		lnCount = -1;
		}
	if (lnCount == -1)
		{
		if (tagList != (PyObject*) NULL) Py_CLEAR(tagList);
		return Py_BuildValue("");	// Return None on error.
		}
	else return(tagList);	
}

/* ***************************************************************************** */
/* Returns a string pointer to a buffer containing the name of the currently     */
/* active index tag of the currently selected table.                             */
/* Returns NULL on error.  Buffer may be an empty string if no tag is set.       */
static PyObject *cbwORDER(PyObject *self, PyObject *args)
{
	TAG4* lpTag;
	long lnTest;
	unsigned char lcTagName[45];
	PyObject *xValue = NULL;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gcStrReturn[0] = (char) 0;
	
	if (gpCurrentTable != NULL)
		{
		lpTag = d4tagSelected(gpCurrentTable);
		if (lpTag != NULL)
			{
			strncpy(lcTagName, t4alias(lpTag), 44);
			lcTagName[44] = (char) 0;
#ifdef IS_PY3K
			lnTest = strlen(lcTagName);			
			xValue = PyUnicode_DecodeLatin1((const char*) lcTagName, lnTest, "replace");
			if (xValue == (void*) NULL)
				{
				strcpy(gcErrorMessage, "Illegal character found in tag name, not defined in Code Page 1252.");
				gnLastErrorNumber = -10011;
				PyErr_Format(PyExc_ValueError, gcErrorMessage);
				// error out now to avoid General protection fault on 0 pointer.	
				return NULL;
				}
#endif
#ifdef IS_PY2K					
			xValue = PyString_FromString((unsigned char*) t4alias(lpTag));					
#endif
			}
		else
			{
			return Py_BuildValue("s", "");	
			}
		}
	else
		{
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		return Py_BuildValue("");	// Return None on error.
		}
	return xValue;	
}


/* *********************************************************************************** */
/* Utility function that determines what the field is from a field name passed as a    */
/* char by reference.  On success returns the FIELD4 pointer.  On failure returns      */
/* NULL and sets the appropriate error messages so the calling program doesn't have to.*/
FIELD4* cbxGetFieldFromName(char* lpcFieldName)
{
	FIELD4* lpReturn = NULL;
	DATA4*   lpTable;
	long lnReturn = TRUE;
	char lcAliasName[MAXALIASNAMESIZE + 1];
	char lcFieldName[MAXALIASNAMESIZE + 30];
	char lcWorking[MAXALIASNAMESIZE + 30];
	char* lcPtrDot;
	char* lcPtrDash;
	char *lcPtr = NULL;
	long nOffset = 1;
	long lnRecord;
	
	strncpy(lcWorking, lpcFieldName, 129);
	lcWorking[129] = (char) 0;
	
	lcPtrDot = strstr(lcWorking, ".");
	lcPtrDash = strstr(lcWorking, "->");
	if (lcPtrDot != NULL) lcPtr = lcPtrDot;
	if (lcPtrDash != NULL)
		{
		lcPtr = lcPtrDash;
		nOffset = 2;
		}
	if (lcPtr != NULL)
		{
		*lcPtr = (char) 0;
		strncpy(lcFieldName, lcPtr + nOffset, MAXALIASNAMESIZE + 29);
		lcFieldName[MAXALIASNAMESIZE + 29] = (char) 0;
		lpTable = code4data(&codeBase, lcWorking); // Now has only the alias name.
		if (lpTable == NULL)
			{
			strcpy(gcErrorMessage, "Unidentified Table Alias");
			gnLastErrorNumber = -9984;
			lnReturn = FALSE;
			}			
		}
	else
		{
		if (gpCurrentTable)
			{
			lpTable = gpCurrentTable;
			strncpy(lcFieldName, lcWorking, MAXALIASNAMESIZE + 29);
			lcFieldName[MAXALIASNAMESIZE + 29] = (char) 0;			
			}			
		else	
			{
			strcpy(gcErrorMessage, "No Table Open in Selected Area");
			gnLastErrorNumber = -9999;
			lnReturn = FALSE;
			}	
		}
	if (lnReturn)
		{
		lnRecord = d4recNo(lpTable);
		if ((lnRecord < 1) || (lnRecord > d4recCount(lpTable)))
			{
			strcpy(gcErrorMessage, "No Current Record");
			gnLastErrorNumber = -9984;
			lnReturn = FALSE;				
			}
		}		
	if (lnReturn) /* OK so far, now we test the field name */
		{
		StrToLower(lcFieldName);
		lpReturn = d4field(lpTable, lcFieldName);
		if (lpReturn == NULL)
			{
			strcpy(gcErrorMessage, "Field Not Found: ");
			strcat(gcErrorMessage, lcFieldName);
			gnLastErrorNumber = -9985;
			lnReturn = FALSE;	
			}
		}
	return(lpReturn);	
}

/* ********************************************************************************** */
/* Function that works like the SEEK() function in VFP.  Difference is that all three */
/* are required in this version, and that the Codebase rules for passing all match    */
/* information in character form will apply.                                          */
/* lpcMatch is the string representation of what to seek for.  If a string, then      */
/* no problem.                                                                        */
/* If a date, then "YYYYMMDD".                                                        */
/* If a datetime, then "YYYYMMDDHH:MM:SS:TTT" where TTT is thousandths of seconds.    */
/* If a number of any type, then a string representation, e.g. "1004.35".             */
/* Returns TRUE on FOUND, FALSE on not found.  Will leave the record pointer at the   */
/* closest match if the match string is between the top and bottom values of the      */
/* specified index TAG, otherwise at EOF, so you need to test for that.               */
/* As with VFP, this does NOT change the currently selected table, so it's safe to    */
/* use while you're doing some looping through a table and need to look something up  */
/* in a different table.  Returns -1 on ERROR.                                        */
static PyObject *cbwSEEK(PyObject *self, PyObject *args)
{
	long lnReturn = FALSE;
	long lnResult;
	DATA4 *lpTable = NULL;
	INDEX4 *lpIndex;
	TAG4   *lpTag;
	TAG4   *lpOldTag = NULL; /* The tag prior to when we did our seek */
	PyObject *lxMatch;
	PyObject *lxAlias;
	PyObject *lxTagName;
	char lcAlias[MAXALIASNAMESIZE + 1];
	char lcMatch[400];
	char lcTagName[50];
	const unsigned char *cTest;

	if (!PyArg_ParseTuple(args, "OOO", &lxMatch, &lxAlias, &lxTagName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		gnLastErrorNumber = -10000;
		return NULL;
		}	

	cTest = testStringTypes(lxMatch);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcMatch, Unicode2Char(lxMatch), 399);
	lcMatch[399] = (char) 0;
	
	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE); // Alias is always ASCII.
	
	cTest = testStringTypes(lxTagName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTagName, Unicode2Char(lxTagName), 49);
	lcTagName[49] = (char) 0;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gcLastSeekString[0] = (char) 0; /* Initialize it to empty */

	if (strlen(lcAlias) == 0)
		{
		if (gpCurrentTable)
			{
			strncpy(lcAlias, d4alias(gpCurrentTable), MAXALIASNAMESIZE);
			lcAlias[MAXALIASNAMESIZE] = (char) 0;
			lpTable = code4data(&codeBase, lcAlias);
			}
		}
	else
		{
		lpTable = code4data(&codeBase, lcAlias);
		}
	
	if (lpTable == NULL)
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "Alias not found");
		gnLastErrorNumber = -9994;	
		}
	else
		{
		if (strlen(lcMatch) == 0)
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, "No Match Value Provided");
			gnLastErrorNumber = -99955;				
			}
		else
			{
			lpIndex = d4index(lpTable, NULL); /* Looks for the currently open production index */
			if (lpIndex == NULL)
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, "No Index Tags Available");
				gnLastErrorNumber = -9993;	
				}	
			else
				{
				/* We have an index and a table, so now we try to find the specified tag */
				lpOldTag = d4tagSelected(lpTable); /* keep track of this */
				if (strlen(lcTagName) == 0)
					{
					lpTag = lpOldTag; // Have to use the one currently active, if any.	
					}
				else
					{
					lpTag = d4tag(lpTable, lcTagName);	
					//printf("Tag from Name %s \n", lpcTagName);
					}
				
				if (lpTag == NULL)
					{
					lnReturn = -1;
					strcpy(gcErrorMessage, "No Order Set or Tag Not Recognized");
					gnLastErrorNumber = -9992;	
					}
				else
					{
					d4tagSelect(lpTable, lpTag);
					//printf("seeking in table %p \n", lpTable);
					lnResult = d4seek(lpTable, lcMatch);
					strcpy(gcLastSeekString, lcMatch);

					while (TRUE)
						{
						if (lnResult == r4success)
							{
							if (gnDeletedFlag == TRUE)
								{
								lnResult = cbxFINDUNDELETED(lpTable, 1, TRUE);
								if (lnResult == r4success)
									{
									lnReturn = 1;
									break; /* Found an undeleted one. */	
									}
								}
							else
								{
								lnReturn = 1; /* FOUND */
								break;
								}	
							}
						if ((lnResult == r4after) || (lnResult == r4eof))
							{
							lnReturn = 0; /* NOT FOUND, BUT NORMAL */
							break;	
							}
						if ((lnResult == r4entry) || (lnResult == r4locked) || (lnResult == r4unique) || (lnResult == r4noTag))
							{
							lnReturn = -1;
							strcpy(gcErrorMessage, "Seek Failed: record move not allowed.");
							gnLastErrorNumber = -9991;
							break;	
							}
						if (lnResult < 0)
							{
							lnReturn = -1;
							gnLastErrorNumber = codeBase.errorCode;
							strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
	 						break;
							}
						/* Otherwise something weird happened */
						lnReturn = -1;
						gnLastErrorNumber = 99999;
						strcpy(gcErrorMessage, "Unidentified Seek Error.");
						break;
						}
					if (lpOldTag != NULL) d4tagSelect(lpTable, lpOldTag); /* Put things back where they were */	
					}
				}
			}
		}
	if (lnReturn == -1) return Py_BuildValue("");
	else return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}


/* ************************************************************************************ */
/* Function that obtains the value of one field, specified by either just the field     */
/* name (for the currently selected table) or the full alias.fieldname notation.        */
/* Returns the field value in a character buffer to which the return value points.      */
/* The string may be empty, in which case check the last error message to be sure that  */
/* this was not caused by an error.                                                     */
/* New March 18, 2014.  Like cbwSCATTERFIELD, but if lbTrim is TRUE, then trims leading */
/* and trailing blanks from the string.                                                 */
static PyObject *cbwSCATTERFIELDCHAR(PyObject *self, PyObject *args)
{
	char *lcPtr;
	FIELD4 *lpField;
	long lbTrim;
	PyObject *lxRawField;
	static PyObject *lxRet;
	char lcFieldName[MAXALIASNAMESIZE + 30];
	const unsigned char *cTest;	
	long lbGoodData = TRUE;
	long lnTest;
	long bCopyFlag = FALSE;
	unsigned char *cTempBuff;

	if (!PyArg_ParseTuple(args, "Ol", &lxRawField, &lbTrim))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}
		
	cTest = testStringTypes(lxRawField);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxRawField), MAXALIASNAMESIZE + 29);
	lcFieldName[MAXALIASNAMESIZE + 29] = (char) 0;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lpField = cbxGetFieldFromName(lcFieldName);
	
	if (lpField != NULL)
		{
		if (f4null(lpField) == 1)
			{
			lbGoodData = FALSE; // This is OK actually, the field contains a NULL value so we don't set any error flags.
			return Py_BuildValue(""); //None is allowed for nullable fields.	
			}
		else
			{
			lcPtr = (char*) f4memoStr(lpField); /* Handles memo fields identically to non-memo fields */
			if (lcPtr == NULL)
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				lbGoodData = FALSE;
				}
			else
				{
				lnTest = strlen(lcPtr);	
				if (lbTrim)
					{
					cTempBuff = calloc((size_t) (lnTest + 1), sizeof(char));
					strncpy(cTempBuff, lcPtr, lnTest);
					*(cTempBuff + lnTest) = (char) 0;
					MPStrTrim('A', cTempBuff);
					bCopyFlag = TRUE;
					}
				else cTempBuff = lcPtr;
#ifdef IS_PY3K				
				lnTest = strlen(cTempBuff);						
				lxRet = PyUnicode_DecodeLatin1((const char*) cTempBuff, lnTest, "replace");
				if (lxRet == (void*) NULL)
					{
					strcpy(gcErrorMessage, "Illegal character found in field contents, not defined in Code Page 1252.");
					gnLastErrorNumber = -10011;
					PyErr_Format(PyExc_ValueError, gcErrorMessage);
					// error out now to avoid General protection fault on 0 pointer.
					if(bCopyFlag) free(cTempBuff);	
					return NULL;
					}
#endif
#ifdef IS_PY2K					
				lxRet = PyString_FromString((unsigned char*) cTempBuff);					
#endif			
				if(bCopyFlag) free(cTempBuff);
				}
			}	
		}
	else
		{
		gnLastErrorNumber = -25340;
		strcpy(gcErrorMessage, "Unrecognized Field Name");
		lbGoodData = FALSE;
		}
	if (lbGoodData)	return lxRet;
	else return Py_BuildValue(""); // None returned on error.
}


/* *********************************************************************************** */
/* Function that obtains the value of one field, specified by either just the field    */
/* name (for the currently selected table) or the full alias.fieldname notation.       */
/* Returns the field value in a character buffer to which the return value points.     */
/* The string may be empty, in which case check the last error message to be sure that */
/* this was not caused by an error. Only works for char, memo, number fields.          */
/* DOES NOT WORK WITH ANY FORM OF BINARY FIELDS OR DATA.                               */
static PyObject *cbwSCATTERFIELD(PyObject *self, PyObject *args)
{
	// DEPRECATED.  DO NOT USE.  MORE COMPLETE FUNCTIONALITY PROVIDED BY cbwSCATTERFIELDCHAR().
	char *lcPtr;
	FIELD4 *lpField;
	PyObject *lxFieldName;
	char lcFieldName[MAXALIASNAMESIZE + 30];
	const unsigned char *cTest;

	if (!PyArg_ParseTuple(args, "O", &lxFieldName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}

	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 29);
	lcFieldName[MAXALIASNAMESIZE + 29] = (char) 0;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gcStrReturn[0] = (char) 0;
	lpField = cbxGetFieldFromName(lcFieldName);
	
	if (lpField != NULL)
		{
		lcPtr = (char*) f4memoStr(lpField); /* Handles memo fields identically to non-memo fields */
		if (lcPtr == NULL)
			{
			gnLastErrorNumber = codeBase.errorCode;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gcStrReturn[0] = (char) 0; /* We're done... ERROR */
			}
		else
			{
			strcpy(gcStrReturn, lcPtr);
			}			
		}

	return Py_BuildValue("s", gcStrReturn);		
}

/* ***************************************************************************** */
/* Clears the buffers for the currently selected table, which forces a new read  */
/* the next time any part of the table is accessed.  This is more drastic than   */
/* the cbwREFRESHRECORD, in that it forces a complete re-read of the table header*/
/* and other parts of the table.  Will slow down processing if called on every   */
/* read in a loop.                                                               */
/* Returns TRUE for SUCCESS, False on ERROR.                                        */
/* Added 04/10/2014. JSH.                                                        */
static PyObject *cbwREFRESHBUFFERS(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4refresh(gpCurrentTable);
		if (lnResult == 0)
			{
			lnReturn = TRUE;
			}
		else
			{
			if (lnResult > 0) /* Save Failure error. */
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, "Buffer content not refreshed");
				gnLastErrorNumber = -99873;	
				}
			else /* ERROR */
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;					
				}
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}


/* ***************************************************************************** */
/* Refreshes the "current record" of the current table.                          */
/* Returns TRUE for SUCCESS, False on ERROR.                                        */
static PyObject *cbwREFRESHRECORD(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4refreshRecord(gpCurrentTable);
		if (lnResult == 0)
			{
			lnReturn = TRUE;
			}
		else
			{
			if (lnResult > 0) /* Save Failure error. */
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, "Record content not refreshed");
				gnLastErrorNumber = -9943;	
				}
			else /* ERROR */
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;					
				}
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}

/* ***************************************************************************** */
/* INTERNAL VERSION of cbwFLUSH() below.                                         */
long cbxFLUSH(void)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable)
		{
		lnResult = d4flush(gpCurrentTable);
		if (lnResult == 0)
			{
			lnReturn = TRUE;
			}
		else
			{
			if (lnResult > 0) /* Save Failure error. */
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, "Record content not saved");
				gnLastErrorNumber = -9995;	
				}
			else /* ERROR */
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;					
				}
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return lnReturn;	
}

/* ***************************************************************************** */
/* Flushes all data for the currently selected table.                            */
/* Returns TRUE for SUCCESS, False on ERROR.                           */
static PyObject *cbwFLUSH(PyObject *self, PyObject *args)
{
	return(Py_BuildValue("N", PyBool_FromLong(cbxFLUSH())));	
}

/* ***************************************************************************** */
/* Equivalent of RECALL ALL which applies to the currently selected table.  Just */
/* un-deletes ALL deleted records.  Returns TRUE (1) if successful, or -1 on     */
/* error.                                                                        */
static PyObject *cbwRECALLALL(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	TAG4 *saveSelected ;
	long lnCount = 0;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gnProcessTally = 0;
	
	if (gpCurrentTable != NULL)
		{
		saveSelected = d4tagSelected(gpCurrentTable) ;
		d4tagSelect( gpCurrentTable, NULL ) ; /* use record ordering */
		lnResult = d4lockAll( gpCurrentTable ) ;
		if (lnResult == 0)
			{
			for( d4top( gpCurrentTable ) ; !d4eof( gpCurrentTable ) ; d4skip( gpCurrentTable, 1 ) )
				{
				d4recall( gpCurrentTable ) ;
				lnCount++ ;
				}
			d4tagSelect( gpCurrentTable, saveSelected ) ; /* reset the selected tag.*/
			gnProcessTally = lnCount;
			}
		else
			{
			strcpy(gcErrorMessage, "Unable to lock table for recall all");
			gnLastErrorNumber = -9989;
			lnReturn = FALSE;	
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}

/* ***************************************************************************** */
/* Returns the most recent TALLY value like the _TALLY built in variable in VFP. */
static PyObject *cbwTALLY(PyObject *self, PyObject *args)
{
	return Py_BuildValue("l", gnProcessTally);
}


/* ***************************************************************************** */
/* Equivalent of RECALL which applies to the currently selected table.  Just     */
/* un-deletes the current record.  Returns TRUE (1) if successful, or False if      */
/* there is no current record or other problem.                                  */
static PyObject *cbwRECALL(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gnProcessTally = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4recNo(gpCurrentTable);
		if (lnResult < 0)
			{
			/* No current record to delete */
			strcpy(gcErrorMessage, "No current record to recall");
			gnLastErrorNumber = -9990;
			lnReturn = FALSE;	
			}
		else
			{
			d4recall(gpCurrentTable); /* OK, even if not really deleted */
			gnProcessTally = 1;
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}

/* ********************************************************************************** */
/* Flushes buffers like VFP FLUSH ALL where ALL table data is flushed.                */
/* Returns TRUE on success, False on ERROR.                                              */
static PyObject *cbwFLUSHALL(PyObject *self, PyObject *args)
{
	long lnResult;
	long lnReturn = TRUE;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
		
	lnResult = code4flush(&codeBase);
	if (lnResult != 0)
		{
		lnReturn = FALSE;
		gnLastErrorNumber = codeBase.errorCode;
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));			
		}
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));			
}

/* ******************************************************************************* */
/* Equivalent of FCOUNT() which returns the number of fields in the current table. */
/* Returns the number or -1 on error.                                              */
static PyObject *cbwFCOUNT(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4numFields(gpCurrentTable);
		if (lnResult < 1)
			{
			if (lnResult == 0) /* No fields defined -- ERROR, but not reported as such */
				{
				strcpy(gcErrorMessage, "No fields defined for current table.");
				gnLastErrorNumber = -9985;
				lnReturn = -1;	
				}
			else
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				lnReturn = -1;					
				}	
			}
		else
			{
			lnReturn = lnResult ; /* number of fields */	
			}
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return Py_BuildValue("l", lnReturn); // Returns numbers.	
}


/* ************************************************************************** */
/* Equivalent to VFP Function RECCOUNT() returning the number of records in   */ 
/* the table. Returns -1 on ERROR.                                            */
static PyObject *cbwRECCOUNT(PyObject *self, PyObject *args)
{
	long lnReturn = 0;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (gpCurrentTable == NULL)
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	else
		{
		lnResult = d4recCount(gpCurrentTable);
		if (lnResult < 0)
			{
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode;
			lnReturn = -1;				
			}
		else lnReturn = lnResult;	
		}	
	return Py_BuildValue("l", lnReturn); // Returns numbers	
}

/* ***************************************************************************** */
/* Returns EOF status for the currently selected table.                          */
/* Returns TRUE for EOF, FALSE for not or None on ERROR.                           */
static PyObject *cbwEOF(PyObject *self, PyObject *args)
{
	long lnReturn = -1;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4eof(gpCurrentTable);
		if (lnResult == 0) lnReturn = FALSE;
		if (lnResult > 0 ) lnReturn = TRUE;
		if (lnResult < 0 )
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode;				
			}
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	if (lnReturn == -1) return Py_BuildValue(""); // Returning None	
	else return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}

/* ***************************************************************************** */
/* Returns EOF status for the currently selected table.                          */
/* Returns TRUE for EOF, FALSE for not or -1 on ERROR.                           */
/* Internal Version.                                                             */
long cbxEOF(void)
{
	long lnReturn = -1;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4eof(gpCurrentTable);
		if (lnResult == 0) lnReturn = FALSE;
		if (lnResult > 0 ) lnReturn = TRUE;
		if (lnResult < 0 )
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode;				
			}
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return(lnReturn);	
}


/* ******************************************************************************** */
/* Like VFP SCATTER MEMVAR BLANK this method generates a dictionary of blank field  */
/* data to be used to populate some table record.  Takes an optional alias argument */
/* if you want the record from a table not currently selected.                      */
/* Returns the dict() on success or None on failure of any kind.                    */
static PyObject *cbwSCATTERBLANK(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnFldCnt, lnTest;
	long bBinaryAsUnicode = FALSE;
	register long jj;
	char *lpcAlias;
	FIELD4 *lpField;
	DATA4 *lpTable = NULL;
	PyObject *dataDict = NULL;
	PyObject *xValue = NULL;
	PyObject *xKey = NULL;
	PyObject *lxAlias;
	char lcAlias[MAXALIASNAMESIZE + 1];
	char cName[25];
	const unsigned char *cTest;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (!PyArg_ParseTuple(args, "Ol", &lxAlias, &bBinaryAsUnicode))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}

	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);
	
	if (strlen(lcAlias) == 0)
		{
		if (gpCurrentTable)
			{
			lpTable = gpCurrentTable;
			}
		else
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "No Table is Selected, Alias Not Found");	
			gnLastErrorNumber = -9994;
			}
		}
	else
		{
		lpTable = code4data(&codeBase, lcAlias);
		if (lpTable == NULL)
			{
			gnLastErrorNumber = codeBase.errorCode;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			lnReturn = FALSE;				
			}
		}
	
	if (lpTable)
		{
		lnFldCnt = d4numFields(lpTable);
		if (lnFldCnt < 1)
			{
			if (lnFldCnt == 0) /* No fields defined -- ERROR, but not reported as such */
				{
				strcpy(gcErrorMessage, "No fields defined for specified table.");
				gnLastErrorNumber = -9985;
				lnReturn = FALSE;
				}
			else
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				lnReturn = FALSE;
				}	
			}
		else
			{
			dataDict = PyDict_New();
			for (jj = 0;jj < lnFldCnt; jj++)
				{
				lpField = d4fieldJ(lpTable, (short) (jj + 1));				
				strncpy(cName, f4name(lpField), 24);
				cName[24] = (char) 0;
				
				xKey = cbNameToPy(cName);
				if (xKey == (void*) NULL)
					{	
					return NULL;
					}

				switch ((char) f4type(lpField))
				{
					case 'L':	
						xValue = Py_BuildValue("N", PyBool_FromLong(FALSE));
						break;
						
					case 'C':
					case 'M':
						xValue = Py_BuildValue("s", "");
						break;
						
					case 'N':
						xValue = Py_BuildValue("d", 0.0);
						break;
						
					case 'I':
						xValue = Py_BuildValue("l", 0);
						break;
						
					case 'Z':
					case 'X': // binary
#ifdef IS_PY3K	
						if (!bBinaryAsUnicode) xValue = PyBytes_FromStringAndSize("", 0);
						else xValue = Py_BuildValue("s", "");
#endif
#ifdef IS_PY2K					
						if (!bBinaryAsUnicode) xValue = PyByteArray_FromStringAndSize("", 0);
						else xValue = Py_BuildValue("s", "");
#endif
						break;
						
					case 'D':
					case 'T':
						xValue = Py_BuildValue(""); // Both D and T default to None when empty.
						break;
						
					case 'Y':
						xValue = Py_BuildValue("d", 0.0);
						break;
						
					case 'F':
						xValue = Py_BuildValue("d", 0.0);
						break;
						
					default:
						xValue = Py_BuildValue("s", "");
				}
		
				PyDict_SetItem(dataDict, xKey, xValue);	
				Py_DECREF(xValue);
				Py_DECREF(xKey);
				xValue = NULL;
				xKey = NULL;
				}
			}
		}

	if (lnReturn == FALSE)
		{
		if (dataDict != (PyObject*) NULL) Py_CLEAR(dataDict); // Get rid of it.
		return Py_BuildValue(""); // Return None
		}
	else return(dataDict);	
}

/* ******************************************************************************** */
/* Like VFP AFIELDS() in that it returns information about all the fields in        */
/* the currently selected table.  The field data is returned as a list() of dicts() */
/* with the cType, nWidth, nDecimals, cName, and bNulls values defining each field. */
/* Returns None on error.                                                           */
static PyObject *cbwAFIELDS(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	long lnFldCnt;
	long lnTest;
	char lcAlias[MAXALIASNAMESIZE + 1];
	const char *pTest;
	register long jj;
	char cType[2];
	FIELD4 *lpField;
	DATA4 *lpTable;
	FIELD4INFO *lpStruct;
	FIELD4INFO *lpInfo;
	PyObject *fieldList = NULL;
	PyObject *fieldItem;
	PyObject *xValue;
	PyObject *xKey;
	PyObject *lxAlias;
	const unsigned char *cTest;
	
	if (!PyArg_ParseTuple(args, "O", &lxAlias))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}	

	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (strlen(lcAlias) == 0)
		{
		if (gpCurrentTable)
			{
			lpTable = gpCurrentTable;
			}
		else
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "No Table is Selected, Alias Not Found");	
			gnLastErrorNumber = -9994;
			lpTable = NULL;
			}
		}
	else
		{
		lpTable = code4data(&codeBase, lcAlias);
		if (lpTable == NULL)
			{
			gnLastErrorNumber = codeBase.errorCode;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			lnReturn = FALSE;				
			}
		}

	if (lpTable != NULL)
		{
		lnResult = d4numFields(lpTable);
		if (lnResult < 1)
			{
			if (lnResult == 0) /* No fields defined -- ERROR, but not reported as such */
				{
				strcpy(gcErrorMessage, "No fields defined for current table.");
				gnLastErrorNumber = -9985;
				lnReturn = -1;
				}
			else
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				lnReturn = -1;
				}	
			}
		else
			{
			lnFldCnt = lnResult ; /* number of fields */
			lnReturn = lnResult; /* Now fill the gaFields[] array */
			fieldList = PyList_New(0);
			lpStruct = d4fieldInfo(lpTable);
			if (lpStruct != NULL)
				{
				for (jj = 0;jj < lnFldCnt; jj++)
					{
					//lpField = d4fieldJ(lpTable, (short) (jj + 1));
					lpInfo = lpStruct + jj;
					if (lpInfo->name == (char) NULL) break;				
					fieldItem = PyDict_New();

					xValue = cbNameToPy(lpInfo->name);
					if (xValue == (void*) NULL)
						{
						return NULL;
						}
					xKey = stringToPy("cName");							
					PyDict_SetItem(fieldItem, xKey, xValue);
					
					Py_DECREF(xValue);
					Py_DECREF(xKey);

					cType[0] = (char) lpInfo->type;
					cType[1] = (char) 0;
					xValue = Py_BuildValue("s", cType);
					xKey = stringToPy("cType");	 
					PyDict_SetItem(fieldItem, xKey, xValue);
					Py_DECREF(xValue);
					Py_DECREF(xKey);		

					xValue = Py_BuildValue("l", lpInfo->len);
					xKey = stringToPy("nWidth");	
					PyDict_SetItem(fieldItem, xKey, xValue);
					Py_DECREF(xValue);
					Py_DECREF(xKey);
					
					//xValue = Py_BuildValue("l", f4decimals(lpField));
					xValue = Py_BuildValue("l", lpInfo->dec);
					xKey = stringToPy("nDecimals");
					PyDict_SetItem(fieldItem, xKey, xValue);
					Py_DECREF(xValue);
					Py_DECREF(xKey);
								
					lnTest = (lpInfo->nulls == r4null);
					xValue = Py_BuildValue("N", PyBool_FromLong(lnTest));
					xKey = stringToPy("bNulls");

					PyDict_SetItem(fieldItem, xKey, xValue);
					Py_DECREF(xValue);
					Py_DECREF(xKey);
					
					PyList_Append(fieldList, fieldItem);
					Py_DECREF(fieldItem);	
					}
				u4free(lpStruct);	
				}
			}
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		}
	if (lnReturn == -1)
		{
		if (fieldList != (PyObject*) NULL) Py_CLEAR(fieldList); // Don't need it.
		return Py_BuildValue(""); // Return None
		}
	else return(fieldList);	
}

/* ***************************************************************************** */
/* Like VFP AFIELDS() in that it returns information about all the fields in     */
/* the currently selected table.  The field information is stored in the global  */
/* array gaFields a pointer to which is returned by the function.  Pass a long   */
/* int by reference to get the count.  Be sure to grab the field info from the   */
/* array before calling this again, as the array will be overwritten each time.  */
/* You can iterate through the fields and when you come to a cType[0] value of 0,*/
/* you are done.                                                                 */
/* If there is nothing there, then check the lpnCount.  It will be -1 on error.  */
/* INTERNAL VERSION.                                                             */
VFPFIELD* cbxAFIELDS(long* lpnCount)
{
	long lnReturn = TRUE;
	long lnResult;
	long lnFldCnt;
	FIELD4INFO *lpStruct;
	FIELD4INFO *lpInfo;	
	long jj;
	FIELD4 *lpField;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	memset(gaFields, 0, (size_t) 256 * sizeof(VFPFIELD));
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4numFields(gpCurrentTable);
		if (lnResult < 1)
			{
			if (lnResult == 0) /* No fields defined -- ERROR, but not reported as such */
				{
				strcpy(gcErrorMessage, "No fields defined for current table.");
				gnLastErrorNumber = -9985;
				lnReturn = -1;
				*lpnCount = -1;	
				}
			else
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				lnReturn = -1;
				*lpnCount = -1;					
				}	
			}
		else
			{
			lnFldCnt = lnResult ; /* number of fields */
			lnReturn = lnResult; /* Now fill the gaFields[] array */
			*lpnCount = lnFldCnt;
			lpStruct = d4fieldInfo(gpCurrentTable);
			if (lpStruct != NULL)
				{
				for (jj = 0;jj < lnFldCnt; jj++)
					{
					//lpField = d4fieldJ(gpCurrentTable, (short) (jj + 1));	
					lpInfo = lpStruct + jj;			
					//strcpy(gaFields[jj].cName, f4name(lpField));
					strcpy(gaFields[jj].cName, lpInfo->name);
					//gaFields[jj].cType[0]  = (char) f4type(lpField);
					gaFields[jj].cType[0] = (char) lpInfo->type;
					gaFields[jj].cType[1]  = (char) 0;
					//gaFields[jj].nWidth    = f4len(lpField);
					gaFields[jj].nWidth = lpInfo->len;
					//gaFields[jj].nDecimals = f4decimals(lpField);
					gaFields[jj].nDecimals = lpInfo->dec;
					//gaFields[jj].bNulls    = f4null(lpField);
					gaFields[jj].bNulls = lpInfo->nulls;
					}
				u4free(lpStruct);
				}
			else
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, "Insufficient Memory to Determine Structure");
				gnLastErrorNumber = -998124;
				*lpnCount = -1;	
				}
			}
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		*lpnCount = -1;
		}
	return(gaFields);	
}


/* ******************************************************************************** */
/* Specialized, fast field info with just the field types in a dict().              */
/* Returns None on error.                                                           */
static PyObject *cbwAFIELDTYPES(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	long lnFldCnt, lnTest;
	register long jj;
	char cType[2];
	FIELD4 *lpField;
	PyObject *fieldTypes = NULL;
	PyObject *xValue = NULL;
	PyObject *xKey = NULL;
	unsigned char cFldName[40];
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4numFields(gpCurrentTable);
		if (lnResult < 1)
			{
			if (lnResult == 0) /* No fields defined -- ERROR, but not reported as such */
				{
				strcpy(gcErrorMessage, "No fields defined for current table.");
				gnLastErrorNumber = -9985;
				lnReturn = -1;
				}
			else
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				lnReturn = -1;
				}	
			}
		else
			{
			lnFldCnt = lnResult ; /* number of fields */
			lnReturn = lnResult; 
			fieldTypes = PyDict_New();
			for (jj = 0;jj < lnFldCnt; jj++)
				{
				lpField = d4fieldJ(gpCurrentTable, (short) (jj + 1));				
	
				cType[0] = (char) f4type(lpField);
				cType[1] = (char) 0;
				xValue = Py_BuildValue("s", cType);

				strncpy(cFldName, f4name(lpField), 39);
				cFldName[39] = (char) 0;
				xKey = cbNameToPy(cFldName);

				if (xKey == (void*) NULL)
					{	
					return NULL;
					}
				PyDict_SetItem(fieldTypes, xKey, xValue);
				Py_DECREF(xValue);
				Py_DECREF(xKey);
				xValue = NULL;
				xKey = NULL;
				}
			}
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		}
	if (lnReturn == -1)
		{
		if (fieldTypes != (PyObject*) NULL) Py_CLEAR(fieldTypes); // Don't  need it any more.
		return Py_BuildValue(""); // Return None
		}
	else return(fieldTypes);	
}

/* ********************************************************************************** */
/* Converts the contents of the specified 'C' or 'M' field to a Python object var     */
/* of the appropriate type and text coding, if any.                                   */
static PyObject *cbxMakeCharVar(FIELD4 *lppField, unsigned char lcCodeAsc, long lpbStripBlanks, unsigned char cType)
{
	static unsigned char cWorkBuff[1026];
	long nLen;
	PyObject *lxRetVal = NULL;
	unsigned char *cBuffBig;
	long bBigMode = FALSE;
	char *cBuff;
	char *cMsg;
	
	if (cType == 'C') nLen = f4len(lppField);
	else nLen = f4memoLen(lppField);

	if (nLen < 1024)
		{
		memset(cWorkBuff, 0, (size_t) ((nLen + 2) * sizeof(unsigned char)));
		cBuff = cWorkBuff;
		}
	else
		{
		cBuff = calloc((size_t) (nLen + 2), sizeof(char));
		bBigMode = TRUE;
		}
	memcpy(cBuff, f4memoPtr(lppField), (size_t) (nLen * sizeof(unsigned char)));
	
#ifdef IS_PY3K
	switch(lcCodeAsc)
	{
	case 'W':
		lxRetVal = PyUnicode_DecodeMBCS(cBuff, (Py_ssize_t) nLen, "replace");
		cMsg = "Localized Windows Double Byte coding";
		break;
		
	case '8':
		lxRetVal = PyUnicode_DecodeUTF8((const char*) cBuff, (Py_ssize_t) nLen, "replace");
		cMsg = "UTF-8 coding";		
		break;
		
	case 'X':
	default:
		if (lpbStripBlanks) MPStrTrim('R', cBuff);
		lxRetVal = PyUnicode_DecodeLatin1((const char*) cBuff, (Py_ssize_t) nLen, "replace");		
		cMsg = "Extended ASCII with Code Page 1252";
		break;
	}
#endif
#ifdef IS_PY2K
	switch(lcCodeAsc)
	{
	case 'W':
		lxRetVal = PyUnicode_DecodeMBCS(cBuff, (Py_ssize_t) nLen, "replace");
		cMsg = "Localized Windows Double Byte coding";
		break;
		
	case '8':
		lxRetVal = PyUnicode_FromStringAndSize(cBuff, (Py_ssize_t) nLen); // stores as UTF-8.
		cMsg = "UTF-8 coding";	
		break;
		
	case 'X':
	default:
		if (lpbStripBlanks) MPStrTrim('R', cBuff);		
		lxRetVal = PyString_FromString(cBuff); // recognizes the null terminator, regular string.
		cMsg = "Extended ASCII";
		break;
	}
#endif
	if (bBigMode) free(cBuff);
	if (lxRetVal == (void*) NULL)
		{
		strcpy(gcErrorMessage, "Illegal character found, not defined in specified encoding: ");
		strcat(gcErrorMessage, cMsg);
		gnLastErrorNumber = -10011;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		// error out now to avoid General protection fault on 0 pointer.	
		return NULL;
		}	
	return(lxRetVal);
}


/* ********************************************************************************** */
/* Converts the contents of the specified 'C' or 'M' field to a Python object var     */
/* of the appropriate type and text coding, if any.                                   */
static PyObject *cbxMakeCharBinVar(FIELD4 *lppField, unsigned char lcCodeAsc, long lpbStripBlanks, unsigned char cType)
{
	static unsigned char cWorkBuff[1026];
	long nLen;
	PyObject *lxRetVal = NULL;
	unsigned char *cBuffBig;
	long bBigMode = FALSE;
	char *cBuff;
	char cMsg[100];
	
	if (cType == 'C') nLen = f4len(lppField);
	else nLen = f4memoLen(lppField);
	
	if (nLen < 1024)
		{
		memset(cWorkBuff, 0, (size_t) ((nLen + 2) * sizeof(unsigned char)));
		cBuff = cWorkBuff;
		}
	else
		{
		cBuff = calloc((size_t) (nLen + 2), sizeof(char));
		bBigMode = TRUE;
		}
	memcpy(cBuff, f4memoPtr(lppField), (size_t) (nLen * sizeof(unsigned char)));
	
#ifdef IS_PY3K
	switch(lcCodeAsc)
	{
	case 'D':
		lxRetVal = lxRetVal = PyUnicode_DecodeMBCS(cBuff, (Py_ssize_t) nLen, "replace");
		strcpy(cMsg, "Localized Windows Double Byte coding");
		break;
		
	case 'U':
		lxRetVal = PyUnicode_DecodeUTF8((const char*) cBuff, (Py_ssize_t) nLen, "replace");
		strcpy(cMsg, "UTF-8 coding");		
		break;

	case 'C':
		lxRetVal = PyUnicode_Decode((const char*) cBuff, (Py_ssize_t) nLen, gcCustomCodePage, "replace");
		strcpy(cMsg, "Custom Decoding with ");
		strcat(cMsg, gcCustomCodePage);
		break;

	case 'X':
	default:
		lxRetVal = PyBytes_FromStringAndSize((const char*) cBuff, (Py_ssize_t) nLen);
		strcpy(cMsg, "Raw Binary");
		break;
	}
#endif
#ifdef IS_PY2K
	switch(lcCodeAsc)
	{
	case 'D':
		lxRetVal = PyUnicode_DecodeMBCS(cBuff, (Py_ssize_t) nLen, "replace");
		strcpy(cMsg, "Localized Windows Double Byte coding");
		break;
		
	case 'C':
		lxRetVal = PyUnicode_Decode((const char*) cBuff, (Py_ssize_t) nLen, gcCustomCodePage, "replace");
		strcpy(cMsg, "Custom Decoding with ");
		strcat(cMsg, gcCustomCodePage);
		break;
		
	case 'U':
		lxRetVal = PyUnicode_FromStringAndSize(cBuff, (Py_ssize_t) nLen); // stores as UTF-8.
		strcpy(cMsg, "UTF-8 coding");	
		break;

	case 'X':
	default:
		lxRetVal = PyByteArray_FromStringAndSize(cBuff, (Py_ssize_t) nLen); // ignores nulls, copies everything.
		strcpy(cMsg, "Raw Binary");
		break;
	}
#endif
	if (bBigMode) free(cBuff);
	if (lxRetVal == (void*) NULL)
		{
		strcpy(gcErrorMessage, "Illegal character found, not defined in specified encoding: ");
		strcat(gcErrorMessage, cMsg);
		gnLastErrorNumber = -10011;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		// error out now to avoid General protection fault on 0 pointer.	
		return NULL;
		}	
	return(lxRetVal);
}

/* ********************************************************************************** */
/* Converts the contents of the Python string or bytecode object and stores the       */
/* into the specified field using the requested conversion.                           */
long cbxPyToCharFld(FIELD4 *lppField, PyObject *lpxString, unsigned char cType, unsigned char lcCodeAsc)
{
	long nLen;
	char *cMsg;
	size_t nSourceLen;
	long bIsBinary = FALSE; // Any bytearray or bytes value is, by definition and is stored without encoding.
	long lnRet = r4success;
	long bDone = FALSE;
	char cFldName[30];
	PyObject *pyBytes;
	PyObject *pyString;
	char *cBuff;
	Py_ssize_t nDataSize;

#ifdef IS_PY3K

	if(PyBytes_Check(lpxString)) // Alwas raw copy
		{
		nDataSize = PyBytes_Size(lpxString);
		lnRet = f4memoAssignN(lppField, PyBytes_AsString(lpxString), (unsigned long) nDataSize);
		bDone = TRUE;
		}
	if(PyByteArray_Check(lpxString)) // Alwas raw copy
		{
		nDataSize = PyByteArray_Size(lpxString);
		lnRet = f4memoAssignN(lppField, PyByteArray_AsString(lpxString), (unsigned long) nDataSize);
		bDone = TRUE;
		}	

	if (!bDone)
		{
		if (PyUnicode_Check(lpxString))
			{
			switch(lcCodeAsc)
				{
				case 'W':
					cMsg = "Localized Windows Double Byte coding";
					pyString = PyUnicode_AsMBCSString(lpxString);
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 29);
						cFldName[29] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return -1;
						}
					nDataSize = PyBytes_Size(pyString);
					cBuff = PyBytes_AsString(pyString);					
					if (cType == 'C') f4assignN(lppField, cBuff, (unsigned long) nDataSize);
					else
						{
						lnRet = f4memoAssignN(lppField, cBuff, (unsigned long) nDataSize);	
						}					
					break;
					
				case '8':
					cMsg = "UTF-8 coding";
					pyString = PyUnicode_AsUTF8String(lpxString);
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 29);
						cFldName[29] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;;
						}					
					nDataSize = PyBytes_Size(pyString);
					cBuff = PyBytes_AsString(pyString);
					if (cType == 'C') f4assignN(lppField, cBuff, (unsigned long) nDataSize);
					else
						{
						lnRet = f4memoAssignN(lppField, cBuff, (unsigned long) nDataSize);	
						}					
					break;
					
				case 'X':
				default:
					cMsg = "Extended ASCII";  // null terminators
					pyString = PyUnicode_AsEncodedString(lpxString, "cp1252", "replace");
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 29);
						cFldName[29] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return -1;
						}
					cBuff = PyBytes_AsString(pyString);
					if (cType == 'C') f4assign(lppField, cBuff);
					else lnRet = f4memoAssign(lppField, cBuff);				
					break;
				}
			bDone = TRUE;				
			}
		}
		
	if(!bDone)
		{
		// not a string or byte array, so we just convert to a string	
		pyString = PyUnicode_AsLatin1String(PyObject_Str(lpxString));
		if (pyString != NULL)
			{
			cBuff = PyBytes_AsString(pyString);
			lnRet = f4memoAssign(lppField, cBuff);
			}	
		else
			{
			lnRet = -1;
			}
		}
#endif
#ifdef IS_PY2K
	if(PyByteArray_Check(lpxString)) // Alwas raw copy
		{
		nDataSize = PyByteArray_Size(lpxString);
		lnRet = f4memoAssignN(lppField, PyByteArray_AsString(lpxString), (unsigned long) nDataSize);
		}
	else
		{
		if (PyString_CheckExact(lpxString))
			{
			nDataSize = PyString_Size(lpxString);
			cBuff = PyString_AsString(lpxString);				
			switch(lcCodeAsc)
				{	
				case 'R': // raw, ignores embedded nulls.
					cMsg = "Raw string copy";
					if (cType == 'C') f4assignN(lppField, cBuff, (unsigned long) nDataSize);
					else
						{
						lnRet = f4memoAssignN(lppField, cBuff, (unsigned long) nDataSize);	
						} 
					break;
					
				case 'X':
				default: // nulls define end of string
					cMsg = "Extended ASCII";
					if (cType == 'C') f4assign(lppField, cBuff);
					else lnRet = f4memoAssign(lppField, cBuff);
					break;
				}
			bDone = TRUE;			
			}
		if (!bDone && PyUnicode_CheckExact(lpxString))
			{
			switch(lcCodeAsc)
				{
				case 'W':
					cMsg = "Localized Windows Double Byte coding";
					pyString = PyUnicode_AsMBCSString(lpxString);
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 29);
						cFldName[29] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;;
						}					
					nDataSize = PyString_Size(pyString);
					cBuff = PyString_AsString(pyString);					
					if (cType == 'C') f4assignN(lppField, cBuff, (unsigned long) nDataSize);
					else
						{
						lnRet = f4memoAssignN(lppField, cBuff, (unsigned long) nDataSize);	
						}					
					break;
					
				case '8':
					cMsg = "UTF-8 coding";
					pyString = PyUnicode_AsUTF8String(lpxString);
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 29);
						cFldName[29] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;;
						}					
					nDataSize = PyString_Size(pyString);
					cBuff = PyString_AsString(pyString);
					if (cType == 'C') f4assignN(lppField, cBuff, (unsigned long) nDataSize);
					else
						{
						lnRet = f4memoAssignN(lppField, cBuff, (unsigned long) nDataSize);	
						}					
					break;
					
				case 'R':
					cMsg = "Raw string copy";
					nDataSize = PyUnicode_GET_DATA_SIZE(lpxString);
					cBuff = (unsigned char *) PyUnicode_AsUnicode(lpxString);
					if (cType == 'C') f4assignN(lppField, cBuff, (unsigned long) nDataSize);
					else
						{
						lnRet = f4memoAssignN(lppField, cBuff, (unsigned long) nDataSize);	
						}
					break;
					
				case 'X':
				default:
					cMsg = "Extended ASCII";  // null terminators
					pyString = PyUnicode_AsEncodedString(lpxString, "cp1252", "replace");
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 29);
						cFldName[29] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;;
						}
					cBuff = PyString_AsString(pyString);
					if (cType == 'C') f4assign(lppField, cBuff);
					else lnRet = f4memoAssign(lppField, cBuff);				
					break;
				}
			bDone = TRUE;				
			}
		
		if(!bDone)
			{
			// not a string or byte array, so we just convert to a string
			f4memoAssign(lppField, PyString_AsString(PyObject_Str(lpxString)));
			}
		}
#endif
	if (lnRet < 0)
		{
		// something went wrong
		strcpy(gcErrorMessage, "Writing data to field failed: ");
		if (codeBase.errorCode != 0) strcat(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		else strcat(gcErrorMessage, "UNKNOWN REASON.");
		}
	return(lnRet);
}


/* ********************************************************************************** */
/* Converts the contents of the specified 'C' or 'M' field to a Python object var     */
/* of the appropriate type and text coding, if any.                                   */
/* Note that char (binary) fields are marked as type 'Z' and memo (binary) fields are */
/* type code 'X'.                                                                     */
long cbxPyToCharFldBin(FIELD4 *lppField, PyObject *lpxString, unsigned char cType, unsigned char lcCodeAsc)
{
	long nLen;
	char cMsg[150];
	size_t nSourceLen;
	long bIsBinary = FALSE; // Any bytearray or bytes value is, by definition and is stored without encoding.
	long lnRet = r4success;
	long bDone = FALSE;
	char cFldName[100];
	static char cBinBuff[1000];
	char *cWorkBuff;
	long bNeedFree = FALSE;
	long nFldSize = 0;
	PyObject *pyBytes;
	PyObject *pyString;
	char *cBuff;
	Py_ssize_t nDataSize;

	if (cType == 'C')
		{
		nFldSize = f4len(lppField);
		if (nFldSize > 950)
			{
			cWorkBuff = calloc((size_t) nFldSize, sizeof(char)); // padded with nulls
			bNeedFree = TRUE;
			}
		else
			{
			memset(cBinBuff, (size_t) (nFldSize + 2), sizeof(char));
			cWorkBuff = cBinBuff;
			}
		}

#ifdef IS_PY3K

	if(PyBytes_Check(lpxString)) // Alwas raw copy
		{
		nDataSize = PyBytes_Size(lpxString);
		if (cType == 'C')
			{
			memcpy(cWorkBuff, PyBytes_AsString(lpxString), (size_t) nDataSize);
			f4assignN(lppField, cWorkBuff, nFldSize); // result is padded with nulls, not spaces.
			}
		else lnRet = f4memoAssignN(lppField, PyBytes_AsString(lpxString), (unsigned long) nDataSize);
		bDone = TRUE;
		}
	if(PyByteArray_Check(lpxString)) // Alwas raw copy
		{
		nDataSize = PyByteArray_Size(lpxString);
		if (cType == 'C')
			{
			memcpy(cWorkBuff, PyByteArray_AsString(lpxString), (size_t) nDataSize);
			f4assignN(lppField, cWorkBuff, nFldSize); // result is padded with nulls, not spaces.
			}
		else lnRet = f4memoAssignN(lppField, PyByteArray_AsString(lpxString), (unsigned long) nDataSize);
		bDone = TRUE;
		}

	if (!bDone)
		{
		if (PyUnicode_Check(lpxString))
			{
			switch(lcCodeAsc)
				{
				case 'D':
					strcpy(cMsg, "Localized Windows Double Byte coding");
					pyString = PyUnicode_AsMBCSString(lpxString);
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 99);
						cFldName[99] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;;
						}
					nDataSize = PyBytes_Size(pyString);
					memcpy(cWorkBuff, PyBytes_AsString(lpxString), (size_t) nDataSize);				
					if (cType == 'C')
						{
						f4assignN(lppField, cWorkBuff, (unsigned long) nFldSize);
						}
					else
						{
						lnRet = f4memoAssignN(lppField, cWorkBuff, (unsigned long) nDataSize);	
						}					
					break;
					
				case 'U':
					strcpy(cMsg, "UTF-8 coding");
					pyString = PyUnicode_AsUTF8String(lpxString);
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 99);
						cFldName[99] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;;
						}					
					nDataSize = PyBytes_Size(pyString);
					memcpy(cWorkBuff, PyBytes_AsString(lpxString), (size_t) nDataSize);				
					if (cType == 'C')
						{
						f4assignN(lppField, cWorkBuff, (unsigned long) nFldSize);
						}
					else
						{
						lnRet = f4memoAssignN(lppField, cWorkBuff, (unsigned long) nDataSize);	
						}					
					break;

				case 'C':
					strcpy(cMsg, "Custom coding");
					pyString = PyUnicode_AsEncodedString(lpxString, gcCustomCodePage, "replace");
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 99);
						sprintf(gcErrorMessage, "String coding for: %s %s, Field %s, FAILED.", cMsg, gcCustomCodePage, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;;
						}					
					nDataSize = PyBytes_Size(pyString);
					memcpy(cWorkBuff, PyBytes_AsString(lpxString), (size_t) nDataSize);				
					if (cType == 'C')
						{
						f4assignN(lppField, cWorkBuff, (unsigned long) nFldSize);
						}
					else
						{
						lnRet = f4memoAssignN(lppField, cWorkBuff, (unsigned long) nDataSize);	
						}					
					break;
				
				case 'X':
				default:
					strcpy(cMsg, "Extended ASCII");  // null terminators
					pyString = PyUnicode_AsEncodedString(lpxString, "cp1252", "replace");
					// same as Latin1
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 99);
						cFldName[99] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;;
						}
					cBuff = PyBytes_AsString(pyString);
					if (cType == 'C') f4assign(lppField, cBuff);
					else lnRet = f4memoAssign(lppField, cBuff);				
					break;
				}
			bDone = TRUE;				
			}
		}
		
	if(!bDone)
		{
		// not a string or byte array, so we just convert to a string
		pyString = PyUnicode_AsLatin1String(PyObject_Str(lpxString));
		if (pyString != NULL)
			{
			cBuff = PyBytes_AsString(pyString);
			lnRet = f4memoAssign(lppField, cBuff);
			}
		else return -1;
		}
#endif
#ifdef IS_PY2K
	if(PyByteArray_Check(lpxString)) // Alwas raw copy
		{
		nDataSize = PyByteArray_Size(lpxString);
		lnRet = f4memoAssignN(lppField, PyByteArray_AsString(lpxString), (unsigned long) nDataSize);
		}
	else
		{
		if (PyString_CheckExact(lpxString))
			{
			nDataSize = PyString_Size(lpxString);
			cBuff = PyString_AsString(lpxString);				
			switch(lcCodeAsc)
				{	
				case 'R': // raw, ignores embedded nulls.
					strcpy(cMsg, "Raw string copy");
					if (cType == 'C') f4assignN(lppField, cBuff, (unsigned long) nDataSize);
					else
						{
						lnRet = f4memoAssignN(lppField, cBuff, (unsigned long) nDataSize);	
						} 
					break;
					
				case 'X':
				default: // nulls define end of string
					strcpy(cMsg, "Extended ASCII");
					if (cType == 'C') f4assign(lppField, cBuff);
					else lnRet = f4memoAssign(lppField, cBuff);
					break;
				}
			bDone = TRUE;			
			}
		if (!bDone && PyUnicode_CheckExact(lpxString))
			{
			switch(lcCodeAsc)
				{
				case 'D':
					strcpy(cMsg, "Localized Windows Double Byte coding");
					pyString = PyUnicode_AsMBCSString(lpxString);
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 99);
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;
						}
					nDataSize = PyString_Size(pyString);
					memcpy(cWorkBuff, PyString_AsString(lpxString), (size_t) nDataSize);				
					if (cType == 'C')
						{
						f4assignN(lppField, cWorkBuff, (unsigned long) nFldSize);
						}
					else
						{
						lnRet = f4memoAssignN(lppField, cWorkBuff, (unsigned long) nDataSize);	
						}												
					break;
					
				case 'U':
					strcpy(cMsg, "UTF-8 coding");
					pyString = PyUnicode_AsUTF8String(lpxString);
					if (pyString == NULL)
						{
						strncpy(cFldName, f4name(lppField), 99);
						cFldName[99] = (char) 0;
						sprintf(gcErrorMessage, "String coding for: %s, Field %s, FAILED.", cMsg, cFldName);
						PyErr_Format(PyExc_ValueError, gcErrorMessage);	
						return  -1;;
						}					
					nDataSize = PyString_Size(pyString);
					memcpy(cWorkBuff, PyString_AsString(lpxString), (size_t) nDataSize);				
					if (cType == 'C')
						{
						f4assignN(lppField, cWorkBuff, (unsigned long) nFldSize);
						}
					else
						{
						lnRet = f4memoAssignN(lppField, cWorkBuff, (unsigned long) nDataSize);	
						}												
					break;
					
				case 'R':
				case 'X':
					strcpy(cMsg, "Raw string copy - Legacy Unicode");
					nDataSize = PyUnicode_GET_DATA_SIZE(lpxString);
					cBuff = (unsigned char *) PyUnicode_AsUnicode(lpxString);
					if (cType == 'C') f4assignN(lppField, cBuff, (unsigned long) nDataSize);
					else
						{
						lnRet = f4memoAssignN(lppField, cBuff, (unsigned long) nDataSize);	
						}
					break;
				}
			bDone = TRUE;				
			}
		
		if(!bDone)
			{
			// not a string or byte array, so we just convert to a string
			f4memoAssign(lppField, PyString_AsString(PyObject_Str(lpxString)));
			}
		}
#endif
	if (lnRet < 0)
		{
		// something went wrong
		strcpy(gcErrorMessage, "Writing data to field failed: ");
		if (codeBase.errorCode != 0) strcat(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		else strcat(gcErrorMessage, "UNKNOWN REASON.");
		}
	if (bNeedFree) free(cWorkBuff);
	return(lnRet);
}

/* ********************************************************************************** */
/* Utility function that converts the value in a FIELD4 to a Python value for return  */
/* to one of the data output methods.                                                 */
/* NOTE: If lpbConvertTypes is False, then all values are returned as strings.        */
static PyObject *cbxGetPythonValue(FIELD4 *lpField, long lpbConvertTypes, long lpbStripBlanks, unsigned char *lpcCoding)
{
	PyObject *lxValue = NULL;
	PyObject *tempDict;
	PyObject *xTest = NULL;
	PyObject *modDecimal;
	int lcType;
	char lcBuff[500];
	char lcNum[20];
	const char cAscCodes[] = "W8";
	const char cBinCodes[] = "DUCPA";
	double lnDoubleValue;
	char *lcPtr;
	long lnYear;
	long lnMonth;
	long lnDay;
	long lnHour;
	long lnMinute;
	long lnSecond;
	long lnTest;
	unsigned char lcCodeBin = 'X';
	unsigned char lcCodeAsc = 'X';

	lcPtr = strchr(cAscCodes, *lpcCoding);
	if (lcPtr != NULL) lcCodeAsc = *lcPtr;
	else {
		lcPtr = strchr(cAscCodes, *(lpcCoding + 1));
		if (lcPtr != NULL) lcCodeAsc = *lcPtr;
	}
	lcPtr = strchr(cBinCodes, *lpcCoding);
	if (lcPtr != NULL) lcCodeBin = *lcPtr;
	else {
		lcPtr = strchr(cBinCodes, *(lpcCoding + 1));
		if (lcPtr != NULL) lcCodeBin = *lcPtr;
	}

	lcType = f4type(lpField);
	if (f4null(lpField) == 1)
		{
		if (lpbConvertTypes)
			{
			return Py_BuildValue(""); // Nothing to do.  It is null regardless of the type.
			}
		else
			{
			return Py_BuildValue("s", ".NULL."); // Send back as the string representation of NULL in VFP.	
			}
		}

	switch(lcType)
		{
		case 'C':
			lxValue = cbxMakeCharVar(lpField, lcCodeAsc, lpbStripBlanks, 'C');
			break;

		case 'Z':
			lxValue = cbxMakeCharBinVar(lpField, lcCodeAsc, lpbStripBlanks, 'C');
			break;
			
		case 'M':
			lxValue = cbxMakeCharVar(lpField, lcCodeAsc, lpbStripBlanks, 'M');
			break;
			
		case 'X':
			lxValue = cbxMakeCharBinVar(lpField, lcCodeAsc, lpbStripBlanks, 'M');
			break;
			
		case 'L':
			if (!lpbConvertTypes)
				{
				lxValue = Py_BuildValue("s", f4memoStr(lpField));	
				}
			else
				{
				lxValue = Py_BuildValue("N", PyBool_FromLong(f4true(lpField)));	
				}
			break;
				
		case 'I':
			if (!lpbConvertTypes)
				{
				sprintf(lcBuff, "%ld", f4long(lpField));
				lxValue = Py_BuildValue("s", lcBuff);
				}
			else
				{
				lxValue = Py_BuildValue("l", f4long(lpField));	
				}
			break;
			
		case 'B':
		case 'N':
			if (!lpbConvertTypes)
				{
				strcpy(lcBuff, f4str(lpField));
				lxValue = Py_BuildValue("s", lcBuff);	
				}
			else
				{
				lxValue = Py_BuildValue("d", f4double(lpField));	
				}
			break;
			
		case 'T':
			lcPtr = f4dateTime(lpField);
			if (lcPtr != NULL)
				{
				if (!lpbConvertTypes)
					{
					lxValue = Py_BuildValue("s", MPSSStrTran(lcPtr, ":", "", 0));
					}
				else
					{
					strncpy(lcNum, lcPtr, 4);
					lcNum[4] = (char) 0;
					lnYear = atol(lcNum);
					if (lnYear < 200)
						{
						lxValue = Py_BuildValue(""); // Empty or invalid datetime.	
						}
					else
						{
						strncpy(lcNum, lcPtr + 4, 2);
						lcNum[2] = (char) 0;
						lnMonth = atol(lcNum);
						
						strncpy(lcNum, lcPtr + 6, 2);
						lcNum[2] = (char) 0;
						lnDay = atol(lcNum);
						
						strncpy(lcNum, lcPtr + 8, 2);
						lcNum[2] = (char) 0;
						lnHour = atol(lcNum);
						
						strncpy(lcNum, lcPtr + 11, 2);
						lcNum[2] = (char) 0;
						lnMinute = atol(lcNum);
						
						strncpy(lcNum, lcPtr + 14, 2);
						lcNum[2] = (char) 0;
						lnSecond = atol(lcNum);
						
						lxValue = PyDateTime_FromDateAndTime(lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond, 0);
						}
					}
				}
			else
				{
				lxValue = Py_BuildValue(""); // Return None.	
				}
			break;
			
		case 'D':
			lcPtr = (char*) f4str(lpField);
			if (lcPtr != NULL)
				{
				if (!lpbConvertTypes)
					{
					lxValue = Py_BuildValue("s", lcPtr); // Just send back the string yyyymmdd.	
					}
				else
					{
					strncpy(lcNum, lcPtr, 4);
					lcNum[4] = (char) 0;
					lnYear = atol(lcNum);
					
					if (lnYear < 200)
						{
						lxValue = Py_BuildValue(""); // Return empty date.	VFP Dates aren't reliable prior to about the year 200 AD.
						}
					else
						{
						strncpy(lcNum, lcPtr + 4, 2);
						lcNum[2] = (char) 0;
						lnMonth = atol(lcNum);
						
						strncpy(lcNum, lcPtr + 6, 2);
						lcNum[2] = (char) 0;
						lnDay = atol(lcNum);
						
						lxValue = PyDate_FromDate(lnYear, lnMonth, lnDay); 
						}	
					}	
				}
			else
				{
				lxValue = Py_BuildValue(""); // None or Empty.	
				}
			break;
			
		case 'Y':
			if (!lpbConvertTypes)
				{
				lcPtr = (char*) f4currency(lpField, 4);
				if (lcPtr != NULL)
					{
					strcpy(lcBuff, lcPtr);
					lxValue = Py_BuildValue("s", lcBuff); // Just the string representation.
					}	
				else
					{
					lxValue = Py_BuildValue("s", "");	
					}
				}
			else
				{
				lcPtr = (char*) f4currency(lpField, 4);
				modDecimal = PyImport_ImportModule("decimal");
				if (modDecimal != NULL)
					{
					xTest = PyObject_CallMethod(modDecimal, "Decimal", "(s)", lcPtr);
					}
				if (xTest != NULL)
					{
					lxValue = xTest;
					}
				else
					{
					lxValue = Py_BuildValue("s", ""); // Failed.  Nothing to be done about it.	
					}
				
				}
			break;
			
		default:
			lxValue = Py_BuildValue("s", f4memoStr(lpField)); 
		}

	return lxValue;	
}


/* ************************************************************************************ */
/* A bit like VFP SCATTER in that this function returns all field values at once.  But  */
/* it has only one type of output: a Python Dict() with one element for each field keyed*/
/* by the uppercased field name.    */
/* CHANGE AS OF October 17, 2013:                                                       */
/* - New parameter lpcFieldList.  If non-empty, must be a comma-delimited list of       */
/*   field names (upper case), for which values will be returned.                       */
/* - If lpcFieldList is passed, the return buffer will contain just the fields asked    */
/*   for.  The lpcFieldList buffer itself will have a string of field type chars put    */
/*   back into it so that the caller can make the proper conversions.                   */
/* Pass TRUE in lpbByList to force return of a list of values, not a dict.              */

static PyObject *cbwSCATTER(PyObject *self, PyObject *args)
{
	long lpbConvertTypes;
	long lpbStripBlanks;
	char *lpcAlias;
	char *lpcFieldList;
	long lpbByList = FALSE;
	long lnReturn = TRUE;
	long lnFldCnt;
	long lnLen;
	register long jj;
	long lnRecord;
	double lnDoubleValue;
	long lnLongValue, lnTest;
	char *lcPtr;
	FIELD4 *lpField;
	DATA4 *lpTable = NULL;
	char lcName[25];
	char lcBuff[250];
	char cfList[4000];
	char *lpTest;
	long lbIsList = FALSE;
	PyObject *recordDict = NULL;
	PyObject *lxValue;
	PyObject *lxKey;	
	PyObject *lxTemp;
	char lcAlias[MAXALIASNAMESIZE + 1];
	char lcCodes[6];
	PyObject *lxAlias;
	PyObject *lxFieldList;
	PyObject *lxCodes;
	const unsigned char *cTest;
	
	if (!PyArg_ParseTuple(args, "OllOlO", &lxAlias, &lpbConvertTypes, &lpbStripBlanks, &lxFieldList, &lpbByList, &lxCodes))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);
		
	cTest = testStringTypes(lxFieldList);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(cfList, Unicode2Char(lxFieldList), 3999);	
	cfList[3999] = (char) 0;

	cTest = testStringTypes(lxCodes);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcCodes, Unicode2Char(lxCodes), 5);
	lcCodes[5] = (char) 0;
	Conv1252ToASCII(lcCodes, TRUE);

	if (strlen(cfList) > 0)
		{
		lbIsList = TRUE;
		}

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (strlen(lcCodes) == 0) strcpy(lcCodes, "XX");
	else if (strlen(lcCodes) == 1) strcat(lcCodes, "X");
	StrToUpper(lcCodes);
	
	if (strlen(lcAlias) == 0)
		{
		if (gpCurrentTable)
			{
			lpTable = gpCurrentTable;
			}
		else
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "No Table is Selected, Alias Not Found");	
			gnLastErrorNumber = -9994;
			}
		}
	else
		{
		lpTable = code4data(&codeBase, lcAlias);
		if (lpTable == NULL)
			{
			gnLastErrorNumber = codeBase.errorCode;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			lnReturn = FALSE;				
			}
		}	
	
	if (lpTable)
		{
		lnRecord = d4recNo(lpTable);
		if ((lnRecord > 0) && (lnRecord <= d4recCount(lpTable)))
			{
			lnFldCnt = 0;
			//if ((lpcFieldList == NULL) || (strlen(lpcFieldList) == 0))
			if (!lbIsList)
				{
				lnFldCnt = d4numFields(lpTable);
				}
			else
				{
				//lbIsList = TRUE;
				StrToLower(cfList);
				lpTest = strtok(cfList, ",");
				jj = -1;
				while (lpTest)
					{
					jj += 1;
					strncpy(gaFieldList[jj], lpTest, 20); // just to be safe.  buffer size is 30.
					*(gaFieldList[jj] + 20) = (char) 0;
					lpTest = strtok(NULL, ","); 	
					}
				lnFldCnt = jj + 1;
				}
			if (lnFldCnt == 0)
				{
				if (!lbIsList)
					{
					strcpy(gcErrorMessage, "No fields defined for current table.");
					gnLastErrorNumber = -9985;
					}
				else
					{
					gnLastErrorNumber = codeBase.errorCode;
					strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
					}	
				lnReturn = FALSE;
				}	
			else
				{
				if (!lpbByList)	recordDict = PyDict_New();
				else recordDict = PyList_New(0);

				if (!lbIsList)
					{
					// We simply iterate through all the fields in the table
					for (jj = 1; jj <= lnFldCnt; jj++)
						{
						lpField = d4fieldJ(lpTable, (short) jj);
						lxTemp = cbxGetPythonValue(lpField, lpbConvertTypes, lpbStripBlanks, lcCodes);
						if (lxTemp == NULL) return(NULL); // error already set.

						if (!lpbByList)
							{
							lxKey = cbNameToPy(f4name(lpField));
							PyDict_SetItem(recordDict, lxKey, lxTemp);
							Py_DECREF(lxKey);
							lxKey = NULL;							
							} 
						else PyList_Append(recordDict, lxTemp);

						Py_DECREF(lxTemp);
						lxTemp = NULL;
						}
					}
				else
					{
					// Iterate through the field list.
					for (jj = 0; jj < lnFldCnt; jj++)
						{
						lpField = d4field(lpTable, gaFieldList[jj]);
						if (lpField == NULL)
							{
							gnLastErrorNumber = codeBase.errorCode;
							strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
							strcat(gcErrorMessage, " Field: ");
							strcat(gcErrorMessage, gaFieldList[jj]);
							sprintf(lcBuff, "%ld", jj + 1);
							strcat(gcErrorMessage, " Cnt: ");
							strcat(gcErrorMessage, lcBuff);
							lnReturn = FALSE;
							break;								
							}							
						else
							{
							lxTemp = cbxGetPythonValue(lpField, lpbConvertTypes, lpbStripBlanks, lcCodes);
							if (lxTemp == NULL) return(NULL); // error already set.
							
							if (!lpbByList)
								{
								lxKey = cbNameToPy(f4name(lpField));	
								PyDict_SetItem(recordDict, lxKey, lxTemp);
								Py_DECREF(lxKey);
								lxKey = NULL;								
								}
							else PyList_Append(recordDict, lxTemp);
								
							Py_DECREF(lxTemp);
							lxTemp = NULL;
							}
						}		
					}
				}
			}
		else
			{
			gnLastErrorNumber = -9984;
			strcpy(gcErrorMessage, "No current record.");
			lnReturn = FALSE;
			}
		}
	//if (cfList != NULL) free(cfList);
	if (!lnReturn)
		{
		if (recordDict != (PyObject*) NULL) Py_CLEAR(recordDict); // don't need it.
		return Py_BuildValue(""); // Return None
		}
	else return(recordDict);	
}

/* ********************************************************************************** */
/* OLD VERSION REPLACED BY NEW cbwSCATTER() with Python Parameters and Returns.       */
/* ********************************************************************************** */
/* A bit like SCATTER in that this function returns all field values at once.  But,   */
/* big difference, it returns all the values in a text string to which a pointer is   */
/* returned by the function (gcStrReturn buffer).  This returns the values for the    */
/* current record of the currently selected table.  On error the buffer returned has  */
/* zero length.  Get the error message in that case.                                  */
/* Pass a string value by reference containing the delimiter.  If an empty string is  */
/* passed, the delimiter used will be ~~ (two tildes).  Maximum length of delimiter   */
/* is 32 characters.                                                                  */
/* This function works ONLY on the currently selected table.                          */
/* CHANGE AS OF October 17, 2013:                                                     */
/* - New parameter lpcFieldList.  If non-empty, must be a comma-delimited list of     */
/*   field names (upper case), for which values will be returned.                     */
/* - If lpcFieldList is passed, the return buffer will contain just the fields asked  */
/*   for.  The lpcFieldList buffer itself will have a string of field type chars put  */
/*   back into it so that the caller can make the proper conversions.                 */

char* cbwOLDSCATTER(char* lpcDelimiter, long lpbStripBlanks, char* lpcFieldList)
{
	long lnResult;
	long lnFldCnt;
	long jj;
	int  lcType;
	long lnRecord;
	double lnDoubleValue;
	long lnLongValue;
	char lcDelim[33];
	char *lcPtr;
	FIELD4 *lpField;
	char lcNumBuff[50];
	char lcTypeBuff[601]; /* holds a list of field types for the sub-list of fields requested. */
	char *lpTest;
	long lbIsList = FALSE;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	memset(lcNumBuff, 0, (size_t) 50 * sizeof(char));
	
	if (gpCurrentTable != NULL)
		{
		lnRecord = d4recNo(gpCurrentTable);
		if ((lnRecord > 0) && (lnRecord <= d4recCount(gpCurrentTable)))
			{
			if ((lpcFieldList == NULL) || (strlen(lpcFieldList) == 0))
				{
				lnResult = d4numFields(gpCurrentTable);
				lbIsList = FALSE;
				}
			else
				{
				lbIsList = TRUE;
				}
			if (!lbIsList && (lnResult < 1))
				{
				if (lnResult == 0) /* No fields defined -- ERROR, but not reported as such */
					{
					strcpy(gcErrorMessage, "No fields defined for current table.");
					gnLastErrorNumber = -9985;
					gcStrReturn[0] = (char) 0;	
					}
				else
					{
					gnLastErrorNumber = codeBase.errorCode;
					strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
					gcStrReturn[0] = (char) 0;					
					}	
				}
			else
				{
				if (!lbIsList) lnFldCnt = lnResult ; /* number of fields */
				else lnFldCnt = 600; // impossibly large, but reflects our buffer sizes.
				memset(gcStrReturn, 0, (size_t) 500 * sizeof(char));
				memset(lcDelim, 0, (size_t) 32 * sizeof(char));
				if (lpcDelimiter == NULL)
					{
					lcDelim[0] = '~';
					lcDelim[1] = '~';
					}
				else
					{
					if (strlen(lpcDelimiter) > 0)
						{
						strncpy(lcDelim, lpcDelimiter, 32);
						lcDelim[32] = (char) 0;	
						}
					else
						{
						lcDelim[0] = '~';
						lcDelim[1] = '~';	
						}
					}
				//for (jj = 0;jj < lnFldCnt; jj++)
				jj = -1;
				while (TRUE)
					{
					jj += 1;
					if (jj >= lnFldCnt) break; // We're done.  Note that lnFldCnt is set very large for the lbIsList case.
					if (!lbIsList)
						{
						lpField = d4fieldJ(gpCurrentTable, (short) (jj + 1));
						}
					else
						{
						lpTest = strtokX(jj == 0 ? lpcFieldList : NULL, ",");
						if (lpTest == NULL) break; // We're done with the field list.
						lpField = cbxGetFieldFromName(lpTest);
						if (lpField == NULL)
							{
							gnLastErrorNumber = codeBase.errorCode;
							strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
							strcat(gcErrorMessage, " Field: ");
							strcat(gcErrorMessage, lpTest);
							sprintf(lcNumBuff, "%ld", jj);
							strcat(gcErrorMessage, " Cnt: ");
							strcat(gcErrorMessage, lcNumBuff);
							gcStrReturn[0] = (char) 0; /* We're done... ERROR */
							break;								
							}
						}
					if (jj > 0) strcat(gcStrReturn, lcDelim);
					lcType = f4type(lpField);
					if (lbIsList)
						{
						lcTypeBuff[jj] = lcType;
						lcTypeBuff[jj + 1] = (char) NULL;	
						}
					switch(lcType)
						{
						case 'M':
							lcPtr = (char*) f4memoStr(lpField);
							break;
							
						case 'T':
							lcPtr = (char*) f4dateTime(lpField);
							if (lcPtr != NULL)
								{
								lcPtr = MPSSStrTran(lcPtr, ":", "", 0);
								}
							break;
							
						case 'Y':
							lcPtr = (char*) f4currency(lpField, 4);
							break;
							
						case 'B':
							lnDoubleValue = f4double(lpField);
							sprintf(lcNumBuff, "%lf", lnDoubleValue);
							lcPtr = lcNumBuff;
							break;
							
						case 'I':
							lnLongValue = f4long(lpField);
							sprintf(lcNumBuff, "%ld", lnLongValue);
							lcPtr = lcNumBuff;
							break;
							
							
						default:
							lcPtr = (char*) f4str(lpField);	
						}

					if (lcPtr == NULL)
						{
						gnLastErrorNumber = codeBase.errorCode;
						strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
						gcStrReturn[0] = (char) 0; /* We're done... ERROR */
						break;	
						}
					else
						{
						if (lpbStripBlanks == 1)
							{
							c4trimN(lcPtr, strlen(lcPtr) + 1);	
							}
						strcat(gcStrReturn, lcPtr);
						}
					}
				}
			}
		else
			{
			gnLastErrorNumber = -9984;
			strcpy(gcErrorMessage, "No current record.");
			gcStrReturn[0] = (char) 0;	
			}
		}
	else
		{
		gcStrReturn[0] = (char) 0;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	if ((gnLastErrorNumber == 0) && lbIsList)
		{
		strcpy(lpcFieldList, lcTypeBuff); // By definition, the list of types will be smaller than the list of field names.	
		}
	return(gcStrReturn);	
}

/* ************************************************************************************* */
/* Completely generic function that returns a text string representation of a specified  */
/* field.  Returns the type as a character value in the second parm passed by reference. */
/* Like the SCATTER functions, you can pass a table qualified name "MYTABLE.MYFIELD" to  */
/* refer to a table other than the currently selected one.                               */
/* If lpbConvertTypes is TRUE, then returns a single Python value of the correct type,   */
/* otherwise it returns a Python tuple consisting of the string value and the type       */
/* string.  Returns None or (None, None) on error.                                       */
//DOEXPORT char* cbwCURVAL(char* lpcFieldName, char* lpcFieldType)
static PyObject *cbwCURVAL(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	PyObject *lxValue = NULL;
	PyObject *lxReturn = NULL;
	PyObject *lxItem1 = NULL;
	PyObject *lxItem2 = NULL;
	PyObject *lxFieldName;
	PyObject *lxCodes;
	FIELD4 *lpField;
	long lpbConvertTypes;
	char   lcType[2];
	char   lcFieldName[MAXALIASNAMESIZE + 30];
	char   lcCodes[6];
	const unsigned char *cTest;	

	if (!PyArg_ParseTuple(args, "OlO", &lxFieldName, &lpbConvertTypes, &lxCodes))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 29);
	lcFieldName[MAXALIASNAMESIZE + 29] = (char) 0;

	cTest = testStringTypes(lxCodes);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcCodes, Unicode2Char(lxCodes), 5);
	lcCodes[5] = (char) 0;
	Conv1252ToASCII(lcCodes, TRUE);
	StrToUpper(lcCodes);

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lpField = cbxGetFieldFromName(lcFieldName);
	if (lpField != NULL)
		{
		lxValue = cbxGetPythonValue(lpField, lpbConvertTypes, lpbConvertTypes, lcCodes); // If we convert, we also strip blanks.
		if (lxValue == NULL)
			{
			// can return NULL if coding fails...
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "Unidentified field value");	
			gnLastErrorNumber = -95839;
			return(NULL); // triggers a Python error, the type of which has already been set up.
			}
		}
	else lnReturn = FALSE;

	if (!lnReturn)
		{
		if (lpbConvertTypes)
			{
			return(Py_BuildValue(""));	
			}
		else
			{
			lxValue = PyTuple_New((Py_ssize_t) 2);
			lxItem1 = Py_BuildValue("");
			lxItem2 = Py_BuildValue("");
			PyTuple_SET_ITEM(lxValue, (Py_ssize_t) 0, lxItem1);
			PyTuple_SET_ITEM(lxValue, (Py_ssize_t) 1, lxItem2);
			//Py_DECREF(lxItem1);  NEVER... PyTuple steals the reference.
			//Py_DECREF(lxItem2);
			return lxValue;
			}	
		}
	else
		{
		if (lpbConvertTypes)
			{
			lxReturn = lxValue; // Just the Python value	
			}
		else
			{
			lcType[0] = f4type(lpField);
			lcType[1] = (char) 0;
			lxReturn = PyTuple_New((Py_ssize_t) 2);
			PyTuple_SetItem(lxReturn, (Py_ssize_t) 0, lxValue);
			
			lxItem1 = Py_BuildValue("s", lcType);
			PyTuple_SetItem(lxReturn, (Py_ssize_t) 1, lxItem1);
			//Py_CLEAR(lxItem1); NOT NEEDED cause PyTuple_SetItem grabs the reference for itself.	
			//Py_CLEAR(lxValue);
			}	
		return (lxReturn);
		}
}

void cbxClearFieldSpecs(void)
{
	register long jj;
	for (jj = 0;jj < MAXFIELDCOUNT; jj++)
		{
		gaFieldSpecs[jj].type = 0;
		gaFieldSpecs[jj].len = 0;
		gaFieldSpecs[jj].dec = 0;
		gaFieldSpecs[jj].nulls = 0;
		if (gaFieldSpecs[jj].name != (char *) NULL)
			{
			free(gaFieldSpecs[jj].name);
			gaFieldSpecs[jj].name = (char *) NULL;
			}
		}	
}

/* *********************************************************************************** */
/* Function that obtains the value of a Logical field, as per the above functions.     */
/* Returns 1 for a value of TRUE in the field.  Returns 0 for a value of FALSE, and    */
/* returns -1 on any kind of error.  Returns ERROR if the field is not LOGICAL.        */
static PyObject *cbwSCATTERFIELDLOGICAL(PyObject *self, PyObject *args)
{
	char *lcPtr;
	FIELD4 *lpField;
	long lnReturn = TRUE;
	int lcType;
	long lbValue;
	PyObject *lxFieldName;
	const unsigned char *cTest;
	char lcFieldName[MAXALIASNAMESIZE + 40];	
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (!PyArg_ParseTuple(args, "O", &lxFieldName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 38);
	lcFieldName[MAXALIASNAMESIZE + 38] = (char) 0;

	lpField = cbxGetFieldFromName(lcFieldName);
	if (lpField != NULL) /* Test for Date or DateTime type */
		{
		lcType = f4type(lpField);
		if (lcType == (int) 'L')
			{
			lbValue = f4true(lpField);
			if (lbValue < 0)
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				lnReturn = -1; /* We're done... ERROR */
				}
			else
				{
				lnReturn = lbValue;
				}
			}			
		else
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, "Not a Logical Type");
			gnLastErrorNumber = -9982;
			}	
		}
	else
		{
		lnReturn = -1;	
		}

	if (lnReturn == -1) return(Py_BuildValue(""));
	else return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));		
}


/* *********************************************************************************** */
/* Function that obtains the value of a date field and returns the date value as a     */
/* Python datetime.date() object.  Returns None on an empty date or error.             */
/* Returns the date portion only of a datetime field.                                  */
static PyObject *cbwSCATTERFIELDDATE(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	PyObject *lxValue = NULL;
	long lnYear = 0;
	long lnMonth = 0;
	long lnDay = 0;
	long lnHour = 0;
	long lnMinute = 0;
	long lnSecond = 0;
	char *lcPtr;
	int  lcType = 0;
	FIELD4 *lpField;
	char lcNum[30];
	PyObject *lxFieldName;
	const unsigned char *cTest;
	char lcFieldName[MAXALIASNAMESIZE + 40];

	if (!PyArg_ParseTuple(args, "O", &lxFieldName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 38);
	lcFieldName[MAXALIASNAMESIZE + 38] = (char) 0;
	lpField = cbxGetFieldFromName(lcFieldName);

	if (lpField == NULL)
		{
		lnReturn = FALSE;	// Error message already set.
		}
	else
		{
		lcType = f4type(lpField);
		if (lcType == (int) 'T')
			{
			lcPtr = f4dateTime(lpField);
			if (lcPtr != NULL)
				{
				strncpy(lcNum, lcPtr, 4);
				lcNum[4] = (char) 0;
				lnYear = atol(lcNum);
				if (lnYear < 200)
					{
					lxValue = Py_BuildValue(""); // Empty or invalid datetime.	
					}
				else
					{
					strncpy(lcNum, lcPtr + 4, 2);
					lcNum[2] = (char) 0;
					lnMonth = atol(lcNum);
					
					strncpy(lcNum, lcPtr + 6, 2);
					lcNum[2] = (char) 0;
					lnDay = atol(lcNum);
					
					lxValue = PyDate_FromDate(lnYear, lnMonth, lnDay);
					}
				}
			else
				{
				lxValue = Py_BuildValue("");	
				}
			}
		else
			{
			if (lcType == (int) 'D')
				{
				lcPtr = f4str(lpField);
				if (lcPtr != NULL)
					{
					if (strlen(lcPtr) == 0)
						{
						lxValue = Py_BuildValue(""); // Return None.	
						}
					else
						{
						strncpy(lcNum, lcPtr, 4);
						lcNum[4] = (char) 0;
						lnYear = atol(lcNum);
						
						if (lnYear < 200)
							{
							lxValue = Py_BuildValue(""); // Return empty date.	VFP Dates aren't reliable prior to about the year 200 AD.
							}
						else
							{
							strncpy(lcNum, lcPtr + 4, 2);
							lcNum[2] = (char) 0;
							lnMonth = atol(lcNum);
							
							strncpy(lcNum, lcPtr + 6, 2);
							lcNum[2] = (char) 0;
							lnDay = atol(lcNum);
							
							if ((lnMonth >= 1) && (lnMonth <= 12) && (lnDay >= 1) && (lnDay <= 31))
								{							
								lxValue = PyDate_FromDate(lnYear, lnMonth, lnDay);
								}
							else
								{
								lxValue = Py_BuildValue("");
								}
							}
						}
					}
				else
					{
					lxValue = Py_BuildValue("");	
					}
				}					
			else
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, "Not a Date/Time Type");
				gnLastErrorNumber = -9983;
				}
			}
		}
	if (!lnReturn)
		{
		if (lxValue != (PyObject*) NULL) Py_CLEAR(lxValue);
		return(Py_BuildValue(""));
		}
	else return(lxValue);
}

/* *********************************************************************************** */
/* Function that obtains the value of a datetime or date field and returns the Python  */
/* datetime value for the field.  Returns None if the field is empty.  Also returns    */
/* on error, so the user will need to check the error flag status to see if an error   */
/* happened.                                                                           */
static PyObject *cbwSCATTERFIELDDATETIME(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	PyObject *lxValue = NULL;
	long lnYear = 0;
	long lnMonth = 0;
	long lnDay = 0;
	long lnHour = 0;
	long lnMinute = 0;
	long lnSecond = 0;
	char *lcPtr;
	int  lcType = 0;
	FIELD4 *lpField;
	char lcNum[20];
	PyObject *lxFieldName;
	const unsigned char *cTest;
	char lcFieldName[MAXALIASNAMESIZE + 40];

	if (!PyArg_ParseTuple(args, "O", &lxFieldName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 38);
	lcFieldName[MAXALIASNAMESIZE + 38] = (char) 0;
	lpField = cbxGetFieldFromName(lcFieldName);

	if (lpField == NULL)
		{
		lnReturn = FALSE;	// Error message already set.
		}
	else
		{
		lcType = f4type(lpField);
		if (lcType == (int) 'T')
			{
			lcPtr = f4dateTime(lpField);
			if (lcPtr != NULL)
				{
				strncpy(lcNum, lcPtr, 4);
				lcNum[4] = (char) 0;
				lnYear = atol(lcNum);
				if (lnYear < 200)
					{
					lxValue = Py_BuildValue(""); // Empty or invalid datetime.	
					}
				else
					{
					strncpy(lcNum, lcPtr + 4, 2);
					lcNum[2] = (char) 0;
					lnMonth = atol(lcNum);
					
					strncpy(lcNum, lcPtr + 6, 2);
					lcNum[2] = (char) 0;
					lnDay = atol(lcNum);
					
					strncpy(lcNum, lcPtr + 8, 2);
					lcNum[2] = (char) 0;
					lnHour = atol(lcNum);
					
					strncpy(lcNum, lcPtr + 11, 2);
					lcNum[2] = (char) 0;
					lnMinute = atol(lcNum);
					
					strncpy(lcNum, lcPtr + 14, 2);
					lcNum[2] = (char) 0;
					lnSecond = atol(lcNum);
					
					lxValue = PyDateTime_FromDateAndTime(lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond, 0);
					}
				}
			else
				{
				lxValue = Py_BuildValue("");	
				}
			}
		else
			{
			if (lcType == (int) 'D')
				{
				lcPtr = f4str(lpField);
				if (lcPtr != NULL)
					{
					strncpy(lcNum, lcPtr, 4);
					lcNum[4] = (char) 0;
					lnYear = atol(lcNum);
					
					if (lnYear < 200)
						{
						lxValue = Py_BuildValue(""); // Return empty date.	VFP Dates aren't reliable prior to about the year 200 AD.
						}
					else
						{
						strncpy(lcNum, lcPtr + 4, 2);
						lcNum[2] = (char) 0;
						lnMonth = atol(lcNum);
						
						strncpy(lcNum, lcPtr + 6, 2);
						lcNum[2] = (char) 0;
						lnDay = atol(lcNum);
						
						lxValue = PyDateTime_FromDateAndTime(lnYear, lnMonth, lnDay, 0, 0, 0, 0); 
						}
					}
				else
					{
					lxValue = Py_BuildValue("");	
					}
				}					
			else
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, "Not a Date/Time Type");
				gnLastErrorNumber = -9983;
				}
			}
		}
	if (!lnReturn) return(Py_BuildValue(""));
	else return(lxValue);
}

/* *********************************************************************************** */
/* Function that obtains the value of an numeric or float field with decimals as a     */
/* double in a double parameter passed by reference.  Returns the number or None on    */
/* error.  If not a numeric field (N, Y, B, F, I, etc.) returns None.               */
static PyObject *cbwSCATTERFIELDDOUBLE(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	double lnValue = 0.0;
	char cType;
	FIELD4 *lpField;
	PyObject *lxReturn = NULL;
	PyObject *lxFieldName;
	const unsigned char *cTest;
	char lcFieldName[MAXALIASNAMESIZE + 40];

	if (!PyArg_ParseTuple(args, "O", &lxFieldName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}
		
	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 38);
	lcFieldName[MAXALIASNAMESIZE + 38] = (char) 0;
	lpField = cbxGetFieldFromName(lcFieldName);		

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (lpField != NULL)
		{
		cType = f4type(lpField);
		if (strchr("YNIFB", (int) cType) == NULL)
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "Not a numeric field type");
			gnLastErrorNumber = -9984;	
			}
		else
			{
			lxReturn = Py_BuildValue("d", f4double(lpField));
			}
		}
	else lnReturn = FALSE;
	
	if (!lnReturn) return(Py_BuildValue(""));
	else return(lxReturn);	
}

/* *********************************************************************************** */
/* Function that obtains the value of an numeric or currency field as a fixed decimal  */
/* Python object of type decimal.Decimal().  Returns the number or None on error.  If  */
/* not a numeric field (N, Y) returns None.                                            */
static PyObject *cbwSCATTERFIELDDECIMAL(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	double lnValue = 0.0;
	char cType;
	FIELD4 *lpField;
	PyObject *xTest = NULL;
	PyObject *modDecimal;
	PyObject *lxValue = NULL;
	int nDecimals = -1;
	register long jj;
	char cBuff[50];
	char *lcPtr;
	PyObject *lxFieldName;
	const unsigned char *cTest;
	char lcFieldName[MAXALIASNAMESIZE + 40];

	if (!PyArg_ParseTuple(args, "O", &lxFieldName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}
		
	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 38);
	lcFieldName[MAXALIASNAMESIZE + 38] = (char) 0;
	lpField = cbxGetFieldFromName(lcFieldName);			
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (lpField != NULL)
		{
		cType = f4type(lpField);
		switch (cType)
		{
		case 'N':
		case 'F':
			nDecimals = f4decimals(lpField);
			strcpy(cBuff, f4str(lpField));
			MPStrTrim('A', cBuff);
			if (cBuff[0] == (char) 0)
				{
				strcpy(cBuff, "0."); // Can happen in some cases where field isn't properly initialized.
				for (jj = 0; jj < nDecimals;jj++)
					{
					strcat(cBuff, "0");	
					}
				}

			if (strchr(cBuff, '.') == NULL) strcat(cBuff, "."); // The converter needs a decimal point.
			break;
			
		case 'Y':
			nDecimals = 4;
			strcpy(cBuff, f4currency(lpField, nDecimals));
			break;
			
		default:	
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "Not a fixed decimal numeric field type");
			gnLastErrorNumber = -9984;			
		}
		if (nDecimals > -1)
			{
			modDecimal = PyImport_ImportModule("decimal");
			if (modDecimal != NULL)
				{
				xTest = PyObject_CallMethod(modDecimal, "Decimal", "(s)", cBuff);
				if (xTest != NULL)
					{
					lxValue = xTest;
					}
				else
					{
					lnReturn = FALSE;
					strcpy(gcErrorMessage, "Invalid number representation ");
					strcat(gcErrorMessage, cBuff);
					gnLastErrorNumber = -98349;	
					}
				}
			else
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, "Decimal converter failed");
				gnLastErrorNumber = -98348;	
				}
			}
		}
	else lnReturn = FALSE;
	
	if (!lnReturn) return(Py_BuildValue(""));
	else return(lxValue);	
}

/* *********************************************************************************** */
/* Function that obtains the value of an integer or numeric field with 0 decimals as a */
/* long in a long parameter passed by reference.  Returns 1 on OK, -1 on ERROR.        */
/* See cbwSCATTERFIELD() for details on passing field name/table name info.            */
/* NOTE: It is up to the caller to make sure that this field contains a value which    */
/* is appropriately converted to a long.  It will convert a charcter field containing  */
/* a string of numbers, a date (converted to a Julian day number), an integer, or a    */
/* Number field.  The number field's decimal portion will be truncated, as will any    */
/* decimal portion of a text string number.                                            */
/* Using this on a field containing alphabetic character information will return a 0.  */
static PyObject *cbwSCATTERFIELDLONG(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnValue = 0;
	FIELD4 *lpField;
	PyObject *lxValue = NULL;
	PyObject *lxFieldName;
	const unsigned char *cTest;
	char lcFieldName[MAXALIASNAMESIZE + 40];

	if (!PyArg_ParseTuple(args, "O", &lxFieldName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}
		
	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 38);
	lcFieldName[MAXALIASNAMESIZE + 38] = (char) 0;	
	lpField = cbxGetFieldFromName(lcFieldName);	

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (lpField != NULL)
		{
		lxValue = Py_BuildValue("l", f4long(lpField));
		}
	else lnReturn = FALSE;
		
	if (!lnReturn) return(Py_BuildValue(""));
	else return(lxValue);	
}

/* ********************************************************************************** */
/* Generic field updater routine.  This routine takes a FIELD4 pointer and a pointer  */
/* to a string buffer containing the value to be updated.  Uses the type information  */
/* it has to determine how to store the information into the field.  Numbers should   */
/* be provided as a text representation of the decimal value.  Dates as YYYYMMDD      */
/* strings. DateTime values as YYYYMMDDHHMMSS values.  Logical may be either 'T' or   */
/* 'Y' or '1' for TRUE, 'F' or 'N' or '0' (zero) for FALSE.                           */
/* Returns TRUE on success, -1 on failure, in which case the error messages will have */
/* been set.                                                                          */
/* NOTE: Pass a NULL pointer as lpcValue to set the value of the field to NULL.       */
long cbxUpdateField(FIELD4* lppField, char* lpcValue)
{
	int lcType;
	long lnLongValue;
	double lnDoubleValue;
	long lnLen;
	long lnReturn = TRUE;
	char lcDTBuff[50];
	char lcTest;
	
	if (lppField == NULL)
		{
		strcpy(gcErrorMessage, "Invalid field identifier");
		gnLastErrorNumber = -9975;
		return -1; // nothing can be done.
		}
	memset(lcDTBuff, 0, (size_t) 50 * sizeof(char));
	if (lpcValue == NULL)
		{
		f4assignNull(lppField);
		return(TRUE); /* No checking done, just return.  If the field is not set to allow nulls, nothing happens. */
		/* Out of here, we're done. */	
		}
	lcType = f4type(lppField);
	switch(lcType)
		{
		case 'C': /* Character */
			f4assign(lppField, lpcValue); /* No type checking, just string stuffed into the field */
			break;
			
		case 'D': /* Date */
			if ((strcmp(lpcValue, "        ") == 0) || (strcmp(lpcValue, "") == 0))
				{
				lnLongValue = 0;
				f4assignLong(lppField, lnLongValue);
				}
			else
				{
				lnLongValue = date4long(lpcValue);
				if (lnLongValue < 0)
					{
					/* An error condition -- bad date format */
					strcpy(gcErrorMessage, "Bad Date Format >");
					strcat(gcErrorMessage, lpcValue);
					strcat(gcErrorMessage, "<");
					gnLastErrorNumber = -8000;
					lnReturn = -1;	
					}
				else
					{
					f4assignLong(lppField, lnLongValue); /* Can be 0 on an empty date which is valid in a VFP table. */	
					}
				}
			break;
		
		case 'T': /* DateTime */
			if (strcmp(lpcValue, "            ") == 0)
				{
				strcpy(lcDTBuff, "          :  :  ");	
				}
			else
				{
				strncpy(lcDTBuff, lpcValue, 10);
				lcDTBuff[10] = (char) 0;
				strcat(lcDTBuff, ":");
				strncat(lcDTBuff, lpcValue + 10, 2);
				strcat(lcDTBuff, ":");
				strncat(lcDTBuff, lpcValue + 12, 2);
				}
			f4assignDateTime(lppField, lcDTBuff);
			break;
			
		case 'N': /* Number or Numeric */
		case 'F': /* Floating Point */
		case 'B': /* 8-Byte Double Number */
			lnDoubleValue = atof(lpcValue);
			f4assignDouble(lppField, lnDoubleValue);
			break;
			
		case 'Y': /* Currency */
			f4assignCurrency(lppField, lpcValue);
			break;
			
		case 'I': /* 4-Byte Integer */
			lnLongValue = atol(lpcValue);
			f4assignLong(lppField, lnLongValue);
			break;
			
		case 'M': /* Text Memo Field */
			f4memoAssign(lppField, lpcValue);
			break;
			
		case 'L': /* Logical */
			lcTest = lpcValue[0];
			if ((lcTest == '1') || (lcTest == 't') || (lcTest == 'T') || (lcTest == 'Y') || (lcTest == 'y')) lcTest = 'T';
			else lcTest = 'F';
			f4assignChar(lppField, (int) lcTest);
			break;
		
		case 'Z': /* Binary Character */
			lnLen = strlen(lpcValue); /* In this function only character values (no NULLs) can be handled, but we can still save it into binary fields */
			if (f4len(lppField) < lnLen) lnLen = f4len(lppField);
			strncpy(f4assignPtr(lppField), lpcValue, lnLen); /* Raw copy of the bytes into the field data buffer */
			break;
			
		case 'G': /* General - contains OLE objects */
			/* No Action at this time */
			break;
			
		case 'X': /* Binary Memo Field */
			lnLen = strlen(lpcValue);
			f4memoAssignN(lppField, lpcValue, lnLen); /* we copy in plain text, but it's going into a field that can handle binary. */
			break;
				
		default: /* Shouldn't get here with VFP data, but ya never know */
			break;
				
		}
	return(lnReturn);		
}

/* ********************************************************************************** */
/* Rough analog to GATHER MEMVAR in that it applies updates to all or a specified     */
/* set of fields in the table.  Takes three parameters:                               */
/* lpcFields is a delimited list of field names to be updated.  If lpcFields is the   */
/*   empty string, then will attempt to update all fields in the table.  If not       */
/*   enough values are provided in either case, an error occurs, and no update is     */
/*   done.                                                                            */
/* lpcValues is a delimited list of field values to update into the listed fields.    */
/*   There must be a one-to-one relationship between the values provided and the      */
/*   fields specified (or implied).                                                   */
/* lpcDelimiter is a string of up to 30 characters that will be used to delimit       */
/*   the elements of the field and value lists.  If this parameter is the empty       */
/*   string, a delimiter of ~~ (two tilde characters) will be assumed.                */
/* Returns TRUE on success, -1 on ERROR.                                              */
/* Operates on the currently selected table.  If none exists, returns ERROR.          */
/* OBSOLETE BUT FUNCTIONAL.  USE cbwGATHERMEMVARDICT() instead.                       */
/* Continues to function exactly as it did in prior versions.                         */
long cbxGATHERMEMVAR(DATA4 *lppTable, char *lpcFields, char* lpcValues, char* lpcDelimiter)
{
	long lnReturn = TRUE;
	long lnResult;
	char lcDelim[32];
	long lnFldCnt = 0;   /* The number of fields in the table */
	long lnFldInCnt = 0; /* The actual number of field names submitted */
	long lnValueCnt = 0; /* The number of field values found */
	long jj;
	register long kk, nn;
	FIELD4* laFields[256];
	char* laValues[256];
	char  lcFieldName[256];
	char  lcBadFieldName[256];
	char* lpPtr;
	char* lpTemp;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	memset(lcDelim, 0, (size_t) 32 * sizeof(char));
	memset(laFields, 0, (size_t) 256 * sizeof(FIELD4*));
	memset(laValues, 0, (size_t) 256 * sizeof(char*));
	lcBadFieldName[0] = (char) 0;
	lcFieldName[0] = (char) 0;

	if (strlen(lpcDelimiter) == 0)
		{
		lcDelim[0] = '~';
		lcDelim[1] = '~';
		}
	else
		{
		strncpy(lcDelim, lpcDelimiter, 30);
		lcDelim[30] = (char) 0;
		}
	
	if (gpCurrentTable != NULL)
		{
		lnFldCnt = d4numFields(lppTable);
		if (lnFldCnt < 1)
			{
			if (lnFldCnt == 0) /* No fields defined -- ERROR, but not reported as such */
				{
				strcpy(gcErrorMessage, "No fields defined for current table.");
				gnLastErrorNumber = -9985;
				lnReturn = -1;	
				}
			else
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				lnReturn = -1;					
				}
			}
		else
			{
			/* First get the array of values */
			lpTemp = lpcValues;
			nn = 0;
			while(TRUE)
				{
				lpPtr = strtokX(lpTemp, lcDelim);
				lpTemp = NULL; /* from here on out */
				if (lpPtr == NULL) break;
				laValues[lnValueCnt] = malloc(strlen(lpPtr) + 1);
				strcpy(laValues[lnValueCnt], lpPtr);
				lnValueCnt += 1;
				}
			if (strlen(lpcFields) != 0)
				{
				lpTemp = lpcFields;
				nn = 0;
				while(TRUE)
					{
					lpPtr = strtokX(lpTemp, lcDelim);
					lpTemp = NULL;
					if (lpPtr == NULL) break;
					laFields[lnFldInCnt] = d4field(lppTable, lpPtr);
					if (laFields[lnFldInCnt] == NULL)
						{
						strcpy(lcBadFieldName, lpPtr);
						break;
						}
					lnFldInCnt += 1;
					}
				lnFldCnt = lnFldInCnt;	
				}
			else
				{
				for (jj = 1; jj <= lnFldCnt; jj++)
					{
					laFields[jj - 1] = d4fieldJ(lppTable, (short) jj); /* all of the fields will be supplied (we hope) and in their proper order */
					}	
				}
			if (lnFldCnt != lnValueCnt) /* Error condition.  We have to have one-to-one matching */
				{
				lnReturn = -1;
				sprintf(gcErrorMessage, "Field and Value count xx mismatch: %ld and %ld", lnFldCnt, lnValueCnt);
				strcat(gcErrorMessage, lcBadFieldName);
				gnLastErrorNumber = -9980;
				}
			else
				{
				for (jj = 0;jj < lnFldCnt;jj++)
					{
					lnResult = cbxUpdateField(laFields[jj], laValues[jj]);
					if (lnResult == -1)
						{
						lnReturn = -1; /* Something went wrong and we can't continue */
						strcat(gcErrorMessage, " Bad Name: ");
						strcat(gcErrorMessage, f4name(laFields[jj]));
						gnLastErrorNumber = -9981;
						
						break;	
						}	
					}
				/* The record buffer is now changed.  The programmer can now call a flush explicitly or do one of the    */
				/* several things in VFP which can cause the data to get saved automatically like moving off the current */
				/* record, closing the table, etc.                                                                       */
	
				}
			}		
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	for (jj = 0;jj < 256;jj++)
		{
		if (laValues[jj] != NULL) free(laValues[jj]);	
		}
	return(lnReturn);
}

/* *************************************************************************************** */
/* Python version of the above internal method.                                            */
/* NOTE: This function is deprecated.  It works only with 8-bit null-terminated strings    */
/* and is intended for such use only.  The cbwGATHERDICT() function takes any valid        */
/* Python type including Unicode strings and stores them accordingly.                      */
static PyObject *cbwGATHERMEMVAR(PyObject *self, PyObject *args)
{
	long lnReturn;	
	char *lpcFields = NULL;
	char *lpcValues = NULL;
	char *lpcDelimiters = NULL;
	const char *cPtr;
	PyObject *lxFields;
	PyObject *lxValues;
	PyObject *lxDelimiters;
	const unsigned char *cTest;
	size_t lnSize;
	
	if (!PyArg_ParseTuple(args, "OOO", &lxFields, &lxValues, &lxDelimiters))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}	

	cTest = testStringTypes(lxFields);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	cPtr = Unicode2Char(lxFields);
	lnSize = strlen(cPtr);
	lpcFields = (char *) calloc(lnSize + 1, sizeof(char));
	strncpy(lpcFields, cPtr, lnSize);
	*(lpcFields + lnSize) = (char) 0;

	cTest = testStringTypes(lxValues);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		if (lpcFields != NULL) free(lpcFields);		
		return NULL;			
		}
	cPtr = Unicode2Char(lxValues); // need to get out of the static buffer that gets overwritten
	lnSize = strlen(cPtr);
	lpcValues = (char *) calloc(lnSize + 1, sizeof(char));
	strncpy(lpcValues, cPtr, lnSize);
	*(lpcValues + lnSize) = (char) 0;	

	cTest = testStringTypes(lxDelimiters);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		if (lpcValues != NULL) free(lpcValues);
		if (lpcFields != NULL) free(lpcFields);		
		return NULL;			
		}
	cPtr = Unicode2Char(lxDelimiters);
	lnSize = strlen(cPtr);
	lpcDelimiters = (char *) calloc(lnSize + 1, sizeof(char));
	strncpy(lpcDelimiters, cPtr, lnSize);	
	*(lpcDelimiters + lnSize) = (char) 0;
	lnReturn = cbxGATHERMEMVAR(gpCurrentTable, lpcFields, lpcValues, lpcDelimiters);

	if (lpcDelimiters != NULL) free(lpcDelimiters);
	if (lpcValues != NULL) free(lpcValues);
	if (lpcFields != NULL) free(lpcFields);
		
	return(Py_BuildValue("l", lnReturn));
}

/* *************************************************************************************** */
/* Utility function that stores the contents of a Python variable object into a specified  */
/* CodeBase field.  Returns 1 on OK, 0 or NULL on failure.  If it returns -1, a Python     */
/* error has occured and an exception should be declared.                                  */
long cbxPyObjectToField(PyObject *lpoValue, FIELD4 *lpxField, char *lpcCoding)
{
	long lnReturn = 1;
	double lnDouble;
	long lnLong;
	long lbStored = FALSE;
	int cType;
	long lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond;
	char cDateBuff[50];
	const char cAscCodes[] = "W8";
	const char cBinCodes[] = "DUCP";
	Py_ssize_t lnLength;
	PyObject* lxTemp;
	PyObject* pBytes;
	char *pCPtr;
	char *lcPtr;
	const char *pConst;
	unsigned char lcCodeBin = 'X';
	unsigned char lcCodeAsc = 'X';

	lcPtr = strchr(cAscCodes, *lpcCoding);
	if (lcPtr != NULL) lcCodeAsc = *lcPtr;
	else {
		lcPtr = strchr(cAscCodes, *(lpcCoding + 1));
		if (lcPtr != NULL) lcCodeAsc = *lcPtr;
	}
	lcPtr = strchr(cBinCodes, *lpcCoding);
	if (lcPtr != NULL) lcCodeBin = *lcPtr;
	else {
		lcPtr = strchr(cBinCodes, *(lpcCoding + 1));
		if (lcPtr != NULL) lcCodeBin = *lcPtr;
	}
	
	codeBase.errorCode = 0;
	cType = f4type(lpxField);
	
	if (lpoValue == Py_None) // Yes, there is only ever ONE (it's a singleton, silly) Py_None, so anything that is None has the same pointer!
		{
		f4assignNull(lpxField);
		if (f4null(lpxField) != 1)
			{
			switch(cType)
				{
				case 'D':
					f4assignLong(lpxField, 0);
					break;
				
				case 'T':
					f4assignDateTime(lpxField, "          :  :  ");	
					break;
						
				default:
					f4assign(lpxField, ""); // Just shove a blank in there, not sure what else to do.
				}
			}
		lbStored = TRUE;
		codeBase.errorCode = 0; // Just in case, since the above may result in an error condition.	
		}
	if (lbStored) return(1);

	switch(cType)
	{
		case 'C':
			lnLong = cbxPyToCharFld(lpxField, lpoValue, 'C', lcCodeAsc);
			break;
			
		case 'M':
			lnLong = cbxPyToCharFld(lpxField, lpoValue, 'M', lcCodeAsc);
			break;

		case 'G':
			lnLong = cbxPyToCharFldBin(lpxField, lpoValue, 'M', 'X'); // force to pure binary.
			break;

		case 'X':
			lnLong = cbxPyToCharFldBin(lpxField, lpoValue, 'M', lcCodeAsc);
			break;
			
		case 'Z':
			lnLong = cbxPyToCharFld(lpxField, lpoValue, 'C', lcCodeAsc);
			break;
			
		case 'N':
		case 'F':
		case 'B':
			if (PyFloat_Check(lpoValue))
				{
				lnDouble = PyFloat_AsDouble(lpoValue);
				f4assignDouble(lpxField, lnDouble);
				}
			else
				{
				if (PyLong_Check(lpoValue))
					{
					lnLong = PyLong_AsLong(lpoValue);
					f4assignDouble(lpxField, (double) lnLong);	
					}
				else
					{
					lxTemp  = PyObject_Str(lpoValue); 	
#ifdef IS_PY3K		
					pBytes = PyUnicode_AsEncodedString(lxTemp, u8"cp1252", u8"ignore");
					lnDouble = atof(PyBytes_AsString(pBytes));
#endif
#ifdef IS_PY2K
					lnDouble = atof(PyString_AsString(lxTemp));
#endif
					f4assignDouble(lpxField, lnDouble);						
					}
				}
			break;
			
		case 'I':
			if (PyLong_Check(lpoValue))
				{
				f4assignLong(lpxField, PyLong_AsLong(lpoValue));
				}
			else
				{
				lxTemp  = PyObject_Str(lpoValue);	
#ifdef IS_PY3K	
				pBytes = PyUnicode_AsEncodedString(lxTemp, u8"cp1252", u8"ignore");
				lnLong = atol(PyBytes_AsString(pBytes));
#endif
#ifdef IS_PY2K
				lnLong = atol(PyString_AsString(lxTemp));
#endif					
				f4assignLong(lpxField, lnLong);	
				}
			break;
			
		case 'L':
			if (PyBool_Check(lpoValue))
				{
				if (PyObject_IsTrue(lpoValue))
					{
					f4assignChar(lpxField, (int) 'T');
					}
				else
					{
					f4assignChar(lpxField, (int) 'F');	
					}
				}
			else
				{
				cDateBuff[0] = (char) 0;
#ifdef IS_PY3K	
				if (PyUnicode_Check(lpoValue))
					{	
					pBytes = PyUnicode_AsEncodedString(lpoValue, u8"cp1252", u8"ignore");
				    strncpy(cDateBuff, PyBytes_AsString(pBytes), 4);
#endif
#ifdef IS_PY2K
				if (PyString_Check(lpoValue))
					{
				    strncpy(cDateBuff, PyString_AsString(lpoValue), 4);
#endif							
					cDateBuff[1] = (char) 0;
					StrToUpper(cDateBuff);
					if ((cDateBuff[0] == 'Y') || (cDateBuff[0] == 'T') || (cDateBuff[0] == 'J') || (cDateBuff[0] == 'S') || 
						(cDateBuff[0] == 'O') || (cDateBuff[0] == '1') || (cDateBuff[0] == 'V') || (cDateBuff[0] == 'W'))
						{
						// Allows for English 'Yes' or 'True', German 'Ja', French 'Oui', Italian and Spanish 'Si' and the number '1'.
						// Also German 'Wehr' and French 'Vrai' and Italian 'Vero'
						f4assignChar(lpxField, (int) 'T');	
						}
					else
						{
						f4assignChar(lpxField, (int) 'F');
						}
					}
				else
					{
					if (PyLong_Check(lpoValue))
						{
						lnLong = PyLong_AsLong(lpoValue);
						if (lnLong == 0)
							{
							f4assignChar(lpxField, (int) 'F');	
							}
						else
							{
							f4assignChar(lpxField, (int) 'T');	// Any non-zero value is considered True.
							}
						}
					else
						{
						f4assignChar(lpxField, (int) 'F'); // nothing else to be done about it.	
						}
					}	
				}
			break;
			
		case 'D':
			if (PyDate_Check(lpoValue) || PyDateTime_Check(lpoValue))
				{
				lnYear = PyDateTime_GET_YEAR(lpoValue);
				lnMonth = PyDateTime_GET_MONTH(lpoValue);
				lnDay = PyDateTime_GET_DAY(lpoValue);
				sprintf(cDateBuff, "%4ld%02ld%02ld", lnYear, lnMonth, lnDay);
				f4assignLong(lpxField, date4long(cDateBuff));
				}
			else
				{

#ifdef IS_PY3K
				if (PyUnicode_Check(lpoValue))
					{
					pBytes = PyUnicode_AsEncodedString(lpoValue, u8"cp1252", u8"ignore");	
				    strncpy(cDateBuff, PyBytes_AsString(pBytes), 49);
#endif
#ifdef IS_PY2K
				if (PyString_Check(lpoValue))
					{
				    strncpy(cDateBuff, PyString_AsString(lpoValue), 49);
#endif				
					cDateBuff[49] = (char) 0;
					MPStrTrim('A', cDateBuff);
					if (strlen(cDateBuff) == 0)
						{
						f4assignLong(lpxField, 0);	
						}
					else
						{
						lnLong = date4long(cDateBuff);
						if (lnLong <= 0)
							{
							f4assignLong(lpxField, 0);
							lnReturn = 0;
							strcpy(gcErrorMessage, "Invalid string value passed to date field");
							gnLastErrorNumber = -13594;
							}
						else f4assignLong(lpxField, lnLong);
						}	
					}
				else
					{
					if (PyLong_Check(lpoValue))
						{
						lnLong = PyLong_AsLong(lpoValue);
						if (lnLong >= 0)
							{
							f4assignLong(lpxField, lnLong); // Hope they know what they are doing!	
							}
						else
							{
							f4assignLong(lpxField, 0);
							strcpy(gcErrorMessage, "Invalid numeric integer value passed to date field");
							gnLastErrorNumber = -13595;
							lnReturn = 0;	
							}
						}
					else
						{
#ifdef IS_PY3K			
						pBytes = PyUnicode_AsEncodedString(PyObject_Str(lpoValue), u8"cp1252", u8"ignore");
				        strncpy(cDateBuff, PyBytes_AsString(pBytes), 49);
#endif
#ifdef IS_PY2K
				        strncpy(cDateBuff, PyString_AsString(PyObject_Str(lpoValue)), 49);
#endif					
						cDateBuff[49] = (char) 0;		
						if (strcmp(cDateBuff, "None") == 0 || strlen(cDateBuff) == 0);
							{
							f4assignLong(lpxField, 0);	
							}
						}	
					}	
				}
			break;
			
		case 'T':
			if (PyDateTime_Check(lpoValue))
				{
				// process a true datetime.
				lnYear = PyDateTime_GET_YEAR(lpoValue);
				lnMonth = PyDateTime_GET_MONTH(lpoValue);
				lnDay = PyDateTime_GET_DAY(lpoValue);
				lnHour = PyDateTime_DATE_GET_HOUR(lpoValue);
				lnMinute = PyDateTime_DATE_GET_MINUTE(lpoValue);
				lnSecond = PyDateTime_DATE_GET_SECOND(lpoValue);
				sprintf(cDateBuff, "%4ld%02ld%02ld%02ld:%02ld:%02ld", lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond);
				f4assignDateTime(lpxField, cDateBuff);	
				}
			else
				{
				if (PyDate_Check(lpoValue))
					{
					// process a date into a datetime field.
					lnYear = PyDateTime_GET_YEAR(lpoValue);
					lnMonth = PyDateTime_GET_MONTH(lpoValue);
					lnDay = PyDateTime_GET_DAY(lpoValue);
					// format: CCYYMMDDhh:mm:ss
					sprintf(cDateBuff, "%4ld%02ld%02ld%02ld:%02ld:%02ld", lnYear, lnMonth, lnDay, 0L, 0L, 0L);
					f4assignDateTime(lpxField, cDateBuff);
					}
				else
					{	

#ifdef IS_PY3K
					if (PyUnicode_Check(lpoValue))
						{
						pBytes = PyUnicode_AsEncodedString(lpoValue, u8"cp1252", u8"ignore");
				        strncpy(cDateBuff, PyBytes_AsString(pBytes), 49);
#endif
#ifdef IS_PY2K
					if (PyString_Check(lpoValue))
						{
				        strncpy(cDateBuff, PyString_AsString(lpoValue), 49);
#endif
						cDateBuff[49] = (char) 0;							
						MPStrTrim('A', cDateBuff);
						if (strlen(cDateBuff) == 0)
							{
							f4assignDateTime(lpxField, "          :  :  ");	
							}
						else
							{
							f4assignDateTime(lpxField, cDateBuff); // Hope they got it right!
							}	
						}
					else
						{
						if (PyLong_Check(lpoValue))
							{
							lnLong = PyLong_AsLong(lpoValue);
							if (lnLong >= 0)
								{
								if (lnLong == 0)
									{
									f4assignDateTime(lpxField, "          :  :  "); // Empty datetime.	
									}
								else
									{
									f4assignDateTime(lpxField, "          :  :  "); // Empty datetime.
									lnReturn = 0;
									strcpy(gcErrorMessage, "Cannot assign long integer value to datetime field");
									gnLastErrorNumber = -13597;	
									}
								}
							else
								{
								f4assignDateTime(lpxField, "          :  :  ");	
								strcpy(gcErrorMessage, "Invalid numeric integer value passed to date field");
								gnLastErrorNumber = -13595;
								lnReturn = 0;	
								}
							}
						else
							{
#ifdef IS_PY3K				
							pBytes = PyUnicode_AsEncodedString(PyObject_Str(lpoValue), u8"cp1252", u8"ignore");
				            strncpy(cDateBuff, PyBytes_AsString(pBytes), 49);
#endif
#ifdef IS_PY2K
				            strncpy(cDateBuff, PyString_AsString(PyObject_Str(lpoValue)), 49);
#endif
							cDateBuff[49] = (char) 0;
							if (strcmp(cDateBuff, "None") == 0 || strlen(cDateBuff) == 0);
								{
								f4assignDateTime(lpxField, "          :  :  ");	// None >> empty datetime.
								}
							}	
						}	
					}
				}		
			break;
			
		case 'Y':

#ifdef IS_PY3K
			if (PyUnicode_Check(lpoValue))
				{	
				pBytes = PyUnicode_AsEncodedString(lpoValue, u8"cp1252", u8"ignore");
				strncpy(cDateBuff, PyBytes_AsString(pBytes), 49);
#endif
#ifdef IS_PY2K
			if (PyString_Check(lpoValue))
				{
				strncpy(cDateBuff, PyString_AsString(lpoValue), 49);
#endif
				cDateBuff[49] = (char) 0;
				f4assignCurrency(lpxField, cDateBuff);	
				}
			else
				{
				// Just convert to a string and hope for the best.
				lxTemp  = PyObject_Str(lpoValue);	
#ifdef IS_PY3K	
				pBytes = PyUnicode_AsEncodedString(lxTemp, u8"cp1252", u8"ignore");
				f4assignCurrency(lpxField, PyBytes_AsString(pBytes));
#endif
#ifdef IS_PY2K
				f4assignCurrency(lpxField, PyString_AsString(lxTemp));
#endif					
				}
			break;			
				
		case 'W': // Fields intended for pure Unicode content.
		case 'Q':
			// not supported
			break;
			
		default:
			break;
	}
	
	return(lnReturn);
}

/* *************************************************************************************** */
/* New approach to the GATHER function, which takes as its first parameter an optional     */
/* alias and then a Python Dict() with one item for each field it is desired to fill in.   */
/* The lbAppendRecord parameter specifies whether this record data is updated into the     */
/* current record of the specified table (value passed is False) or if a new record is     */
/* first appended to the table and then updated with this data (value passed is True).     */
/* Returns True on OK, False on any error.  Some extreme errors may trigger a Python       */
/* exception.                                                                              */
static PyObject *cbwGATHERDICT(PyObject *self, PyObject *args)
{
	PyObject *dataDict;
	PyObject *key, *value;
	PyObject *testKey;
	char *testString;
	char *cFieldName;
	char cName[51];
	Py_ssize_t pos = 0;
	Py_ssize_t nStrLen = 0;
	register long jj;
	DATA4 *lpTable;
	FIELD4 *lpField;
	long lnReturn = TRUE;
	long lnTest;
	long lbAppendRecord = FALSE;
	long lnResult = 0;
	long lnDictCount = 0;
	long lnSavedCount = 0;
	PyObject *pBytes;
	PyObject *lxCodes;
	PyObject *lxAlias;
	char lcCodes[6];
	char lcAlias[MAXALIASNAMESIZE + 1];
	const unsigned char *cTest;	

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	// Note that the Python docs seem to indicate that the object format identifier 
	// should be a lower-case 'o', but in fact it is an upper case 'O'.
	if (!PyArg_ParseTuple(args, "OOlO", &lxAlias, &dataDict, &lbAppendRecord, &lxCodes))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);		
		
	cTest = testStringTypes(lxCodes);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcCodes, Unicode2Char(lxCodes), 5);
	lcCodes[5] = (char) 0;
	Conv1252ToASCII(lcCodes, TRUE);

	if (!PyDict_Check(dataDict))
		{
		PyErr_Format(PyExc_ValueError, "A dict() is required as the second parameter!");
		return(NULL);
		}
		
	if (strlen(lcAlias) == 0)
		{
		if (gpCurrentTable)
			{
			lpTable = gpCurrentTable;
			}
		else
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "No Table is Selected, Alias Not Found");	
			gnLastErrorNumber = -9994;
			}
		}
	else
		{
		lpTable = code4data(&codeBase, lcAlias);
		if (lpTable == NULL)
			{
			gnLastErrorNumber = codeBase.errorCode;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			lnReturn = FALSE;				
			}
		}	

	if (lbAppendRecord)
		{
		if (lpTable != NULL)
			{
			lnResult = d4appendStart(lpTable, 0); /* no memo content carryover from the current record */
			if (lnResult != 0)
				{
				if (lnResult > 0)
					{
					strcpy(gcErrorMessage, "Unable to Move Record Pointer");
					gnLastErrorNumber = -9980;
					lnReturn = FALSE; /* One of the several locking and duplicate key errors.  In any event, FAILURE. */	
					}
				else
					{
					gnLastErrorNumber = codeBase.errorCode;
					strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
					strcat(gcErrorMessage, " Code Base Error");
					lnReturn = FALSE; /* We're done... ERROR */					
					}					
				}
			else
				{
				d4blank(lpTable); // clears the record buffer to prepare for appending a new record.	
				}
			}
		}
	if (lnReturn) // we have an identified table and, if required, we've started appending a record successfully.
		{
		while (PyDict_Next(dataDict, &pos, &key, &value))
			{
			cName[50] = (char) 0;
			cTest = testStringTypes(key);
			if (*cTest != (const unsigned char) 0)
				{
				strcpy(gcErrorMessage, cTest);
				gnLastErrorNumber = -10000;
				PyErr_Format(PyExc_ValueError, gcErrorMessage);
				return NULL;			
				}
			strncpy(cName, Unicode2Char(key), 50);
			cName[50] = (char) 0;

			lnDictCount += 1;

			// Theoretically, CodeBase doesn't support field names over 10 characters anyway.

			StrToLower(cName);
			lpField = d4field(lpTable, cName);
			if (lpField == NULL)
				{
				codeBase.errorCode = 0;
				continue; // It's ok if a field isn't in the table, but we better have at least one saved!
				}
			else
				{
				lnTest = cbxPyObjectToField(value, lpField, lcCodes);
				if (lnTest < 0)
					{
					PyErr_Format(PyExc_TypeError, "Couldn't save data to field %s: %s", cName, gcErrorMessage);	
					return NULL;
					}
				lnSavedCount += 1;
				}
			}

		if (lnSavedCount == 0)
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "No fields in Dict matched table structure");
			gnLastErrorNumber = -9985;	
			}
		}
	if (lnReturn && lbAppendRecord)
		{
		// Finish the insertion since all is OK.
		lnResult = d4append(lpTable);
		d4recall(lpTable); // Seems to be required sometimes for the above.
		if (lnResult != 0)
			{
			if (lnResult > 0)
				{
				strcpy(gcErrorMessage, "Unable to Move Record Pointer");
				gnLastErrorNumber = -9980;
				lnReturn = FALSE; /* One of the several locking and duplicate key errors.  In any event, FAILURE. */	
				}
			else
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				strcat(gcErrorMessage, " ERR from CodeBase");
				lnReturn = FALSE; /* We're done... ERROR */					
				}						
			}			
		}
	return (Py_BuildValue("N", PyBool_FromLong(lnReturn)));

}

/* ****************************************************************************** */
/* Internal version of cbwPREPAREFILTER().                                        */
long cbxPREPAREFILTER(char *lpcFilter)
{
	long lnReturn = TRUE;
	EXPR4 *loExpr;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		loExpr = expr4parse(gpCurrentTable, lpcFilter);
		if (loExpr == NULL)
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode;		
			}
		else
			{
			if (gpCurrentFilterExpr != NULL)
				{
				expr4free(gpCurrentFilterExpr);
				gpCurrentFilterExpr = NULL;
				}
			gpCurrentFilterExpr = loExpr;
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return(lnReturn);		
}

/* ****************************************************************************** */
/* This function prepares a filter string against the characteristics of the      */
/* currently selected table.  Returns 1 on OK, 0 (False) on failure of any kind.  */
static PyObject *cbwPREPAREFILTER(PyObject *self, PyObject *args)
{
	long lnReturn;
	char *lpcFilter;
	PyObject *lxResult;
	unsigned char lcFilter[1000];
	PyObject *lxFilter;
	const unsigned char *cTest;	
	
	if (!PyArg_ParseTuple(args, "O", &lxFilter))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}
		
	cTest = testStringTypes(lxFilter);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFilter, Unicode2Char(lxFilter), 998);	
	lcFilter[998] = (char) 0;
		
	lnReturn = cbxPREPAREFILTER(lcFilter);
	return Py_BuildValue("l", lnReturn);
}

/* ***************************************************************************** */
/* Companion for the above which applies the filter to the actual current record */
/* Returns 1 if filter expression evaluates to TRUE, 0 if it evaluates to FALSE, */
/* otherwise triggers a Python ERROR.                                            */
long cbxTESTFILTER(void)
{
	char lcResultBuffer[200];
	char *lpcResult;
	int lnLen;
	long lnReturn = -1;
	int lnType;
	int lnValue;
	
	if (gpCurrentFilterExpr == NULL)
		{
		lnReturn = -2;
		strcpy(gcErrorMessage, "No Filter Expression Active");
		gnLastErrorNumber = -8932;
		//return PyErr_Format(PyExc_TypeError, "Filter Expression has not been defined");
		}
	else
		{
		lnLen = expr4vary(gpCurrentFilterExpr, &lpcResult);
		memset(lcResultBuffer, 0, (size_t) 200 * sizeof(char));
		if ((lnLen > 0) && (lnLen < 200))
			{
			lnType = expr4type(gpCurrentFilterExpr);
			if (lnType == r4log)
				{
				lnValue = expr4true(gpCurrentFilterExpr);
				if (lnValue > 0)
					{
					if ((gnDeletedFlag == TRUE) && (d4deleted(gpCurrentTable) != 0))
						lnReturn = 0;
					else
						lnReturn = 1;
					}
				else
					if (lnValue == 0)
						lnReturn = 0;
					else
						{
						lnReturn = -1;
						strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
						gnLastErrorNumber = codeBase.errorCode;	
						}
				}
			else
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, "Filter expression did not produce a LOGICAL value");
				gnLastErrorNumber = -8932;	
				}	
			}
		else
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode;	
			}
		}
	return lnReturn;
}

static PyObject *cbwTESTFILTER(PyObject *self, PyObject *args)
{
	long lnReturn;
	
	lnReturn = cbxTESTFILTER();
	if (lnReturn == -1)
		{
		return PyErr_Format(PyExc_ValueError, "Filter evaluation failed: %s", gcErrorMessage);
		}
	else
		{
		if (lnReturn == -2)
			{
			return PyErr_Format(PyExc_TypeError, "Filter Expression has not been defined");	
			}
		else
			{
			return Py_BuildValue("l", lnReturn);
			}
		}
}

/* ******************************************************************************* */
/* Clean up after the filter work is done.                                         */
static PyObject *cbwCLEARFILTER(PyObject *self, PyObject *args)
{
	if (gpCurrentFilterExpr != NULL)
		{
		expr4free(gpCurrentFilterExpr);
		gpCurrentFilterExpr = NULL;	
		}	
	return(Py_BuildValue("l", 1));
}

/* ********************************************************************************** */
/* Generic function call to replace one field's contents with the specified value.    */
/* Returns True on success, False on ERROR.                                           */
static PyObject *cbwREPLACE_FIELD(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	PyObject *lxValue;
	char lcFieldName[MAXALIASNAMESIZE + 30];
	FIELD4 *lpField;
	PyObject *lxCodes;
	PyObject *lxFieldName;
	char lcCodes[6];
	const unsigned char *cTest;	
	
	if (!PyArg_ParseTuple(args, "OOO", &lxFieldName, &lxValue, &lxCodes))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 29);
	lcFieldName[MAXALIASNAMESIZE + 29] = (char) 0;
		
	cTest = testStringTypes(lxCodes);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcCodes, Unicode2Char(lxCodes), 5);
	lcCodes[5] = (char) 0;
	Conv1252ToASCII(lcCodes, TRUE);

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	lpField = cbxGetFieldFromName(lcFieldName);
	if (lpField == NULL)
		{
		lnReturn = FALSE; /* Error messages are already set */
		}
	else
		{
		/* Ok, we have a valid field in the current table */
		lnResult = cbxPyObjectToField(lxValue, lpField, lcCodes);
		if (lnResult < 1)
			{
			PyErr_Format(PyExc_ValueError, "Field Save Failed for %s: %s", lcFieldName, gcErrorMessage);
			return NULL;
			}
		else lnReturn = lnResult;
		}

	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));
}

/* ************************************************************************** */
/* Internal version of CLOSETABLE().                                          */
long cbxCLOSETABLE(char *lpcAlias)
{
	long lnReturn = TRUE;
	long lbResult;
	DATA4 *lpTable;
	long jj;
	char lcAlias[240];
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	//printf("Closing Table %s \n", lpcAlias);	
	if (strlen(lpcAlias) > 0)
		{
		strcpy(lcAlias, lpcAlias);
		lpTable = code4data(&codeBase, lcAlias);	
		//lpTable = GetTableFromAlias(lcAlias);
		}
	else
		{
		lpTable = gpCurrentTable;
		if (gpCurrentTable != NULL)
			{
			strcpy(lcAlias, d4alias(gpCurrentTable));
			//strcpy(lcAlias, GetTableAlias(gpCurrentTable)); // TEMP
			}	
		else
			{
			lcAlias[0] = (char) 0;	
			}
		}
	//StrToUpper(lcAlias);
	//ClearAliasFromTracker(lcAlias);
	StrToLower(lcAlias);

	if (gnOpenTempIndexes > 0)
		{
		// We'll look for an associated index, even if there's been an error, just in case.
		// NOTE: We have to close the indexes before we close the table.
		for (jj = 0; jj < gnMaxTempIndexes; jj++)
			{
			if (!gaTempIndexes[jj].bOpen) continue;
			if (strcmp(gaTempIndexes[jj].cTableAlias, lcAlias) == 0)
				{
				gaTempIndexes[jj].pTable = NULL; // Tell the function that the table is already closed.
				cbxTEMPINDEXCLOSE(jj);
				}
			}	
		}
	
	if (gnExclCnt > 0)
		{
		for (jj = 0; jj < gnExclLimit; jj++)
			{
			if (strcmp(lcAlias, gaExclList[jj]) == 0)
				{
				strcpy(gaExclList[jj], "");
				gnExclCnt -= 1;
				break;	
				}	
			}
		}
	
	if (lpTable == NULL)
		{
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		gnLastErrorNumber = codeBase.errorCode;
		lnReturn = FALSE;	
		}
	else
		{
		if (lpTable == gpCurrentTable)
			{
			gpCurrentTable = NULL;
			gpCurrentIndexTag = NULL;
			}
		lbResult = d4close(lpTable);
		
		if (lbResult < 0)
			{
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode;
			lnReturn = FALSE;	
			}	
		}
	return(lnReturn);	
}

/* ************************************************************************** */
/* Close one specified table.  Pass the Alias Name.                           */
/* Returns TRUE (1) on OK, FALSE (0) if alias not found                       */
static PyObject *cbwCLOSETABLE(PyObject *self, PyObject *args)
{
	char lcAlias[MAXALIASNAMESIZE + 1];
	PyObject *lxAlias;
	long lnReturn;
	const unsigned char *cTest;

	if (!PyArg_ParseTuple(args, "O", &lxAlias))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}
		
	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);
	
	lnReturn = cbxCLOSETABLE(lcAlias);
	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));		
}

/* ****************************************************************************** */
/* Closes all tables and flushes all buffers.                                     */
/* Function like VFP CLOSE DATABASES ALL.                                         */
/* Keeps the engine alive for future table opens.                                 */
/* Returns True on success, False on failure.                                     */
static PyObject *cbwCLOSEDATABASES(PyObject *self, PyObject *args)
{
	long lbReturn = TRUE;
	long lnResult;
	register long jj;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (gnOpenTempIndexes > 0)
		{
		// We'll look for all open temporary index files.
		for (jj = 0; jj < gnMaxTempIndexes; jj++)
			{
			gaTempIndexes[jj].pTable = NULL; // Tell the function that the table is already closed.
			if (gaTempIndexes[jj].bOpen) cbxTEMPINDEXCLOSE(jj);
			}
		}
	if (gnExclCnt > 0)
		{
		for (jj = 0; jj < gnExclMax; jj++)
			{
			if (strlen(gaExclList[jj]) > 0)
				{
				strcpy(gaExclList[jj], "");
				gnExclCnt -= 1;
				if (gnExclCnt == 0) break;
				}
			}	
		}	
	lnResult = code4close(&codeBase);
	if (lnResult < 0)
		{
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error, but we close anyway. */
		gnLastErrorNumber = codeBase.errorCode;
		lbReturn = FALSE;	
		}
	gpCurrentTable = NULL;
	gpCurrentIndexTag = NULL;
	
	return(Py_BuildValue("N", PyBool_FromLong(lbReturn)));
}

/* ********************************************************************************** */
/* This function is very similar to the VFP CREATE command which takes as a           */
/* parameter the path name of the table to create and a text string with              */
/* a comma delimited list of the field descriptions.  Each field must have 5 values   */
/* in the list:                                                                       */
/* Field Name (limited in this version to 10 characters, starting with a letter or_)  */
/* Field Type (one of C, D, T, M, I, Y, B, G, X, L, N, Z -- see VFP or CodeBase doc   */
/*  NOTE: cb types V and W are NOT supported.                                         */
/* Field Size (number of chars for type C, total digits for N, etc.)                  */
/* Field Decimals (number of decimal points for Y:currency, N:number types)           */
/* Field Nulls OK (TRUE or FALSE -- or blank = FALSE indicating if NULLs are allowed) */
/* Each Field is listed on a separate line in the text file.  Newline characters sep- */
/* the lines.                                                                         */
/* Returns True on OK, False on CodeBase errors.  No indexes are created.             */
/* The default file name extension is .DBF, which will be added if none is specified. */
/* You may specify other file name extensions if you want to.                         */
/* Replaces any table with the same name.                                             */
/* On successful completion, the table is open and is the currently selected table.   */
/* NOTE: Invalid field specifications or a blank table name will trigger Python       */
/* exceptions which are propagated to the interpreter.  Creation errors from inside   */
/* CodeBase are reported by returning False.                                          */
static PyObject *cbwCREATETABLE(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult = 0;

	Py_ssize_t jj;
	Py_ssize_t lnFldCount;
	register long kk;
	char lcBuff[20];
	char lcTempName[31];
	char *lpPtr;
	long lnFldCnt;
	DATA4 *lpTable;
	PyObject *lxFieldSpecs;
	PyObject *lxFieldItem;
	PyObject *lxFieldProp;
	PyObject *pBytes;
	PyObject *lxTableName;
	char lcTableName[400];
	const unsigned char *cTest;

	if (!PyArg_ParseTuple(args, "OO", &lxTableName, &lxFieldSpecs))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	codeBase.errorCode = 0;
	codeBase.safety = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	cTest = testStringTypes(lxTableName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTableName, Unicode2Char(lxTableName), 399);
	lcTableName[399] = (char) 0;


	if (!PyList_Check(lxFieldSpecs))
		{
		PyErr_Format(PyExc_ValueError, "Second parameter must be a list() object.");
		return NULL;
		}

	if (strlen(MPStrTrim('A', lcTableName)) == 0)
		{
		PyErr_Format(PyExc_ValueError, "Table name may not be blank");	
		return NULL;
		}
		
	lpPtr = strstr(lcTableName, ".");
	if (lpPtr == NULL)
		{
		strcat(lcTableName, ".DBF");	
		}
	
	lnFldCount = PyList_Size(lxFieldSpecs);
	if (lnFldCount == 0)
		{
		PyErr_Format(PyExc_ValueError, "No field specs in input list");	
		return NULL;
		}
	if (lnFldCount > MAXFIELDCOUNT)
		{
		PyErr_Format(PyExc_ValueError, "Too many fields defined: %ld. Limit is %ld", lnFldCount, MAXFIELDCOUNT);	
		return NULL;
		}
		
	for (jj = 0;jj < lnFldCount;jj++)
		{
		lxFieldItem = PyList_GetItem(lxFieldSpecs, jj);
		// The above is a Python object of type VFPFIELD.
		lxFieldProp = PyObject_GetAttrString(lxFieldItem, "cName");
#ifdef IS_PY3K	
				pBytes = PyUnicode_AsEncodedString(lxFieldProp, u8"cp1252", u8"ignore");
				strncpy(lcTempName, PyBytes_AsString(pBytes), 20);
#endif
#ifdef IS_PY2K
				strncpy(lcTempName, PyString_AsString(lxFieldProp), 20);
#endif	
		lcTempName[20] = (char) 0;
		Py_DECREF(lxFieldProp);
		lxFieldProp = NULL;
		
		if (strlen(lcTempName) > 10)
			{
			PyErr_Format(PyExc_ValueError, "Field name %s is too long", lcTempName);	
			return NULL;
			}
		gaFieldSpecs[jj].name = calloc((size_t) 13, sizeof(char)); /* longer than needed */
		strncpy(gaFieldSpecs[jj].name, lcTempName, 11);
		
		for (kk = 0; kk < jj; kk++)
			{
			if (strcmp(lcTempName, gaFieldSpecs[kk].name) == 0) /* This is a duplicate name. BAD */
				{
				PyErr_Format(PyExc_ValueError, "Field name %s is DUPLICATED", lcTempName);
				return NULL;
				}	
			}
		lxFieldProp = PyObject_GetAttrString(lxFieldItem, "cType");
#ifdef IS_PY3K
				pBytes = PyUnicode_AsEncodedString(lxFieldProp, u8"cp1252", u8"ignore");
				strncpy(lcBuff, PyBytes_AsString(pBytes), 4);
#endif
#ifdef IS_PY2K
				strncpy(lcBuff, PyString_AsString(lxFieldProp), 4);
#endif			
		lcBuff[4] = (char) 0;
		Py_DECREF(lxFieldProp);
		lxFieldProp = NULL;
		gaFieldSpecs[jj].type = (short) lcBuff[0];
		
		lxFieldProp = PyObject_GetAttrString(lxFieldItem, "nWidth");
		gaFieldSpecs[jj].len = PyLong_AsLong(lxFieldProp);
		Py_DECREF(lxFieldProp);
		lxFieldProp = NULL;
		
		lxFieldProp = PyObject_GetAttrString(lxFieldItem, "nDecimals");
		gaFieldSpecs[jj].dec = PyLong_AsLong(lxFieldProp);
		Py_DECREF(lxFieldProp);
		lxFieldProp = NULL;
		
		lxFieldProp = PyObject_GetAttrString(lxFieldItem, "bNulls");
		if (PyObject_IsTrue(lxFieldProp))
			gaFieldSpecs[jj].nulls = r4null;
		else
			gaFieldSpecs[jj].nulls = 0;
		Py_DECREF(lxFieldProp);
		lxFieldProp = NULL;
		}
	
	if (lnReturn == TRUE) /* All the fields were correctly specified */
		{
		lpTable = d4create(&codeBase, lcTableName, gaFieldSpecs, NULL);
		if (lpTable == NULL)
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			strcat(gcErrorMessage, " File Name: ");
			strcat(gcErrorMessage, lcTableName);
			gnLastErrorNumber = codeBase.errorCode;
			gpCurrentTable = NULL;	
			}
		else
			{
			gpCurrentTable = lpTable;	
			}
		}
	cbxClearFieldSpecs();

	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));	
}

/* ********************************************************************************** */
/* An INTERNAL VERSION of cbwCREATETABLE() callable from within this C module.        */
/* This assumes that the gaFields global array has already been populated with the*/
/* data required for producing the required table by passing to d4create().           */
/* There is NO error checking because there will never be any external field spec     */
/* data passed to this function.                                                      */
long cbxCREATETABLE(char* lpcTableName)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	long lbFieldOK;
	long jj;
	long kk;
	double ldStart, ldEnd;
	char lcFullName[410];
	char lcBuff[20];
	char lcTempName[20];
	long lbDupeNameError = FALSE;
	char *lpPtr;
	char *lpTemp;
	long lnFldCnt = -1;
	DATA4 *lpTable;

	codeBase.errorCode = 0;
	codeBase.safety = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	memset(lcFullName, 0, (size_t) 410 * sizeof(char)); // space for the .dbf
	strncpy(lcFullName, lpcTableName, 399);
	lpPtr = strstr(lcFullName, ".");
	if (lpPtr == NULL)
		{
		strcat(lcFullName, ".DBF");	
		}
	// gaFieldSpecs[] array has already been populated.	
	if (lnReturn == TRUE) /* All the fields were correctly specified */
		{
		lpTable = d4create(&codeBase, lcFullName, gaFieldSpecs, NULL);
		if (lpTable == NULL)
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			strcat(gcErrorMessage, " File Name: ");
			strcat(gcErrorMessage, lcFullName);
			gnLastErrorNumber = codeBase.errorCode;
			gpCurrentTable = NULL;	
			}
		else
			{
			gpCurrentTable = lpTable;	
			}
		}
	return(lnReturn);	
}

/* ********************************************************************************** */
/* Utility function that parses a date string using a pattern string as a guide.      */
long cbxParseDateByFormat(char *lpcDateStr, char *lpcPattern, char* lpcDateSep, char* lpcPattSep)
{
	char lcYear[10];
	char lcMonth[10];
	char lcDay[10];
	char lcWork[20];
	
	char cDateSep;
	char cPattSep;
	char cSchema[6];
	char cPrevChar = (char) 0;
	char cTemp;
	char* pPtr;
	long jj;
	long lnLen;
	long lnPos = 0;
	long lnTypePos = 0;
	long cTypeChar = 0;
	long bGoodSchema = TRUE;
	long lnDateReturn = 0;
	long bGoodDate = TRUE;
	long nSegCount = 0;
	
	cDateSep = lpcDateSep[0];
	cPattSep = lpcPattSep[0];
	
	memset(lcYear, 0, (size_t) (10 * sizeof(char)));
	memset(lcMonth, 0, (size_t) (10 * sizeof(char)));
	memset(lcDay, 0, (size_t) (10 * sizeof(char)));	
	memset(cSchema, 0, (size_t) (6 * sizeof(char)));
	//printf("Parsing by Format: %s ## %s ## %s ## %s \n", lpcDateStr, lpcPattern, lpcDateSep, lpcPattSep);
	lnLen = strlen(lpcPattern); 
	// First we parse the pattern to get the underlying schema: MDY, DMY or YMD
	for (jj = 0; jj < lnLen; jj++)
		{
		cTemp = lpcPattern[jj];
		if (cTemp == cPattSep) continue;
		pPtr = strchr("MDY", cTemp);
		if (pPtr == NULL)
			{
			bGoodSchema = FALSE;
			break;	
			}
		if (cTemp != cPrevChar)
			{
			cPrevChar = cTemp;
			cSchema[lnPos] = cTemp;
			lnPos += 1;	
			if (lnPos > 3)
				{
				bGoodSchema = FALSE;
				break;	
				}
			}	
		}
	//printf("Schema: %s  good flag %ld \n", cSchema, bGoodSchema);
	if (bGoodSchema)
		{
		pPtr = strchr(lpcDateStr, ' ');
		if (pPtr != NULL) *pPtr = (char) 0; // Truncate at the first space character encountered.
		lnLen = strlen(lpcDateStr);
		// Now we parse this into Year, Month and Day strings...
		lnPos = 0;
		memset(lcWork, 0, (size_t) (20 * sizeof(char)));
		for (jj = 0; jj <= lnLen; jj++)
			{
			cTemp = lpcDateStr[jj];
			if ((cTemp != cDateSep) && (cTemp != (char) 0))
				{
				if (!isdigit(cTemp))
					{
					bGoodDate = FALSE;
					break;	
					}
				lcWork[lnPos] = cTemp;
				lnPos += 1;
				if (lnPos > 19)
					{
					bGoodDate = FALSE;
					//printf("lnPos Too Big %ld \n", lnPos);
					break;	
					}
				}	
			else // It IS the separator or the terminating NULL, so we do a copy as appropriate.
				{
				nSegCount += 1;
				if (nSegCount > 3)
					{
					// invalid format...
					bGoodDate = FALSE;
					//printf("Too many segments: %ld \n", nSegCount);
					break;	
					}
				switch (cSchema[nSegCount - 1])
					{
					case 'M':
						strncpy(lcMonth, lcWork, 5);
						lcMonth[5] = (char) 0;
						break;
						
					case 'D':
						strncpy(lcDay, lcWork, 5);
						lcDay[5] = (char) 0;
						break;
						
					case 'Y':
						strncpy(lcYear, lcWork, 5);
						lcYear[5] = (char) 0;
						break;
						
					default:
						bGoodSchema = FALSE;
						break;	
					}
				memset(lcWork, 0, (size_t) (20 * sizeof(char)));
				lnPos = 0;
				}	
			}
		if (bGoodDate)
			{
			switch(strlen(lcYear))
				{
				case 1:
				case 3:
					bGoodDate = FALSE;
					break;
					
				case 2:
					if (strcmp(lcYear, "30") < 0)
						{
						strcpy(lcWork, "20");
						strcat(lcWork, lcYear);	
						}	
					else
						{
						strcpy(lcWork, "19");
						strcat(lcWork, lcYear);	
						}
					break;
					
				case 4:
					strcpy(lcWork, lcYear);
					
				default:
					bGoodDate = FALSE;
					break;
				}
			if (bGoodDate)
				{
				if (strlen(lcMonth) == 2)
					{
					strcat(lcWork, lcMonth);
					}
				else
					{
					strcat(lcWork, "0");
					strncat(lcWork, lcMonth, 1);
					}	
				if (strlen(lcDay) == 2)
					strcat(lcWork, lcDay);
				else
					{
					strcat(lcWork, "0");
					strncat(lcWork, lcDay, 1);
					}
				
				lnDateReturn = date4long(lcWork);
				if (lnDateReturn < 0) lnDateReturn = 0;
				}
			}
		}
	return lnDateReturn;	
}

/* ********************************************************************************** */
/* A reasonably flexible routine for parsing a date string value into a long to store */
/* in a DBF table DATE type field from.  Takes a single pointer to a character string */
/* as its parameter and returns the long value which can be stored by f4assignLong()  */
/* into the DATE field.  This is a Julian date.  Returns 0 on error or if for any     */
/* reason it is unable to figure out what the date should be.                         */
/* This expects a string on one of several formats, but all consisting of numbers with*/
/* or without separating punctuation.  For example: 08/15/15 for August 15, 2015.     */
/* Uses the currently set dateformat value to determine what pattern of MDY it should */
/* use to interpret the parts.                                                        */
long cbxParseDate(char* lpcDateStr)
{
	long lnStrLen;
	long lnDateReturn = 0;
	long lbPunctuation = FALSE;
	long lbPatternPunc = FALSE;
	char* lpPtr = NULL;
	char lcSeparator[2];
	char lcPattSep[2];
	char lcPattern[20];
	char lcYear[10];
	char lcMonth[10];
	char lcDay[10];
	char lcStdDate[20];
	char lcWorkDate[120];
	char lcTest[50]; // Big just in case.
	
	memset(lcYear, 0, (size_t) (10 * sizeof(char)));
	memset(lcMonth, 0, (size_t) (10 * sizeof(char)));
	memset(lcDay, 0, (size_t) (10 * sizeof(char)));
	lcPattSep[0] = lcPattSep[1] = lcSeparator[0] = lcSeparator[1] = (char) 0;

	strcpy(lcWorkDate, MPSSStrTran(lpcDateStr, " ", "", 0)); // Removes spaces throughout the string, as they are meaningless.
	lnStrLen = strlen(lcWorkDate);

	lpPtr = strpbrk(lcWorkDate, ".-/");
	lbPunctuation = (lpPtr != NULL);
	if (lbPunctuation) lcSeparator[0] = *lpPtr;
	strcpy(lcPattern, code4dateFormat(&codeBase));
	
	lpPtr = strpbrk(lcPattern, ".-/");
	lbPatternPunc = (lpPtr != NULL);
	if (lbPatternPunc) lcPattSep[0] = *lpPtr;

	while (TRUE)
		{
		if (lbPunctuation && lbPatternPunc)
			{
			// Whether the same or different punctuation, we'll still use the
			// pattern as the basis for parsing.
			lnDateReturn = cbxParseDateByFormat(lcWorkDate, lcPattern, lcSeparator, lcPattSep);
			break;
			}	
		
		if (lbPunctuation && !lbPatternPunc)
			{
			// A hard one.  The pattern doesn't have punctuation.  Normally
			// that means YYYYMMDD format or YYMMDD format.  But the string
			// DOES have punctuation.  We'll try the USA Default pattern.
			lnDateReturn = cbxParseDateByFormat(lcWorkDate, "MM/DD/YYYY", lcSeparator, "/");
			break;	
			}	
		
		if (!lbPunctuation)
			{
			// There is no punctuation in the incoming string.  We'll try parsing
			// based on YYMMDD or YYYYMMDD.  If they fail, we'll try MMDDYY or
			// MMDDYYYY.

			switch(lnStrLen)
				{
				case 6: // 6 character dates are never a good idea for Y2K reasons, but they are still with us.
					strncpy(lcYear, lcWorkDate, 2);
					lcYear[2] = (char) 0;
					strncpy(lcMonth, lcWorkDate + 2, 2);
					lcMonth[2] = (char) 0;
					strncpy(lcDay, lcWorkDate + 4, 2);
					lcDay[2] = (char) 0;
					if (strcmp(lcYear, "30") < 0)
						{
						strcpy(lcStdDate, "19");
						}
					else
						{
						strcpy(lcStdDate, "20");	
						}
					strcat(lcStdDate, lcYear);
					strcat(lcStdDate, lcMonth);
					strcat(lcStdDate, lcDay);	
					lnDateReturn = date4long(lcStdDate);
					if (lnDateReturn < 0)
						{
						// Try MMDDYY format
						strncpy(lcMonth, lcWorkDate, 2);
						lcMonth[2] = (char) 0;
						strncpy(lcDay, lcWorkDate + 2, 2);
						lcDay[2] = (char) 0;
						strncpy(lcYear, lcWorkDate + 4, 2);
						lcYear[2] = (char) 0;
						if (strcmp(lcYear, "30") < 0)
							{
							strcpy(lcStdDate, "19");
							}	
						else
							{
							strcpy(lcStdDate, "20");	
							}
						strcat(lcStdDate, lcYear);
						strcat(lcStdDate, lcMonth);
						strcat(lcStdDate, lcDay);
						lnDateReturn = date4long(lcStdDate);
						if (lnDateReturn < 0) lnDateReturn = 0; // We're done.	
						}				
					break;
					
				case 8:
					// We can only hope that this is in YYYYMMDD format.
					lnDateReturn = date4long(lcWorkDate);
					if (lnDateReturn < 0) lnDateReturn = 0;
					break;
					
				default:
					// No idea of what to do with this.  Bad date string.
					break;					
				}
				
					
			break;	
			}	
		break;
		}
	
	return lnDateReturn;
}

/* ********************************************************************************** */
/* Generic function for converting source data in string/char format to a CodeBase    */
/* field type.  Used in the append functions below.                                   */
long cbxStoreCharValue(long lpnType, FIELD4* lppField, char* lpcString)
{
	double nDbl;
	char lcWorkStr[1000];
	char cChar;
	long nTest;
	
	switch(lpnType)
		{
		case 'C':
			f4assign(lppField, lpcString);
			break;
			
		case 'M':
		case 'X':
			f4memoAssign(lppField, lpcString);
			break;	
		
		case 'N':
		case 'B':
		case 'F':
			nDbl = atof(lpcString);
			f4assignDouble(lppField, nDbl);										
			break;
		
		case 'Y':
			f4assignCurrency(lppField, lpcString);
			break;
			
//		case 'Q':
//		case 'W':
//			f4assignUnicode(lppField, lpcString);
			
		case 'T':
			strcpy(lcWorkStr, lpcString);
			MPSSStrTran(lcWorkStr, " ", "", 0);
			f4assignDateTime(lppField, lcWorkStr);
			break;
			
		case 'L':
			cChar = 'F'; // False;
			strcpy(lcWorkStr, lpcString);
			MPSSStrTran(lcWorkStr, " ", "", 0);
			if ((lcWorkStr[0] == 'Y') || (lcWorkStr[0] == 'y') || (lcWorkStr[0] == 'T') || (lcWorkStr[0] == 't') || (lcWorkStr[0] == '1'))
				cChar = 'T';
			f4assignChar(lppField, cChar);
			break;
			
		case 'D':
			nTest = cbxParseDate(lpcString);
			f4assignLong(lppField, nTest);
			break;
			
		case 'I':
			nTest = atol(lpcString);
			f4assignLong(lppField, nTest);
			break;
			
		default:
			break;
		}
	return (TRUE);
}

/* ***************************************************************************** */
/* Appends a blank record to the end of the currently selected table.            */
/* Returns TRUE for OK, -1 on ERROR.                                             */
/* Internal function.                                                            */
long cbxAPPENDBLANK(void)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4appendBlank(gpCurrentTable);
		if (lnResult != 0) /* Something went wrong */
			{
			if (lnResult > 0)
				{
				strcpy(gcErrorMessage, "Unable to Move Record Pointer");
				gnLastErrorNumber = -9980;
				lnReturn = -1; /* One of the several locking and duplicate key errors.  In any event, FAILURE. */	
				}
			else
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				lnReturn = -1; /* We're done... ERROR */					
				}
			}
		}
	else
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return(lnReturn);	
}

/* ********************************************************************************** */
/* Python aware version of the routine above.                                         */
static PyObject *cbwAPPENDBLANK(PyObject *self, PyObject *args)
{
	long lnReturn = cbxAPPENDBLANK();
	if (lnReturn == -1) lnReturn = 0;
	return Py_BuildValue("N", PyBool_FromLong(lnReturn));	
}

/* ***************************************************************************** */
/* Function that works like VFP SELECT myAlias.  Pass the Alias Name.            */
/* Returns 1 on success, 0 on "can't find alias"                                 */
/* INTERNAL VERSION.                                                             */
long cbxSELECT(char* lpcAlias)
{
	long lbReturn = TRUE;
	DATA4 *lpTable;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
		
	lpTable = code4data(&codeBase, lpcAlias);
	if (lpTable == NULL)
		{
		lbReturn = FALSE;
		strcpy(gcErrorMessage, "Alias not found: ");
		strcat(gcErrorMessage, lpcAlias);
		gnLastErrorNumber = -4949;
		}
	else
		{
		gpCurrentTable = lpTable;	
		}
	return(lbReturn);
}


/* ********************************************************************************** */
/* Helper function for the cbwAPPENDFROM function that handles the import from DBF.   */
long cbxAppendFromDBF(DATA4* lpTargetTable, char* lpcSource, char* lpcTestExpr)
{
	long lnReturn = 0;
	long lbOK = TRUE;
	VFPFIELD* lpFld;
	DATA4* lpSourceTable;
	FIELD4* lpSrcField;
	FIELD4* lpTrgField;
	long lnCnt = 0;
	long lnAppendCount = 0;
	long jj;
	long nTest = 1;
	long bNeedTest = FALSE;
	long kk;
	long mCnt;
	long nLen;
	long nSrcLen;
	long nTrgLen;
	long nSrcRecs;
	double nDbl;
	char cChar;
	char lcTargAlias[40];
	char lcSrcAlias[40];
	char lcWorkStr[300];
	FieldPlus fpSrc;
	FieldPlus fpTrg;
	
	strcpy(lcTargAlias, cbxALIAS());
	strcpy(lcSrcAlias, "CBWSRC");
	strcat(lcSrcAlias, MPStrRand(6));
	StrToUpper(lcSrcAlias);
	lbOK = cbxUSE(lpcSource, lcSrcAlias, TRUE, FALSE, FALSE);
	if (lbOK)
		{
		cbxSELECT(lcSrcAlias);
		if ((lpcTestExpr != NULL) && (lpcTestExpr[0] != (char) 0))
			{
			nTest = cbxPREPAREFILTER(lpcTestExpr); // If this fails, we're done!	
			if (nTest == 1) bNeedTest = TRUE;
			}
		if (nTest == 1)
			{
			lpSourceTable = gpCurrentTable;
			// Now we get the field type information and write it into the public arrays.
			lpFld = cbxAFIELDS(&lnCnt);
			gnSrcCount = lnCnt;
			mCnt = 0;
			for (jj = 0; jj < lnCnt; jj++)
				{
				gaSrcFields[jj].xField = gaFields[jj];
				gaSrcFields[jj].nFldType = (long) gaFields[jj].cType[0];
				gaSrcFields[jj].bMatched = FALSE;			
				gaSrcFields[jj].pCBfield = d4field(lpSourceTable, gaSrcFields[jj].xField.cName);
				for (kk = 0; kk < gnTrgCount; kk++)
					{
					if (strcmp(gaTrgFields[kk].xField.cName, gaSrcFields[jj].xField.cName) == 0)
						{
						mCnt += 1;
						gaMatchFields[mCnt - 1] = gaTrgFields[kk];
						gaMatchFields[mCnt - 1].bMatched = TRUE;
						gaMatchFields[mCnt - 1].nSource = jj;
						gaMatchFields[mCnt - 1].nTarget = kk;
						gaMatchFields[mCnt - 1].bIdentical = FALSE;
						if ((gaTrgFields[kk].xField.nWidth == gaSrcFields[jj].xField.nWidth) && 
							(gaTrgFields[kk].xField.nDecimals == gaSrcFields[jj].xField.nDecimals) &&
							(gaTrgFields[kk].nFldType == gaSrcFields[jj].nFldType) && 
							(gaTrgFields[kk].nFldType != (long) 'M') && 
							(gaTrgFields[kk].nFldType != (long) 'G') && // most of these require a binary copy
							(gaTrgFields[kk].nFldType != (long) 'X') &&
							(gaTrgFields[kk].nFldType != (long) 'Z') &&
							(gaSrcFields[kk].nFldType != (long) 'M') &&
							(gaSrcFields[kk].nFldType != (long) 'G') &&
							(gaSrcFields[kk].nFldType != (long) 'Z') &&
							(gaSrcFields[kk].nFldType != (long) 'X'))
							gaMatchFields[mCnt - 1].bIdentical = TRUE; // Only test this once for the fastest way to copy field info.
						break;	
						}
					}
				}
			}
		else
			{
			strcpy(gcErrorMessage, "Invalid Test Expression");
			gnLastErrorNumber = -7945;	
			lnReturn = -1;
			}	
		if ((mCnt > 0) && (nTest == 1))
			{
			// We have at least one matching field, so we can do the append.
			kk = cbxGOTO("TOP", 0);
			
			while (cbxEOF() == 0)
				{
				cbxSELECT(lcSrcAlias);
				if (bNeedTest)
					{
					if (cbxTESTFILTER() != 1)
						{
						if (cbxGOTO("NEXT", 0) < 1) break;
						if (cbxEOF() != 0) break;
						continue;
						}
					}
				cbxSELECT(lcTargAlias);
				if (cbxAPPENDBLANK() == TRUE)
					{
					lnAppendCount += 1;
					for (kk = 0; kk < mCnt; kk++)
						{
						fpTrg = gaMatchFields[kk];
						fpSrc = gaSrcFields[gaMatchFields[kk].nSource];
						lpSrcField = fpSrc.pCBfield;
						lpTrgField = fpTrg.pCBfield;
						while (TRUE) // Not a loop really, but a glorified switch.
							{
							if (f4null(lpSrcField))
								{
								if (fpTrg.xField.bNulls) f4assignNull(lpTrgField);
								// No else. We leave it blank if it isn't allowed to be null.
								break;	
								}	
								
							if (fpTrg.bIdentical) // Exactly identical, so we can use this very efficient function.
								{
								f4assignField(lpTrgField, lpSrcField);								
								break;	
								}
								
							if (((fpTrg.nFldType == (long) 'N') && (fpSrc.nFldType == (long) 'N')) ||
								((fpTrg.nFldType == (long) 'N') && (fpSrc.nFldType == (long) 'F')) ||
								((fpTrg.nFldType == (long) 'F') && (fpSrc.nFldType == (long) 'N')) ||
								((fpTrg.nFldType == (long) 'F') && (fpSrc.nFldType == (long) 'B')) ||
								((fpTrg.nFldType == (long) 'B') && (fpSrc.nFldType == (long) 'F')) ||
								((fpTrg.nFldType == (long) 'N') && (fpSrc.nFldType == (long) 'B')) ||
								((fpTrg.nFldType == (long) 'B') && (fpSrc.nFldType == (long) 'N')))// d4assignField() can also be used in a few other cases...
								{
								f4assignField(lpTrgField, lpSrcField);
								break;									
								}
								
							if ((fpTrg.nFldType == (long) 'C') && (fpSrc.nFldType == (long) 'C'))
								{
								f4assign(lpTrgField, f4str(lpSrcField));
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'M') && (fpSrc.nFldType == (long) 'M'))
								{
								nSrcLen = f4memoLen(lpSrcField);
								f4memoAssignN(lpTrgField, f4memoPtr(lpSrcField), nSrcLen);
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'C') && (fpSrc.nFldType == (long) 'M'))
								{
								f4assign(lpTrgField, f4memoStr(lpSrcField)); // Assumes a text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'M') && (fpSrc.nFldType == (long) 'C'))
								{
								f4memoAssign(lpTrgField, f4str(lpSrcField)); // Assumes a text string
								break;	
								}

							if ((fpTrg.nFldType == (long) 'Z') && (fpSrc.nFldType == (long) 'C'))
								{
								nSrcLen = f4len(lpSrcField);
								f4assignN(lpTrgField, f4ptr(lpSrcField), nSrcLen); // Assumes a text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'C') && (fpSrc.nFldType == (long) 'Z'))
								{
								nSrcLen = f4len(lpSrcField);								
								f4assignN(lpTrgField, f4ptr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'Z') && (fpSrc.nFldType == (long) 'Z'))
								{
								nSrcLen = f4len(lpSrcField);								
								f4assignN(lpTrgField, f4ptr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'Z') && (fpSrc.nFldType == (long) 'M'))
								{
								nSrcLen = f4memoLen(lpSrcField);
								f4assignN(lpTrgField, f4memoPtr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}

							if ((fpTrg.nFldType == (long) 'M') && (fpSrc.nFldType == (long) 'Z'))
								{
								nSrcLen = f4len(lpSrcField);
								f4memoAssignN(lpTrgField, f4ptr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'Z') && ((fpSrc.nFldType == (long) 'X') || (fpSrc.nFldType == (long) 'X')))
								{
								nSrcLen = f4memoLen(lpSrcField);									
								f4assignN(lpTrgField, f4memoPtr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'X') && (fpSrc.nFldType == (long) 'Z'))
								{
								nSrcLen = f4len(lpSrcField);									
								f4memoAssignN(lpTrgField, f4ptr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}
								
							if (((fpTrg.nFldType == (long) 'X') && (fpSrc.nFldType == (long) 'X')) ||
								((fpTrg.nFldType == (long) 'G') && (fpSrc.nFldType == (long) 'X')) ||
								((fpTrg.nFldType == (long) 'G') && (fpSrc.nFldType == (long) 'G')) ||
								((fpTrg.nFldType == (long) 'X') && (fpSrc.nFldType == (long) 'G')))
								{
								nSrcLen = f4memoLen(lpSrcField);									
								f4memoAssignN(lpTrgField, f4memoPtr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'X') && (fpSrc.nFldType == (long) 'M'))
								{
								nSrcLen = f4memoLen(lpSrcField);									
								f4memoAssignN(lpTrgField, f4memoPtr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'M') && (fpSrc.nFldType == (long) 'X'))
								{
								nSrcLen = f4memoLen(lpSrcField);									
								f4memoAssignN(lpTrgField, f4memoPtr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'C') && (fpSrc.nFldType == (long) 'X'))
								{
								nSrcLen = f4memoLen(lpSrcField);									
								f4assignN(lpTrgField, f4memoPtr(lpSrcField), nSrcLen); // Assumes a binary text string
								break;	
								}
								
							if ((fpTrg.nFldType == (long) 'X') && (fpSrc.nFldType == (long) 'C'))
								{
								nSrcLen = f4len(lpSrcField);									
								f4memoAssignN(lpTrgField, f4ptr(lpSrcField), nSrcLen); // Assumes a text string
								break;	
								}
// default
							if (fpSrc.nFldType == (long) 'C')
								{
								cbxStoreCharValue(fpTrg.nFldType, lpTrgField, f4str(lpSrcField));
								break;	
								}
							break;							
							}	
						}	
					}
				else
					{
					break; // Something bad happened.	
					}
				// Now we've got the buffer filled and the record is ready to be saved.
				//d4flush(lpTargetTable); // TOO SLOW
				cbxSELECT(lcSrcAlias);
				if (cbxGOTO("NEXT", 0) < 1) break;
				if (cbxEOF() != 0) break;
				}// end of table loop
			lnReturn = lnAppendCount;
			}
		else
			{
			strcpy(gcErrorMessage, "No Common Data Fields, can't append");
			gnLastErrorNumber = -7943;	
			lnReturn = -1;
			}

		d4flush(lpTargetTable);
		cbxCLOSETABLE(lcSrcAlias);	
		}
	else
		{
		strcpy(gcErrorMessage, "Unable to open Source Table");
		gnLastErrorNumber = -7944;
		lnReturn = -1;
		}
	
	return lnReturn;
}

/* ********************************************************************************** */
/* Helper function for the cbwAPPENDFROM function that handles the import from SDF.   */
long cbxAppendFromSDF(DATA4* lpTargetTable, char* lpcSource, char* lpcTestExpr)
{
	long laFromMark[300];
	long laFldLength[300];
	long lnWidth = 0;
	long lnReturn = 0;
	long lnFromPos = 0;
	long jj;
	VFPFIELD* lpFld;
	char lcTargAlias[30];
	FILE* pHandle;
	char cLineBuff[10000];
	char cFieldStr[500];
	char* pPtr = NULL;
	long lnLineLen;
	long lnTestLen;
	char nType;
	long lnAppendCount = 0;
	long lnCurrPos = 0;
	FIELD4* lpTrgField;
	long bGoodLoad = TRUE;
	
	strcpy(lcTargAlias, cbxALIAS());
	//printf("SDF APPENDING to %s\n", lcTargAlias);
	for (jj = 0; jj < gnTrgCount; jj++)
		{
		lnWidth = gaTrgFields[jj].xField.nWidth;
		if (gaTrgFields[jj].xField.cType[0] == 'I') lnWidth = 11; // Forced wider to handle full width of max long int
		if (gaTrgFields[jj].xField.cType[0] == 'B') lnWidth = 16; // Forced wide for double sized values.
		if (gaTrgFields[jj].xField.cType[0] == 'T') lnWidth = 16; // Date time text has to be wider than 8, the native width.
		laFromMark[jj] = lnFromPos;
		lnFromPos += lnWidth;
		laFldLength[jj] = lnWidth;
		if ((gaTrgFields[jj].nFldType == (long) 'M') || (gaTrgFields[jj].nFldType == (long) 'G'))
			{
			bGoodLoad = FALSE;
			strcpy(gcErrorMessage, "SDF Type Import Cannot target table with a MEMO or GENERAL type fields.");
			gnLastErrorNumber = -7956;	
			lnReturn = -1;
			}
		if (lnWidth > 498)
			{
			bGoodLoad = FALSE;
			strcpy(gcErrorMessage, "SDF Type Import Max Field Size is 498 bytes (VFP Max 254).");
			gnLastErrorNumber = -7957;
			lnReturn = -1;
			}
		}
	if (bGoodLoad)
		{
		memset(cLineBuff, 0, (size_t) (10000 * sizeof(char)));
		memset(cFieldStr, 0, (size_t) (500 * sizeof(char)));
		pHandle = fopen(lpcSource, "r");
		if (pHandle != NULL)
			{
			while (!feof(pHandle))
				{
				pPtr = fgets(cLineBuff, 9999, pHandle);
				if (pPtr == NULL) break;
				lnLineLen = strlen(cLineBuff);
				if (lnLineLen > laFldLength[0])  // Has to be at least as much data as in the first field.
					{
					cbxSELECT(lcTargAlias);
					if (cbxAPPENDBLANK() == TRUE)
						{
						lnAppendCount += 1;
						lnCurrPos = 0;
						for (jj = 0; jj < gnTrgCount; jj++)
							{
							lpTrgField = gaTrgFields[jj].pCBfield;
							nType = gaTrgFields[jj].nFldType;
							if (lnCurrPos < lnLineLen)
								{
								lnTestLen = laFldLength[jj];
								if (lnTestLen > 499) lnTestLen = 499;
								strncpy(cFieldStr, cLineBuff + laFromMark[jj], lnTestLen);
								cFieldStr[lnTestLen] = (char) 0;
								cbxStoreCharValue(nType, gaTrgFields[jj].pCBfield, cFieldStr);
								}
							else
								{
								break;	
								}
							}
						}					
					}
				else
					{
					break;	
					}
				}
			cbxSELECT(lcTargAlias);
			cbxFLUSH();
			fclose(pHandle);
			lnReturn = lnAppendCount;
			}
		else
			{
			strcpy(gcErrorMessage, "Source File missing or unavailable");
			gnLastErrorNumber = -7955;
			lnReturn = -1;
			}
		}
	return lnReturn;
}

/* ********************************************************************************** */
/* Helper function for the cbwAPPENDFROM function that handles the import from CSV.   */
/* Implemented in Python due to the powerful tools available.                         */
long cbxAppendFromCSV(DATA4* lpTargetTable, char* lpcSource, char* lpcTestExpr)
{
	long lnReturn = -1;
		
	return lnReturn;
}


/* ********************************************************************************** */
/* This is an in-DLL tool for appending the contents of one table to the end of       */
/* another.  The goal is speed by avoiding passing data into the PY wrapper and doing */
/* all the data translations during the process.                                      */
/*                                                                                    */
/* Parameters:                                                                        */
/* - lpcAlias - alias of any currently open table.  Pass an empty string to refer     */
/*   to the currently selected table.  If there isn't a selected table or the alias   */
/*   isn't recognized, reports an error.  This is the table into which the external   */
/*   data will be appended.                                                           */
/*                                                                                    */
/* - lpcSource - fully qualified path name of the source data file.  The file name    */
/*   extension will be the primary means of determining what type of file it is.  DBF */
/*   will signify a table to be opened by CodeBase.  CSV will be recognized as text   */
/*   comma separated values.  SDF or DAT will be assumed to be text data with fixed   */
/*   length fields such that their lengths are identical to those of the fields in the*/
/*   target table.  If the extension does NOT reflect the type of the file, then the  */
/*   type string MUST be provided in the lpcType parameter.                           */
/*                                                                                    */
/* - lpcTestExpr - any expression evaluating to a logical result which is made up of  */
/*   field names in the targeted alias table.  If this expression evaluates to TRUE,  */
/*   the record will be appended into the table, otherwise the source data will be    */
/*   skipped. CURRENTLY ONLY IMPLEMENTED FOR DBF TYPE APPENDS.                        */
/*                                                                                    */
/* - lpcType - If the lpcSource file name has a correct file name extension reflecting*/
/*   the type, then pass this as NULL or an empty string.  Otherwise this must have   */
/*   a string with one of these values: DBF, CSV, DAT or SDF.                         */
/*                                                                                    */
/* Returns:                                                                           */
/* The number of records appended (0 to n records in the source) or -1 on error.      */
/*                                                                                    */
/* Comments:                                                                          */
/* For DBF type sources, there MUST be at least one field name in the source which    */
/* matches the target alias table exactly.  Otherwise an error will be returned.      */
/* Fields with matching names and types will be copied exactly from the source to the */
/* target if there is sufficient room in the target field.  If the target field is too*/
/* small, data will be truncated without warning.                                     */
/*                                                                                    */
/* In the case where the DBF field names are the same, but the types are different,   */
/* the source data will be matched as well as possible.  For example a source field   */
/* which is a DATE type or a DATETIME type if paired with a field in the target of    */
/* type CHAR will be stored as a raw string representing the date.  If paired with a  */
/* LONG type field, the result would be a numeric value of the Julian Date.           */
/*                                                                                    */
/* DBF source fields which permit NULL values and which are NULL will be stored in the*/
/* target table as if the target field allows NULLS, otherwise as an empty value of   */
/* the target field's type.                                                           */
/*                                                                                    */
/* In the case of SDF, CSV, and DAT type files, the source will always be a string.   */
/* NO data conversions will be done explicitly except those which DBF tables permit   */
/* by design.  So a value of T will be translated to True if the target field is      */
/* type LOGICAL, for example, but a value of G will produce undefined results or      */
/* possibly False.                                                                    */
long cbxAPPENDFROM(char* lpcAlias, char* lpcSource, char* lpcTestExpr, char* lpcType)
{
	long lnResult = 0;	
	DATA4 *lpTargetTable = NULL;
	DATA4 *lpOldTable = NULL;
	char lcExt[20];
	char lcWorkName[300];
	long lnPt = 0;
	long lnType = -1; // Types 1 = DBF, 2 = SDF or DAT, 3 = CSV
	char *lpChar = NULL;
	char *lpTest = NULL;
	long lbByType = FALSE;
	char lcWorkAlias[50];
	VFPFIELD* lpFld;
	long lnCnt = 0;
	long jj;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	lpOldTable = gpCurrentTable;
	
	if ((lpcAlias == NULL) || (lpcAlias[0] == '\0'))
		{
		lpTargetTable = gpCurrentTable;	
		}
	else
		{
		lpTargetTable = code4data(&codeBase, lpcAlias);	
		}
		
	if (lpTargetTable == NULL)
		{
		strcpy(gcErrorMessage, "No Target Table Available or Alias Not Found");
		gnLastErrorNumber = -7940;
		lnResult = -1;
		}
	else
		{
		if ((lpcSource == NULL) || (lpcSource[0] == '\0'))
			{
			strcpy(gcErrorMessage, "Source Append File not specified");
			gnLastErrorNumber = -7941;
			lnResult = -1;
			}
		else
			{
			strcpy(lcWorkAlias, d4alias(lpTargetTable));
			strncpy(lcWorkName, lpcSource, 299);
			lcWorkName[299] = (char) 0;
			StrToUpper(lcWorkName);
			lpChar = strrchr(lcWorkName, '.');
			if (lpChar != NULL)
				{
				strncpy(lcExt, lpChar + 1, 19);
				lcExt[19] = (char) 0;
				}
			if ((lpcType != NULL) && (lpcType[0] != '\0'))
				{
				// They have passed an lpcType string, so it takes precedence.
				strncpy(lcExt, lpcType, 19);
				lcExt[19] = (char) 0;
				StrToUpper(lcExt);	
				}
				
			if (strcmp(lcExt, "DBF") == 0) lnType = 1;
			else
				{
				if ((strcmp(lcExt, "SDF") == 0)	|| (strcmp(lcExt, "DAT") == 0)) lnType = 2;
				else
					{
					if (strcmp(lcExt, "CSV") == 0) lnType = 3;
					}
				}
			if (lnType == -1)
				{
				// Can't determine the type, either from the lpcType parm or from the
				// file name extension.
				strcpy(gcErrorMessage, "Source file type unidentified: >>");
				strcat(gcErrorMessage, lcExt);
				strcat(gcErrorMessage, "<<");
				gnLastErrorNumber = -7942;	
				lnResult = -1;
				}
			else
				{
				cbxSELECT(lcWorkAlias);
				// Now we get the field type information and write it into the public arrays.
				lpFld = cbxAFIELDS(&lnCnt);
				gnTrgCount = lnCnt;
				for (jj = 0; jj < lnCnt; jj++)
					{
					gaTrgFields[jj].xField = gaFields[jj];
					gaTrgFields[jj].nFldType = (long) gaFields[jj].cType[0];
					gaTrgFields[jj].bMatched = FALSE;
					gaTrgFields[jj].pCBfield = d4field(lpTargetTable, gaTrgFields[jj].xField.cName);	
					}

				switch(lnType)
					{
					case 1:
						lnResult = cbxAppendFromDBF(lpTargetTable, lpcSource, lpcTestExpr);
						break;
						
					case 2:
						lnResult = cbxAppendFromSDF(lpTargetTable, lpcSource, lpcTestExpr);
						break;
						
					case 3:
						lnResult = cbxAppendFromCSV(lpTargetTable, lpcSource, lpcTestExpr);
						break;
						
					default:
						// shouldn't ever get here.
						break;
					}	
				}
			}
		}
	gpCurrentTable = lpOldTable;
	return lnResult;
}

/* ************************************************************************************ */
/* Loads another data file in one of several types into an open DBF table.              */
static PyObject *cbwAPPENDFROM(PyObject *self, PyObject *args)
{
	char lcAlias[MAXALIASNAMESIZE];
	char lcSource[255];
	char lcTestExpr[400];
	char lcType[6];
	PyObject *lxAlias;
	PyObject *lxSource;
	PyObject *lxTestExpr;
	PyObject *lxType;
	const unsigned char *cTest;

	if (!PyArg_ParseTuple(args, "OOOO", &lxAlias, &lxSource, &lxTestExpr, &lxType))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE - 1);
	Conv1252ToASCII(lcAlias, TRUE);		

	cTest = testStringTypes(lxSource);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcSource, Unicode2Char(lxSource), 254);
	lcSource[254] = (char) 0;

	cTest = testStringTypes(lxTestExpr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTestExpr, Unicode2Char(lxTestExpr), 399);
	lcTestExpr[399] = (char) 0;
	
	cTest = testStringTypes(lxType);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcType, Unicode2Char(lxType), 5);
	lcType[5] = (char) 0;
		
	return(Py_BuildValue("l", cbxAPPENDFROM(lcAlias, lcSource, lcTestExpr, lcType)));
}

/* ********************************************************************************** */
/* Support functions for the cbwCOPYTO() function.                                    */

long cbxCopyToDBF(DATA4* lppSource, char* lpcOutput, char* lpcTestExpr)
{
	long lnResult = 0;
	long bOK;
	char lcTargAlias[50];
	char lcSrceAlias[50];
	char lcSourceDBF[350];
	register long jj;
	long kk = 0;
	
	strcpy(lcSrceAlias, cbxALIAS());
	strcpy(lcSourceDBF, cbxDBF(lcSrceAlias));

	// First we have to create the table... based on the gaSrcFields array:
	for (jj = 0; jj < gnSrcCount; jj++)
		{
		if (!gaSrcFields[jj].bMatched) continue;
		
		gaFieldSpecs[kk].name = (char*) calloc((size_t) 22, sizeof(char));
		strncpy(gaFieldSpecs[kk].name, gaSrcFields[jj].xField.cName, 20);
		StrToUpper(gaFieldSpecs[kk].name);
		gaFieldSpecs[kk].type = gaSrcFields[jj].xField.cType[0];
		gaFieldSpecs[kk].len = gaSrcFields[jj].xField.nWidth;
		gaFieldSpecs[kk].dec = gaSrcFields[jj].xField.nDecimals;
		if (gaSrcFields[jj].xField.bNulls)
			{
			gaFieldSpecs[kk].nulls = r4null;	
			}
		else
			{
			gaFieldSpecs[kk].nulls = 0;	
			}
		
		kk += 1;
		}

	bOK = cbxCREATETABLE(lpcOutput);
	
	cbxClearFieldSpecs();  // Undoes the loading of gaFieldSpecs above including the malloc().
	
	if (bOK != TRUE)
		{
		// Error already is set.
		lnResult = -1; // just in case we forgot it above.	
		}
	else
		{
		strcpy(lcTargAlias, cbxALIAS());
		if (bOK == TRUE)
			{
			lnResult = cbxAPPENDFROM(lcTargAlias, lcSourceDBF, lpcTestExpr, "DBF");
			}
		}
	cbxCLOSETABLE(lcTargAlias);
	cbxSELECT(lcSrceAlias);
	
	return lnResult;	
}

/* *************************************************************************************************** */
/* Utility function fo the cbxCopyToTextFile() function that does all the thinking about how to turn   */
/* the field data into a string for output.                                                            */
/* NOTE: Unicode and strings with embedded nulls are NOT supported.                                    */
char *cbxMakeTextOutputField(VFPFIELD* pStructFld, long lpnOutType, FIELD4* pContentFld, char* cBuff, long lpnCnt, long lpbStrip)
{
	char cWorkBuff[500];
	long nWidth;
	long nDecimals;
	long jj;
	double dValue;
	char cType;
	char cDelimiter = (char) 0;
	char cPattern[50];
	long nLen;
	long nNum;
	char cTempNum[50];
	
	nWidth = pStructFld->nWidth;
	nDecimals = pStructFld->nDecimals;
	cType = pStructFld->cType[0];
	cPattern[0] = (char) 0;
	
	*cBuff = (char) 0; // To start	
	switch(cType)
		{
		case 'C':
			strcpy(cWorkBuff, f4str(pContentFld));
			if (lpbStrip) MPStrTrim('R', cWorkBuff);
			break;
			
		case 'N':
		case 'F':
			dValue = f4double(pContentFld);
			sprintf(cPattern, "%%%ld.%ldlf", nWidth, nDecimals);
			sprintf(cWorkBuff, cPattern, dValue);
			
			if (lpnOutType != 2) MPStrTrim('A', cWorkBuff);		
			break;
		
		case 'B':
			dValue = f4double(pContentFld);
			sprintf(cTempNum, "%lf", dValue);
			if (lpnOutType == 2)
				{
				nLen = strlen(cTempNum);
				memset(cWorkBuff, 0, (size_t) (50 * sizeof(char)));
				if (nLen < 16)
					{
					
					for (jj = 0; jj < (16 - nLen); jj++)
						{
						cWorkBuff[jj] = ' ';	
						}	
					strcat(cWorkBuff, cTempNum);
					}
				else
					{
					strncpy(cWorkBuff, cTempNum, 16);
					cWorkBuff[16] = (char) 0;
					}
				}
			else
				{
				strcpy(cWorkBuff, MPStrTrim('A', cTempNum));	
				}
			break;
		
		case 'Y':
			strcpy(cTempNum, f4currency(pContentFld, 4));
			if (lpnOutType == 2)
				{
				nLen = strlen(cTempNum);
				memset(cWorkBuff, 0, (size_t) (50 * sizeof(char)));
				if (nLen < 16)
					{
					
					for (jj = 0; jj < (16 - nLen); jj++)
						{
						cWorkBuff[jj] = ' ';	
						}	
					strcat(cWorkBuff, cTempNum);
					}
				else
					{
					strncpy(cWorkBuff, cTempNum, 16); // Chop it off!
					cWorkBuff[16] = (char) 0;
					}
				}
			else
				{
				strcpy(cWorkBuff, MPStrTrim('A', cTempNum));
				}
			break;
			
		case 'T':
			strcpy(cWorkBuff, f4dateTime(pContentFld));
			break;
			
		case 'L':
			cWorkBuff[0] = 'F';
			cWorkBuff[1] = (char) 0;
			if (f4true(pContentFld)) cWorkBuff[0] = 'T';
			break;
			
		case 'D':
			nNum = f4long(pContentFld);
			if (nNum > 0)
				{
				date4assign(cWorkBuff, nNum);
				cWorkBuff[8] = (char) 0;
				}
			else strcpy(cWorkBuff, "        ");
			break;
			
		case 'I':
			nNum = f4long(pContentFld);
			sprintf(cTempNum, "%11ld", nNum);
			if (lpnOutType != 2)
				{
				strcpy(cWorkBuff, MPStrTrim('A', cTempNum));	
				}
			else
				{
				strcpy(cWorkBuff, cTempNum);	
				}
			break;
			
		default:
			break;
		}
		
	switch (lpnOutType)
		{
		case 2: // Fixed field length
			strcpy(cBuff, cWorkBuff); // straight copy, full width/
			//strcat(cBuff, "|"); // Debug ONLY.
			break;
			
		case 3: // CSV
			if (lpnCnt > 0) strcpy(cBuff, ",");
			else strcpy(cBuff, "");
			if (cType == 'C')
				{
				strcat(cBuff, "\"");
				strcat(cBuff, MPSSStrTran(cWorkBuff, "\"", "\"\"", 0));
				strcat(cBuff, "\"");
				}
			else
				{
				strcat(cBuff, cWorkBuff);	
				}
			break;
			
		case 4: // TAB
			if (lpnCnt > 0) strcpy(cBuff, "\t");
			else strcpy(cBuff, "");
			strcat(cBuff, cWorkBuff);
			break;
			
		default:
			break;
		}
	
	return cBuff;	
}

/* *************************************************************************************************** */
/* Utility function for the cbwCOPYTO() function.                                                      */
long cbxCopyToTextFile(DATA4* lppSource, char* lpcOutput, long lpnOutType, char* lpcTestExpr, long lpbHdr, long lpbStrip)
{
	long lnResult = 0;
	long lnTest = 1;
	long bNeedTest = FALSE;
	long lnFieldCnt;
	long lnOutCnt;
	long lnRowCnt = 0;
	char lcOutBuff[20000];
	FILE* pHandle;
	char* pPtr;
	char lcFldBuff[500];
	char lcWorkFlds[5000];
	register long jj;
	char cDelim[2];
	long lnType;
	
	gnLastErrorNumber = 0;
	
	memset(lcOutBuff, 0, (size_t) (20000 * sizeof(char)));
	memset(lcFldBuff, 0, (size_t) (500 * sizeof(char)));
	lcWorkFlds[0] = (char) 0;
	//printf("IN COPY TO TEXT FILE\n");
	pHandle = fopen(lpcOutput, "w");
	if (pHandle == NULL)
		{
		strcpy(gcErrorMessage, "Unable to open destination file for output.");
		gnLastErrorNumber = -4891;	
		lnResult = -1;
		}
	else
		{
		gpCurrentTable = lppSource;
		if ((lpcTestExpr != NULL) && (lpcTestExpr[0] != (char) 0))
			{
			bNeedTest = TRUE;
			lnTest = cbxPREPAREFILTER(lpcTestExpr);
			if (lnTest != 1)
				{
				strcpy(gcErrorMessage, "Invalid record selection expression");
				gnLastErrorNumber = -4892;	
				lnResult = -1;
				}
			}
		//printf("bNeed Test %ld  ErrNo %ld\n", bNeedTest, gnLastErrorNumber);
		
		if (lpbHdr && (lpnOutType != 2))
			{
			cDelim[0] = ((lpnOutType == 3) ? ',' : '\t');
			cDelim[1] = (char) 0;
			lcOutBuff[0] = (char) 0;
			for (jj = 0; jj < gnSrcCount; jj++)
				{
				if (gaSrcFields[jj].bMatched == FALSE) continue;
				if (jj > 0) strcat(lcOutBuff, cDelim);
				strcat(lcOutBuff, gaSrcFields[jj].xField.cName);	
				}
			strcat(lcOutBuff, "\n");
			fputs(lcOutBuff, pHandle);
			}
		
		if (gnLastErrorNumber == 0) // OK to continue.
			{
			cbxGOTO("TOP", 0);
			while (!cbxEOF())
				{
				if (bNeedTest)
					{
					if (cbxTESTFILTER() != 1)
						{
						if (cbxGOTO("NEXT", 0) < 1) break;
						if (cbxEOF() != 0) break;	
						continue;
						}	
					}
				lcOutBuff[0] = (char) 0;
				lnOutCnt = 0;

				for (jj = 0; jj < gnSrcCount; jj++)
					{
					if (!gaSrcFields[jj].bMatched) continue;
					lnType = gaSrcFields[jj].nFldType;
					lcFldBuff[0] = (char) 0;
					strcat(lcOutBuff, cbxMakeTextOutputField(&(gaSrcFields[jj].xField), lpnOutType, gaSrcFields[jj].pCBfield, lcFldBuff, lnOutCnt, lpbStrip));
					lnOutCnt += 1;
					}
				strcat(lcOutBuff, "\n");
				fputs(lcOutBuff, pHandle);
				lnRowCnt += 1;
				if (cbxGOTO("NEXT", 0) < 1) break;
				if (cbxEOF() != 0) break;
				}	
			lnResult = lnRowCnt;
			}	
		fclose(pHandle);	
		}
	return lnResult;	
}

/* ********************************************************************************** */
/* An output routine similar to the VFP COPY TO.  It takes any currently open table   */
/* and writes out a text file in one of three possible formats with the table data    */
/* which meets an optional selection criteria.                                        */
/* Parameters:                                                                        */
/* - cAlias -- The alias name of the open table with contents to be written out.      */
/*             If blank or null, the currently selected table will be used.           */
/* - cOutput -- The fully qualified path name of the output file.                     */
/* - cType -- One of four string values: "SDF", "CSV", "TAB", "DBF", standing for:    */
/*       SDF = System Data Format -- Fixed width fields in records delimited by CR/LF */
/*       CSV = Comma Separated Value -- Plain text with fields separated by commas    */
/*             unless they are text, in which case, they are set off with double quote*/
/*             signs thus: "joe's Diner","Hambergers To Go",35094,"Another Name".     */
/*             If the field has an embedded quote, the quote is transformed into a    */
/*             double-double quote: "Joe's ""Best"" Hambergers","Pete's Grill".       */
/*             Embedded line breaks within character/text fields are NOT supported.   */
/*             If encountered, they WILL BE REMOVED from the data prior to output.    */
/*       TAB - Tab separated values.  Plain text with fields separated by a single    */
/*             Tab (ASCII Character 9) character.  Character and number data is given */
/*             without any quotations other than what might exist in the data.  In    */
/*             this implementation embedded line breaks in character fields and em-   */
/*             bedded TAB characters in text fields are NOT supported and WILL BE     */
/*             REMOVED prior to output of the data.                                   */
/*       DBF - Standard VFP-type DBF table.  Data structure will be identical to the  */
/*             source unless a cFieldList is supplied.  No indexes will be copied.    */
/* - bHeader -- If True, then CSV and TAB outputs will include a first line consisting*/
/*             of the names of each field, separated by the appropriate separator,    */
/*             either Comma or TAB respectively.  Otherwise no field name data will   */
/*             be included.                                                           */
/* - cFieldList -- A list of field names to include in the output.  If omitted, ALL   */
/*             fields will be included (except as indicated in NOTES).                */
/* - cTestExpr -- Text string containing a VFP Logical expression which must be True  */
/*             for the record to be output.  If empty or null, ALL records are output.*/
/* - bStripBlanks -- If True, trailing blanks will be stripped at the end of records. */
/*             Defaults to False.                                                     */
/* NOTES:                                                                             */
/* - For all types except DBF, fields which cannot be represented by simple text      */
/*   strings without line breaks will be skipped.  These include types MEMO, GENERAL, */
/*   CHARACTER(BINARY), and BLOB.                                                     */
/* - Date fields will be output as MM/DD/YYYY or as specified by the current setting  */
/*   of cbwSETDATEFORMAT() (Non-DBF formats only).                                    */
/* - Date-Time fields will be represented as CCYYMMDDhh:mm:ss, which is re-importable */
/*   back into CodeBase DBF tables via cbwAPPENDFROM(). (Non-DBF formats only).       */
/* - The Deleted status of records will be respected based on the cbwSETDELETED()     */
/*   setting.                                                                         */
/* - Current Order will be respected.                                                 */
/* - The TYPE of the file will be determined by the cType parameter regardless of the */
/*   name of the output file!!                                                        */
/* - Field lengths for the SDF type output are, in general determined by the length   */
/*   of the data fields themselves.  Rules for length and output are as follows by    */
/*   type:                                                                            */
/*    - C - Width of field.                                                           */
/*    - D - 8 characters in the form YYYYMMDD                                         */
/*    - T - 16 characters in the form YYYYMMDDhh:mm:ss                                */
/*    - L - 1 character ('F' or 'T')                                                  */
/*    - N - Determined by size and decimals of field.                                 */
/*    - F (Float) - like N                                                            */
/*    - B (Double) - 16 bytes, left-filled                                            */
/*    - I (Integer) - 11 bytes, left-filled                                           */
/*    - Y (Currency) - 16 bytes, 4 decimals, left filled.                             */
/* RETURNS:                                                                           */
/* Returns number of records output or -1 on error.                                   */
//DOEXPORT long cbwCOPYTO(char* lpcAlias, char* lpcOutput, char* cType, long bHeader, char* cFieldList, char* cTestExpr, long bStripBlanks)
static PyObject *cbwCOPYTO(PyObject *self, PyObject *args)
{
	long lnResult = 0;	// Number of records output or -1 on error.
	DATA4 *lpSourceTable = NULL;
	char lcWorkName[300];
	long lnPt = 0;
	long lnType = -1; // Types 1 = DBF, 2 = SDF or DAT, 3 = CSV, 4 = TAB
	char *lpChar = NULL;
	char *lpTest = NULL;
	char *pPtr = NULL;
	long lbByType = FALSE;
	char lcWorkAlias[MAXALIASNAMESIZE + 1];
	long lnOutFldCnt = 0;
	long bFieldTest = FALSE;
	VFPFIELD* lpFld;
	long lnCnt = 0;
	long jj;
	DATA4 *lpOldTable = NULL;
	const unsigned char *cTest;
	char lcFldMatch[5000];
	char lcFldTest[25];
	long bStripBlanks;
	long bHeader;

	char lcAlias[MAXALIASNAMESIZE + 1];
	char lcOutput[255];	
	char lcType[6];		
	char lcFieldList[3000];
	char lcTestExpr[500];

	PyObject *lxAlias;
	PyObject *lxOutput;
	PyObject *lxType;
	PyObject *lxFieldList;
	PyObject *lxTestExpr;
	
	if (!PyArg_ParseTuple(args, "OOOlOOl", &lxAlias, &lxOutput, &lxType, &bHeader, &lxFieldList, &lxTestExpr, &bStripBlanks))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}

	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);		
	
	cTest = testStringTypes(lxOutput);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcOutput, Unicode2Char(lxOutput), 254);
	lcOutput[254] = (char) 0;
		
	cTest = testStringTypes(lxType);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcType, Unicode2Char(lxType), 5);
	lcType[5] = (char) 0;
	Conv1252ToASCII(lcType, TRUE);		

	cTest = testStringTypes(lxFieldList);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldList, Unicode2Char(lxFieldList), 2999);
	lcFieldList[2999] = (char) 0;
	
	cTest = testStringTypes(lxTestExpr);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcTestExpr, Unicode2Char(lxTestExpr), 498);
	lcTestExpr[498] = (char) 0;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpOldTable = gpCurrentTable;
		}
	
	if ((lcAlias == NULL) || (lcAlias[0] == '\0'))
		{
		lpSourceTable = gpCurrentTable;	
		}
	else
		{
		lpSourceTable = code4data(&codeBase, lcAlias);	
		}
		
	if (lpSourceTable == NULL)
		{
		strcpy(gcErrorMessage, "No Source Table Available or Alias Not Found");
		gnLastErrorNumber = -7980;
		lnResult = -1;
		}
	else
		{
		if ((lcOutput == NULL) || (lcOutput[0] == '\0'))
			{
			strcpy(gcErrorMessage, "Output File Name not specified");
			gnLastErrorNumber = -7981;
			lnResult = -1;
			}
		else
			{
			strncpy(lcWorkAlias, d4alias(lpSourceTable), MAXALIASNAMESIZE);
			lcWorkAlias[MAXALIASNAMESIZE] = (char) 0;
			cbxSELECT(lcWorkAlias);
			while (TRUE)
				{
				if (strcmp(lcType, "DBF") == 0)
					{
					lnType = 1;
					break;
					}
				
				if (strcmp(lcType, "SDF") == 0)
					{
					lnType = 2;
					break;
					}
				
				if (strcmp(lcType, "DAT") == 0) {lnType = 2; break;}
				
				if (strcmp(lcType, "CSV") == 0) {lnType = 3; break;}
				
				if (strcmp(lcType, "TAB") == 0) {lnType = 4; break;}
				
				break;
				}
			}
		if (lnType != -1)
			{			
			// Now we get the field type information and write it into the public arrays.
			if ((lcFieldList != NULL) && (lcFieldList[0] != (char) 0))
				{
				strcpy(lcFldMatch, ",");
				strcat(lcFldMatch, lcFieldList);
				strcat(lcFldMatch, ","); // Ensures we can search for ',MYFIELD,' in this string.
				StrToUpper(lcFldMatch); // In place uppercase.
				bFieldTest = TRUE;
				}			
			lcFldTest[0] = (char) 0;
			
			lpFld = cbxAFIELDS(&lnCnt);
			gnSrcCount = lnCnt;
			for (jj = 0; jj < lnCnt; jj++)
				{
				gaSrcFields[jj].xField = gaFields[jj];
				gaSrcFields[jj].nFldType = (long) gaFields[jj].cType[0];
				gaSrcFields[jj].bMatched = TRUE;
				if (bFieldTest)
					{
					strcpy(lcFldTest, ",");
					strcat(lcFldTest, gaSrcFields[jj].xField.cName);
					strcat(lcFldTest, ",");
					StrToUpper(lcFldTest);
					pPtr = strstr(lcFldMatch, lcFldTest);
					if (pPtr == NULL) gaSrcFields[jj].bMatched = FALSE;					
					}
				if (lnType > 1)
					{
					// A text type output doesn't allow memo and general type fields.
					if (strchr("MGXZW", gaSrcFields[jj].nFldType)) gaSrcFields[jj].bMatched = FALSE;
					//if (gaSrcFields[jj].bMatched == FALSE) printf("Skipped: %s  Type %s \n", gaSrcFields[jj].xField.cName, gaSrcFields[jj].xField.cType);	
					}
				if (gaSrcFields[jj].bMatched) lnOutFldCnt += 1;				
				gaSrcFields[jj].pCBfield = d4field(lpSourceTable, gaSrcFields[jj].xField.cName);
				}
			if (lnOutFldCnt == 0)
				{
				strcpy(gcErrorMessage, "No Output Fields in List found in Source Table.");
				gnLastErrorNumber = -7983;	
				lnResult = -1;				
				}
			else
				{
				//printf("outfldcount %ld \n", lnOutFldCnt);
				switch(lnType)
					{
					case 1:
						lnResult = cbxCopyToDBF(lpSourceTable, lcOutput, lcTestExpr);
						break;
	
					case 2:
						lnResult = cbxCopyToTextFile(lpSourceTable, lcOutput, 2, lcTestExpr, FALSE, FALSE); // Fixed sizes, so no headers or blank stripping.
						break;
						
					case 3:
						lnResult = cbxCopyToTextFile(lpSourceTable, lcOutput, 3, lcTestExpr, bHeader, bStripBlanks);
						break;
						
					case 4:
						lnResult = cbxCopyToTextFile(lpSourceTable, lcOutput, 4, lcTestExpr, bHeader, bStripBlanks);
						
					default:
						// shouldn't ever get here.
						break;
					}
				}	
			}
		else
			{
			strcpy(gcErrorMessage, "Unrecognized File Output Type");
			gnLastErrorNumber = -7982;
			lnResult = -1;
			}
		}
	gpCurrentTable = lpOldTable;			

	//return lnResult;	
	return Py_BuildValue("l", lnResult);
}

/* *********************************************************************************** */
/* Rough analog to INSERT FROM SQL code in VFP.  This takes the same lpcValue string   */
/* as cbwGATHERMEMVAR() but it MUST include values for ALL fields.  If there is a mis- */
/* match between the number of fields passed and the number of fields in the table, it */
/* triggers an error condition and the function returns -1.  For general fields you    */
/* must include an empty string as the value.  General fields are not updated by this  */
/* function, but memo fields are.  As with cbwGATHERMEMVAR() you may optionally define */
/* the delimiter characters used to separate field data in the value string. "~~" is   */
/* the default delimiter string.                                                       */
/* On Feb. 24, added the lpcFldNames parameter which is an exact analog to the same    */
/* named field for the cbwGATHERMEMVAR() function which is used here.  This allows the */
/* insert to use strings which may have the same fields, but those fields may not be in*/
/* the same order in the target table.                                                 */
/*                                                                                     */
/* If all the data is available and correct, the function starts the append process,   */
/* executes cbxGATHERMEMVAR() and, if it is successful, then completes the append      */
/* process.  At this point a new record exists at the end of the current table, but    */
/* flush will not yet have been called.  The current record pointer will be positioned */
/* on the new record, so individual field-by-field data changes can still be made to   */
/* the record.  The actual data may or may not have been physically written to the     */
/* disk.  To force a write (time consuming and not recommended for lots of sequential  */
/* appends until the end of the process), issue a cbwFLUSH() or cbwFLUSHALL() command. */
/* In this latest version accepts an alias name into which the insertion should be     */
/* performed.                                                                          */
static PyObject *cbwINSERT(PyObject *self, PyObject *args)
{
	long lnResult;
	long lnReturn = TRUE;
	char lcAlias[MAXALIASNAMESIZE + 1];
	char lcValues[65000];
	char lcFldNames[5000];
	char lcDelimiter[30];
	DATA4 *lpTable;
	PyObject *lxAlias;
	PyObject *lxValues;
	PyObject *lxFldNames;
	PyObject *lxDelimiter;	
	const unsigned char *cTest;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (!PyArg_ParseTuple(args, "OOOO", &lxAlias, &lxValues, &lxFldNames, &lxDelimiter))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);		
		return NULL;
		}

	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);
	
	cTest = testStringTypes(lxValues);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcValues, Unicode2Char(lxValues), 64999);
	lcValues[64999] = (char) 0;
	
	cTest = testStringTypes(lxFldNames);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFldNames, Unicode2Char(lxFldNames), 4999);
	lcFldNames[4999] = (char) 0;
	
	cTest = testStringTypes(lxDelimiter);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcDelimiter, Unicode2Char(lxDelimiter), 29);
	lcDelimiter[29] = (char) 0;
	Conv1252ToASCII(lcDelimiter, TRUE);	

	if (strlen(lcAlias) == 0)
		{
		if (gpCurrentTable)
			{
			lpTable = gpCurrentTable;
			}
		else
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "No Table is Selected, Alias Not Found");	
			gnLastErrorNumber = -9994;
			}
		}
	else
		{
		lpTable = code4data(&codeBase, lcAlias);
		if (lpTable == NULL)
			{
			gnLastErrorNumber = codeBase.errorCode;
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			lnReturn = FALSE;				
			}
		}
	
	if (lpTable != NULL)
		{
		lnResult = d4appendStart(lpTable, 0); /* no memo content carryover from the current record */
		if (lnResult == 0)
			{
			d4blank(lpTable); // Clear the buffer for a brand new record.
			lnResult = cbxGATHERMEMVAR(lpTable, lcFldNames, lcValues, lcDelimiter);
			if (lnResult == TRUE)
				{
				lnResult = d4append(lpTable);
				d4recall(lpTable); // Seems to be required sometimes for the above.
				if (lnResult != 0)
					{
					if (lnResult > 0)
						{
						strcpy(gcErrorMessage, "Unable to Move Record Pointer");
						gnLastErrorNumber = -9980;
						lnReturn = FALSE; /* One of the several locking and duplicate key errors.  In any event, FAILURE. */	
						}
					else
						{
						gnLastErrorNumber = codeBase.errorCode;
						strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
						strcat(gcErrorMessage, " ERR from CodeBase");
						lnReturn = FALSE; /* We're done... ERROR */					
						}						
					}	
				}
			else
				{
				lnReturn = FALSE; /* The function cbwGATHERMEMVAR() has already set the error messages */
				strcat(gcErrorMessage, " From Gather Memvar");
				}	
			}
		else
			{
			if (lnResult > 0)
				{
				strcpy(gcErrorMessage, "Unable to Move Record Pointer");
				gnLastErrorNumber = -9980;
				lnReturn = FALSE; /* One of the several locking and duplicate key errors.  In any event, FAILURE. */	
				}
			else
				{
				gnLastErrorNumber = codeBase.errorCode;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				strcat(gcErrorMessage, " Code Base Error");
				lnReturn = FALSE; /* We're done... ERROR */					
				}				
			}	
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}

	//return(lnReturn);	
	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));
}


/* ***************************************************************************** */
/* Equivalent of DELETE which applies to the currently selected table.  Just     */
/* deletes the current record.  Returns TRUE (1) if successful, or -1 if         */
/* there is no current record or other problem.                                  */
static PyObject *cbwDELETE(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gnProcessTally = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4recNo(gpCurrentTable);
		if (lnResult < 1)
			{
			/* No current record to delete */
			strcpy(gcErrorMessage, "No current record to delete");
			gnLastErrorNumber = -9990;
			lnReturn = FALSE;	
			}
		else
			{
			d4delete(gpCurrentTable);
			gnProcessTally = 1;
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));	
}

/* ***************************************************************************** */
/* Equivalent of RLOCK OR LOCK which applies to the currently selected table.    */
/* Just locks the current record.  Returns TRUE (1) if successful, or -1 if      */
/* there is no current record or other problem.                                  */
static PyObject *cbwLOCK(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	long lnRecord;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gnProcessTally = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnRecord = d4recNo(gpCurrentTable);
		if (lnRecord < 0)
			{
			/* No current record to lock */
			strcpy(gcErrorMessage, "No current record to lock");
			gnLastErrorNumber = -9990;
			lnReturn = FALSE;	
			}
		else
			{
			lnResult = d4lock(gpCurrentTable, lnRecord); 
			if (lnResult != 0)
				{
				lnReturn = FALSE;
				strcpy(gcErrorMessage, "Unable to lock");
				gnLastErrorNumber = -9998;
				}
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}

	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));
}

/* ***************************************************************************** */
/* Equivalent of FLOCK in VFP, which locks the entire table, preventing changes  */
/* by others, but supporting read access without problems.                       */
/* Returns TRUE, if successful, or FALSE if there is no current record or other  */
/* problem.                                                                      */
static PyObject *cbwFLOCK(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	long lnRecord;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gnProcessTally = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4lockFile(gpCurrentTable); 
		if (lnResult != 0)
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "Unable to lock entire Table");
			gnLastErrorNumber = -99991;
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}

	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));
}

/* ***************************************************************************** */
/* Equivalent of UNLOCK which applies to the currently selected table.           */
/* Unlocks ALL RECORDS.  Returns True if successful, or False on ERROR.         */
static PyObject *cbwUNLOCK(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gnProcessTally = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4unlock(gpCurrentTable); /* OK, even if not really deleted */
		if (lnResult != 0)
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "Unable to unlock");
			gnLastErrorNumber = -9997;
			}
		}
	else
		{
		lnReturn = FALSE;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));	
}

/* ************************************************************************** */
/* Function equivalent to VFP PACK Command on the current table.  Removes     */
/*   all deleted records.  Returns True on success, False on failure.         */
static PyObject *cbwPACK(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
		
	if (gpCurrentTable == NULL)
		{
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		lnReturn = FALSE;	
		}
	else
		{
		lnResult = d4pack(gpCurrentTable);
		if (lnResult != 0)
			{
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode;
			lnReturn = FALSE;	
			}
		else
			{
			lnResult = d4memoCompress(gpCurrentTable);
			if (lnResult == 0)
				{
				if (d4recCount(gpCurrentTable) > 0) d4top(gpCurrentTable);	
				}
			else
				{
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;
				lnReturn = FALSE;
				}	
			}	
		}
	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));	
}

/* *************************************************************************** */
/* Like the VFP REINDEX command, reindexes all CDX index tags in the currently
   selected data table.  Returns 1 on OK, 0 on failure for any reason.         */
static PyObject *cbwREINDEX(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
		
	if (gpCurrentTable == NULL)
		{
		strcpy(gcErrorMessage, "No Current Table Selected");
		lnReturn = FALSE;
		gnLastErrorNumber = -9999;	
		}
	else
		{
		lnResult = d4reindex(gpCurrentTable);
		if (lnResult != 0)
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage,error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode; 	
			}	
		}
	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));	
}

/* ************************************************************************** */
/* Function equivalent to VFP ZAP Command on the current table.  Removes      */
/*   ALL records.  Returns 1 on success, 0 on failure.                        */
static PyObject *cbwZAP(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gnProcessTally = 0;
	
	if (gpCurrentTable == NULL)
		{
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		lnReturn = FALSE;	
		}
	else
		{
		lnResult = d4zap(gpCurrentTable, 1, d4recCount(gpCurrentTable));
		if (lnResult != 0)
			{
			strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
			gnLastErrorNumber = codeBase.errorCode;
			lnReturn = FALSE;	
			}
		else
			{
			lnResult = d4memoCompress(gpCurrentTable);
			/* We don't report an error here.  The data is orphaned anyway at this point. */	
			}
		}
	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));	
}
/* ******************************************************************************** */
/* Tests the current READONLY status of the currently selected table or the one     */
/* specified by the alias given.                                                    */
/* Returns 1 TRUE for IS readonly, 0 FALSE for NOT readonly, and -1 on error.       */
static PyObject *cbwISREADONLY(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	DATA4 *lpTable = NULL;
	char lcAlias[MAXALIASNAMESIZE + 1];
	PyObject *lxAlias;
	const unsigned char *cTest;

	if (!PyArg_ParseTuple(args, "O", &lxAlias))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxAlias);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcAlias, Unicode2Char(lxAlias), MAXALIASNAMESIZE);
	lcAlias[MAXALIASNAMESIZE] = (char) 0;
	Conv1252ToASCII(lcAlias, TRUE);	

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gcStrReturn[0] = (char) 0;
		
	if (strlen(lcAlias) > 0)
		{
		lpTable = code4data(&codeBase, lcAlias);	
		}
	else
		{
		if (gpCurrentTable == NULL)
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, "No Table Selected");
			gnLastErrorNumber = -99999;
			}
		else
			{
			lpTable = gpCurrentTable;
			}	
		}
	if (lpTable != NULL)
		{
		lnReturn = lpTable->readOnly;
		}
	else
		{
		strcpy(gcErrorMessage, "Alias Name Not Found");
		gnLastErrorNumber = -99999;
		lnReturn = -1;
		}
	if (lnReturn == -1)
		{
		return Py_BuildValue(""); // Returns None
		}
	else
		{
		return Py_BuildValue("N", PyBool_FromLong(lnReturn));
		}
}

/* ************************************************************************************* */
/* Exported function that takes a field name (with an optional alias prefix) and returns */
/* a tuple with five elements: Single character string indicating the type, Integer      */
/* giving the length, Integer givung the decimals count (if any), a Bool indicating      */
/* if nulls are allowed, a Bool indicating if this a binary field.                       */
static PyObject *cbwFIELDINFO(PyObject *self, PyObject *args)
{
	PyObject *lxTuple;
	char lcFieldName[MAXALIASNAMESIZE + 30];
	FIELD4 *lpField;
	PyObject *lxFieldName;
	PyObject *lxType;
	const unsigned char *cTest;
	long nFldLen;
	long nFldDec;
	char cFldType[3];
	long bNulls;
	long bBinary;
	
	if (!PyArg_ParseTuple(args, "O", &lxFieldName))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 29);
	lcFieldName[MAXALIASNAMESIZE + 29] = (char) 0;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	lpField = cbxGetFieldFromName(lcFieldName);		
	if (lpField == NULL)
		{
		strcpy(gcErrorMessage, "Field not recognized: ");
		strcat(gcErrorMessage, lcFieldName);
		gnLastErrorNumber = -9999;
		return Py_BuildValue(""); // Returns None
		}
	lxTuple = PyTuple_New((Py_ssize_t) 5);
	cFldType[0] = (char) f4type(lpField);
	cFldType[1] = (char) 0;
	
#ifdef IS_PY3K	
				lxType = PyUnicode_FromString(cFldType);
#endif
#ifdef IS_PY2K
				lxType = PyString_FromString(cFldType);
#endif	
	PyTuple_SetItem(lxTuple, 0, lxType);  // no need to decrement ref.
	
	nFldLen = f4len(lpField);
	nFldDec = f4decimals(lpField);
	bNulls = (long) lpField->null;
	bBinary = (long) lpField->binary;
	if (bBinary != 1) bBinary = 0;
	PyTuple_SetItem(lxTuple, 1, Py_BuildValue("l", nFldLen));  
	PyTuple_SetItem(lxTuple, 2, Py_BuildValue("l", nFldDec));  
	PyTuple_SetItem(lxTuple, 3, PyBool_FromLong(bNulls));  
	PyTuple_SetItem(lxTuple, 4, PyBool_FromLong(bBinary));
	return(lxTuple);
}

///* ********************************************************************************** */
///* One of a set of functions that works as the REPLACE command in VFP but for just    */
///* one field in the currently selected table.  Each one is geared to one type of      */
///* field content.                                                                     */
///* Returns 1 on success, -1 on ERROR.                                                 */
///* This takes the date value as a julian day number as recognized by VFP.  The julian */
///* day number of Jan. 1, 1990 is 2447893, for example.                                */
///* If the value passed is < 0 or otherwise invalid, the result is undefined.  A value */
///* of 0 results in an empty date field.                                               */
//DOEXPORT long cbwREPLACE_DATEN(char* lpcFieldName, long lpnJulian)
//{
//	long lnReturn = TRUE;
//	long lnResult = 0;
//	FIELD4 *lpField;
//	long lnLongValue;
//			
//	codeBase.errorCode = 0;
//	gcErrorMessage[0] = (char) 0;
//	gnLastErrorNumber = 0;
// 
//	if (gpCurrentTable != NULL)
//		{
//		lpField = cbxGetFieldFromName(lpcFieldName);
//		if (lpField == NULL)
//			{
//			lnReturn = -1; /* Error messages are already set */
//			}
//		else
//			{
//			if (f4type(lpField) != 'D')
//				{
//				lnReturn = -1;
//				strcpy(gcErrorMessage, "Not a Date Field");
//				gnLastErrorNumber = -8025;
//				}
//			else
//				{
//				f4assignLong(lpField, lpnJulian); /* Can be 0 on an empty date which is valid in a VFP table. */
//				}
//			}			
//		}
//	else
//		{
//		lnReturn = -1;
//		strcpy(gcErrorMessage, "No Table Open in Selected Area");
//		gnLastErrorNumber = -9999;			
//		}
//	return(lnReturn);		
//}
//
//
///* ********************************************************************************** */
///* One of a set of functions that works as the REPLACE command in VFP but for just    */
///* one field in the currently selected table.  Each one is geared to one type of      */
///* field content.                                                                     */
///* Returns 1 on success, -1 on ERROR.                                                 */
///* This takes the date value as a character string in the form:                       */
///* YYYYMMDD                                                                           */
///* IF THE VALUE PASSED IS NOT A VALID DATE REPRESENTATION, AN ERROR IS REPORTED.      */
//DOEXPORT long cbwREPLACE_DATEC(char* lpcFieldName, char* lpcValue)
//{
//	long lnReturn = TRUE;
//	long lnResult = 0;
//	FIELD4 *lpField;
//	long lnLongValue;
//			
//	codeBase.errorCode = 0;
//	gcErrorMessage[0] = (char) 0;
//	gnLastErrorNumber = 0;
//
// 
//	if (gpCurrentTable != NULL)
//		{
//		lpField = cbxGetFieldFromName(lpcFieldName);
//		if (lpField == NULL)
//			{
//			lnReturn = -1; /* Error messages are already set */
//			}
//		else
//			{
//			if (strcmp(lpcValue, "        ") == 0)
//				{
//				lnLongValue = 0;
//				f4assignLong(lpField, lnLongValue);
//				}
//			else
//				{
//				lnLongValue = date4long(lpcValue);
//				if (lnLongValue < 0)
//					{
//					/* An error condition -- bad date format */
//					strcpy(gcErrorMessage, "Bad Date Format");
//					gnLastErrorNumber = -8000;
//					lnReturn = -1;	
//					}
//				else
//					{
//					if (f4type(lpField) != 'D')
//						{
//						strcpy(gcErrorMessage, "Not a Date Field");
//						gnLastErrorNumber = -8025;
//						lnReturn = -1;	
//						}
//					else
//						{
//						f4assignLong(lpField, lnLongValue); /* Can be 0 on an empty date which is valid in a VFP table. */	
//						}
//					}
//				}	
//			}			
//		}
//	else
//		{
//		lnReturn = -1;
//		strcpy(gcErrorMessage, "No Table Open in Selected Area");
//		gnLastErrorNumber = -9999;			
//		}
//	return(lnReturn);		
//}
//
//
///* ********************************************************************************** */
///* One of a set of functions that works as the REPLACE command in VFP but for just    */
///* one field in the currently selected table.  Each one is geared to one type of      */
///* field content.                                                                     */
///* Returns 1 on success, -1 on ERROR.                                                 */
///* This takes the datetime value as a character string in the form:                   */
///* YYYYMMDDHHMMSS                                                                     */
///* IF THE VALUE PASSED IS NOT A VALID DATETIME REPRESENTATION, THE FIELD IS SET TO    */
///* THE VALUE OF AN EMPTY DATETIME AND NO ERROR IS REPORTED.  PLAN ACCORDINGLY!        */
//DOEXPORT long cbwREPLACE_DATETIMEC(char* lpcFieldName, char* lpcValue)
//{
//	long lnReturn = TRUE;
//	long lnResult = 0;
//	FIELD4 *lpField;
//	char lcDTBuff[20];
//		
//	codeBase.errorCode = 0;
//	gcErrorMessage[0] = (char) 0;
//	gnLastErrorNumber = 0;
//	strncpy(lcDTBuff, lpcValue, 10);
//	lcDTBuff[10] = ':';
//	strncpy(lcDTBuff + 11, lpcValue + 10, 2);
//	lcDTBuff[13] = ':';
//	strncpy(lcDTBuff + 14, lpcValue + 12, 2);
//	lcDTBuff[16] = (char) 0;
// 
//	if (gpCurrentTable != NULL)
//		{
//		lpField = cbxGetFieldFromName(lpcFieldName);
//		if (lpField == NULL)
//			{
//			lnReturn = -1; /* Error messages are already set */
//			}
//		else
//			{
//			/* Ok, we have a valid field in the current table */
//			f4assignDateTime(lpField, lcDTBuff);	
//			}			
//		}
//	else
//		{
//		lnReturn = -1;
//		strcpy(gcErrorMessage, "No Table Open in Selected Area");
//		gnLastErrorNumber = -9999;			
//		}
//	return(lnReturn);		
//}

/* ********************************************************************************** */
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
/* This takes the datetime value as a set of 6 long values representing the Year,     */
/* Month, Day, Hour, Minute, and Second.  If these values are not a valid represen-   */
/* tation of a datetime value, an ERROR is generated (return value of -1).            */
static PyObject *cbwREPLACE_DATETIMEN(PyObject *self, PyObject *args)
{
	char lcFieldName[MAXALIASNAMESIZE + 30];
	FIELD4 *lpField;
	PyObject *lxFieldName;
	const unsigned char *cTest;
	long lpnYr;
	long lpnMo;
	long lpnDa;
	long lpnHr;
	long lpnMi;
	long lpnSe;
	char lcDTBuff[20];
	long lnReturn = 1;
	
	if (!PyArg_ParseTuple(args, "Ollllll", &lxFieldName, &lpnYr, &lpnMo, &lpnDa, &lpnHr, &lpnMi, &lpnSe))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 29);
	lcFieldName[MAXALIASNAMESIZE + 29] = (char) 0;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	if ((lpnYr < 10) || (lpnYr > 4000) || (lpnMo < 1) || (lpnMo > 12) || 
		(lpnDa < 1) || (lpnDa > 31) || (lpnHr < 0) || (lpnHr > 24) ||
		(lpnMi < 0) || (lpnMi > 59) || (lpnSe < 0) || (lpnSe > 59))
		{
		lnReturn = -1;
		gnLastErrorNumber = -8059;
		strcpy(gcErrorMessage, "Invalid Datetime input values");	
		}
 	else
 		{
		lpField = cbxGetFieldFromName(lcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			/* Ok, we have a valid field in the current table */
			sprintf(lcDTBuff, "%04ld%02ld%02ld%02ld:%02ld:%02ld", lpnYr, lpnMo, lpnDa, lpnHr, lpnMi, lpnSe);
			f4assignDateTime(lpField, lcDTBuff);	
			}
		}
	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));		
}

///* ********************************************************************************** */
///* One of a set of functions that works as the REPLACE command in VFP but for just    */
///* one field in the currently selected table.  Each one is geared to one type of      */
///* field content.                                                                     */
///* Returns 1 on success, -1 on ERROR.                                                 */
//DOEXPORT long cbwREPLACE_CURRENCY(char* lpcFieldName, char* lpcValText)
//{
//	long lnReturn = TRUE;
//	long lnResult = 0;
//	FIELD4 *lpField;
//		
//	codeBase.errorCode = 0;
//	gcErrorMessage[0] = (char) 0;
//	gnLastErrorNumber = 0;
//	
//	if (gpCurrentTable != NULL)
//		{
//		lpField = cbxGetFieldFromName(lpcFieldName);
//		if (lpField == NULL)
//			{
//			lnReturn = -1; /* Error messages are already set */
//			}
//		else
//			{
//			if (f4type(lpField) != 'Y')
//				{
//				strcpy(gcErrorMessage, "Not a Currency Field");
//				gnLastErrorNumber = -8025;
//				lnReturn = -1;
//				}
//			else
//				{	
//				/* Ok, we have a valid field in the current table */
//				f4assignCurrency(lpField, lpcValText);
//				}	
//			}			
//		}
//	else
//		{
//		lnReturn = -1;
//		strcpy(gcErrorMessage, "No Table Open in Selected Area");
//		gnLastErrorNumber = -9999;			
//		}
//	return(lnReturn);		
//}
//
/* ********************************************************************************** */
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
static PyObject *cbwREPLACE_LONG(PyObject *self, PyObject *args)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
	char lcFieldName[MAXALIASNAMESIZE + 30];
	PyObject *lxFieldName;
	long lpnValue;
	const unsigned char *cTest;
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	if (!PyArg_ParseTuple(args, "Ol", &lxFieldName, &lpnValue))
		{
		strcpy(gcErrorMessage, "Bad or missing parameter passed");
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;
		}

	cTest = testStringTypes(lxFieldName);
	if (*cTest != (const unsigned char) 0)
		{
		strcpy(gcErrorMessage, cTest);
		gnLastErrorNumber = -10000;
		PyErr_Format(PyExc_ValueError, gcErrorMessage);
		return NULL;			
		}
	strncpy(lcFieldName, Unicode2Char(lxFieldName), MAXALIASNAMESIZE + 29);
	lcFieldName[MAXALIASNAMESIZE + 29] = (char) 0;

	lpField = cbxGetFieldFromName(lcFieldName);
	if (lpField != NULL)
		{
		/* Ok, we have a valid field in the current table */
		f4assignLong(lpField, lpnValue); /* Allowed to update integer and number fields */				
		}
	else
		{
		lnReturn = FALSE; /* Error messages are already set */
		}
	return(Py_BuildValue("N", PyBool_FromLong(lnReturn)));		
}
//
///* ********************************************************************************** */
///* One of a set of functions that works as the REPLACE command in VFP but for just    */
///* one field in the currently selected table.  Each one is geared to one type of      */
///* field content.                                                                     */
///* Returns 1 on success, -1 on ERROR.                                                 */
//DOEXPORT long cbwREPLACE_DOUBLE(char* lpcFieldName, double lpnValue)
//{
//	long lnReturn = TRUE;
//	long lnResult = 0;
//	FIELD4 *lpField;
//		
//	codeBase.errorCode = 0;
//	gcErrorMessage[0] = (char) 0;
//	gnLastErrorNumber = 0;
//	
//	if (gpCurrentTable != NULL)
//		{
//		lpField = cbxGetFieldFromName(lpcFieldName);
//		if (lpField == NULL)
//			{
//			lnReturn = -1; /* Error messages are already set */
//			}
//		else
//			{
//			/* Ok, we have a valid field in the current table */
//			f4assignDouble(lpField, lpnValue);	
//			}			
//		}
//	else
//		{
//		lnReturn = -1;
//		strcpy(gcErrorMessage, "No Table Open in Selected Area");
//		gnLastErrorNumber = -9999;			
//		}
//	return(lnReturn);		
//}
//
///* ********************************************************************************** */
///* One of a set of functions that works as the REPLACE command in VFP but for just    */
///* one field in the currently selected table.  Each one is geared to one type of      */
///* field content.                                                                     */
///* Returns 1 on success, -1 on ERROR.                                                 */
//DOEXPORT long cbwREPLACE_LOGICAL(char* lpcFieldName, long lpnValue)
//{
//	long lnReturn = TRUE;
//	long lnResult = 0;
//	FIELD4 *lpField;
//	char lcFlag;
//		
//	codeBase.errorCode = 0;
//	gcErrorMessage[0] = (char) 0;
//	gnLastErrorNumber = 0;
//	
//	if (gpCurrentTable != NULL)
//		{
//		lpField = cbxGetFieldFromName(lpcFieldName);
//		if (lpField == NULL)
//			{
//			lnReturn = -1; /* Error messages are already set */
//			}
//		else
//			{
//			/* Ok, we have a valid field in the current table */
//			lcFlag = (lpnValue == 1 ? 'T' : 'F');
//			f4assignChar(lpField, (int) lcFlag);			
//			}			
//		}
//	else
//		{
//		lnReturn = -1;
//		strcpy(gcErrorMessage, "No Table Open in Selected Area");
//		gnLastErrorNumber = -9999;			
//		}
//	return(lnReturn);		
//}
//
///* ********************************************************************************** */
///* One of a set of functions that works as the REPLACE command in VFP but for just    */
///* one field in the currently selected table.  Each one is geared to one type of      */
///* field content.                                                                     */
///* Returns 1 on success, -1 on ERROR.                                                 */
//DOEXPORT long cbwREPLACE_MEMO(char* lpcFieldName, char* lpcValue)
//{
//	long lnReturn = TRUE;
//	long lnResult = 0;
//	FIELD4 *lpField;
//		
//	codeBase.errorCode = 0;
//	gcErrorMessage[0] = (char) 0;
//	gnLastErrorNumber = 0;
//	
//	if (gpCurrentTable != NULL)
//		{
//		lpField = cbxGetFieldFromName(lpcFieldName);
//		if (lpField == NULL)
//			{
//			lnReturn = -1; /* Error messages are already set */
//			}
//		else
//			{
//			/* Ok, we have a valid field in the current table */
//			f4memoAssign(lpField, lpcValue);	
//			}			
//		}
//	else
//		{
//		lnReturn = -1;
//		strcpy(gcErrorMessage, "No Table Open in Selected Area");
//		gnLastErrorNumber = -9999;			
//		}
//	return(lnReturn);		
//}
//

/******* METHOD TABLE *******/

/* The PyMethodDef array is the Method Table (http://www.python.org/doc/current/ext/methodTable.html).
   The Method Table must have an entry for every function
   that is to be called by the Python interpretor. */
static PyMethodDef cbMethods[] =
{
   { "setdateformat", cbwSETDATEFORMAT, METH_VARARGS, "Sets the Date Format" },
   { "getdateformat", cbwGETDATEFORMAT, METH_NOARGS, "Returns the Date Format" },
   { "initdatasession", cbwINITDATASESSION, METH_VARARGS, "Initializes an Available Data Session" },
   { "switchdatasession", cbwSWITCHDATASESSION, METH_VARARGS, "Switch to Another Existing Data Session" },
   { "closedatasession", cbwCLOSEDATASESSION, METH_VARARGS, "Close an Active Data Session" },
   { "geterrormessage", cbwERRORMSG, METH_NOARGS, "Retrieve the Last Error Message" },
   { "geterrornumber", cbwERRORNUM, METH_NOARGS, "Retrieve the Last Error Number" },
   { "getcurrentsession", cbwGETSESSIONNUMBER, METH_VARARGS, "Retrieve the Current Data Session Index" },
   { "getlargemode", cbwISLARGEMODE, METH_NOARGS, "Returns Status of Large Table Mode" },
   { "setdeleted", cbwSETDELETED, METH_VARARGS, "Set and Retrieve Deleted Status" },
   { "use", cbwUSE, METH_VARARGS, "Open/Use a Table" }, 
   { "alias", cbwALIAS, METH_NOARGS, "Retrieve the Alias of the Current Table" },
   { "dbf", cbwDBF, METH_VARARGS, "Retrieve the Full Path Name of the Current Table" },
   { "useexcl", cbwUSEEXCL, METH_VARARGS, "Open/Use a Table Exclusively" },
   { "isexclusive", cbwISEXCL, METH_NOARGS, "Indicates if the Current Table was Opened Exclusively" },
   { "tempindex", cbwTEMPINDEX, METH_VARARGS, "Creates a Temporary Index" },
   { "selecttempindex", cbwTEMPINDEXSELECT, METH_VARARGS, "Selects an Existing Temporary Index" },
   { "closetempindex", cbwTEMPINDEXCLOSE, METH_VARARGS, "Closes and Destroys a Temporary Index" },
   { "indexon", cbwINDEXON, METH_VARARGS, "Adds an Index Tag to the Main Index" },
   { "select", cbwSELECT, METH_VARARGS, "Select an Open Table as the Current One" },
   { "tagcount", cbwTAGCOUNT, METH_NOARGS, "Gets the Number of Index Tags" },
   { "used", cbwUSED, METH_VARARGS, "Indicates if the Specified Table Alias Name is In Use" },
   { "bof", cbwBOF, METH_NOARGS, "Returns Beginning of File Status" },
   { "goto", cbwGOTO, METH_VARARGS, "Goes to a Specified Place in a Table" },
   { "calcstats", cbwCALCSTATS, METH_VARARGS, "Calculate Statistics from a Table" },
   { "deleted", cbwDELETED, METH_NOARGS, "Indicates of a Record has been Deleted" },
   { "recno", cbwRECNO, METH_NOARGS, "Returns the Current Record Number" },
   { "count", cbwCOUNT, METH_VARARGS, "Counts Records for a Given Condition" },
   { "locate", cbwLOCATE, METH_VARARGS, "Locate a Record Based on an Expression" },
   { "locatecontinue", cbwLOCATECONTINUE, METH_NOARGS, "Continues the Locate Search" },
   { "locateclear", cbwLOCATECLEAR, METH_NOARGS, "Terminates a Locate Sequence" },
   { "deletetag", cbwDELETETAG, METH_VARARGS, "Deletes an Index Tag from the Main Index" },
   { "setorderto", cbwSETORDERTO, METH_VARARGS, "Set the Active Index Order on a Table" },
   { "ataginfo", cbwATAGINFO, METH_NOARGS, "Returns List of Dicts of Index Tag Information" },
   { "order", cbwORDER, METH_NOARGS, "Returns Tag Name of Current Index Tag" },
   { "scatterfield", cbwSCATTERFIELD, METH_VARARGS, "Returns text string contents of the specified Field" },
   { "scatterfieldex", cbwSCATTERFIELDCHAR, METH_VARARGS, "Returns text string from field recognizing nulls."}, 
   { "seek", cbwSEEK, METH_VARARGS, "Searches a Table by an Index Value" },
   { "skip", cbwSKIP, METH_VARARGS, "Skips forward or back through the table" },
   { "refreshbuffers", cbwREFRESHBUFFERS, METH_NOARGS, "Rereads all record data from disk" },
   { "refreshrecord", cbwREFRESHRECORD, METH_NOARGS, "Rereads one record's data from disk" },
   { "flush", cbwFLUSH, METH_NOARGS, "Flushes current table buffers to disk" },
   { "recallall", cbwRECALLALL, METH_NOARGS, "Undeletes all deleted records"} , 
   { "gettally", cbwTALLY, METH_NOARGS, "Returns latest activity count tally"}, 
   { "recall", cbwRECALL, METH_NOARGS, "Undeletes the current record"}, 
   { "flushall", cbwFLUSHALL, METH_NOARGS, "Flushes all table buffers"}, 
   { "fcount", cbwFCOUNT, METH_NOARGS, "Returns the number of Fields in the current table"},
   { "reccount", cbwRECCOUNT, METH_NOARGS, "Returns number of records in the current table"}, 
   { "eof", cbwEOF, METH_NOARGS, "Returns True if at the end of the table"}, 
   { "afields", cbwAFIELDS, METH_VARARGS, "Returns a list of field structure dicts"},
   { "afieldtypes", cbwAFIELDTYPES, METH_NOARGS, "Returns a dict of types by name"},
   { "scatterblank", cbwSCATTERBLANK, METH_VARARGS, "Returns a dict of blank field values"},
   { "scatter", cbwSCATTER, METH_VARARGS, "Returns the current record data as a dict"}, 
   { "curval", cbwCURVAL, METH_VARARGS, "Returns a field value as Python or String"}, 
   { "scatterfieldlogical", cbwSCATTERFIELDLOGICAL, METH_VARARGS, "Returns value of a logical field"},
   { "scatterfielddatetime", cbwSCATTERFIELDDATETIME, METH_VARARGS, "Returns value of a datetime field"},
   { "scatterfielddate", cbwSCATTERFIELDDATE, METH_VARARGS, "Returns value of a date field"},
   { "scatterfielddouble", cbwSCATTERFIELDDOUBLE, METH_VARARGS, "Returns value of a numeric field as a float"},
   { "scatterfielddecimal", cbwSCATTERFIELDDECIMAL, METH_VARARGS, "Returns numeric value as a decimal.Decimal() value"},
   { "scatterfieldlong", cbwSCATTERFIELDLONG, METH_VARARGS, "Returns numeric value as a long integer"}, 
   { "gathermemvar", cbwGATHERMEMVAR, METH_VARARGS, "Stores multiple values back into table fields"},
   { "gatherdict", cbwGATHERDICT, METH_VARARGS, "Stores values from a dict() into the current record"},
   { "preparefilter", cbwPREPAREFILTER, METH_VARARGS, "Prepares a record filter expression for future evaluation"}, 
   { "testfilter", cbwTESTFILTER, METH_NOARGS, "Tests the current record against the current filter"}, 
   { "clearfilter", cbwCLEARFILTER, METH_NOARGS, "Clears the current filter"}, 
   { "replace", cbwREPLACE_FIELD, METH_VARARGS, "Replaces field contents with the specified value"},
   { "closetable", cbwCLOSETABLE, METH_VARARGS, "Closes the specified table"},
   { "closedatabases", cbwCLOSEDATABASES, METH_NOARGS, "Closes all open tables and indexes"},
   { "createtable", cbwCREATETABLE, METH_VARARGS, "Creates a DBF-type data table."},
   { "appendblank", cbwAPPENDBLANK, METH_NOARGS, "Appends a blank record to the current table."},
   { "appendfrom", cbwAPPENDFROM, METH_VARARGS, "Appends data to an existing and open table."},
   { "copyto", cbwCOPYTO, METH_VARARGS, "Copies an open table to another table or a text file."},
   { "insertintotable", cbwINSERT, METH_VARARGS, "Stores data from a string into a table, OBSOLETE."},
   { "delete", cbwDELETE, METH_NOARGS, "Deletes the current record of the currently selected table"},
   { "rlock", cbwLOCK, METH_NOARGS, "Locks the current record."},
   { "flock", cbwFLOCK, METH_NOARGS, "Locks the current table."}, 
   { "unlock", cbwUNLOCK, METH_NOARGS, "Unlocks the current record."},
   { "pack", cbwPACK, METH_NOARGS, "Packs out deleted records and the memo file"},
   { "reindex", cbwREINDEX, METH_NOARGS, "Reindexes the current table"},
   { "zap", cbwZAP, METH_NOARGS, "Clears all records from the current table"},
   { "replacelong", cbwREPLACE_LONG, METH_VARARGS, "Replaces contents of a long/integer field"},
   { "copytags", cbwCOPYTAGS, METH_VARARGS, "Copies index tags from one table to another"},
   { "setcustomencoding", cbwSETCODEPAGE, METH_VARARGS, "Sets the Custom Code Page to a built-in Python codec"},
   { "isreadonly", cbwISREADONLY, METH_VARARGS, "Returns True if the specified table is opened in readonly mode"},
   { "fieldinfo", cbwFIELDINFO, METH_VARARGS, "Returns a tuple of field characteristics"},
   { "replacedatetimen", cbwREPLACE_DATETIMEN, METH_VARARGS, "Replace datetime field with element numbers"},
   { 0, 0, 0, 0 }
};

#ifdef IS_PY2K
PyMODINIT_FUNC initCodeBasePYWrapper(void)
{
#endif
#ifdef IS_PY3K
PyMODINIT_FUNC PyInit_CodeBasePYWrapper(void)
{
	PyObject* modReturn;
#endif
	long kk;
	long jj;
	CODEBASESTATUS *pStat;

	PyDateTime_IMPORT;
		
	for (kk = 0; kk < MAXDATASESSIONS; kk++)
		{
		pStat = &(gaCodeBaseEnvironments[kk]);
		pStat->pCurrentTable = NULL;
		pStat->pCurrentIndexTag = NULL;
		pStat->pCurrentFilterExpr = NULL;
		pStat->pLastField = NULL;
		pStat->pCurrentQuery = NULL;
		pStat->bLargeMode = FALSE;
		pStat->cQueryAlias[0] = (char) 0;
		pStat->nDeletedFlag = FALSE;
		pStat->cErrorMessage[0] = (char) 0;
		pStat->nLastErrorNumber = 0;
		strcpy(pStat->cDelimiter, "||");
		pStat->cLastSeekString[0] = (char) 0;
		strcpy(pStat->cDateFormatCode, "MM/DD/YY");
		for (jj = 0; jj < MAXEXCLCOUNT; jj++)
			{
			pStat->aExclList[jj] = calloc((size_t) (MAXALIASNAMESIZE + 1), sizeof(char));	
			}
		memset(pStat->aTempIndexes, 0, (size_t) (MAXTEMPINDEXES * sizeof(TEMPINDEX)));
		memset(pStat->aIndexTags, 0, (size_t) (MAXINDEXTAGS * sizeof(VFPINDEXTAG)));
		}
	
	memset(gnCodeBaseFlags, 0, (size_t) (MAXDATASESSIONS * sizeof(long)));
	for (kk = 0; kk < MAXDATASESSIONS; kk++)
		{
		gnCodeBaseFlags[kk] = -1;	
		}
	memset(gaCodeBases, 0, (size_t) (MAXDATASESSIONS * sizeof(CODE4)));
	//memset(gaExclList, 0, (size_t) (MAXEXCLCOUNT * sizeof(char*)));
	for (kk = 0; kk < MAXEXCLCOUNT; kk++)
		{
		gaExclList[kk] = calloc((size_t) (MAXALIASNAMESIZE + 1), sizeof(char));	
		}
	for (kk = 0; kk < MAXFIELDCOUNT; kk++)
		{
		gaFieldList[kk] = calloc(30, sizeof(char));
		}
	cbxInitGlobals();
	memset(gaFieldSpecs, 0, (size_t) (MAXFIELDCOUNT * sizeof(FIELD4INFO)));
	//memset(gaAliasTrack, 0, (size_t) (MAXOPENTABLECOUNT * sizeof(AliasTrack)));
#ifdef IS_PY2K
   (void) Py_InitModule("CodeBasePYWrapper", cbMethods);
#endif
#ifdef IS_PY3K
    static struct PyModuleDef moduledef = {
        PyModuleDef_HEAD_INIT,
        "CodeBasePYWrapper",     /* m_name */
        "Tools for accessing DBF tables using CodeBase functions.",  /* m_doc */
        -1,                  /* m_size */
        cbMethods,           /* m_methods */
        NULL,                /* m_reload */
        NULL,                /* m_traverse */
        NULL,                /* m_clear */
        NULL,                /* m_free */
    };
    modReturn = PyModule_Create(&moduledef);
    return(modReturn);
#endif
}
