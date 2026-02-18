/* --------------------------------------- November 12, 2008 - 12:40pm -------------------------
 * C language program providing a wrapper for higher level functions to access the CodeBase 
 * c4dll.dll library.  As of 2018, this program is designed to be compiled for 32-bit applications.
 * --------------------------------------------------------------------------------------------- */
/* *****************************************************************************
Copyright (C) 2008-2018 by M-P Systems Services, Inc.,
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
   
#define TRUE 1
#define FALSE 0
#define DOEXPORT __declspec(dllexport)
#define BUFFER_LEN 8920000
#define _CRT_SECURE_NO_WARNINGS 1
#include "d:\\codebase\\WorkingSource\\d4all.h"
#include <time.h>

long DOEXPORT cbwDELETED(void);

long glInitDone = FALSE;
long gbStrictAliasMode = FALSE;
long gnMaxTempTags = 20;
long gnMaxTempIndexes = 200;
long gnOpenTempIndexes = 0;
long gbLargeMode = FALSE;
CODE4 codeBase;

DATA4 *gpCurrentTable;
TAG4  *gpCurrentIndexTag; /* Future use */
EXPR4 *gpCurrentFilterExpr = NULL;
FIELD4 *gpLastField = NULL;
RELATE4 *gpCurrentQuery = NULL;
char gcQueryAlias[50]; /* Stores the table alias to which the gpCurrentQuery applies.  If cbwCONTINUE() is
called and a different table is currently selected, an error is reported. */

long gnDeletedFlag = FALSE; /* If TRUE, then will attempt to skip over deleted records. */
char* cTextBuffer = NULL; /* Allocate a bunch of space to this when starting up. */
char gcErrorMessage[1000]; /* Contains the most recent error message.  Guaranteed to be less than 1000 characters in length. */
long gnLastErrorNumber = 0;
char gcStrReturn[BUFFER_LEN]; /* Where we put stuff for static string return */
long gnProcessTally = 0; /* Keeps track of the TALLY of last records processed like VFP _TALLY. */
char gcDelimiter[3];
char gcLastSeekString[1000]; /* We store the last SEEK string value for future needs. */
char gcDateFormatCode[20];

typedef struct VFPFIELDtd {
	char cName[132];
	char cType[4];
	long nWidth;
	long nDecimals;
	long bNulls;	
} VFPFIELD;
VFPFIELD gaFields[256];

typedef struct FieldPlusTd {
	VFPFIELD xField;
	long nFldType;
	long bMatched;
	long nSource; // Index into the Source array. Used in gaMatchFields
	long nTarget; // Index into the Target array.  Used in gaMatchFields
	long bIdentical; // Set to TRUE if the type, width, and decimals are all identical.
	FIELD4 *pCBfield;
} FieldPlus;
FieldPlus gaSrcFields[300];
long gnSrcCount;
long gnTrgCount;
FieldPlus gaTrgFields[300];
FieldPlus gaMatchFields[300];

typedef struct VFPINDEXTAGtd {
	char cTagName[130];
	char cTagExpr[256];
	char cTagFilt[256];
	long nDirection;
	long nUnique;
} VFPINDEXTAG;
VFPINDEXTAG gaIndexTags[512]; /* The max we will allow, and well over the expected typical numbers */

typedef struct TEMPINDEXtd {
	INDEX4 *pIndex;
	char cFileName[255];
	char cTableAlias[100];
	VFPINDEXTAG aTags[20];
	long bOpen;
	long nTagCount;
	DATA4 *pTable;
} TEMPINDEX;
TEMPINDEX gaTempIndexes[200];


FIELD4INFO gaFieldSpecs[256];
//CODE4  codeBase;
//DATA4  *dataFile = 0;
//FIELD4 *fName,*lName,*address,*age,*birthDate,*married ,*amount,*comment;
//
//FIELD4INFO  fieldInfo [] =
//{
//   {"F_NAME",r4str,10,0},
//   {"L_NAME",r4str,10,0},
//   {"ADDRESS",r4str,35,0},
//   {"AGE",r4num,2,0},
//   {"BIRTH_DATE",r4date,8,0},
//   {"MARRIED",r4log,1,0},
//   {"AMOUNT",r4num,7,2},
//   {"COMMENT",r4memo,10,0},
//   {0,0,0,0},
//};	

long DOEXPORT cbwGOTO(char*, long);
long DOEXPORT cbwRECNO(void);

char* gaExclList[300]; // A list of alias values of currently open tables with the Exclusive property set to TRUE.
// We do this since the CodeBase engine doesn't give us a good way to test whether a table was opened in exclusive
// mode or not. Added 10/25/2013. JSH.  Alias strings are ALWAYS lower CASE.
long gnExclMax = 300;
long gnExclLimit = 20;
long gnExclCnt = 0;

DOEXPORT long cbwCREATETABLE(char*, char*); // Needed before it is defined.

/*-----------------1/30/2006 9:17PM-----------------
 * This DllMain doesn't seem to get fired off, even tho this kind of thing
 * does in the PowerBasic world.  Who knows why.
 * --------------------------------------------------*/
long DllMain(long hInstance, long fwdReason, long lpvReserved)
{
	return(1);
};

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
	lnLen = (long) strlen(lcName);
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
	
	lnLen = (long) strlen(cString);
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
			if ((*(cString + jj) == ' ') || (*(cString + jj) == (char) NULL))	*(cString + jj) = (char) NULL;
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
		lnVal = (long) floor((double) rand() * lnChrCnt / (double) RAND_MAX);
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
DOEXPORT char* MPSSStrTran(char* lpcSource, char* lpcTarget, char* lpcAlternative, long lpnCount)
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
	lnFromLen = (long) strlen(lpcTarget);
	lnToLen = (long) strlen(lpcAlternative);
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
		lnNewBuffLen = (long) strlen(lpcSource) + 1 + (lnOccurances * lnDiffLen); /* worst case */
		}
	else
		{
		lnNewBuffLen = (long) strlen(lpcSource) + 1;	
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

/* ******************************************************************************************* */
/* Wrapper for the d4index() function that doesn't depend on the table having been opened
   with an alias name the same as the table base name.  Added 02/20/2012. JSH. */
/* As of Oct. 25, 2013, we've discoverd that this will NOT produce a correct return value if   */
/* the table base name exceed 32 characters.                                                   */
/* TECH SUPPORT ISSUE WT102913-103426                                                          */
/* As of November 4, 2013, we have new source code from CodeBase that fixes this problem with  */
/* d4index() and allows us to call it regardless of what the alias has been set to and also    */
/* regardless of the length of the file names.                                                 */
INDEX4 *cbxD4Index( DATA4 *lppTable)
{
	INDEX4* pRet;
	// This code is no longer needed.
//	char lcTableName[500];
//	char lcIndexName[500];
//	char lcIndexPathName[500];
//	char lcIndexPathNameExtU[500];
//	char lcIndexPathNameExtL[500];
//	strncpy(lcTableName, d4fileName(lppTable), (size_t) 499);
//	
//	if (!gbStrictAliasMode)
//		{
//		u4namePiece(lcIndexName, 499, lcTableName, FALSE, FALSE );
//		StrToUpper(lcIndexName);
//		u4namePiece(lcIndexPathName, 499, lcTableName, TRUE, FALSE);
//		StrToUpper(lcIndexPathName);
//		strcpy(lcIndexPathNameExtL, lcIndexPathName);
//		strcpy(lcIndexPathNameExtU, lcIndexPathName);
//		strcat(lcIndexPathNameExtL, ".cdx");
//		strcat(lcIndexPathNameExtU, ".CDX");
//	
//		pRet = d4index(lppTable, lcIndexName); // Works around some, but not all of the peculiarities of CodeBase.
//		if (pRet == NULL)
//			{
//			pRet = d4index(lppTable, lcIndexPathName);	
//			}
//		if (pRet == NULL)
//			{
//			pRet = d4index(lppTable, lcIndexPathNameExtL);	
//			}
//		if (pRet == NULL)
//			{
//			pRet = d4index(lppTable, lcIndexPathNameExtU);	
//			}
//		}
//	else // For tables with base names longer than 32 characters, this gbStrictAliasMode must be turned on and NO alias defined.
//		{
		pRet = d4index(lppTable, NULL);
//		}
	return pRet;
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
long DOEXPORT cbwSETDATEFORMAT(char* lcFormat)
{
	long lnReturn = 1;
	char cDateFmat[20];
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
	return lnReturn;
}

/* ****************************************************************************** */
/* Retrieves the current dateformat code value.                                   */
DOEXPORT char* cbwGETDATEFORMAT(void)
{
return gcDateFormatCode; // We set this whenever we change it in CodeBase, so it has the
// latest status info.	
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
			//printf("The result of seek %ld \n", lnResult);
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

/* ****************************************************************************** */
/* Function that sets the gbStrictAliasMode property to True or False.  This has
   been added as of 10/26/2013 because d4index() doesn't work correctly to find
   the associated CDX INDEX4* value for tables with more than 32 characters in the
   file name.  In such cases, this must be turned ON and the table opened without
   a specified alias name in order for any index-related functions to work properly.
   Be sure to turn gbStrictAliasMode FALSE once it is not needed any longer.      */
long DOEXPORT cbwSETSTRICTALIASMODE(long bMode)
{
	gbStrictAliasMode = bMode;
	return(bMode);
	// Code left in for backwards compatibility, but this construct is not needed
	// after Nov. 4, 2013. bug fix from CodeBase. JSH.
}

/* ****************************************************************************** */
/* This function prepares a filter string against the characteristics of the      */
/* currently selected table.  Returns 1 on OK, 0 (False) on failure of any kind.  */
long DOEXPORT cbwPREPAREFILTER(char* lpcFilter)
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

/* ***************************************************************************** */
/* Companion for the above which applies the filter to the actual current record */
/* Returns 1 if filter expression evaluates to TRUE, 0 if it evaluates to FALSE, */
/* otherwise returns -1 on error.                                                */
long DOEXPORT cbwTESTFILTER(void)
{
	char lcResultBuffer[200];
	char *lpcResult;
	int lnLen;
	long lnReturn = -1;
	int lnType;
	int lnValue;
	
	if (gpCurrentFilterExpr == NULL)
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Filter Expression Active");
		gnLastErrorNumber = -8932;	
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

/* ******************************************************************************* */
/* Clean up after the filter work is done.                                         */
long DOEXPORT cbwCLEARFILTER(void)
{
	if (gpCurrentFilterExpr != NULL)
		{
		expr4free(gpCurrentFilterExpr);
		gpCurrentFilterExpr = NULL;	
		}	
	return(1);
}

/* ****************************************************************************** */
/* Call this first before you call anything else or the results will be BAD.      */
/* Pass a 1 in lnLargeFlag, and this will initialize the engine to use the large  */
/* table support.  Otherwise pass 0 as lnLargeFlag.  WARNING!  Tables opened with */
/* Large Table Support can NOT be opened simultaneously by Visual FoxPro.         */
long DOEXPORT cbwInit(long lnLargeFlag)
{
	long lnResult = 0;
	if (glInitDone) return(glInitDone); /* Don't allow this to be called a second time. */

	lnResult = code4init(&codeBase);
	if (lnResult >= 0)
		{
		glInitDone = TRUE;
		gnLastErrorNumber = 0;
		gpCurrentTable = NULL;
		gpCurrentIndexTag = NULL;
		gnProcessTally = 0;
		gcDelimiter[0] = '|';
		gcDelimiter[1] = '|';
		gcDelimiter[2] = (char) 0;
		
		memset(gcErrorMessage, 0, (size_t) 1000 * sizeof(char));
		memset(gaFields, 0, (size_t) 256 * sizeof(VFPFIELD));
		memset(gaIndexTags, 0, (size_t) 512 * sizeof(VFPINDEXTAG));
		memset(gcLastSeekString, 0, (size_t) 1000 * sizeof(char));
		memset(gaTempIndexes, 0, (size_t) gnMaxTempIndexes * sizeof(TEMPINDEX));
		codeBase.errOff = 1; /* No error messages displayed... just return error codes and let us handle them. */
		codeBase.safety = 0; /* allow us to overwrite things */
		code4dateFormatSet(&codeBase, "MM/DD/YYYY");
		codeBase.compatibility = 30; // Visual Foxpro 6.0 and above.
		codeBase.optimize = OPT4ALL;
		codeBase.memExpandData = 3;
		codeBase.memStartBlock = 6;
		//codeBase.memStartMax = 268435456; // This and two above added 07/28/2011 for memory control (otherwise used 1GB of memory). JSH.
		codeBase.memStartMax = 67108864; // The above still wastes a lot of memory. 02/20/2012. JSH.
		codeBase.codePage = cp1252; // Added this 11/19/2011 to force Visual FoxPro code page.
		codeBase.lockAttempts = 2; // Tries twice to lock a record.  Up to us to try again beyond that.  Added 02/07/2012. JSH.
		codeBase.lockDelay = 2; // 2/100ths of a second. Added 02/07/2012. JSH.
		codeBase.singleOpen = FALSE;
		codeBase.autoOpen = TRUE;
		codeBase.ignoreCase = TRUE;
		if (lnLargeFlag == 1)
			{
			code4largeOn(&codeBase);
			gbLargeMode = TRUE;
			}
		srand((unsigned int) time(NULL));
		memset(gaExclList, 0, (size_t) gnExclMax * sizeof(char*));
		strcpy(gcDateFormatCode, "MM/DD/YY"); // The CodeBase Default
		}
	else
		{
		glInitDone = FALSE;
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		gnLastErrorNumber = codeBase.errorCode;	
		}
return(glInitDone);	
}

/* ****************************************************************************** */
/* Call this to determine the current setting of "Large Table Mode".  Returns     */
/* 1 if large table mode is ON, otherwise 0.                                      */
long DOEXPORT cbwIsLargeMode(void)
{
return (gbLargeMode);	
}

/* ****************************************************************************** */
/* Call this at the end when you are shutting down or to put the module in a state
   of suspension.  Once this is called, cbwInit() must be called again            */
long DOEXPORT cbwShutDown(void)
{
	long lnResult, lbResult, jj;
	if (glInitDone == FALSE) return(TRUE); /* Nothing to do. */
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
				
	lnResult = code4close(&codeBase);
	lnResult = code4initUndo(&codeBase);
	lbResult = TRUE;
	if (lnResult < 0)
		{
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error, but we shut down anyway. */
		gnLastErrorNumber = codeBase.errorCode;
		lbResult = FALSE;	
		}

	if (gnExclCnt > 0)
		{
		for (jj = 0; jj < gnExclMax; jj++)
			{
			if (gaExclList[jj] != NULL)
				{
				free(gaExclList[jj]);
				gaExclList[jj] = NULL;
				gnExclCnt -= 1;
				if (gnExclCnt == 0) break;
				}
			}	
		}

	glInitDone = FALSE;
	gpCurrentTable = NULL;
	return(lbResult);	
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
long DOEXPORT cbwCALCSTATS(long lpnStat, char *lpcFieldExpr, char *lpcForVFPexpr, double *lpnValue)
{
	double lnReturn = 0.0;
	double lnTemp = 0.0;
	long lnOK = TRUE;
	long lnResult = 0;
	long lnStart = 1;
	long lnCount = 0;
	long lnOldRecord = 0;
	RELATE4 *xpQuery = NULL;
	EXPR4 *xpExpr = NULL;
	double lnVeryLargeValue = 9999999999999999.9;
	double lnVerySmallValue = -9999999999999999.9;
	double lnSumTrack = 0.0;
	
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
				lnResult = relate4querySet(xpQuery, lpcForVFPexpr);
				if (lnResult != r4success)
					{
					strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
					gnLastErrorNumber = codeBase.errorCode;
					lnReturn = -1.0;					
					}
				}
				
			if (lnReturn > -1.0)
				{
				xpExpr = expr4parse(gpCurrentTable, lpcFieldExpr);
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
					lnOldRecord = cbwRECNO();	
					do
						{
						if (lnStart == 1) lnResult = relate4top(xpQuery);
						else lnResult = relate4skip(xpQuery, 1);
						lnStart = 0;
						if (lnResult == r4success)
							{
							if ((gnDeletedFlag == FALSE) || ((gnDeletedFlag == TRUE) && (cbwDELETED() == FALSE)))
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
					cbwGOTO("RECORD", lnOldRecord); // Go back to where we were.
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
		lnOK = FALSE;
		*lpnValue = lnReturn;
		}
	else
		{
		lnOK = TRUE;
		*lpnValue = lnReturn;	
		}
	return(lnOK);
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
long DOEXPORT cbwCOUNT(char *lpcVFPexpr)
{
	long lnReturn = 0;
	long lnResult = 0;
	long lnStart = 1;
	long lnCount = 0;
	long lnOldRecord = 0;
	RELATE4 *xpQuery = NULL;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
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
				lnResult = relate4querySet(xpQuery, lpcVFPexpr);
				if (lnResult != r4success)
					{
					strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error */
					gnLastErrorNumber = codeBase.errorCode;
					lnReturn = -1;					
					}
				else
					{
					lnOldRecord = cbwRECNO();	
					do
						{
						if (lnStart == 1) lnResult = relate4top(xpQuery);
						else lnResult = relate4skip(xpQuery, 1);
						lnStart = 0;
						if (lnResult == r4success)
							{
							if (gnDeletedFlag == TRUE)
								{
								if (!cbwDELETED()) lnCount += 1;	
								}
							else
								{
								lnCount += 1; // always.	
								}
							}
						else break;
						} while(TRUE);
					lnResult = relate4free(xpQuery, 0);
					cbwGOTO("RECORD", lnOldRecord); // Go back to where we were.
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
	return(lnReturn);
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
   SET DELETED setting.  Returns -1 on ERROR. */
long DOEXPORT cbwLOCATE(char* lpcVFPexpr)
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
			lnResult = relate4querySet(gpCurrentQuery, lpcVFPexpr);
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
		if (cbwDELETED())
			{
			// This "FOUND" record has been deleted, so we try for the next non-deleted one...
			lnReturn = FALSE; // be pessimistic
			while (relate4skip(gpCurrentQuery, 1) == r4success)
				{
				if (!cbwDELETED())
					{
					lnReturn = TRUE;
					break;	
					}	
				}
			}	
		}
			
	return(lnReturn);	
}

/* ***************************************************************************** */
/* Appends a blank record to the end of the currently selected table.            */
/* Returns TRUE for OK, -1 on ERROR.                                             */
long DOEXPORT cbwAPPENDBLANK(void)
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


/* ******************************************************************************* */
/* Function that terminates a cbwLOCATE process.  Call this after a series of
   cbwLOCATE and cbwCONTINUE commands to end the cycle and free up memory.
   Returns TRUE unless there is no current LOCATE process active, in which case
   it returns FALSE and the error values are set.                                  */
long DOEXPORT cbwLOCATECLEAR(void)
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
	return(lnReturn);
}

/* ******************************************************************************* */
/* Equivalent to the VFP CONTINUE that gets the next matching record after a
   successful cbwLOCATE() function.  Returns FALSE if nothing else found. -1 on
   any kind of error.  EOF() or BOF() return FALSE, not error.                     */
long DOEXPORT cbwCONTINUE(void)
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
			lnReturn = -1;			
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
					lnReturn = -1;							
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
		if (cbwDELETED())
			{
			// This "FOUND" record has been deleted, so we try for the next non-deleted one...
			lnReturn = FALSE; // be pessimistic
			while (relate4skip(gpCurrentQuery, 1) == r4success)
				{
				if (!cbwDELETED())
					{
					lnReturn = TRUE;
					break;	
					}	
				}
			}	
		}
			
	return(lnReturn);	
}

/* ******************************************************************************* */
/* Sets the status of the DELETED flag meaning that if TRUE, deleted records are   */
/* invisible.  If FALSE, then deleted records are visible but can be detected with */
/* the cbwDELETED() function.                                                      */
/* Operates on all tables.                                                         */
/* NOTE ADDED FEATURE 08/04/2014.  Pass -1, and it will return the current state of*/
/* the deleted flag.                                                               */
long DOEXPORT cbwSETDELETED(long lpnDeletedOn)
{
	if (lpnDeletedOn == -1)
		{
		return(gnDeletedFlag);	
		}
	else
		{
		gnDeletedFlag = lpnDeletedOn;
		return(1);
		}	
}

/* ******************************************************************************* */
/* Attempts to USE the table passed by name.  Returns 1 on OK, 0 on failure.       */
/* Accepts an optional 2nd parm to specify a non-standard alias.  Pass an empty    */
/* string to use the default.                                                      */
long DOEXPORT cbwUSE(char* lpcTable, char* lpcAlias, long lpnReadOnlyFlag)
{
	DATA4 *lpTable;
	long lbReturn = TRUE;
	char cWorkTable[300];
	
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
	
	codeBase.accessMode = OPEN4DENY_NONE; /* Full shared mode */
	strcpy(cWorkTable, lpcTable);
	StrToUpper(cWorkTable); // Force the DLL to store the table name as upper no matter what.  OK for Windows.
	lpTable = d4open(&codeBase, cWorkTable);
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
		gpCurrentTable = lpTable;
		gpCurrentIndexTag = NULL;
		if (strlen(lpcAlias) > 0)
			{
			d4aliasSet(gpCurrentTable, lpcAlias);	
			}
		d4top(gpCurrentTable); /* Always position at the first physical record on open */	
		}
	return(lbReturn);
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
			if (gaExclList[jj] == NULL) continue;
			if (strcmp(gaExclList[jj], cMatchAlias) == 0)
				{
				nAccessMode = OPEN4DENY_RW;
				nMatchPoint = jj;
				break;
				}	
			}	
		}
	
	codeBase.accessMode = (char) nAccessMode;
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
			if (gbStrictAliasMode)
				{
				// No alias name force.  Just in case, we re-store the original name to whatever the natural alias is.;
				if (strcmp(cMatchAlias, cTableAlias) != 0)
					{
					if (nMatchPoint > -1)
						{
						free(gaExclList[jj]);
						gaExclList[jj] = calloc(strlen(cMatchAlias) + 2, sizeof(char));
						strcpy(gaExclList[jj], StrToLower(cMatchAlias));
						}	
					}	
				}
			else
				{
				if (strlen(cTableAlias) > 0)
					{
					d4aliasSet(gpCurrentTable, cTableAlias); // Forcing alias to its original value	
					}
				}
			d4top(gpCurrentTable); /* Always position at the first physical record on open */	
			}
		}
	if (nAccessMode != OPEN4DENY_RW)
		{
		lbReturn = FALSE;
		strcpy(gcErrorMessage, "EXCLUSIVE NOT RESTORED");
		gnLastErrorNumber = -394249;
		}
	return(lbReturn);	
}

/* ******************************************************************************* */
/* Uses the gaExclList array to get the exclusive status of the currently selected */
/* table.  Returns 1 (TRUE) on yes, it is exclusive, 0 on NOT exclusive, -1 on err.*/
long DOEXPORT cbwISEXCL(void)
{
	char cAlias[400];
	long jj;
	long nRet = 0;
	if (gpCurrentTable == NULL)
		{
		strcpy(gcErrorMessage, "No Table Selected");
		gnLastErrorNumber = -99999;
		return(-1);			
		}	
	
	strcpy(cAlias, d4alias(gpCurrentTable));
	StrToLower(cAlias);
	for (jj = 0;jj < gnExclLimit; jj++)
		{
		if (gaExclList[jj] == NULL) continue;
		if (strcmp(gaExclList[jj], cAlias) == 0)
			{
			nRet = 1;
			break;	
			}	
		}
	return(nRet);	
}

/* ******************************************************************************* */
/* Attempts to USE the table passed by name.  Returns 1 or 0 on OK, -1 on failure. */
/* Accepts an optional 2nd parm to specify a non-standard alias.  Pass an empty    */
/* string to use the default.                                                      */
/* Like cbwUSE except that the table is opened EXCLUSIVE (never read only!)        */
long DOEXPORT cbwUSEEXCL(char* lpcTable, char* lpcAlias)
{
	DATA4 *lpTable;
	long lbReturn = TRUE;
	char cWorkAlias[300];
	char cWorkTable[400];
	long nAliasLen = 0;
	long jj = 0;
	
	codeBase.readOnly = FALSE;
	codeBase.accessMode = OPEN4DENY_RW; /* i.e. EXCLUSIVE */
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	strcpy(cWorkTable, lpcTable);
	StrToUpper(cWorkTable);
	
	lpTable = d4open(&codeBase, cWorkTable);
	if (lpTable == NULL)
		{
		lbReturn = FALSE;
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		gnLastErrorNumber = codeBase.errorCode;	
		}
	else
		{
		gpCurrentTable = lpTable;
		gpCurrentIndexTag = NULL;
		if (strlen(lpcAlias) > 0)
			{
			d4aliasSet(gpCurrentTable, lpcAlias);	
			}
		d4top(gpCurrentTable);
		strcpy(cWorkAlias, d4alias(gpCurrentTable));
		nAliasLen = (long) strlen(cWorkAlias);
		for (jj = 0;(jj < (gnExclLimit + 1)) && (jj < gnExclMax) ; jj++)
			{
			if (gaExclList[jj] == NULL)
				{
				gaExclList[jj] = calloc(nAliasLen + 2, sizeof(char));
				strcpy(gaExclList[jj], StrToLower(cWorkAlias));
				gnExclCnt += 1;
				if (jj > gnExclLimit) gnExclLimit = jj + 1;
				break;	
				}
			}
		}
	return(lbReturn);
}

/* ***************************************************************************** */
/* Function that works like VFP SELECT myAlias.  Pass the Alias Name.            */
/* Returns 1 on success, 0 on "can't find alias"                                 */
long DOEXPORT cbwSELECT(char* lpcAlias)
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
		//strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		//gnLastErrorNumber = codeBase.errorCode;	
		}
	else
		{
		gpCurrentTable = lpTable;	
		}
	return(lbReturn);
}

/* ********************************************************************************** */
/* Counterpart to the cbwTEMPINDEX() function.  Hand it the index code of a temp      */
/* index, and it will close the index.  If the table to which this was associated     */
/* is still open, it will re-associate it with its CDX and set the order to NULL      */
/* meaning "no order".  Returns 1 on OK, -1 on failure.                               */
DOEXPORT long cbwTEMPINDEXCLOSE(long lnIndx)
{
	long lnReturn = 1;
	long lnResult;
	
	if ((lnIndx > gnMaxTempIndexes) || (lnIndx < 0))
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "Temp Index Code is out of range");
		gnLastErrorNumber = -9194;	
		}
	else
		{
		if (gaTempIndexes[lnIndx].bOpen)
			{
			gaTempIndexes[lnIndx].bOpen = FALSE;
			if (gaTempIndexes[lnIndx].pIndex != NULL)
				{
				lnResult = i4close(gaTempIndexes[lnIndx].pIndex);
				if (lnResult != r4success)
					{
					lnReturn = -1;	
					}
				}
			if (gaTempIndexes[lnIndx].pTable != NULL)
				{
				d4tagSelect(gaTempIndexes[lnIndx].pTable, NULL);
				}
			if (gnOpenTempIndexes > 0) gnOpenTempIndexes -= 1;
	
			remove(gaTempIndexes[lnIndx].cFileName);
			memset(gaTempIndexes[lnIndx].cFileName, 0, (size_t) 255 * sizeof(char));
			memset(gaTempIndexes[lnIndx].cTableAlias, 0, (size_t) 100 * sizeof(char));
			gaTempIndexes[lnIndx].nTagCount = 0;
			gaTempIndexes[lnIndx].pIndex = NULL;
			}	
		}
	return lnReturn;
}

/* *************************************************************************** */
/* Function like VFP CLOSE DATABASES ALL.                                      */
/* Keeps the engine alive for future table opens.                              */
long DOEXPORT cbwCLOSEDATABASES(void)
{
	long lbReturn, lnResult;
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
			if (gaTempIndexes[jj].bOpen) cbwTEMPINDEXCLOSE(jj);
			}
		}
	if (gnExclCnt > 0)
		{
		for (jj = 0; jj < gnExclMax; jj++)
			{
			if (gaExclList[jj] != NULL)
				{
				free(gaExclList[jj]);
				gaExclList[jj] = NULL;
				gnExclCnt -= 1;
				if (gnExclCnt == 0) break;
				}
			}	
		}	
	lnResult = code4close(&codeBase);
	lbReturn = TRUE;
	if (lnResult < 0)
		{
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode)); /* Error, but we close anyway. */
		gnLastErrorNumber = codeBase.errorCode;
		lbReturn = FALSE;	
		}
	gpCurrentTable = NULL;
	gpCurrentIndexTag = NULL;
	
	return(lbReturn);
}

/* ************************************************************************** */
/* Close one specified table.  Pass the Alias Name.                           */
/* Returns TRUE (1) on OK, FALSE (0) if alias not found                       */
long DOEXPORT cbwCLOSETABLE(char *lpcAlias)
{
	long lnReturn = TRUE;
	long lbResult;
	DATA4 *lpTable;
	long jj;
	char lcAlias[200];
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
		
	if (strlen(lpcAlias) > 0)
		{
		lpTable = code4data(&codeBase, lpcAlias);
		strcpy(lcAlias, lpcAlias);	
		}
	else
		{
		lpTable = gpCurrentTable;
		if (gpCurrentTable != NULL)
			{
			strcpy(lcAlias, d4alias(gpCurrentTable));
			}	
		else
			{
			lcAlias[0] = (char) 0;	
			}
		}
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
				cbwTEMPINDEXCLOSE(jj);
				}
			}	
		}
	
	if (gnExclCnt > 0)
		{
		for (jj = 0; jj < gnExclLimit; jj++)
			{
			if (gaExclList[jj] == NULL) continue;
			if (strcmp(lcAlias, gaExclList[jj]) == 0)
				{
				free(gaExclList[jj]);
				gaExclList[jj] = NULL;
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
/* Function like VFP DBF() which returns the fully qualified path name of     */
/* the currently selected table or an empty string.                           */
/* Accepts an alias or an empty string.                                       */
DOEXPORT char* cbwDBF(char *lpcAlias)
{
	long lnReturn = TRUE;
	DATA4 *lpTable;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
		
	if (strlen(lpcAlias) > 0)
		{
		lpTable = code4data(&codeBase, lpcAlias);	
		}
	else
		{
		if (gpCurrentTable == NULL)
			{
			strcpy(gcStrReturn, "");
			strcpy(gcErrorMessage, "No Table Selected");
			gnLastErrorNumber = -99999;
			return(FALSE);	
			}
		lpTable = gpCurrentTable;	
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
		gcStrReturn[0] = (char) 0;	
		}
	return(gcStrReturn);
}

/* *************************************************************************** */
/* Returns the value of the last error message                                 */
DOEXPORT char* cbwERRORMSG(void)
{
	return(gcErrorMessage);	
}

/* *************************************************************************** */
/* Returns the value of the last error number.                                 */
long DOEXPORT cbwERRORNUM(void)
{
	return(gnLastErrorNumber);	
}

/* *************************************************************************** */
/* Like the VFP ALIAS() function returns the name of the currently selected    */
/* ALIAS.                                                                      */
DOEXPORT char* cbwALIAS(void)
{
	char lcName[200];
	if (gpCurrentTable == NULL) lcName[0] = (char) 0;
	else strcpy(lcName, d4alias(gpCurrentTable));
	
	strcpy(gcStrReturn, lcName);
	return(gcStrReturn);	
}

/* *************************************************************************** */
/* Like the VFP REINDEX command, reindexes all CDX index tags in the currently
   selected data table.  Returns 1 on OK, 0 on failure for any reason.         */
long DOEXPORT cbwREINDEX(void)
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
	return(lnReturn);	
}

/* ************************************************************************** */
/* Tests the current READONLY status of the currently selected table.         */
/* Returns 1 TRUE for IS readonly, 0 FALSE for NOT readonly, and -1 on error. */
long DOEXPORT cbwISREADONLY(void)
{
	long lnReturn = -1;
	if (gpCurrentTable != NULL)
		{
		lnReturn = gpCurrentTable->readOnly;	
		}
	return(lnReturn);
}

/* ************************************************************************** */
/* Function equivalent to VFP PACK Command on the current table.  Removes     */
/*   all deleted records.  Returns 1 on success, 0 on failure.                */
long DOEXPORT cbwPACK(void)
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
	return(lnReturn);	
}

/* ************************************************************************** */
/* Equivalent to VFP Function RECCOUNT() returning the number of records in   */ 
/* the table. Returns -1 on ERROR.                                            */
long DOEXPORT cbwRECCOUNT(void)
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
	return(lnReturn);	
}

/* ***************************************************************************** */
/* Equivalent to VFP Function RECNO() returning the number of the current record */ 
/* in the table. Returns -1 on ERROR.                                            */
long DOEXPORT cbwRECNO(void)
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


/* ************************************************************************** */
/* Function equivalent to VFP ZAP Command on the current table.  Removes      */
/*   ALL records.  Returns 1 on success, 0 on failure.                        */
long DOEXPORT cbwZAP(void)
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
	return(lnReturn);	
}

/* ***************************************************************************** */
/* Returns EOF status for the currently selected table.                          */
/* Returns TRUE for EOF, FALSE for not or -1 on ERROR.                           */
long DOEXPORT cbwEOF(void)
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

/* ***************************************************************************** */
/* Equivalent of RECALL ALL which applies to the currently selected table.  Just */
/* un-deletes ALL deleted records.  Returns TRUE (1) if successful, or -1 on     */
/* error.                                                                        */
long DOEXPORT cbwRECALLALL(void)
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
			lnReturn = -1;	
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
/* Returns the most recent TALLY value like the _TALLY built in variable in VFP. */
long DOEXPORT cbwTALLY(void)
{
	return(gnProcessTally);
}

/* ***************************************************************************** */
/* Equivalent of UNLOCK which applies to the currently selected table.           */
/* Unlocks ALL RECORDS.  Returns TRUE (1) if successful, or -1 on ERROR.         */
long DOEXPORT cbwUNLOCK(void)
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
			lnReturn = -1;
			strcpy(gcErrorMessage, "Unable to unlock");
			gnLastErrorNumber = -9997;
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
/* Equivalent of RLOCK OR LOCK which applies to the currently selected table.    */
/* Just locks the current record.  Returns TRUE (1) if successful, or -1 if      */
/* there is no current record or other problem.                                  */
long DOEXPORT cbwLOCK(void)
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
			lnReturn = -1;	
			}
		else
			{
			lnResult = d4lock(gpCurrentTable, lnRecord); 
			if (lnResult != 0)
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, "Unable to lock");
				gnLastErrorNumber = -9998;
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


/* ***************************************************************************** */
/* Equivalent of RECALL which applies to the currently selected table.  Just     */
/* un-deletes the current record.  Returns TRUE (1) if successful, or -1 if      */
/* there is no current record or other problem.                                  */
long DOEXPORT cbwRECALL(void)
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
			lnReturn = -1;	
			}
		else
			{
			d4recall(gpCurrentTable); /* OK, even if not really deleted */
			gnProcessTally = 1;
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
/* Equivalent of DELETED() which applies to the currently selected table and     */
/* current record.  Returns 1 TRUE if deleted or 0 FALSE.  -1 on ERROR.          */
long DOEXPORT cbwDELETED(void)
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
/* Equivalent of DELETE which applies to the currently selected table.  Just     */
/* deletes the current record.  Returns TRUE (1) if successful, or -1 if         */
/* there is no current record or other problem.                                  */
long DOEXPORT cbwDELETE(void)
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
			lnReturn = -1;	
			}
		else
			{
			d4delete(gpCurrentTable);
			gnProcessTally = 1;
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
/* Equivalent of VFP SET ORDER TO which specifies at TAG name.  Returns 1 on OK  */
/* or -1 on Not OK.  Sets error code on why not ok.                              */
long DOEXPORT cbwSETORDERTO(char *lpcTagName)
{
	long lnReturn = TRUE;
	TAG4 *lpTag;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		if ((lpcTagName == NULL) || (strlen(lpcTagName) == 0))
			{
			lpTag = NULL; /* This is legal.  They just want record number order */
			}
		else
			{
			lpTag = d4tag(gpCurrentTable, lpcTagName);	
			if (lpTag == NULL)
				{
				lnReturn = -1;
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
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return(lnReturn);	
}

/* ***************************************************************************** */
/* Returns BOF status for the currently selected table.                          */
/* Returns TRUE for BOF, FALSE for not or -1 on ERROR.                           */
long DOEXPORT cbwBOF(void)
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
	return(lnReturn);	
}

/* Utility function for cbwGOTO */
long cbwxSetReturn(DATA4* lpbTable, long lpnResult)
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
long DOEXPORT cbwGOTO(char* lpcHow, long lpnCount)
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
				lnReturn = cbwxSetReturn(gpCurrentTable, lnResult);	
				break;	
				}
			if ((*lpcHow == 'B')	|| (*lpcHow == 'b')) /* Bottom */
				{
				lnResult = d4bottom(gpCurrentTable);
				if ((gnDeletedFlag == TRUE) && (lnResult == r4success))
					lnResult = cbxFINDUNDELETED(gpCurrentTable, -1, FALSE);
				lnReturn = cbwxSetReturn(gpCurrentTable, lnResult);	
				break;	
				}
			if ((*lpcHow == 'N')	|| (*lpcHow == 'n')) /* Next */
				{
				lnResult = d4skip(gpCurrentTable, 1);
				if ((gnDeletedFlag == TRUE) && (lnResult == r4success))
					lnResult = cbxFINDUNDELETED(gpCurrentTable, 1, FALSE);
				lnReturn = cbwxSetReturn(gpCurrentTable, lnResult);	
				break;	
				}
			if ((*lpcHow == 'P')	|| (*lpcHow == 'p')) /* Previous */
				{
				lnResult = d4skip(gpCurrentTable, -1);
				if ((gnDeletedFlag == TRUE) && (lnResult == r4success))
					lnResult = cbxFINDUNDELETED(gpCurrentTable, -1, FALSE);				
				lnReturn = cbwxSetReturn(gpCurrentTable, lnResult);	
				break;	
				}		
			if ((*lpcHow == 'S')	|| (*lpcHow == 's')) /* Skip */
				{
				if (lpnCount == 0) lpnCount = 1; /* Just in case, we at least skip 1 */
				lnResult = d4skip(gpCurrentTable, lpnCount);
				if ((gnDeletedFlag == TRUE) && (lnResult == r4success))
					lnResult = cbxFINDUNDELETED(gpCurrentTable, (lpnCount > 0 ? 1 : -1), FALSE);				
				lnReturn = cbwxSetReturn(gpCurrentTable, lnResult);	
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
						lnReturn = cbwxSetReturn(gpCurrentTable, lnResult);	
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
long DOEXPORT cbwSEEK(char* lpcMatch, char* lpcAlias, char* lpcTagName)
{
	long lnReturn = FALSE;
	long lnResult;
	DATA4 *lpTable;
	INDEX4 *lpIndex;
	TAG4   *lpTag;
	TAG4   *lpOldTag; /* The tag prior to when we did our seek */
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gcLastSeekString[0] = (char) 0; /* Initialize it to empty */

	//printf("the pointers: %p %p %p \n", lpcMatch, lpcAlias, lpcTagName);
	lpTable = code4data(&codeBase, lpcAlias);
	if (lpTable == NULL)
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "Alias not found");
		gnLastErrorNumber = -9994;	
		}
	else
		{
		//lpIndex = d4index(lpTable, NULL); /* Looks for the currently open production index */
		lpIndex = cbxD4Index(lpTable); /* Knows about tables with custom aliases. */
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
			lpTag = d4tag(lpTable, lpcTagName);	
			if (lpTag == NULL)
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, "Index Tag Does Not Exist");
				gnLastErrorNumber = -9992;	
				}
			else
				{
				//printf("Ready to Tag Select %p from name %s \n", lpTag, lpcTagName);
				d4tagSelect(lpTable, lpTag);
				//printf("Ready to seek in %p for %s\n", lpTable, lpcMatch);
				lnResult = d4seek(lpTable, lpcMatch);
				strcpy(gcLastSeekString, lpcMatch);
				//printf("Last Seek String %s\n", lpcMatch);
				while (TRUE)
					{
					if (lnResult == r4success)
						{
						if (gnDeletedFlag == TRUE)
							{
							//printf("Looking for undeleted\n");
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
	return(lnReturn);	
}

/* ********************************************************************************** */
/* Equivalent to USED() function which returns TRUE if there is an alias by the       */
/* specified name, else FALSE.                                                        */
long DOEXPORT cbwUSED(char* lpcAlias)
{
	long lnReturn = FALSE;
	DATA4 *lpTable;
	StrToLower(lpcAlias);
	lpTable = code4data(&codeBase, lpcAlias);
	if (lpTable != NULL) lnReturn = TRUE;
	return(lnReturn);	
}

/* ********************************************************************************** */
/* Returns TRUE if engine is active, else FALSE.                                      */
long DOEXPORT cbwEngineActive(void)
{
	return(glInitDone);	
}

/* ********************************************************************************** */
/* Flushes buffers like VFP FLUSH ALL where ALL table data is flushed.                */
/* Returns TRUE on success, -1 on ERROR.                                              */
long DOEXPORT cbwFLUSHALL(void)
{
	long lnResult;
	long lnReturn = TRUE;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
		
	lnResult = code4flush(&codeBase);
	if (lnResult != 0)
		{
		lnReturn = -1;
		gnLastErrorNumber = codeBase.errorCode;
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));			
		}
	return(lnReturn);		
}

/* ******************************************************************************* */
/* Equivalent of FCOUNT() which returns the number of fields in the current table. */
/* Returns the number or -1 on error.                                              */
long DOEXPORT cbwFCOUNT(void)
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
	return(lnReturn);	
}

/* ***************************************************************************** */
/* Clears the buffers for the currently selected table, which forces a new read  */
/* the next time any part of the table is accessed.  This is more drastic than   */
/* the cbwREFRESHRECORD, in that it forces a complete re-read of the table header*/
/* and other parts of the table.  Will slow down processing if called on every   */
/* read in a loop.                                                               */
/* Returns TRUE for SUCCESS, -1 on ERROR.                                        */
/* Added 04/10/2014. JSH.                                                        */
long DOEXPORT cbwREFRESHBUFFERS(void)
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
				lnReturn = -1;
				strcpy(gcErrorMessage, "Buffer content not refreshed");
				gnLastErrorNumber = -99873;	
				}
			else /* ERROR */
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;					
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


/* ***************************************************************************** */
/* Refreshes the "current record" of the current table.                          */
/* Returns TRUE for SUCCESS, -1 on ERROR.                                        */
long DOEXPORT cbwREFRESHRECORD(void)
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
				lnReturn = -1;
				strcpy(gcErrorMessage, "Record content not refreshed");
				gnLastErrorNumber = -9943;	
				}
			else /* ERROR */
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;					
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

/* ***************************************************************************** */
/* Flushes all data for the currently selected table.                            */
/* Returns TRUE for SUCCESS, -1 on ERROR.                           */
long DOEXPORT cbwFLUSH(void)
{
	long lnReturn = TRUE;
	long lnResult;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
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
				lnReturn = -1;
				strcpy(gcErrorMessage, "Record content not saved");
				gnLastErrorNumber = -9995;	
				}
			else /* ERROR */
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
				gnLastErrorNumber = codeBase.errorCode;					
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

/* ***************************************************************************** */
/* Like VFP AFIELDS() in that it returns information about all the fields in     */
/* the currently selected table.  The field information is stored in the global  */
/* array gaFields a pointer to which is returned by the function.  Pass a long   */
/* int by reference to get the count.  Be sure to grab the field info from the   */
/* array before calling this again, as the array will be overwritten each time.  */
/* You can iterate through the fields and when you come to a cType[0] value of 0,*/
/* you are done.                                                                 */
/* If there is nothing there, then check the lpnCount.  It will be -1 on error.  */
DOEXPORT VFPFIELD* cbwAFIELDS(long* lpnCount)
{
	long lnReturn = TRUE;
	long lnResult;
	long lnFldCnt;
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
			for (jj = 0;jj < lnFldCnt; jj++)
				{
				lpField = d4fieldJ(gpCurrentTable, (short) (jj + 1));				
				strcpy(gaFields[jj].cName, f4name(lpField));
				gaFields[jj].cType[0]  = (char) f4type(lpField);
				gaFields[jj].cType[1]  = (char) 0;
				gaFields[jj].nWidth    = f4len(lpField);
				gaFields[jj].nDecimals = f4decimals(lpField);
				// printf("getting a null content for afields()\n");
				gaFields[jj].bNulls    = f4null(lpField);	
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

/* *********************************************************************************** */
/* Utility function that determines what the field is from a field name passed as a    */
/* char by reference.  On success returns the FIELD4 pointer.  On failure returns      */
/* NULL and sets the appropriate error messages so the calling program doesn't have to.*/
FIELD4* cbxGetFieldFromName(char* lpcFieldName)
{
	FIELD4* lpReturn = NULL;
	DATA4*   lpTable = NULL;
	long lnReturn = TRUE;
	char lcAliasName[200];
	char lcFieldName[130];
	char lcOriginal[130];
	char* lcPtr;
	long lnRecord;	
	strcpy(lcOriginal, lpcFieldName);
	
	lcPtr = strstr(lcOriginal, ".");
	if (lcPtr != NULL)
		{
		*lcPtr = (char) 0;
		strcpy(lcAliasName, lcOriginal);
		strcpy(lcFieldName, lcPtr + 1);
		lpTable = code4data(&codeBase, lcAliasName);
		if (lpTable == NULL)
			{
			strcpy(gcErrorMessage, "Unidentified Table Alias");
			gnLastErrorNumber = -9984;
			lnReturn = FALSE;
			}			
		}
	else
		{
		if (gpCurrentTable == NULL)
			{
			strcpy(gcErrorMessage, "No Table Open in Selected Area");
			gnLastErrorNumber = -9999;
			lnReturn = FALSE;
			}
		else
			{
			lpTable = gpCurrentTable;
			strcpy(lcFieldName, lcOriginal);
			}	
		}
	if (lnReturn)
		{
		if (lpTable != NULL) lnRecord = d4recNo(lpTable);
		else lnRecord = -1;
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
			strcpy(gcErrorMessage, "Field Not Found");
			gnLastErrorNumber = -9985;
			lnReturn = FALSE;	
			}
		}
	if (lpReturn == NULL) gcStrReturn[0] = (char) 0;
	return(lpReturn);	
}

long cbxParseDateByFormat(char*, char*, char*, char*);

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
	
	memset(lcYear, 0, (size_t) (10 * sizeof(char)));
	memset(lcMonth, 0, (size_t) (10 * sizeof(char)));
	memset(lcDay, 0, (size_t) (10 * sizeof(char)));
	lcPattSep[0] = lcPattSep[1] = lcSeparator[0] = lcSeparator[1] = (char) 0;

	strcpy(lcWorkDate, MPSSStrTran(lpcDateStr, " ", "", 0)); // Removes spaces throughout the string, as they are meaningless.
	lnStrLen = (long) strlen(lcWorkDate);

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
			//printf("The length is %ld \n", lnStrLen);
			switch(lnStrLen)
				{
				case 6: // 6 character dates are never a good idea for Y2K reasons, but they are still with us.
					strncpy(lcYear, lcWorkDate, 2);
					strncpy(lcMonth, lcWorkDate + 2, 2);
					strncpy(lcDay, lcWorkDate + 4, 2);
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
						strncpy(lcDay, lcWorkDate + 2, 2);
						strncpy(lcYear, lcWorkDate + 4, 2);
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
	char cPrevChar = 0;
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
	lnLen = (long) strlen(lpcPattern); 
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
		lnLen = (long) strlen(lpcDateStr);
		// Now we parse this into Year, Month and Day strings...
		lnPos = 0;
		memset(lcWork, 0, (size_t) (20 * sizeof(char)));
		for (jj = 0; jj <= lnLen; jj++)
			{
			cTemp = lpcDateStr[jj];
			if ((cTemp != cDateSep) && (cTemp != 0))
				{
				if (!isdigit(cTemp))
					{
					bGoodDate = FALSE;
					//printf("Not Is Digit %c \n", cTemp);
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
						break;
						
					case 'D':
						strncpy(lcDay, lcWork, 5);
						break;
						
					case 'Y':
						strncpy(lcYear, lcWork, 5);
						break;
						
					default:
						bGoodSchema = FALSE;
						//printf("Case is BAD %c \n", cSchema[nSegCount - 1]);
						break;	
					}
				memset(lcWork, 0, (size_t) (20 * sizeof(char)));
				lnPos = 0;
				}	
			}
		//printf("Good Date is %ld \n", bGoodDate);
		if (bGoodDate)
			{
			//printf("Year %s  Month %s  Day %s \n", lcYear, lcMonth, lcDay);
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
				//printf("Evaluating >%s< \n", lcWork);					
				lnDateReturn = date4long(lcWork);
				if (lnDateReturn < 0) lnDateReturn = 0;
				}
			}
		}
	return lnDateReturn;	
}

/* ********************************************************************************** */
/* Generic function for converting source data in string/char format to a CodeBase    */
/* field type.  Used in the append functions below.                                   */
void cbxStoreCharValue(long lpnType, FIELD4* lppField, char* lpcString)
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
			//printf("parsed date to %ld \n", nTest);
			f4assignLong(lppField, nTest);
			break;
			
		case 'I':
			nTest = atol(lpcString);
			f4assignLong(lppField, nTest);
			break;
			
		default:
			break;
		}	
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
	char lcTargAlias[40];
	char lcSrcAlias[40];
	FieldPlus fpSrc;
	FieldPlus fpTrg;
	
	strcpy(lcTargAlias, cbwALIAS());
	strcpy(lcSrcAlias, "CBWSRC");
	strcat(lcSrcAlias, MPStrRand(6));
	StrToUpper(lcSrcAlias);
	lbOK = cbwUSE(lpcSource, lcSrcAlias, TRUE);
	if (lbOK)
		{
		cbwSELECT(lcSrcAlias);
		if ((lpcTestExpr != NULL) && (lpcTestExpr[0] != (char) 0))
			{
			nTest = cbwPREPAREFILTER(lpcTestExpr); // If this fails, we're done!	
			if (nTest == 1) bNeedTest = TRUE;
			}
		if (nTest == 1)
			{
			lpSourceTable = gpCurrentTable;
			// Now we get the field type information and write it into the public arrays.
			lpFld = cbwAFIELDS(&lnCnt);
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
							(gaTrgFields[kk].nFldType == gaSrcFields[jj].nFldType) && (gaTrgFields[kk].nFldType != (long) 'M') && (gaTrgFields[kk].nFldType != (long)'G'))
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
			kk = cbwGOTO("TOP", 0);
			
			while (cbwEOF() == 0)
				{
				cbwSELECT(lcSrcAlias);
				if (bNeedTest)
					{
					if (cbwTESTFILTER() != 1)
						{
						if (cbwGOTO("NEXT", 0) < 1) break;
						if (cbwEOF() != 0) break;
						continue;
						}
					}
				cbwSELECT(lcTargAlias);
				if (cbwAPPENDBLANK() == TRUE)
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
							//printf("About to test for null of pointer %p \n", lpSrcField);
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
								nLen = f4memoLen(lpSrcField);
								f4memoAssignN(lpTrgField, f4memoPtr(lpSrcField), nLen);
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
				cbwSELECT(lcSrcAlias);
				if (cbwGOTO("NEXT", 0) < 1) break;
				if (cbwEOF() != 0) break;
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
		cbwCLOSETABLE(lcSrcAlias);	
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
	char lcTargAlias[30];
	FILE* pHandle;
	char cLineBuff[10000];
	char cFieldStr[500];
	char* pPtr = NULL;
	long lnLineLen;
	char nType;
	long lnAppendCount = 0;
	long lnCurrPos = 0;
	FIELD4* lpTrgField;
	long bGoodLoad = TRUE;
	
	strcpy(lcTargAlias, cbwALIAS());
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
				lnLineLen = (long) strlen(cLineBuff);
				if (lnLineLen > laFldLength[0])  // Has to be at least as much data as in the first field.
					{
					cbwSELECT(lcTargAlias);
					if (cbwAPPENDBLANK() == TRUE)
						{
						lnAppendCount += 1;
						lnCurrPos = 0;
						for (jj = 0; jj < gnTrgCount; jj++)
							{
							lpTrgField = gaTrgFields[jj].pCBfield;
							nType = (char) gaTrgFields[jj].nFldType;
							if (lnCurrPos < lnLineLen)
								{
								strncpy(cFieldStr, cLineBuff + laFromMark[jj], laFldLength[jj]);	
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
			cbwSELECT(lcTargAlias);
			cbwFLUSH();
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


DOEXPORT long cbwAPPENDFROM(char* lpcAlias, char* lpcSource, char* lpcTestExpr, char* lpcType)
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
			StrToUpper(lcWorkName);
			lpChar = strrchr(lcWorkName, '.');
			memset(lcExt, 20, sizeof(char));
			if (lpChar != NULL)
				{
				strncpy(lcExt, lpChar + 1, 19);
				}
			if ((lpcType != NULL) && (lpcType[0] != '\0'))
				{
				// They have passed an lpcType string, so it takes precedence.
				strncpy(lcExt, lpcType, 19);
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
				cbwSELECT(lcWorkAlias);
				// Now we get the field type information and write it into the public arrays.
				lpFld = cbwAFIELDS(&lnCnt);
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

/* ********************************************************************************** */
/* Support functions for the cbwCOPYTO() function.                                    */

long cbxCopyToDBF(DATA4* lppSource, char* lpcOutput, char* lpcTestExpr)
{
	long lnResult = 0;
	long bOK;
	long jj;
	char lcTargAlias[50];
	char lcSrceAlias[50];
	char lcStructure[10000];
	char lcFldSpec[100];
	char lcNumber[20];
	char lcSourceDBF[300];
	
	strcpy(lcSrceAlias, cbwALIAS());
	strcpy(lcSourceDBF, cbwDBF(lcSrceAlias));

	// First we have to create the table...
	memset(lcStructure, 0, (size_t) (10000 * sizeof(char)));

	for (jj = 0; jj < gnSrcCount; jj++)
		{
		if (!gaSrcFields[jj].bMatched) continue;
		memset(lcFldSpec, 0, (size_t) (100 * sizeof(char)));
		strcpy(lcFldSpec, gaSrcFields[jj].xField.cName);
		StrToUpper(lcFldSpec);
		strcat(lcFldSpec, ",");
		strcat(lcFldSpec, gaSrcFields[jj].xField.cType);
		strcat(lcFldSpec, ",");
		sprintf(lcNumber, "%ld", gaSrcFields[jj].xField.nWidth);
		strcat(lcFldSpec, lcNumber);
		strcat(lcFldSpec, ",");
		sprintf(lcNumber, "%ld", gaSrcFields[jj].xField.nDecimals);
		strcat(lcFldSpec, lcNumber);
		strcat(lcFldSpec, ",");
		if (gaSrcFields[jj].xField.bNulls != 0) strcat(lcFldSpec, "TRUE\n");
		else strcat(lcFldSpec, "FALSE\n");
		strcat(lcStructure, lcFldSpec);
		}

	bOK = cbwCREATETABLE(lpcOutput, lcStructure);	
	if (bOK != TRUE)
		{
		// Error already is set.
		lnResult = -1; // just in case we forgot it above.	
		}
	else
		{
		strcpy(lcTargAlias, cbwALIAS());
		if (bOK == TRUE)
			{
			lnResult = cbwAPPENDFROM(lcTargAlias, lcSourceDBF, lpcTestExpr, "DBF");
			}
		}
	cbwCLOSETABLE(lcTargAlias);
	
	return lnResult;	
}

/* *************************************************************************************************** */
/* Utility function fo the cbxCopyToTextFile() function that does all the thinking about how to turn   */
/* the field data into a string for output.                                                            */
char *cbxMakeTextOutputField(VFPFIELD* pStructFld, long lpnOutType, FIELD4* pContentFld, char* cBuff, long lpnCnt, long lpbStrip)
{
	char cWorkBuff[500];
	long nWidth;
	long nDecimals;
	long jj;
	double dValue;
	char cType;
	char cDelimiter = 0;
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
				nLen = (long) strlen(cTempNum);
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
				nLen = (long) strlen(cTempNum);
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
			//strcat(cBuff, "|");
			break;
			
		case 3: // CSV
			//printf("WorkBuff >%s<\n", cWorkBuff);
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
	long lnOutCnt;
	long lnRowCnt = 0;
	char lcOutBuff[20000];
	FILE* pHandle;
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
			lnTest = cbwPREPAREFILTER(lpcTestExpr);
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
			cbwGOTO("TOP", 0);
			while (!cbwEOF())
				{
				if (bNeedTest)
					{
					if (cbwTESTFILTER() != 1)
						{
						if (cbwGOTO("NEXT", 0) < 1) break;
						if (cbwEOF() != 0) break;	
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
					//printf("Field %s  Type %c \n", gaSrcFields[jj].xField.cName, (char) lnType);
					strcat(lcOutBuff, cbxMakeTextOutputField(&(gaSrcFields[jj].xField), lpnOutType, gaSrcFields[jj].pCBfield, lcFldBuff, lnOutCnt, lpbStrip));
					lnOutCnt += 1;
					}
				strcat(lcOutBuff, "\n");
				fputs(lcOutBuff, pHandle);
				lnRowCnt += 1;
				if (cbwGOTO("NEXT", 0) < 1) break;
				if (cbwEOF() != 0) break;
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
DOEXPORT long cbwCOPYTO(char* lpcAlias, char* lpcOutput, char* cType, long bHeader, char* cFieldList, char* cTestExpr, long bStripBlanks)
{
	long lnResult = 0;	
	DATA4 *lpSourceTable = NULL;
	long lnPt = 0;
	long lnType = -1; // Types 1 = DBF, 2 = SDF or DAT, 3 = CSV, 4 = TAB
	char *lpChar = NULL;
	char *lpTest = NULL;
	char *pPtr = NULL;
	long lbByType = FALSE;
	char lcWorkAlias[50];
	long lnOutFldCnt = 0;
	long bFieldTest = FALSE;
	VFPFIELD* lpFld;
	long lnCnt = 0;
	long jj;
	DATA4 *lpOldTable = NULL;
	char lcFldMatch[5000];
	char lcFldTest[25];
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpOldTable = gpCurrentTable;
		}
	
	if ((lpcAlias == NULL) || (lpcAlias[0] == '\0'))
		{
		lpSourceTable = gpCurrentTable;	
		}
	else
		{
		lpSourceTable = code4data(&codeBase, lpcAlias);	
		}
		
	if (lpSourceTable == NULL)
		{
		strcpy(gcErrorMessage, "No Source Table Available or Alias Not Found");
		gnLastErrorNumber = -7980;
		lnResult = -1;
		}
	else
		{
		if ((lpcOutput == NULL) || (lpcOutput[0] == '\0'))
			{
			strcpy(gcErrorMessage, "Output File Name not specified");
			gnLastErrorNumber = -7981;
			lnResult = -1;
			}
		else
			{
			strcpy(lcWorkAlias, d4alias(lpSourceTable));
			cbwSELECT(lcWorkAlias);
			while (TRUE)
				{
				if (strcmp(cType, "DBF") == 0)
					{
					lnType = 1;
					break;
					}
				
				if (strcmp(cType, "SDF") == 0)
					{
					lnType = 2;
					break;
					}
				
				if (strcmp(cType, "DAT") == 0) {lnType = 2; break;}
				
				if (strcmp(cType, "CSV") == 0) {lnType = 3; break;}
				
				if (strcmp(cType, "TAB") == 0) {lnType = 4; break;}
				
				break;
				}
			}
		if (lnType != -1)
			{			
			// Now we get the field type information and write it into the public arrays.
			if ((cFieldList != NULL) && (cFieldList[0] != (char) 0))
				{
				strcpy(lcFldMatch, ",");
				strcat(lcFldMatch, cFieldList);
				strcat(lcFldMatch, ","); // Ensures we can search for ',MYFIELD,' in this string.
				StrToUpper(lcFldMatch); // In place uppercase.
				bFieldTest = TRUE;
				}			
			lcFldTest[0] = (char) 0;
			
			lpFld = cbwAFIELDS(&lnCnt);
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
				switch(lnType)
					{
					case 1:
						lnResult = cbxCopyToDBF(lpSourceTable, lpcOutput, cTestExpr);
						break;
	
					case 2:
						lnResult = cbxCopyToTextFile(lpSourceTable, lpcOutput, 2, cTestExpr, FALSE, FALSE); // Fixed sizes, so no headers or blank stripping.
						break;
						
					case 3:
						lnResult = cbxCopyToTextFile(lpSourceTable, lpcOutput, 3, cTestExpr, bHeader, bStripBlanks);
						break;
						
					case 4:
						lnResult = cbxCopyToTextFile(lpSourceTable, lpcOutput, 4, cTestExpr, bHeader, bStripBlanks);
						
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

	return lnResult;	
}

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

DOEXPORT char* cbwSCATTER(char* lpcDelimiter, long lpbStripBlanks, char* lpcFieldList)
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
						lcTypeBuff[jj + 1] = 0;	
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
							c4trimN(lcPtr, (int) strlen(lcPtr) + 1);	
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


/* *********************************************************************************** */
/* Exported function that taks a field name like cbxGetFieldFromName() and returns     */
/* the string pointer to a text value containing the field type code 'C', 'D', etc.    */
/* Returns an empty string if the field isn't found.  If it is, the FIELD4 pointer     */
/* value is stored in gpLastField for future retrieval.                                */
DOEXPORT char* cbwGETFIELDTYPE(char* lpcFieldName)
{
	FIELD4* lpField;
	char lcType;
	
	gcStrReturn[0] = (char) 0;
	gpLastField = NULL;
	
	lpField = cbxGetFieldFromName(lpcFieldName);
	if (lpField != NULL)
		{
		lcType = f4type(lpField);
		gcStrReturn[0] = lcType;
		gcStrReturn[1] = (char) 0;
		gpLastField = lpField;	
		}
	return(gcStrReturn);	
}

/* ************************************************************************************* */
/* Completely generic function that returns a text string representation of a specified  */
/* field.  Returns the type as a character value in the second parm passed by reference. */
/* Like the SCATTER functions, you can pass a table qualified name "MYTABLE.MYFIELD" to  */
/* refer to a table other than the currently selected one.                               */

DOEXPORT char* cbwCURVAL(char* lpcFieldName, char* lpcFieldType)
{
	char *lcPtr = NULL;
	FIELD4 *lpField;
	int  lcType;
	double lnDoubleValue = 0.0;
	long   lnLongValue = 0;
	char   lcNumBuff[50];
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	memset(lcNumBuff, 0, (size_t) 50 * sizeof(char));
	gcStrReturn[0] = (char) 0;

	lpField = cbxGetFieldFromName(lpcFieldName);
	
	if (lpField != NULL)
		{
		lcType = f4type(lpField);
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
			*lpcFieldType = (char) 0;	
			}
		else
			{
			strcpy(gcStrReturn, lcPtr);
			*lpcFieldType = lcType;
			*(lpcFieldType + 1) = (char) 0;
			}		
		}
	else
		{
		//gnLastErrorNumber = codeBase.errorCode;
		//strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
		gcStrReturn[0] = (char) 0; /* We're done... ERROR */
		*lpcFieldType = (char) 0;
		}

	return(gcStrReturn);	
}

/* ************************************************************************************ */
/* Function that obtains the value of one field, specified by either just the field     */
/* name (for the currently selected table) or the full alias.fieldname notation.        */
/* Returns the field value in a character buffer to which the return value points.      */
/* The string may be empty, in which case check the last error message to be sure that  */
/* this was not caused by an error.                                                     */
/* New March 18, 2014.  Like cbwSCATTERFIELD, but if lbTrim is TRUE, then trims leading */
/* and trailing blanks from the string.                                                 */
DOEXPORT char* cbwSCATTERFIELDCHAR(char* lpcFieldName, long lbTrim)
{
	char *lcPtr;
	FIELD4 *lpField;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lpField = cbxGetFieldFromName(lpcFieldName);
	
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
			if (lbTrim) MPStrTrim('A', gcStrReturn);				
			}			
		}

	return(gcStrReturn);	
}

/* *********************************************************************************** */
/* Function that obtains the value of one field, specified by either just the field    */
/* name (for the currently selected table) or the full alias.fieldname notation.       */
/* Returns the field value in a character buffer to which the return value points.     */
/* The string may be empty, in which case check the last error message to be sure that */
/* this was not caused by an error.                                                    */
DOEXPORT char* cbwSCATTERFIELD(char* lpcFieldName)
{
	char *lcPtr;
	FIELD4 *lpField;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lpField = cbxGetFieldFromName(lpcFieldName);
	
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

	return(gcStrReturn);	
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
DOEXPORT long cbwSCATTERFIELDLONG(char* lpcFieldName, long* lpnValue)
{
	long lnReturn = TRUE;
	long lnValue = 0;
	FIELD4 *lpField;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lpField = cbxGetFieldFromName(lpcFieldName);

	if (lpField != NULL)
		{
		lnValue = f4long(lpField);
		}
	else lnReturn = -1;
	*lpnValue = lnValue;
	return(lnReturn);	
}

/* *********************************************************************************** */
/* Function that obtains the value of an numeric or float field with decimals as a     */
/* double in a double parameter passed by reference.  Returns 1 on OK, -1 on ERROR.    */
/* See cbwSCATTERFIELD() for details on passing field name/table name info.            */
/* NOTE: It is up to the caller to make sure that this field contains a value which    */
/* is appropriately converted to a double.  It will convert a charcter field containing  */
/* a string of numbers, a date (converted to a Julian day number), an integer, or a    */
/* Number field.                                                                       */
/* Using this on a field containing alphabetic character information will return a 0.  */
DOEXPORT long cbwSCATTERFIELDDOUBLE(char* lpcFieldName, double* lpnValue)
{
	long lnReturn = TRUE;
	double lnValue = 0.0;
	FIELD4 *lpField;

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lpField = cbxGetFieldFromName(lpcFieldName);

	if (lpField != NULL)
		{
		lnValue = f4double(lpField);
		}
	else lnReturn = -1;
	*lpnValue = lnValue;
	return(lnReturn);	
}

/* *********************************************************************************** */
/* Function that obtains the value of a date field and returns the date value in three */
/* parms year, month, day passed by reference.  Returns 1 on OK, -1 on ERROR.          */
/* See cbwSCATTERFIELD() for details on passing field name/table name info.            */
/* NOTE: This function must be applied ONLY to a DATE or DATETIME field.  If the field */
/* is any other type, an error "Bad Field Type" is returned and the parm values are    */
/* undefined.  If field is DATETIME, only the DATE portion is returned.                */
DOEXPORT long cbwSCATTERFIELDDATE(char* lpcFieldName, long* lpnYear, long* lpnMonth, long* lpnDay)
{
	long lnReturn = TRUE;
	long lnYear = 0;
	long lnMonth = 0;
	long lnDay = 0;
	char *lcPtr;
	int  lcType;
	FIELD4 *lpField;
	char lcDateBuff[50];
	char lcNumBuff[20];

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lcNumBuff[0] = (char) 0;
	memset(lcDateBuff, 0, (size_t) 50 * sizeof(char));
	
	lpField = cbxGetFieldFromName(lpcFieldName);
	
	if (lpField != NULL) /* Test for Date or DateTime type */
		{
		lcType = f4type(lpField);
		if ((lcType == (int) 'D') || (lcType == (int) 'T'))
			lnReturn = TRUE;
		else
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "Not a Date/Time Type");
			gnLastErrorNumber = -9983;
			}	
		}
	else lnReturn = FALSE;

	if (lnReturn)
		{
		if (lcType == (int) 'D') /* A DATE, so extract accordingly */
			{
			lcPtr = f4str(lpField);
			}
		else /* A DATETIME, so extract that differently */
			{
			lcPtr = f4dateTime(lpField);
			}
		if (lcPtr == NULL) /* Error Condition */
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "Invalid Date Returned");
			gnLastErrorNumber = -9982;
			}
		else
			{
			strcpy(lcDateBuff, lcPtr);
			strncpy(lcNumBuff, lcDateBuff, 4);
			lcNumBuff[4] = (char) 0;
			lnYear = strtol(lcNumBuff, NULL, 10);
			strncpy(lcNumBuff, lcDateBuff + 4, 2);
			lcNumBuff[2] = (char) 0;
			lnMonth = strtol(lcNumBuff, NULL, 10);
			strncpy(lcNumBuff, lcDateBuff + 6, 2);
			lcNumBuff[2] = (char) 0;
			lnDay = strtol(lcNumBuff, NULL, 10);
			}		
		}
	*lpnYear  = lnYear;
	*lpnMonth = lnMonth;
	*lpnDay   = lnDay;
	return(lnReturn ? TRUE : -1);	
}

/* *********************************************************************************** */
/* Function that obtains the value of a datetime or date field and returns the datetime*/
/* value in six parms year, month, day, hours, minutes, seconds passed by reference.   */
/* Returns 1 on OK, -1 on ERROR.                                                       */
/* See cbwSCATTERFIELD() for details on passing field name/table name info.            */
/* NOTE: This function must be applied ONLY to a DATE or DATETIME field.  If the field */
/* is any other type, an error "Bad Field Type" is returned and the parm values are    */
/* undefined.  If field is DATE, the hours, minutes, and seconds are returned as 0.    */
DOEXPORT long cbwSCATTERFIELDDATETIME(char* lpcFieldName, long* lpnYear, long* lpnMonth, long* lpnDay, long* lpnHour, long* lpnMinute, long* lpnSecond)
{
	long lnReturn = TRUE;
	long lnYear = 0;
	long lnMonth = 0;
	long lnDay = 0;
	long lnHour = 0;
	long lnMinute = 0;
	long lnSecond = 0;
	char *lcPtr;
	int  lcType = (char) 0;
	FIELD4 *lpField;
	char lcDateBuff[50];
	char lcNumBuff[20];

	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lcNumBuff[0] = (char) 0;
	memset(lcDateBuff, 0, (size_t) 50 * sizeof(char));
	
	lpField = cbxGetFieldFromName(lpcFieldName);

	if (lpField != NULL) /* Test for Date or DateTime type */
		{
		lcType = f4type(lpField);
		if ((lcType == (int) 'D') || (lcType == (int) 'T'))
			{
			lnReturn = TRUE;
			}
		else
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "Not a Date/Time Type");
			gnLastErrorNumber = -9983;
			}	
		}
	else
		{
		lnReturn = FALSE;
		}

	if (lnReturn)
		{
		if (lcType == (int) 'D') /* A DATE, so extract accordingly */
			{
			lcPtr = f4str(lpField);
			}
		else /* A DATETIME, so extract that differently */
			{
			lcPtr = f4dateTime(lpField);
			}
		if (lcPtr == NULL) /* Error Condition */
			{
			lnReturn = FALSE;
			strcpy(gcErrorMessage, "Invalid Date Returned");
			gnLastErrorNumber = -9982;
			}
		else
			{
			strcpy(lcDateBuff, MPSSStrTran(lcPtr, ":", "", 0));
			if (lcType == (int) 'D') strcat(lcDateBuff, "000000"); /* The time values are all 0 on date types */
			strncpy(lcNumBuff, lcDateBuff, 4);
			lcNumBuff[4] = (char) 0;
			lnYear = strtol(lcNumBuff, NULL, 10);
			
			strncpy(lcNumBuff, lcDateBuff + 4, 2);
			lcNumBuff[2] = (char) 0;
			lnMonth = strtol(lcNumBuff, NULL, 10);
			
			strncpy(lcNumBuff, lcDateBuff + 6, 2);
			lcNumBuff[2] = (char) 0;
			lnDay = strtol(lcNumBuff, NULL, 10);
			
			strncpy(lcNumBuff, lcDateBuff + 8, 2);
			lcNumBuff[2] = (char) 0;
			lnHour = strtol(lcNumBuff, NULL, 10);
			
			strncpy(lcNumBuff, lcDateBuff + 10, 2);
			lcNumBuff[2] = (char) 0;
			lnMinute = strtol(lcNumBuff, NULL, 10);
			
			strncpy(lcNumBuff, lcDateBuff + 12, 2);
			lcNumBuff[2] = (char) 0;
			lnSecond = strtol(lcNumBuff, NULL, 10);			
			}		
		}
	*lpnYear  = lnYear;
	*lpnMonth = lnMonth;
	*lpnDay   = lnDay;
	*lpnHour  = lnHour;
	*lpnMinute = lnMinute;
	*lpnSecond = lnSecond;
	return(lnReturn ? TRUE : -1);	
}

/* *********************************************************************************** */
/* Function that obtains the value of a Logical field, as per the above functions.     */
/* Returns 1 for a value of TRUE in the field.  Returns 0 for a value of FALSE, and    */
/* returns -1 on any kind of error.  Returns ERROR if the field is not LOGICAL.        */
DOEXPORT long cbwSCATTERFIELDLOGICAL(char* lpcFieldName)
{
	FIELD4 *lpField;
	long lnReturn = -1;
	long lbResult = FALSE;
	int lcType;
	long lbValue;
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lpField = cbxGetFieldFromName(lpcFieldName);
	if (lpField != NULL) /* Test for Date or DateTime type */
		{
		lcType = f4type(lpField);
		if (lcType == (int) 'L')
			lbResult = TRUE;
		else
			{
			lbResult = FALSE;
			strcpy(gcErrorMessage, "Not a Logical Type");
			gnLastErrorNumber = -9982;
			}	
		}
			
	if (lbResult)
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

	return(lnReturn);	
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
	unsigned long lnULen;
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
			lnULen = (unsigned long) strlen(lpcValue); /* In this function only character values (no NULLs) can be handled, but we can still save it into binary fields */
			if (f4len(lppField) < lnULen) lnULen = f4len(lppField);
			strncpy(f4assignPtr(lppField), lpcValue, (size_t) lnULen); /* Raw copy of the bytes into the field data buffer */
			break;
			
		case 'G': /* General - contains OLE objects */
			/* No Action at this time */
			break;
			
		case 'X': /* Binary Memo Field */
			lnLen = (long) strlen(lpcValue);
			f4memoAssignN(lppField, lpcValue, lnLen); /* we copy in plain text, but it's going into a field that can handle binary. */
			break;
				
		default: /* Shouldn't get here with VFP data, but ya never know */
			break;
				
		}
	return(lnReturn);		
}

/* ********************************************************************************** */
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
DOEXPORT long cbwREPLACE_CHARACTER(char* lpcFieldName, char* lpcValue)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpField = cbxGetFieldFromName(lpcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			/* Ok, we have a valid field in the current table */
			f4assign(lpField, lpcValue);	
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
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
/* This takes the date value as a julian day number as recognized by VFP.  The julian */
/* day number of Jan. 1, 1990 is 2447893, for example.                                */
/* If the value passed is < 0 or otherwise invalid, the result is undefined.  A value */
/* of 0 results in an empty date field.                                               */
DOEXPORT long cbwREPLACE_DATEN(char* lpcFieldName, long lpnJulian)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
			
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
 
	if (gpCurrentTable != NULL)
		{
		lpField = cbxGetFieldFromName(lpcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			if (f4type(lpField) != 'D')
				{
				lnReturn = -1;
				strcpy(gcErrorMessage, "Not a Date Field");
				gnLastErrorNumber = -8025;
				}
			else
				{
				f4assignLong(lpField, lpnJulian); /* Can be 0 on an empty date which is valid in a VFP table. */
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
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
/* This takes the date value as a character string in the form:                       */
/* YYYYMMDD                                                                           */
/* IF THE VALUE PASSED IS NOT A VALID DATE REPRESENTATION, AN ERROR IS REPORTED.      */
DOEXPORT long cbwREPLACE_DATEC(char* lpcFieldName, char* lpcValue)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
	long lnLongValue;
			
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

 
	if (gpCurrentTable != NULL)
		{
		lpField = cbxGetFieldFromName(lpcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			if (strcmp(lpcValue, "        ") == 0)
				{
				lnLongValue = 0;
				f4assignLong(lpField, lnLongValue);
				}
			else
				{
				lnLongValue = date4long(lpcValue);
				if (lnLongValue < 0)
					{
					/* An error condition -- bad date format */
					strcpy(gcErrorMessage, "Bad Date Format");
					gnLastErrorNumber = -8000;
					lnReturn = -1;	
					}
				else
					{
					if (f4type(lpField) != 'D')
						{
						strcpy(gcErrorMessage, "Not a Date Field");
						gnLastErrorNumber = -8025;
						lnReturn = -1;	
						}
					else
						{
						f4assignLong(lpField, lnLongValue); /* Can be 0 on an empty date which is valid in a VFP table. */	
						}
					}
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
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
/* This takes the datetime value as a character string in the form:                   */
/* YYYYMMDDHHMMSS                                                                     */
/* IF THE VALUE PASSED IS NOT A VALID DATETIME REPRESENTATION, THE FIELD IS SET TO    */
/* THE VALUE OF AN EMPTY DATETIME AND NO ERROR IS REPORTED.  PLAN ACCORDINGLY!        */
DOEXPORT long cbwREPLACE_DATETIMEC(char* lpcFieldName, char* lpcValue)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
	char lcDTBuff[20];
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	strncpy(lcDTBuff, lpcValue, 10);
	lcDTBuff[10] = ':';
	strncpy(lcDTBuff + 11, lpcValue + 10, 2);
	lcDTBuff[13] = ':';
	strncpy(lcDTBuff + 14, lpcValue + 12, 2);
	lcDTBuff[16] = (char) 0;
 
	if (gpCurrentTable != NULL)
		{
		lpField = cbxGetFieldFromName(lpcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			/* Ok, we have a valid field in the current table */
			f4assignDateTime(lpField, lcDTBuff);	
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
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
/* This takes the datetime value as a set of 6 long values representing the Year,     */
/* Month, Day, Hour, Minute, and Second.  If these values are not a valid represen-   */
/* tation of a datetime value, an ERROR is generated (return value of -1).            */
DOEXPORT long cbwREPLACE_DATETIMEN(char* lpcFieldName, long lpnYr, long lpnMo, long lpnDa, long lpnHr, long lpnMi, long lpnSe)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
	char lcDTBuff[20];
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	if ((lpnYr < 10) || (lpnYr > 4000) || (lpnMo < 1) || (lpnMo > 12) || 
		(lpnDa < 0) || (lpnDa > 31) || (lpnHr < 0) || (lpnHr > 24) ||
		(lpnMi < 0) || (lpnMi > 59) || (lpnSe < 0) || (lpnSe > 59))
		{
		lnReturn = -1;
		gnLastErrorNumber = -8059;
		strcpy(gcErrorMessage, "Invalid Datetime input values");	
		}
 	else
 		{
		if (gpCurrentTable != NULL)
			{
			lpField = cbxGetFieldFromName(lpcFieldName);
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
		else
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, "No Table Open in Selected Area");
			gnLastErrorNumber = -9999;			
			}
		}
	return(lnReturn);		
}

/* ********************************************************************************** */
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
DOEXPORT long cbwREPLACE_CURRENCY(char* lpcFieldName, char* lpcValText)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpField = cbxGetFieldFromName(lpcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			if (f4type(lpField) != 'Y')
				{
				strcpy(gcErrorMessage, "Not a Currency Field");
				gnLastErrorNumber = -8025;
				lnReturn = -1;
				}
			else
				{	
				/* Ok, we have a valid field in the current table */
				f4assignCurrency(lpField, lpcValText);
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
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
DOEXPORT long cbwREPLACE_LONG(char* lpcFieldName, long lpnValue)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpField = cbxGetFieldFromName(lpcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			/* Ok, we have a valid field in the current table */
			f4assignLong(lpField, lpnValue); /* Allowed to update integer and number fields */
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
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
DOEXPORT long cbwREPLACE_DOUBLE(char* lpcFieldName, double lpnValue)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpField = cbxGetFieldFromName(lpcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			/* Ok, we have a valid field in the current table */
			f4assignDouble(lpField, lpnValue);	
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
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
DOEXPORT long cbwREPLACE_LOGICAL(char* lpcFieldName, long lpnValue)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
	char lcFlag;
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpField = cbxGetFieldFromName(lpcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			/* Ok, we have a valid field in the current table */
			lcFlag = (lpnValue == 1 ? 'T' : 'F');
			f4assignChar(lpField, (int) lcFlag);			
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
/* One of a set of functions that works as the REPLACE command in VFP but for just    */
/* one field in the currently selected table.  Each one is geared to one type of      */
/* field content.                                                                     */
/* Returns 1 on success, -1 on ERROR.                                                 */
DOEXPORT long cbwREPLACE_MEMO(char* lpcFieldName, char* lpcValue)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	FIELD4 *lpField;
		
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lpField = cbxGetFieldFromName(lpcFieldName);
		if (lpField == NULL)
			{
			lnReturn = -1; /* Error messages are already set */
			}
		else
			{
			/* Ok, we have a valid field in the current table */
			f4memoAssign(lpField, lpcValue);	
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
DOEXPORT long cbwGATHERMEMVAR(char* lpcFields, char* lpcValues, char* lpcDelimiter)
{
	long lnReturn = TRUE;
	long lnResult;
	char lcDelim[32];
	long lnFldCnt = 0;   /* The number of fields in the table */
	long lnFldInCnt = 0; /* The actual number of field names submitted */
	long lnValueCnt = 0; /* The number of field values found */
	register long jj;
	long nn;
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
	memset(laFields, 0, (size_t) 256 * sizeof(char*));
	memset(laValues, 0, (size_t) 256 * sizeof(char*));
	lcBadFieldName[0] = (char) 0;
	lcFieldName[0] = (char) 0;
	//printf("The Delim %s \n", lpcDelimiter);
	if (strlen(lpcDelimiter) == 0)
		{
		lcDelim[0] = '~';
		lcDelim[1] = '~';
		}
	else
		{
		strncpy(lcDelim, lpcDelimiter, 30);
		}
	
	if (gpCurrentTable != NULL)
		{
		lnFldCnt = d4numFields(gpCurrentTable);
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
			//printf("Big Value String: \n%s\n", lpcValues);
			nn = 0;
			while(TRUE)
				{
				lpPtr = strtokX(lpTemp, lcDelim);
				lpTemp = NULL; /* from here on out */
				if (lpPtr == NULL) break;
				laValues[lnValueCnt] = malloc(strlen(lpPtr) + 1);
				strcpy(laValues[lnValueCnt], lpPtr);
				//nn += 1;
				//printf("Value %ld >%s< \n", nn, lpPtr);
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
					laFields[lnFldInCnt] = d4field(gpCurrentTable, lpPtr);
					if (laFields[lnFldInCnt] == NULL)
						{
						strcpy(lcBadFieldName, lpPtr);	
						}
//					strcpy(lcFieldName, lpPtr);
//					laFields[lnFldInCnt] = d4field(gpCurrentTable, lcFieldName);
//					if (laFields[lnFldInCnt] == NULL)
//						{
//						strcpy(lcBadFieldName, lcFieldName);	
//						}
					lnFldInCnt += 1;	
					}
				lnFldCnt = lnFldInCnt;	
				}
			else
				{
				for (jj = 1; jj <= lnFldCnt; jj++)
					{
					laFields[jj - 1] = d4fieldJ(gpCurrentTable, (short) jj); /* all of the fields will be supplied (we hope) and in their proper order */
					}	
				}
			if (lnFldCnt != lnValueCnt) /* Error condition.  We have to have one-to-one matching */
				{
				lnReturn = -1;
				//strcpy(gcErrorMessage, "Field and Value count mismatch");
				sprintf(gcErrorMessage, "Field and Value count mismatch: %ld and %ld", lnFldCnt, lnValueCnt);
				gnLastErrorNumber = -9980;
				}
			else
				{
				for (jj = 0;jj < lnFldCnt;jj++)
					{
					lnResult = cbxUpdateField(laFields[jj], laValues[jj]);
					//printf("Fld %s  Value %s \n", f4name(laFields[jj]), laValues[jj]);
					if (lnResult == -1)
						{
						lnReturn = -1; /* Something went wrong and we can't continue */
						//gnLastErrorNumber = codeBase.errorCode;
						//strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
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
DOEXPORT long cbwINSERT(char* lpcValues, char* lpcFldNames, char* lpcDelimiter)
{
	long lnResult;
	long lnReturn = TRUE;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		lnResult = d4appendStart(gpCurrentTable, 0); /* no memo content carryover from the current record */
		if (lnResult == 0)
			{
			lnResult = cbwGATHERMEMVAR(lpcFldNames, lpcValues, lpcDelimiter);
			if (lnResult == TRUE)
				{
				lnResult = d4append(gpCurrentTable);
				d4recall(gpCurrentTable); // Seems to be required sometimes for the above.
				if (lnResult != 0)
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
						strcat(gcErrorMessage, " ERR from CodeBase");
						lnReturn = -1; /* We're done... ERROR */					
						}						
					}	
				}
			else
				{
				lnReturn = -1; /* The function cbwGATHERMEMVAR() has already set the error messages */
				strcat(gcErrorMessage, " From Gather Memvar");
				}	
			}
		else
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
				strcat(gcErrorMessage, " Code Base Error");
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

/* ************************************************************************************************* */
/* Replacement for VFP TAGCOUNT() which returns the number of tags in the CDX index associated       */
/* the currently selected table.  Returns the number or 0 if there aren't any.  Returns -1 on error. */
DOEXPORT long cbwTAGCOUNT(void)
{
	long lnReturn = -1;
	INDEX4* lpIndex;
	TAG4INFO* laTags;
	long lnCount = 0;
	long jj;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentTable != NULL)
		{
		//lpIndex = d4index(gpCurrentTable, NULL);
		lpIndex = cbxD4Index(gpCurrentTable);
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

/* ****************************************************************************** */
/* Equivalent of VFP ATAGINFO() which returns an array of information about the   */
/* index tags in the currently selected table.  Also returns the count of tags    */
/* in the long parameter passed by reference.  Returns NULL if there aren't any   */
/* tags.  Returns -1 in the count parameter on error and NULL as the return.      */
DOEXPORT VFPINDEXTAG* cbwATAGINFO(long* lpnCount)
{
	long lnCount = 0;
	long jj;
	INDEX4* lpIndex;
	TAG4INFO* laTags;
	TAG4INFO* laTagPtr;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	memset(gaIndexTags, 0, (size_t) 512 * sizeof(VFPINDEXTAG));	
	
	if (gpCurrentTable != NULL)
		{
		//lpIndex = d4index(gpCurrentTable, NULL);
		lpIndex = cbxD4Index(gpCurrentTable);
		if (lpIndex == NULL)
			{
			*lpnCount = 0;
			}
		else
			{
			laTags = i4tagInfo(lpIndex);
			if (laTags != NULL)
				{
				for (jj = 0; jj < 1000;jj++) /* Well beyond practical limit */
					{
					laTagPtr = laTags + jj;						
					if (laTagPtr->name == NULL) break;
					strcpy(gaIndexTags[jj].cTagName, laTagPtr->name);
					strcpy(gaIndexTags[jj].cTagExpr, laTagPtr->expression);
					strcpy(gaIndexTags[jj].cTagFilt, laTagPtr->filter);
					gaIndexTags[jj].nDirection = laTagPtr->descending;
					gaIndexTags[jj].nUnique    = laTagPtr->unique;
					lnCount += 1;
					}
				*lpnCount = lnCount;	
				}
			else
				{
				strcpy(gcErrorMessage, "Memory error");
				gnLastErrorNumber = -8001;
				*lpnCount = -1;	
				}				
			}
		}
	else
		{
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		*lpnCount = -1;
		}
	return(gaIndexTags);	
}

/* ***************************************************************************** */
/* Returns a string pointer to a buffer containing the name of the currently     */
/* active index tag of the currently selected table.                             */
/* Returns NULL on error.  Buffer may be an empty string if no tag is set.       */
DOEXPORT char* cbwORDER(void)
{
	char *lpcReturn = NULL;
	TAG4* lpTag;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gcStrReturn[0] = (char) 0;
	
	if (gpCurrentTable != NULL)
		{
		lpTag = d4tagSelected(gpCurrentTable);
		if (lpTag != NULL)
			{
			strcpy(gcStrReturn, t4alias(lpTag));
			}
		lpcReturn = gcStrReturn;		
		}
	else
		{
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return(lpcReturn);	
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
DOEXPORT long cbwTEMPINDEX(char* lpcTagExpr, char* lpcTagFltr, long lpnDescending)
{
	long lnIndexCode = -1;
	char lcFileName[300];
	char lcTableName[300];
	char lcTagName[20];
	TAG4INFO laTagInfo[3];
	TAG4* lpTag = NULL;
	INDEX4 *lpIndex;
	long jj;
	long lnDescending = 0;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	if (gpCurrentTable != NULL)
		{
		memset(lcFileName, 0, (size_t) 300 * sizeof(char)); // Start out with an empty name string
		strcpy(lcTableName, d4fileName(gpCurrentTable));

		strcpy(lcFileName, justFilePath(lcTableName));
		strcat(lcFileName, "\\X");
		strcat(lcFileName, MPStrRand(8));
		strcat(lcFileName, ".CDX");
		memset(laTagInfo, 0, (size_t) 3 * sizeof(TAG4INFO));
		strcpy(lcTagName, "X");
		strcat(lcTagName, MPStrRand(8));

		if (lpnDescending != 0) lnDescending = r4descending; // the official code for descending.
		laTagInfo[0].name = lcTagName; // Can be anything, since there will be only one.
		laTagInfo[0].expression = lpcTagExpr;
		if (strlen(lpcTagFltr) == 0) laTagInfo[0].filter = NULL;
		else laTagInfo[0].filter = (const char*) lpcTagFltr;
		laTagInfo[0].unique = 0;
		laTagInfo[0].descending = (int) lnDescending;
		
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
				strcpy(gaTempIndexes[lnIndexCode].aTags[0].cTagExpr, lpcTagExpr);
				strcpy(gaTempIndexes[lnIndexCode].aTags[0].cTagFilt, lpcTagFltr);
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
	return lnIndexCode;	
}

/* ************************************************************************************** */
/* If you select a different order while a temp index is in effect, you can return to it  */
/* with this function.  In theory, changes made to the table will be reflected in the temp*/
/* indexes while they are open, so they should stay up to date.                           */
/* Pass this the temp index code.  Returns 1 on OK, -1 on any failure.                    */
/* Note this does NOT set the associated table as currently selected.  It just changes    */
/* the current ordering on the table associated with this temp index.                     */
DOEXPORT long cbwTEMPINDEXSELECT(long lnIndx)
{
	long lnReturn = 1;
	TAG4* lpTag = NULL;
	
	if ((lnIndx > gnMaxTempIndexes) || (lnIndx < 0))
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "Temp Index Code is out of range");
		gnLastErrorNumber = -9194;	
		}
	else
		{
		if ((gaTempIndexes[lnIndx].bOpen == FALSE) || (gaTempIndexes[lnIndx].pTable == NULL))
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, "Temp Index has already been closed");
			gnLastErrorNumber = -9294;	
			}
		else
			{
			lpTag = d4tag(gaTempIndexes[lnIndx].pTable, gaTempIndexes[lnIndx].aTags[0].cTagName);
			d4tagSelect(gaTempIndexes[lnIndx].pTable, lpTag);				
			}
		}	
	return lnReturn;
}

/* *********************************************************************************** */
/* Removes one specified index tag from the currently selected table.  Returns 1 on    */
/* success, 0 or -1 on failure.                                                        */
DOEXPORT long cbwDELETETAG(char* lpcTagName)
{
	long lnReturn = TRUE;
	int  lnResult = 0;

	TAG4* lpTag = NULL;
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gcStrReturn[0] = (char) 0;
	
	if (gpCurrentTable == NULL)
		{
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;
		return(0);
		}
		
	if ((lpcTagName == NULL) || (strlen(lpcTagName) == 0))
		{
		strcpy(gcErrorMessage, "NO Tag Specified");
		gnLastErrorNumber = -9993;
		return(0);
		}

	lpTag = d4tag(gpCurrentTable, lpcTagName);	
	if (lpTag == NULL)
		{
		strcpy(gcErrorMessage, "Index Tag Does Not Exist");
		gnLastErrorNumber = -9992;
		return(0);
		}

	lnResult = i4tagRemove(lpTag);
	if (lnResult == 0)
		{
		lnReturn = 1;	
		}
	else
		{
		gnLastErrorNumber = codeBase.errorCode;
		strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));	
		lnReturn = 0;		
		}	
	
	return(lnReturn);	
}

/* *********************************************************************************** */
/* Equivalent to VFP INDEX ON command.  Operates on the currently selected table.      */
/* The index expression may contain VFP type functions recognized by the CodeBase      */
/* engine.  See the CodeBase documentation for more information.                       */
/* Returns TRUE on success, -1 on failure.                                             */
DOEXPORT long cbwINDEXON(char* lpcTagName, char* lpcTagExpr, char* lpcTagFltr, long lpnDescending)
{
	long lnReturn = TRUE;
	int  lnResult = 0;
	long lnTagCount = 0;
	long lnTagLength = 0;
	TAG4INFO laTagInfo[3];
	TAG4* lpTag = NULL;
	INDEX4* lpIndex;
	char lcTableName[256];
	char lcIndexName[256];
	
	codeBase.errorCode = 0;
	gcErrorMessage[0] = (char) 0;
	gcStrReturn[0] = (char) 0;
	gnLastErrorNumber = 0;
	memset(&laTagInfo, 0, (size_t) 3 * sizeof(TAG4INFO));
	memset(lcTableName, 0, (size_t) 256 * sizeof(char));
	memset(lcIndexName, 0, (size_t) 256 * sizeof(char));
	
	if (gpCurrentTable != NULL)
		{
		lnTagLength = (long) strlen(lpcTagName);
		if ((lnTagLength == 0) || (lnTagLength > 10))
			{
			lnReturn = -1;
			strcpy(gcErrorMessage, "Tag Name exceeds 10 characters");
			gnLastErrorNumber = -8002;	
			}
		else
			{
			lnTagCount = cbwTAGCOUNT();
			if (lnTagCount > 0)
				{
				/* There is a CDX file active.  Now check to see if there is a tag by the requested name already */
				lpTag = d4tag(gpCurrentTable, lpcTagName);
				}
			codeBase.errorCode = 0; /* Reset this, per CodeBase Tech */
			laTagInfo[0].name = lpcTagName;
			laTagInfo[0].expression = lpcTagExpr;
			laTagInfo[0].filter = lpcTagFltr;
			laTagInfo[0].unique = 0; // Always.
			laTagInfo[0].descending = (int) lpnDescending;
			
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
						lnReturn = -1;							
						}
					else
						{
						lnReturn = cbxReopenCurrentTable();	
						}
					break;	
					}
					
				if ((lnTagCount > 0) && (lpTag == NULL)) /* There is an index CDX file, but not with that tag.  So we Add it. */
					{
					//lpIndex = d4index(gpCurrentTable, d4alias(gpCurrentTable));
					lpIndex = cbxD4Index(gpCurrentTable);
					if (lpIndex == NULL)
						{
						gnLastErrorNumber = codeBase.errorCode;
						strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
						strcat(gcErrorMessage, " - Finding Index CDX Failed");
						lnReturn = -1;
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
							lnReturn = -1; /* We're done... ERROR */					
							}							
						}

					break;
					}
					
				if ((lnTagCount > 0) && (lpTag != NULL)) /* There is a index CDX file AND the tag exists, so we remove and recreate it. */ 
					{
					lpIndex = cbxD4Index(gpCurrentTable);
					if (lpIndex == NULL)
						{
						gnLastErrorNumber = codeBase.errorCode;
						strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
						strcat(gcErrorMessage, " - Finding Index CDX Failed");
						lnReturn = -1;
						break;							
						}
					lpTag = i4tag(lpIndex, lpcTagName); // Looking this up in the current index file.
					if (lpTag != NULL)
						{
						lnResult = i4tagRemove(lpTag);
						if ((lnResult != 0) && (codeBase.errorCode != e4tagName))
							{
							/* can't figure out why this happens, but it can... for no reason */
							lnResult = cbxReopenCurrentTable();
							if (lnResult == TRUE)
								{
								lpIndex = cbxD4Index(gpCurrentTable);
								if (lpIndex != NULL)
									{
									lpTag = i4tag(lpIndex, lpcTagName);
									if (lpTag != NULL) lnResult = i4tagRemove(lpTag);	
									}	
								if ((lnResult != 0) && (codeBase.errorCode != e4tagName))
									{
									gnLastErrorNumber = codeBase.errorCode;
									strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
									strcat(gcErrorMessage, " - Removing Old Tag Failed a second time");
									lnReturn = -1; /* We're done... ERROR */
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
								strcat(gcErrorMessage, lpcTagName);
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
								lnReturn = -1; /* One of the several locking and duplicate key errors.  In any event, FAILURE. */	
								}
							else
								{
								gnLastErrorNumber = codeBase.errorCode;
								strcpy(gcErrorMessage, error4text(&codeBase, codeBase.errorCode));
								strcat(gcErrorMessage, " - Adding Tag After Removing Old Tag");
								lnReturn = -1; /* We're done... ERROR */					
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
							lnReturn = -1; /* We're done... ERROR */
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
		lnReturn = -1;
		strcpy(gcErrorMessage, "No Table Open in Selected Area");
		gnLastErrorNumber = -9999;	
		}
	return(lnReturn);	
}

/* ********************************************************************************** */
/* This function is very similar to the VFP CREATE command which takes as a           */
/* parameter the path name of the table to create and a text string with              */
/* a comma delimited list of the field descriptions.  Each field must have 5 values   */
/* in the list:                                                                       */
/* Field Name (limited in this version to 10 characters, starting with a letter or_)  */
/* Field Type (one of C, D, T, M, I, Y, B, G, X, L, N -- see VFP or CodeBase docs)    */
/* Field Size (number of chars for type C, total digits for N, etc.)                  */
/* Field Decimals (number of decimal points for Y:currency, N:number types)           */
/* Field Nulls OK (TRUE or FALSE -- or blank = FALSE indicating if NULLs are allowed) */
/* Each Field is listed on a separate line in the text file.  Newline characters sep- */
/* the lines.                                                                         */
/* Returns TRUE on OK, -1 on ERROR of any kind.  No indexes are created.              */
/* The default file name extension is .DBF, which will be added if none is specified. */
/* You may specify other file name extensions if you want to.                         */
/* Replaces any table with the same name.                                             */
/* On successful completion, the table is open and is the currently selected table.   */
DOEXPORT long cbwCREATETABLE(char* lpcTableName, char* lpcFieldList)
{
	long lnReturn = TRUE;
	long lnResult = 0;
	long lbFieldOK;
	long jj;
	long kk;
	char lcFullName[512];
	char lcBuff[20];
	char lcTempName[20];
	long lbDupeNameError = FALSE;
	char *lpPtr;
	char *lpTemp;
	long lnFldCnt = -1;
	char *laFieldSpecs[256];
	DATA4 *lpTable;

	codeBase.errorCode = 0;
	codeBase.safety = 0;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	memset(lcFullName, 0, (size_t) 512 * sizeof(char));
	memset(laFieldSpecs, 0, (size_t) 256 * sizeof(char*));
	//printf("Creating 1 \n");
	strncpy(lcFullName, lpcTableName, 500);
	lpPtr = strstr(lcFullName, ".");
	if (lpPtr == NULL)
		{
		strcat(lcFullName, ".DBF");	
		}
	memset(gaFieldSpecs, 0, (size_t) 256 * sizeof(FIELD4INFO));
	
	lpTemp = lpcFieldList;
	while(TRUE)
		{
		lpPtr = strtokX(lpTemp, "\n");
		lpTemp = NULL;
		if (lpPtr == NULL) break;
		if (strlen(lpPtr) == 0) continue; /* Just in case */
		lnFldCnt += 1; /* Starts at -1, so 0 based */
		laFieldSpecs[lnFldCnt] = malloc((strlen(lpPtr) * sizeof(char)) + 1);
		strcpy(laFieldSpecs[lnFldCnt], lpPtr);
		}
	lnFldCnt += 1; /* Bring to 1 based count */
	if (lnFldCnt == 0)
		{
		lnReturn = -1;
		strcpy(gcErrorMessage, "No fields defined");
		gnLastErrorNumber = -8026;	
		}
	else
		{
		for (jj = 0;jj < lnFldCnt;jj++)
			{
			lbFieldOK = FALSE;
			while(TRUE)
				{
				lpTemp = laFieldSpecs[jj];
				lpPtr = strtokX(lpTemp, ",");
				lpTemp = NULL;
				if (lpPtr == NULL) break;
				gaFieldSpecs[jj].name = malloc(15 * sizeof(char)); /* longer than needed */

				strncpy(lcTempName, lpPtr, 11);
				strcpy(lcTempName, MPSSStrTran(lcTempName, "\n", "", 1));
				strcpy(lcTempName, MPSSStrTran(lcTempName, "\r", "", 1));
				for (kk = 0; kk < jj; kk++)
					{
					if (strcmp(lcTempName, gaFieldSpecs[kk].name) == 0) /* This is a duplicate name. BAD */
						{
						lbDupeNameError = TRUE;
						lbFieldOK = FALSE;
						break;	
						}	
					}
				if (lbDupeNameError == TRUE) break;
					
				strcpy(gaFieldSpecs[jj].name, lcTempName);
				lpPtr = strtokX(lpTemp, ",");
				if (lpPtr == NULL) break;
				gaFieldSpecs[jj].type = (short) *lpPtr; /* One character type */
				
				lpPtr = strtokX(lpTemp, ",");
				if (lpPtr == NULL) break;
				gaFieldSpecs[jj].len = atoi(lpPtr);
				
				lpPtr = strtokX(lpTemp, ",");
				if (lpPtr == NULL) break;
				gaFieldSpecs[jj].dec = atoi(lpPtr);
				lbFieldOK = TRUE; /* We are done if the decimals are provided.  Nulls are assumed to be false, unless they tell us. */
								
				lpPtr = strtokX(lpTemp, ",");
				if (lpPtr == NULL) break;
				strncpy(lcBuff, lpPtr, 10);
				strcpy(lcBuff, MPSSStrTran(lcBuff, "\n", "", 1));
				strcpy(lcBuff, MPSSStrTran(lcBuff, "\r", "", 1));
				if (strcmp(lcBuff, "TRUE") == 0)
					gaFieldSpecs[jj].nulls = r4null;
				else
					gaFieldSpecs[jj].nulls = 0;
				
				break;
				}
			if (lbFieldOK == FALSE)
				{
				lnReturn = -1;
				if (lbDupeNameError == TRUE)
					{
					strcpy(gcErrorMessage, "Duplicate Field Name: ");
					strcat(gcErrorMessage, lcTempName);
					}
				else
					{					
					strcpy(gcErrorMessage, "Incomplete Field Specification for field: ");
					if (gaFieldSpecs[jj].name != NULL)
						strcat(gcErrorMessage, gaFieldSpecs[jj].name);
					}
				gnLastErrorNumber = -8040;
				break;
				}
			}
		//printf("Creating 2 \n");
		if (lnReturn == TRUE) /* All the fields were correctly specified */
			{
			//ldStart = ((double) clock() / (double) CLOCKS_PER_SEC);
			lpTable = d4create(&codeBase, lcFullName, gaFieldSpecs, NULL);
			//ldEnd = ((double) clock() / (double) CLOCKS_PER_SEC);
			//printf("Total time: %lf  Result %p fld cnt %ld \n", ldEnd - ldStart, lpTable, lnFldCnt);
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
		}
	//printf("Creating 3 \n");
	for (jj = 0; jj < 256; jj++)
		{
		if (laFieldSpecs[jj] != NULL) free(laFieldSpecs[jj]);
			else break; /* we're done */	
		}
	for (jj = 0;jj < 256; jj++)
		{
		if (gaFieldSpecs[jj].name != NULL) free(gaFieldSpecs[jj].name);
			else break; /* done here too */			
		}
	//printf("Returning %ld \n", lnReturn);
	return(lnReturn);	
}


