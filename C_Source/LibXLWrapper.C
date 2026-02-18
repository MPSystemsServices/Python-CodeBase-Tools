/*---------------- Created November 12, 2008 - 12:40pm -----------------
 * C language program for accessing the LibXL low-level library for reading
 * writing, and modifying Excel XLS and XLSX files.  Note that a legal
 * licensed copy of libxl.dll must be available and a license key be readable
 * by this module before it will work for production purposes.
 * ---------------------------------------------------------------------*/
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

#define CRT_SECURE_NO_DEPRECATE 1
#include <libxl.h>
#include "stdlib.h"
#include "stdio.h"
#include "string.h"
#include "wctype.h"

#define TRUE 1
#define FALSE 0
#define DOEXPORT __declspec(dllexport)
#define _BUFFER_LEN 8920000
#define _FORMAT_COUNT 2000
#define _FONT_COUNT 2000
#define _SHEET_COUNT 500
#define _PRINT_PORTRAIT 1
#define _PRINT_LANDSCAPE 2
#define _MARGIN_LEFT 1
#define _MARGIN_RIGHT 2
#define _MARGIN_TOP 3
#define _MARGIN_BOTTOM 4
#define _MARGIN_ALL 99
#define _SIZE_PAPER_LETTER 1
#define _SIZE_PAPER_LEGAL 2
#define _SIZE_PAPER_A3 3

typedef struct XLFORMATtd { // NOTE: This structure is NOT used by LibXL.  It is used to simplify the management of format info.
	int nNumFormat; // Code for the Num format of the cell.  Applies also to dates and times.
	int nAlignH; // Alignment code
	int nAlignV; // Ditto for vertical alignment
	int nWordWrap; // 1 = Wrap, 0 = No Wrap
	int nRotation; // 0 to 180.  255 is Vertical Text
	int nIndentLevel; // 0 to 15
	int nBorderTop; // One of the border type codes
	int nBorderLeft;
	int nBorderRight;
	int nBorderBottom;
	int nBackColor;
	int nFillPattern;
	int bLocked; // if TRUE then cell is locked
	int bHidden; // if TRUE then cell is hidden
	char cFontName[45];
	int bItalic;
	int bBold;
	int nColor;
	int bUnderlined;
	int nPointSize;	
} XLFORMAT;

BookHandle gpCurrentBook = NULL; /* Only ONE workbook can be open at one time.  This is its pointer. */
SheetHandle gpSheets[_SHEET_COUNT];
FormatHandle gpFormats[_FORMAT_COUNT];
XLFORMAT gpWrapFormats[_FORMAT_COUNT];
FontHandle gpFonts[_FONT_COUNT];
FormatHandle gpDefaultDateFormat = NULL;
FormatHandle gpDefaultDateTimeFormat = NULL; // Added 02/08/2013
FormatHandle gpDefaultTimeFormat = NULL; // Added 02/08/2013.  These support more control over formats for XLS output.

char gcErrorMessage[1000]; /* Contains the most recent error message.  Guaranteed to be less than 1000 characters in length. */
long gnLastErrorNumber = 0;
char gcStrReturn[_BUFFER_LEN]; /* Where we put stuff for static string return */
wchar_t gcWideString[500];
char gcCurrentFileName[500];
long gbBookChanged = FALSE;
long gnNextSheet = 0;
long gnNextFormat = 0;
long gnDefaultFormat = -1;
char gcCellDelimiter[11];

// Made KeyName and KeyCode to be supplied externally to facilitate release of
// this code into Open Source, as the license for LIBXL is a commercial product.
char gcKeyName[50];
char gcKeyCode[100];

char* gcMessageBadSheet = "Bad Sheet Pointer Value";
long  gnMessageNumberBadSheet = -100;

/*-----------------1/30/2006 9:17PM-----------------
 * This DllMain doesn't seem to get fired off, even tho this kind of thing
 * does in the PowerBasic world.  Who knows why.
 * --------------------------------------------------*/
long DllMain(long hInstance, long fwdReason, long lpvReserved)
{
	return(1);
};

/* *********************************************************************************************************** */
/* Function that pads the source string to length lnLen in place.  Returns a pointer to the source string      */
/* The source string must have been allocated with enough space to hold this result                            */
/* The obligatory chr(0) value is placed at the end of the string.  Truncates source strings longer then lnLen */
/* COPIED HERE FROM mpssClib.c to avoid having to link another DLL just for one function.                      */
char* PadStringRight(char* lcSrc, long lnLen, char lcFillChar)
{
	long lnSrcLen;
	lnSrcLen = strlen(lcSrc);
	if ( lnSrcLen > lnLen )
		{
		*(lcSrc + lnLen) = (char) 0; /* truncation */
		}
	else
		{
		memset(lcSrc + lnSrcLen, (int) lcFillChar, (size_t) (lnLen - lnSrcLen) * sizeof(char));
		*(lcSrc + lnLen) = (char) 0; /* the final null */
		}
	
	return(lcSrc);
}


/* ****************************************************************************** */
/* Utility function that sets the error codes                                     */
void SetErrorCodes(long lpnErrNum)
{
	strcpy(gcErrorMessage, xlBookErrorMessage(gpCurrentBook));
	gnLastErrorNumber = lpnErrNum;	
}


long ShowEnums(void)
{
// FOR DEBUG AND VALUE TESTING ONLY.
//	printf("BORDERSTYLE_NONE = %ld\n", BORDERSTYLE_NONE);
//	printf("BORDERSTYLE_THIN = %ld\n", BORDERSTYLE_THIN);
//	printf("BORDERSTYLE_MEDIUM = %ld\n", BORDERSTYLE_MEDIUM);
//	printf("BORDERSTYLE_DASHED = %ld\n", BORDERSTYLE_DASHED);
//	printf("COLOR_BLACK = %ld\n", COLOR_BLACK);
//	printf("COLOR_WHITE = %ld\n", COLOR_WHITE);
//	printf("COLOR_RED = %ld\n", COLOR_RED);
//	printf("COLOR_BRIGHTGREEN = %ld\n", COLOR_BRIGHTGREEN);
//	printf("COLOR_BLUE = %ld\n", COLOR_BLUE);
//	printf("COLOR_YELLOW = %ld\n", COLOR_YELLOW);
//	printf("COLOR_PINK = %ld\n", COLOR_PINK);
//	printf("COLOR_TURQUOISE = %ld\n", COLOR_TURQUOISE);
//	printf("COLOR_DARKRED = %ld\n", COLOR_DARKRED);
//	printf("COLOR_GREEN = %ld\n",  COLOR_GREEN);
//	printf("ALIGNH_GENERAL = %ld\n", ALIGNH_GENERAL);
//	printf("ALIGNH_LEFT	= %ld\n", ALIGNH_LEFT);
//	printf("NUMFORMAT_GENERAL = %ld\n", NUMFORMAT_GENERAL);
//	printf("NUMFORMAT_NUMBER_SEP_NEGBRARED = %ld\n", NUMFORMAT_NUMBER_SEP_NEGBRARED);
return(1);
}

/* ****************************************************************************** */
/* Utility function for storing an ASCII string into gcWideString.                */
/* Returns gcWideString.                                                          */
//wchar_t* MakeWideString(char* lpcSourceStr)
//{
//	long jj;
//	long lnLen;
//	lnLen = strlen(lpcSourceStr);
//	for (jj = 0;jj < lnLen;jj++)
//		{
//		gcWideString[jj] = btowc((int) *(lpcSourceStr + jj));	
//		}
//	gcWideString[lnLen] = 0;
//	return gcWideString;	
//}

/* ****************************************************************************** */
/* Utility function for changing all letters in a string to lower case            */
void StrToLower(char* lpcStr)
{
	for (lpcStr; *lpcStr != (char) 0; lpcStr++)
		{
		*lpcStr = tolower(*lpcStr);	
		}	
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
		lpReturn = NULL;
		lpPtr = NULL; /* don't try calling this again, cause we're done. */
		}
	else
		{
		lpTest = strstr(lpPtr, lpcDelim);
		if (lpTest == NULL)
			{
			lpReturn = lpPtr;
			lpPtr += strlen(lpPtr);	
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

/* ************************************************************************************** */
/* Internal function that initializes the gpWrapFormats[] array with all values = -1.     */
long InitWrapFormats(void)
{
	long jj;
	XLFORMAT *lpFmat;
	
	for (jj = 0; jj < _FORMAT_COUNT; jj++)
		{
		lpFmat = &(gpWrapFormats[jj]);
		lpFmat->nNumFormat = -1;
		lpFmat->nAlignH = -1;
		lpFmat->nAlignV = -1;
		lpFmat->nWordWrap = -1;
		lpFmat->nRotation = -1;
		lpFmat->nIndentLevel = -1;
		lpFmat->nBorderTop = -1;
		lpFmat->nBorderLeft = -1;
		lpFmat->nBorderRight = -1;
		lpFmat->nBorderBottom = -1;
		lpFmat->nBackColor = -1;
		lpFmat->nFillPattern = -1;
		lpFmat->bLocked = -1;
		lpFmat->bHidden = -1;
		lpFmat->cFontName[0] = (char)0;
		lpFmat->bItalic = -1;
		lpFmat->bBold = -1;
		lpFmat->nColor = -1;
		lpFmat->bUnderlined = -1;
		lpFmat->nPointSize = -1;	
		}
	gnNextFormat = 0;
	return(1);	
}

/* ************************************************************************************** */
/* Internal function that sets all the characteristics of a LibXL FormatHandle based on   */
/* the elements of a WrapFormat structure.                                                */
void PostWrapFormat(XLFORMAT *lpFmat, FormatHandle lpFHandle, FontHandle lpNHandle)
{
	if (lpFmat->nNumFormat != -1)
		{
		xlFormatSetNumFormat(lpFHandle, lpFmat->nNumFormat);	
		}
	if (lpFmat->nAlignH != -1)
		{
		xlFormatSetAlignH(lpFHandle, lpFmat->nAlignH);	
		}
	if (lpFmat->nAlignV != -1)
		{
		xlFormatSetAlignV(lpFHandle, lpFmat->nAlignV);	
		}
	if (lpFmat->nWordWrap != -1)
		{
		xlFormatSetWrap(lpFHandle, lpFmat->nWordWrap);	
		}
	if (lpFmat->nRotation != -1)
		{
		xlFormatSetRotation(lpFHandle, lpFmat->nRotation);	
		}
	if (lpFmat->nIndentLevel != -1)
		{
		xlFormatSetIndent(lpFHandle, lpFmat->nIndentLevel);	
		}
	if (lpFmat->nBorderTop != -1)
		{
		xlFormatSetBorderTop(lpFHandle, lpFmat->nBorderTop);	
		}
	if (lpFmat->nBorderLeft != -1)
		{
		xlFormatSetBorderLeft(lpFHandle, lpFmat->nBorderLeft);	
		}
	if (lpFmat->nBorderRight != -1)
		{
		xlFormatSetBorderRight(lpFHandle, lpFmat->nBorderRight);	
		}
	if (lpFmat->nBorderBottom != -1)
		{
		xlFormatSetBorderBottom(lpFHandle, lpFmat->nBorderBottom);	
		}
	if (lpFmat->nBackColor != -1)
		{
		xlFormatSetPatternForegroundColor(lpFHandle, lpFmat->nBackColor);	
		}
	if (lpFmat->nFillPattern != -1)
		{
		xlFormatSetFillPattern(lpFHandle, lpFmat->nFillPattern);	
		}
	if (lpFmat->bLocked != -1)
		{
		xlFormatSetLocked(lpFHandle, lpFmat->bLocked);	
		}
	if (lpFmat->bHidden != -1)
		{
		xlFormatSetHidden(lpFHandle, lpFmat->bHidden);	
		}
	if (lpFmat->cFontName[0] != (char) 0)
		{
		xlFontSetName(lpNHandle, lpFmat->cFontName);	
		}
	if (lpFmat->bItalic != -1)
		{
		xlFontSetItalic(lpNHandle, lpFmat->bItalic);	
		}
	if (lpFmat->bBold != -1)
		{
		xlFontSetBold(lpNHandle, lpFmat->bBold);	
		}
	if (lpFmat->nColor != -1)
		{
		xlFontSetColor(lpNHandle, lpFmat->nColor);	
		}
	if (lpFmat->bUnderlined != -1)
		{
		xlFontSetUnderline(lpNHandle, lpFmat->bUnderlined);	
		}
	if (lpFmat->nPointSize != -1)
		{
		xlFontSetSize(lpNHandle, lpFmat->nPointSize);
		}
}

/* ************************************************************************************** */
/* Called by an external program that has access to the license user name and the license
   key supplied by LIBXL.com for use of their LIBXL.DLL product.
                                                                                          */
long DOEXPORT xlwSetLicense(char* lpcName, char* lpcKey)
{
	memset(gcKeyName, 0, (size_t) (50 * sizeof(char)));	
	memset(gcKeyCode, 0, (size_t) (100 * sizeof(char)));
	strncpy(gcKeyName, lpcName, 48);
	strncpy(gcKeyCode, lpcKey, 98);
	return(1);
}


/* ************************************************************************************** */
/* Sets the default value of the delimiter string used for writing and getting a row's 
   worth of data at one time.  The default is "\t", but it is suggested that when the data
   are ill-behaved, something more exotic like "\b\v\b" be used to minimize conflict with
   actual cell data.                                                                      */
long DOEXPORT xlwSetDelimiter(char* lpcDelim)
{
	long lnLen = 0;
	if (lpcDelim == NULL)
		{
		strcpy(gcCellDelimiter, "\t"); // Default in case a NULL is passed.
		return(1);
		}
	lnLen = strlen(lpcDelim);
	if ((lnLen == 0) || (lnLen > 10))
		return 0;
	strcpy(gcCellDelimiter, lpcDelim);
	return(1);	
}

/* ************************************************************************************** */
/* Adds a worksheet to the current WorkBook but copies it from the specified sheet in an
   external XLS/XLSX file.  Returns the index of the new sheet or -1 on failure.
   NOTE: This basically just copies the data, and ignores the formatting except for column
   widths and row heights.  You'll need to format the sheet after copying.                */
long DOEXPORT xlwImportSheetFrom(char* lpcFileName, char* lpcSheetName)
{
	long lnReturn = -1;
	long lnSheetCount = 0;
	long lnTargetSheet = -1;
	BookHandle xTempBook = NULL;
	SheetHandle xSrcSheet = NULL;
	SheetHandle lpDest = NULL;
	register long jj;
	register long kk;
	register long nn;
	long lnTest;
	long lnFirstCol;
	long lnLastCol;
	long lnFirstRow;
	long lnLastRow;
	long lnValue;
	long lnWidth = 0;
	long lnHidden = 0;
	double ldValue;
	long lnCellType;
	FormatHandle lppFmat = NULL;
	FormatHandle lpCellFmat = NULL;

	char lcFileName[255];
	const char *lpTest = NULL;
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	strcpy(lcFileName, lpcFileName);
	StrToLower(lcFileName);
	
	if (gpCurrentBook == NULL)
		{
		strcpy(gcErrorMessage, "No Current Workbook to Import Into");
		gnLastErrorNumber = 29479;
		lnReturn = -1;
		}
	else
		{
		if (strstr(lcFileName, ".xlsx") != NULL)
			{
			xTempBook = xlCreateXMLBook();	
			}
		else
			{
			xTempBook = xlCreateBook();	
			}
		xlBookSetKey(xTempBook, gcKeyName, gcKeyCode);
		
		lnTest = xlBookLoad(xTempBook, lcFileName);
		if (lnTest == 0)
			{
			SetErrorCodes(-2);
			lnReturn = -1;	
			}
		else
			{
			lnSheetCount = xlBookSheetCount(xTempBook);
			for (jj = 0;jj < lnSheetCount;jj++)
				{
				xSrcSheet = xlBookGetSheet(xTempBook, jj);
				if (strcmp(lpcSheetName, xlSheetName(xSrcSheet)) == 0)
					{
					lnTargetSheet = jj;
					break;	
					}
				}
			if (lnTargetSheet == -1)
				{
				strcpy(gcErrorMessage, "Source Sheet Name Not Found");
				lnReturn = -1;
				gnLastErrorNumber = 29480;	
				}
			else
				{
				// Found the source sheet in the target
				lpDest = xlBookInsertSheet(gpCurrentBook, gnNextSheet, lpcSheetName, NULL);

				if (lpDest != NULL)
					{
					gpSheets[gnNextSheet] = lpDest;
					gbBookChanged = TRUE;
					lnReturn = gnNextSheet;
					gnNextSheet += 1;
					// Now we can copy cells from the origin sheet to the destination sheet.
					lnFirstCol = xlSheetFirstCol(xSrcSheet);
					lnLastCol  = xlSheetLastCol(xSrcSheet) - 1; // For some reason, this function returns a one-based column number.
					lnFirstRow = xlSheetFirstRow(xSrcSheet);
					lnLastRow  = xlSheetLastRow(xSrcSheet);
					for (jj = lnFirstRow; jj <= lnLastRow; jj++)
						{
						ldValue = xlSheetRowHeight(xSrcSheet, (int) jj);
						lnHidden = xlSheetRowHidden(xSrcSheet, (int) jj);
						lnValue = xlSheetSetRow(lpDest, (int) jj, ldValue, NULL, lnHidden);
						for (kk = lnFirstCol; kk <= lnLastCol; kk++)
							{
							if (jj == lnFirstRow)
								{
								// we can set the column widths first time through.
								ldValue = xlSheetColWidth(xSrcSheet, (int) kk);
								lnHidden = xlSheetColHidden(xSrcSheet, (int) kk);
								lnValue = xlSheetSetCol(lpDest, (int) kk, (int) kk + 1, ldValue, NULL, lnHidden);	
								}
							lnCellType = xlSheetCellType(xSrcSheet, jj, kk);
							lpCellFmat = xlSheetCellFormat(xSrcSheet, jj, kk);
							if (lpCellFmat == NULL)
								{
								continue;
								}

							switch (lnCellType)
								{
								case CELLTYPE_EMPTY:
									break;
					
								case CELLTYPE_NUMBER:
									ldValue = xlSheetReadNum(xSrcSheet, jj, kk, &lppFmat);
									if (lppFmat != NULL)
										{
										xlSheetWriteNum(lpDest, jj, kk, ldValue, NULL);
										}
									break;
				
								case CELLTYPE_STRING:
									lpTest = xlSheetReadStr(xSrcSheet, jj, kk, &lppFmat);
									if ((lpTest != NULL) && (lppFmat != NULL))
										{
										xlSheetWriteStr(lpDest, jj, kk, lpTest, NULL);
										}
									break;
									
								case CELLTYPE_BOOLEAN:
									lnValue = xlSheetReadBool(xSrcSheet, jj, kk, &lppFmat);
									if (lppFmat != NULL)
										{
										xlSheetWriteBool(lpDest, jj, kk, lnValue, NULL); 
										}
									break;
				
								case CELLTYPE_BLANK:
									lnValue = xlSheetWriteBlank(lpDest, jj, kk, NULL);
									break;
								
								case CELLTYPE_ERROR:	
									break;	
									
								default:
									break;	
								}
							}
						} 
					}
				else
					{
					SetErrorCodes(-5);
					lnReturn = -1;				
					}				
				}
			xlBookRelease(xTempBook);
			}
		}
	
	return(lnReturn);
}


/* ************************************************************************************** */
/* Opens an existing XLS or XLSX for reading and writing.  Decides what format to work
   on based on the file name extension: XLS or XLSX.  Returns the number of sheets found
   as the result or 0 on failure.  Sets gcErrorMessage on failure.                        */
long DOEXPORT xlwOpenWorkbook(char* lpcFileName)
{
	char lcExt[30];
	char lcFileName[1200];
	long lnReturn = 0;
	long lnTest;
	long jj;

	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentBook != NULL)
		{
		strcpy(gcErrorMessage, "Workbook already open");
		gnLastErrorNumber = -1;
		lnReturn = 0;
		}
	else
		{
		gbBookChanged = FALSE;
		memset(gpSheets, 0, (size_t) _SHEET_COUNT * sizeof(SheetHandle));
		memset(gpFormats, 0, (size_t) _FORMAT_COUNT * sizeof(FormatHandle));
		memset(gpFonts, 0, (size_t) _FONT_COUNT * sizeof(FontHandle));
		gcCurrentFileName[0] = (char) 0;
		strcpy(lcFileName, lpcFileName);
		StrToLower(lcFileName);
		gnNextSheet = 0;
		gpDefaultDateFormat = NULL;
		gpDefaultDateTimeFormat = NULL;
		gpDefaultTimeFormat = NULL;
		InitWrapFormats();
		
		if (strstr(lcFileName, ".xlsx") != NULL)
			{
			gpCurrentBook = xlCreateXMLBook();	
			}
		else
			{
			gpCurrentBook = xlCreateBook();	
			}

		xlBookSetKey(gpCurrentBook, gcKeyName, gcKeyCode);
		
		lnTest = xlBookLoad(gpCurrentBook, lcFileName);
		if (lnTest == 0)
			{
			SetErrorCodes(-2);
			lnReturn = 0;	
			}
		else
			{
			lnReturn = xlBookSheetCount(gpCurrentBook);
			if (lnReturn == 0)
				{
				SetErrorCodes(-3);	
				}
			else
				{
				strcpy(gcCurrentFileName, lcFileName);
				for (jj = 0;jj < lnReturn;jj++)
					{
					gpSheets[jj] = xlBookGetSheet(gpCurrentBook, jj);	
					}
				gnNextSheet = lnReturn;
				}	
			}
		}
	if (lnReturn == 0)
		{
		if (gpCurrentBook != NULL) xlBookRelease(gpCurrentBook);
		gpCurrentBook = NULL;	
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Closes the open existing or in-memory new workbook.  Releases the workbook handle.
   If changes have been made that you want saved, pass TRUE (1) as the parameter, otherwise
   the workbook will be closed without saving any changes.  Returns 1 if OK.  Always
   returns 1 if no save is requested.  If save is requested, returns 1 if save was OK,
   otherwise returns 0.                                                                   */
long DOEXPORT xlwCloseWorkbook(long lbSave)
{
	long lnReturn = 1;
	long lnResult;
	
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentBook != NULL)
		{
		if (lbSave == TRUE)
			{
			if (gbBookChanged)
				{
				if (strcmp(gcCurrentFileName, "") != 0)
					{
					if (xlBookSave(gpCurrentBook, gcCurrentFileName) == 0)
						{
						lnReturn = 0;
						SetErrorCodes(-4);
						}
					}	
				}	
			}
		xlBookRelease(gpCurrentBook);
		gpCurrentBook = NULL;
		gnNextSheet = 0;
		gnNextFormat = 0;
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Creates a new empty workbook.  This exists only in memory until save is invoked.  Pass
   the full path name of the file to be created, as with the xlwOpenWorkbook() this 
   determines the type of workbook to create from the extension: XLS or XLSX.  No file
   is actually created.  You must issue an explicit xlwSaveWorkbook() to write it to the
   disk.  The second parameter must be the name of the first Worksheet to be added.  Returns
   the sheet code of that first sheet.  If the sheet name passed is empty or NULL, gives
   the sheet the name of "Sheet1".  Returns 1 on Success, 0 on failure.                   */
/* Modified 01/19/2013 with addition of lbNoSheet parm.  If that value is non-zero and
   no value (NULL or empty) is passed as lpcFirstSheet, then will NOT insert an empty sheet.
   Returns 1 in this case too on success. */
long DOEXPORT xlwCreateWorkbook(char* lpcFileName, char* lpcFirstSheet, long lbNoSheet)
{
	char lcExt[30];
	char lcFileName[1200];
	char lcSheetName[254];
	long lnReturn = 0;
	long lnTest;
	

	SheetHandle lpTestSheet = NULL;
	lcSheetName[0] = (char) 0;
	//printf("Creating book: %s \n", lpcFileName);
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	gpDefaultDateFormat = NULL;
	gpDefaultDateTimeFormat = NULL;
	gpDefaultTimeFormat = NULL;
	
	if (gpCurrentBook != NULL)
		{
		strcpy(gcErrorMessage, "Workbook already open");
		gnLastErrorNumber = -1;
		lnReturn = 0;
		}
	else
		{
		gbBookChanged = FALSE;
		memset(gpSheets, 0, (size_t) _SHEET_COUNT * sizeof(SheetHandle));
		memset(gpFormats, 0, (size_t) _FORMAT_COUNT * sizeof(FormatHandle));
		memset(gpFonts, 0, (size_t) _FONT_COUNT * sizeof(FontHandle));
		InitWrapFormats();
		
		strcpy(lcFileName, lpcFileName);
		StrToLower(lcFileName);
		if (strstr(lcFileName, ".xlsx") != NULL)
			{
			gpCurrentBook = xlCreateXMLBook();	
			}
		else
			{
			gpCurrentBook = xlCreateBook();	
			}

		xlBookSetKey(gpCurrentBook, gcKeyName, gcKeyCode);
		
		if ((lpcFirstSheet == NULL) || (strlen(lpcFirstSheet) == 0))
			{
			if (lbNoSheet != 1)
				{
				strcpy(lcSheetName, "Sheet1");
				}
			}
		else
			{
			strcpy(lcSheetName, lpcFirstSheet);
			}
		
		if (lcSheetName[0] != (char) 0)
			{	
			lpTestSheet = xlBookAddSheet(gpCurrentBook, lcSheetName, NULL);
			if (lpTestSheet == NULL)
				{
				SetErrorCodes(-3);
				lnReturn = 0;	
				}
			else
				{
				lnReturn = xlBookSheetCount(gpCurrentBook);
				gbBookChanged = TRUE;
				gpSheets[0] = lpTestSheet;
				strcpy(gcCurrentFileName, lcFileName);
				gnNextSheet = 1;
				}
			}
		else
			{
			gnNextSheet = 0;
			gpSheets[0] = NULL;
			gbBookChanged = TRUE;
			strcpy(gcCurrentFileName, lcFileName);
			lnReturn = 1; /* Empty, but OK, so we return 1 */	
			}
		}
	if (lnReturn == 0)
		{
		if (gpCurrentBook != NULL) xlBookRelease(gpCurrentBook);
		gpCurrentBook = NULL;	
		}
	//ShowEnums();
	return(lnReturn);	
}



/* ************************************************************************************** */
/* Sets column widths for one or more columns.                                            */
long DOEXPORT xlwSetColWidths(long lpnSheet, long lpnFromCol, long lpnToCol, double lpnWidth)
{
	FormatHandle lpFormat;
	SheetHandle lpSheet;
	long lnReturn = 0;
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetSetCol(lpSheet, (int) lpnFromCol, (int) lpnToCol, lpnWidth, NULL, 0);
		if (lnReturn == FALSE) SetErrorCodes(-27);
		else gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return (lnReturn);	
}

/* ************************************************************************************** */
/* Sets the Row Grouping Summary Row to the BOTTOM.  Pass True (1) to go BOTTOM, False
   (0) to go TOP.                                                                        */
long DOEXPORT xlwSetRowGroupBottom(long lpnSheet, long lpbBottom)
{
	SheetHandle lpSheet;
	long lnReturn = 1;
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		xlSheetSetGroupSummaryBelow(lpSheet, (int) lpbBottom);	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;
		lnReturn = 0;	
		}
	return (lnReturn);	
}

/* ************************************************************************************** */
/* Sets the Column Grouping Summary Column to the RIGHT.  Pass True (1) to go right, False
   (0) to go left.                                                                        */
long DOEXPORT xlwSetColGroupRight(long lpnSheet, long lpbRight)
{
	SheetHandle lpSheet;
	long lnReturn = 1;
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		xlSheetSetGroupSummaryRight(lpSheet, (int) lpbRight);	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;
		lnReturn = 0;	
		}
	return (lnReturn);	
}

/* ************************************************************************************** */
/* Sets a group of rows to be "Grouped" for opening/closing the group.                    */
long DOEXPORT xlwGroupRows(long lpnSheet, long lpnFromRow, long lpnToRow, long lpbCollapse)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetGroupRows(lpSheet, (int) lpnFromRow, (int) lpnToRow, (int) lpbCollapse);	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return (lnReturn);	
	
}

/* ************************************************************************************** */
/* Sets a group of columns to be "Grouped" for opening/closing the group.                 */
long DOEXPORT xlwGroupCols(long lpnSheet, long lpnFromCol, long lpnToCol, long lpbCollapse)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetGroupCols(lpSheet, (int) lpnFromCol, (int) lpnToCol, (int) lpbCollapse);	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return (lnReturn);
}


/* ************************************************************************************** */
/* Tests one row to see if hidden.                                                        */
/* Returns 1 if hidden, 0 if not hidden.                                                  */
long DOEXPORT xlwIsRowHidden(long lpnSheet, long lpnRow)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetRowHidden(lpSheet, (int) lpnRow);	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return (lnReturn);
}

/* ************************************************************************************** */
/* Sets the hidden status of one row.                                                     */
/* Returns 1 if successful, 0 on failure.                                                 */
long DOEXPORT xlwSetRowHidden(long lpnSheet, long lpnRow, long lpbHidden)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	char lcTheNum[25];
	
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;

	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetSetRowHidden(lpSheet, (int) lpnRow, (int) lpbHidden);
		if (lnReturn == FALSE)
			{
			SetErrorCodes(-28);
			strcat(gcErrorMessage, ": ");
			sprintf(lcTheNum, "%d Set to %d", (int) lpnRow, (int) lpbHidden);
			strcat(gcErrorMessage, lcTheNum);
			}
		else gbBookChanged = TRUE;	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return (lnReturn);
}

/* ************************************************************************************** */
/* Tests one column to see if hidden.                                                     */
/* Returns 1 if hidden, 0 if not hidden.                                                  */
long DOEXPORT xlwIsColHidden(long lpnSheet, long lpnCol)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetColHidden(lpSheet, (int) lpnCol);	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}		
	return (lnReturn);
}

/* ************************************************************************************** */
/* Sets the hidden status of one Column.                                                  */
/* Returns 1 if successful, 0 on failure.                                                 */
long DOEXPORT xlwSetColHidden(long lpnSheet, long lpnCol, long lpbHidden)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetSetColHidden(lpSheet, (int) lpnCol, (int) lpbHidden);
		if (lnReturn == FALSE) SetErrorCodes(-29);
		else gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}		
	return (lnReturn);
}

/* ************************************************************************************** */
/* Deletes the specified row(s) from the specified worksheet.                             */
long DOEXPORT xlwDeleteRows(long lpnSheet, long lpnFromRow, long lpnToRow)
{
	SheetHandle lpSheet;
	long lnReturn = 0;	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetRemoveRow(lpSheet, (int) lpnFromRow, (int) lpnToRow);
		if (lnReturn == FALSE) SetErrorCodes(-393);
		else gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;			
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Deletes the specified cols(s) from the specified worksheet.                             */
long DOEXPORT xlwDeleteCols(long lpnSheet, long lpnFromCol, long lpnToCol)
{
	SheetHandle lpSheet;
	long lnReturn = 0;	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetRemoveCol(lpSheet, (int) lpnFromCol, (int) lpnToCol);
		if (lnReturn == FALSE) SetErrorCodes(-394);
		else gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;			
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Inserts the specified row(s) into the specified worksheet.                             */
long DOEXPORT xlwInsertRows(long lpnSheet, long lpnFromRow, long lpnToRow)
{
	SheetHandle lpSheet;
	long lnReturn = 0;	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetInsertRow(lpSheet, (int) lpnFromRow, (int) lpnToRow);
		if (lnReturn == FALSE) SetErrorCodes(-395);
		else gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;			
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Inserts the specified cols(s) into the specified worksheet.                            */
long DOEXPORT xlwInsertCols(long lpnSheet, long lpnFromCol, long lpnToCol)
{
	SheetHandle lpSheet;
	long lnReturn = 0;	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetInsertCol(lpSheet, (int) lpnFromCol, (int) lpnToCol);
		if (lnReturn == FALSE) SetErrorCodes(-396);
		else gbBookChanged = TRUE;	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;			
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Sets one or more rows entirely to one format.  Checks for the height of each row    
   before making the change so it only changes the format, not the height.                */
long DOEXPORT xlwSetRowFormats(long lpnSheet, long lpnFromRow, long lpnToRow, long lpnFmatIndex)
{
	FormatHandle lpFormat;
	SheetHandle lpSheet;
	long lnReturn = 0;
	int jR, kC;
	double ldOldHeight = 0.0;
	
	lpFormat = gpFormats[lpnFmatIndex];
	lpSheet = gpSheets[lpnSheet];
	if ((lpFormat != NULL) && (lpSheet != NULL))
		{
		for (jR = lpnFromRow; jR < lpnToRow; jR++)
			{
			ldOldHeight = xlSheetRowHeight(lpSheet, jR);
			lnReturn = xlSheetSetRow(lpSheet, jR, ldOldHeight, lpFormat, 0);	
			}
		gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}	
	return(lnReturn);	
	
}

/* ************************************************************************************** */
/* Attach a pre-defined format to a range of cells.                                       */
long DOEXPORT xlwFormatCells(long lpnSheet, long lpnFromRow, long lpnToRow, long lpnFromCol, long lpnToCol, long lpnFmatIndex)
{
	FormatHandle lpFormat;
	SheetHandle lpSheet;
	long lnReturn = 1;
	int jR, kC;
	
	lpFormat = gpFormats[lpnFmatIndex];
	lpSheet = gpSheets[lpnSheet];
	if ((lpFormat != NULL) && (lpSheet != NULL))
		{
		for (jR = lpnFromRow; jR < lpnToRow; jR++)
			{
			for (kC = lpnFromCol; kC < lpnToCol; kC++)
				{
				xlSheetSetCellFormat(lpSheet, jR, kC, lpFormat);	
				}	
			}	
		gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return(lnReturn);	
}

/* *************************************************************************************** */
/* Applies the LibXL "split()" function                                                    */
long DOEXPORT xlwSplitSheet(long lpnSheet, long lpnRow, long lpnCol)
{
	SheetHandle lpSheet;
	long lnReturn = 1;
	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		xlSheetSplit(lpSheet, lpnRow, lpnCol);	// No return value.
		gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;
		lnReturn = 0;	
		}	
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Merge two or more cells into one big cell.                                             */
long DOEXPORT xlwMergeCells(long lpnSheet, long lpnFromRow, long lpnToRow, long lpnFromCol, long lpnToCol)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		lnReturn = xlSheetSetMerge(lpSheet, lpnFromRow, lpnToRow, lpnFromCol, lpnToCol);
		if (lnReturn) gbBookChanged = TRUE;	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}	
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Apply cell background color to one or more cells.                                      */
long DOEXPORT xlwSetCellColor(long lpnSheet, long lpnFromRow, long lpnToRow, long lpnFromCol, long lpnToCol, int lpnColor)
{
	FormatHandle lpFormat;
	SheetHandle lpSheet;
	long lnReturn = 1;
	int jR, kC;
	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		for (jR = lpnFromRow; jR < lpnToRow; jR++)
			{
			for (kC = lpnFromCol; kC < lpnToCol; kC++)
				{
				lpFormat = xlSheetCellFormat(lpSheet, jR, kC);
				if (lpFormat == NULL)
					if (gnDefaultFormat != -1)
						{
						xlSheetSetCellFormat(lpSheet, jR, kC, gpFormats[gnDefaultFormat]);
						lpFormat = xlSheetCellFormat(lpSheet, jR, kC);
						}
					else
						return(0); 
				xlFormatSetFillPattern(lpFormat, FILLPATTERN_SOLID);
				xlFormatSetPatternForegroundColor(lpFormat, lpnColor);
				xlSheetSetCellFormat(lpSheet, jR, kC, lpFormat);
				//xlSheetSetCellFormat(lpSheet, jR, kC, lpFormat);	
				}	
			}
		gbBookChanged = TRUE;	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return(lnReturn);	
}


/* ************************************************************************************** */
/* Adds a format to the gpWrapFormats[] array and returns the index to it.                */
long DOEXPORT xlwAddWrapFormat(XLFORMAT *lpFmat)
{
	long lnReturn = gnNextFormat;
	FormatHandle lpFHandle;
	FontHandle lpNHandle;
	
	lpFHandle = xlBookAddFormat(gpCurrentBook, NULL);
	lpNHandle = xlBookAddFont(gpCurrentBook, NULL);
	if ((lpFHandle != NULL) && (lpNHandle != NULL))
		{
		gpFormats[lnReturn] = lpFHandle;
		gpFonts[lnReturn] = lpNHandle;		
		gpWrapFormats[lnReturn] = *lpFmat;
		xlFormatSetShrinkToFit(lpFHandle, 0);
		PostWrapFormat(lpFmat, lpFHandle, lpNHandle);
		xlFormatSetFont(lpFHandle, lpNHandle);
		gnNextFormat += 1;
		if (gnDefaultFormat == -1) gnDefaultFormat = lnReturn;						
		}
	else
		{
		lnReturn = -1;
		SetErrorCodes(-31);
		}
	return(lnReturn);	
}


/* ************************************************************************************** */
/* Adds a sheet to the current workbook.  Returns the sheet index number in the gpSheets  
   array.  Returns 0 on error.  Note that the first element gpSheets[0] is assigned when
   the sheet is created, so this value will always be > 0.                                */
long DOEXPORT xlwAddSheet(char* lpcSheetName)
{
	SheetHandle lpTest = NULL;
	long lnReturn = 0;
	
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	if (gpCurrentBook != NULL)
		{
		lpTest = xlBookAddSheet(gpCurrentBook, lpcSheetName, NULL);
		if (lpTest != NULL)
			{
			gpSheets[gnNextSheet] = lpTest;
			gbBookChanged = TRUE;
			lnReturn = gnNextSheet;
			gnNextSheet += 1;	
			}
		else
			{
			SetErrorCodes(-5);
			lnReturn = 0;				
			}
		}
	else
		{
		strcpy(gcErrorMessage, "No Workbook Active");
		gnLastErrorNumber = -6;
		lnReturn = 0;	
		}
	return(lnReturn);	
}

/* ************************************************************************************* */
/* Returns a string representation of the value of the specified cell in the specified
   Worksheet.  Returns NULL on error.  The string may be empty if cell is empty.  The
   typecode ("D" = Date, "N" = Number, "L" = Logical, "S" = String, "E" = 
   Empty, "B" = Blank, "X" = Error or undefined) for the field content is returned in
   the lpcTypeReturn parameter.                                                           */
DOEXPORT char* xlwGetCellValue(long lpnSheetPtr, long lpnRow, long lpnCol, char* lpcTypeReturn)
{
	char* lpcPtrReturn = NULL;
	SheetHandle lpSheet = NULL;
	long lnType = 0;
	const char* lpTest = NULL;
	FormatHandle lppFmat = NULL;
	long lnValue;
	double ldValue;
	long lbDate = FALSE;
	static char lcRetBuff[10000]; // Larger than we're likely to need.
	int lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond, lnMillisecond;
	int lnFormatNumber;
	
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	lpSheet = gpSheets[lpnSheetPtr];
	if (lpSheet != NULL)
		{
		lcRetBuff[0] = (char) 0;
		lnType = xlSheetCellType(lpSheet, lpnRow, lpnCol);
		//printf("The Type: %ld\n", lnType);
		switch(lnType)
			{
			case CELLTYPE_EMPTY:
				strcpy(lpcTypeReturn, "E");
				lpcPtrReturn = lcRetBuff; // Just an empty string.
				break;
					
			case CELLTYPE_NUMBER:
				strcpy(lpcTypeReturn, "N");
				ldValue = xlSheetReadNum(lpSheet, lpnRow, lpnCol, &lppFmat);
				if (lppFmat == NULL)
					{
					SetErrorCodes(-7);
					//printf("Error Code for ldValue = %lf \n", ldValue);	
					}
				else
					{
					lnFormatNumber = xlFormatNumFormat(lppFmat);
					//printf("%d Is The Format\n", lnFormatNumber);
					if ((lnFormatNumber <= 49) || ((lnFormatNumber >= 164) && (lnFormatNumber <= 190)))// Normal type handling
						{
						if (xlSheetIsDate(lpSheet, lpnRow, lpnCol))
							{
							strcpy(lpcTypeReturn, "D");
							// We'll return a string in the form "YYYYMMDDHHMMSSmmm"  where mm is milliseconds
							lnValue = xlBookDateUnpack(gpCurrentBook, ldValue, &lnYear, &lnMonth, &lnDay, &lnHour, &lnMinute, &lnSecond, &lnMillisecond);
							if (lnValue == 0)
								{
								SetErrorCodes(-7);
								}
							else
								{
								sprintf(lcRetBuff, "%04d%02d%02d%02d%02d%02d%03d", lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond, lnMillisecond);
								lpcPtrReturn = lcRetBuff;	
								}
							}
						else
							{
							sprintf(lcRetBuff, "%lf", ldValue); // Just return a string representation of the number.
							//printf("Returning Double %lf \n", ldValue);
							lpcPtrReturn = lcRetBuff;	
							}
						}
					else
						{
						// Probably some exotic number or date format we don't recognize, so we'll return it as a number.  Added 06/07/2011. JSH.
						sprintf(lcRetBuff, "%lf", ldValue);
						//printf("Returning unknown format double %lf \n", ldValue);
						lpcPtrReturn = lcRetBuff;	
						}
					}
				break;
				
			case CELLTYPE_STRING:
				strcpy(lpcTypeReturn, "S");
				lpTest = xlSheetReadStr(lpSheet, lpnRow, lpnCol, &lppFmat);
				//printf("Type is String \n");
				if ((lpTest != NULL) && (lppFmat != NULL))
					{
					strcpy(lcRetBuff, lpTest);
					lpcPtrReturn = lcRetBuff;
					}
				else
					SetErrorCodes(-7);
				break;
									
			case CELLTYPE_BOOLEAN:
				strcpy(lpcTypeReturn, "L");
				lnValue = xlSheetReadBool(lpSheet, lpnRow, lpnCol, &lppFmat);
				if (lppFmat == NULL)
					{
					SetErrorCodes(-7);	
					}
				else
					{
					if (lnValue == 1)
						strcpy(lcRetBuff, "TRUE");
					else
						strcpy(lcRetBuff, "FALSE");
					lpcPtrReturn = lcRetBuff;
					}
				break;
				
				
			case CELLTYPE_BLANK:
				strcpy(lpcTypeReturn, "B");
				lpcPtrReturn = lcRetBuff; // Just an empty string.
				break;
								
			case CELLTYPE_ERROR:	
				strcpy(lpcTypeReturn, "X");	
				// Returns NULL too.
				strcpy(gcErrorMessage, "Cell contains ERROR");
				gnLastErrorNumber = -14939;
				break;
				
			default:
				break;
			}	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	//if (lpcPtrReturn != NULL) printf("RETURNING: %s\n", lpcPtrReturn);
	return(lpcPtrReturn);	
}

DOEXPORT char* xlwGetErrorMessage(void)
{
	return gcErrorMessage;	
}

long DOEXPORT xlwGetErrorNumber(void)
{
	return gnLastErrorNumber;
}

/* ******************************************************************************** */
/* Returns the used range of the specified sheet in the parms passed by reference.  */
/* returns char pointer to string containing the sheet name or NULL on error.       */
DOEXPORT char* xlwGetSheetStats(long lpnSheetPtr, long *lpnFirstCol, long *lpnLastCol, long *lpnFirstRow, long *lpnLastRow)
{
	char *lpStrReturn = NULL;
	long lnVal = 0;
	SheetHandle lpSheet = NULL;
	static char lcTheName[200];
	const char *lpTest;
	
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	lpSheet = gpSheets[lpnSheetPtr];
	if (lpSheet != NULL)
		{
		lpTest = xlSheetName(lpSheet);
		if (lpTest != NULL)
			{
			strcpy(lcTheName, lpTest);
			*lpnFirstCol = xlSheetFirstCol(lpSheet);
			*lpnLastCol  = xlSheetLastCol(lpSheet) - 1; // For some reason, this function returns a one-based column number.
			*lpnFirstRow = xlSheetFirstRow(lpSheet);
			*lpnLastRow  = xlSheetLastRow(lpSheet);
			lpStrReturn = lcTheName;
			}
		else
			{
			SetErrorCodes(-10);	
			}
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
		
	return(lpStrReturn);	
}

/* ********************************************************************************** */
/* Returns the sheet pointer index based on a sheet name passed.                      */
/* Returns the index into the gpSheets[] array or -1 on error.                        */
long DOEXPORT xlwGetSheetFromName(char *lpcSheetName)
{
	long lnReturn = -1;
	long jj;
	SheetHandle lpSheet = NULL;
	
	
	for (jj = 0;jj < gnNextSheet;jj++)
		{
		lpSheet = gpSheets[jj];
		if (lpSheet != NULL)
			{
			if (strcmp(lpcSheetName, xlSheetName(lpSheet)) == 0)
				{
				lnReturn = jj;
				break;	
				}	
			}	
		}	
	return(lnReturn);	
}

/* *********************************************************************************** */
/* Writes a value into a specified cell of a specified sheet.  Pass the value as a
   string representation of the value.  Returns 1 on success, 0 on failure.  In the
   second parm pass a type specifier: S, N, L, D, B, F meaning:
   - String
   - Number
   - Logical
   - Date
   - Blank
   - Formula
   
   Date values must be of the format YYYYMMDDHHMMSSmmm.
   Logical values must be TRUE or FALSE.
   Numbers may be any string representation that can be passed to the atof() function. */
long DOEXPORT xlwWriteCellValue(long lpnSheetIndx, long lpnRow, long lpnCol, char* lpcValue, char* lpcType)
{
	SheetHandle lpSheet = NULL;
	FormatHandle lpFormat = NULL;
	FormatHandle lpDefaultFormat = NULL;
	long lnTest = 0;
	long lnType;
	long lnNumFormat = 0;
	long lnReturn = 0;
	char lcTempType;
	long lbIsDateTime = FALSE;
	long lbIsTimeOnly = FALSE;
	long lbEmptyDate = FALSE;
	double ldValue = 0.0;
	char lcTempStr[1024];
	int lnYear = 0;
	int lnMonth = 0;
	int lnDay = 0;
	int lnHour = 0;
	int lnMinute = 0;
	int lnSecond = 0;
	int lnMilliseconds = 0;
			
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	lpSheet = gpSheets[lpnSheetIndx];
	if (gnDefaultFormat != -1)
		{
		lpDefaultFormat = gpFormats[gnDefaultFormat];
		}
//	else
//		{
//		lpDefaultFormat = lpFormat = xlSheetCellFormat(lpSheet, lpnRow, lpnCol);
//		}
//	printf("defa format = %ld \n", gnDefaultFormat);
	if (lpSheet != NULL)
		{
		lcTempType = *lpcType;
		switch (lcTempType)
			{
			case 'S':
				lnTest = xlSheetWriteStr(lpSheet, lpnRow, lpnCol, lpcValue, lpDefaultFormat); // No format specified at this point.
				if (lnTest == 0)
					SetErrorCodes(-12);
				else
					lnReturn = 1;
				break;
				
			case 'L':
				lnType = 0;
				if ((*lpcValue == 'T') || (*lpcValue == 't') || (*lpcValue == 'Y') || (*lpcValue == 'y'))
					lnType = 1;
				lnTest = xlSheetWriteBool(lpSheet, lpnRow, lpnCol, lnType, lpDefaultFormat);
				if (lnTest == 0)
					SetErrorCodes(-13);
				else
					lnReturn = 1;
				break;
				
			case 'N':
				ldValue = atof(lpcValue);
				lnTest = xlSheetWriteNum(lpSheet, lpnRow, lpnCol, ldValue, lpDefaultFormat);
				if (lnTest == 0)
					SetErrorCodes(-13);
				else
					lnReturn = 1;
				break;
				
			case 'D':
				lnMilliseconds = 0; // Always in this version
				lbEmptyDate = FALSE;
				//printf("Source Str >%s< \n", lpcValue);
				if ((strlen(lpcValue) == 0) || (strcmp(lpcValue, "        ") == 0) || (strcmp(lpcValue, "00000000") == 0) || (strcmp(lpcValue, "        000000") == 0))
					{
					ldValue = 0.0;
					lbEmptyDate = TRUE;
					lnYear = 0;
					lnMonth = 0;
					lnDay = 0;
					//printf("EMPTY DATE\n");	
					}
				else
					{
					if (strlen(lpcValue) == 8)
						{
						lbIsDateTime = FALSE;
						lbIsTimeOnly = FALSE;
						sscanf(lpcValue, "%4d%2d%2d", &lnYear, &lnMonth, &lnDay);
						//printf("DATE: %ld %ld %ld \n", lnYear, lnMonth, lnDay);
						}
					else
						{
						strcpy(lcTempStr, lpcValue);
						if (strncmp(lcTempStr, "        ", 8) == 0) // Blank date, but it is a time
							{
							lbIsTimeOnly = TRUE;
							sscanf(lcTempStr + 8, "%2d%2d%2d", &lnHour, &lnMinute, &lnSecond); 
							lnYear = 0;
							lnMonth = 0;
							lnDay = 0;	
							//printf("TIME: %ld %ld %ld \n", lnHour, lnMinute, lnSecond);
							}
						else
							{
							sscanf(lcTempStr, "%4d%2d%2d%2d%2d%2d", &lnYear, &lnMonth, &lnDay, &lnHour, &lnMinute, &lnSecond);
							lbIsDateTime = TRUE;
							//printf("DateTime %ld %ld %ld %ld %ld %ld\n", lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond);	
							}
						}
					}
				if (lbEmptyDate == FALSE)
					{
					ldValue = xlBookDatePack(gpCurrentBook, lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond, lnMilliseconds);
					//printf("Values %ld %ld %ld %ld %ld %ld %ld \n",  lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond, lnMilliseconds);
					if (xlSheetIsDate(lpSheet, lpnRow, lpnCol))
						{
						// Already a date, so we shouldn't have to mess with the formatting
						//printf("Already a date \n");
						if (lbEmptyDate == FALSE)
							lnTest = xlSheetWriteNum(lpSheet, lpnRow, lpnCol, ldValue, NULL);
						else
							lnTest = xlSheetWriteBlank(lpSheet, lpnRow, lpnCol, NULL);
						}
					else
						{
						// Not a date cell, so we need to set up the appropriate format.
						if (lbIsTimeOnly)
							{
							// Evidently a time only
							if (gpDefaultTimeFormat != NULL)
								{
								lpFormat = gpDefaultTimeFormat;
								}
							else
								{
								lpFormat = xlBookAddFormat(gpCurrentBook, lpDefaultFormat);
								xlFormatSetNumFormat(lpFormat, NUMFORMAT_CUSTOM_HMM);
								gpDefaultTimeFormat = lpFormat;	
								}
							}
						else
							{
							if (lbIsDateTime == FALSE)
								{
								// A Date only.
								//printf("Format Date Only %ld\n", NUMFORMAT_DATE);
								if (gpDefaultDateFormat != NULL)
									{
									lpFormat = gpDefaultDateFormat;
									}
								else
									{
									lpFormat = xlBookAddFormat(gpCurrentBook, lpDefaultFormat);
									xlFormatSetNumFormat(lpFormat, NUMFORMAT_DATE);
									gpDefaultDateFormat = lpFormat;
									}
								}
							else
								{
								// Likely a DateTime value.
								//printf("Format Datetime \n");
								if (gpDefaultDateTimeFormat != NULL)
									{
									lpFormat = gpDefaultDateTimeFormat;
									}
								else
									{
									lpFormat = xlBookAddFormat(gpCurrentBook, lpDefaultFormat);
									xlFormatSetNumFormat(lpFormat, NUMFORMAT_CUSTOM_MDYYYY_HMM);
									gpDefaultDateTimeFormat = lpFormat;
									}
								}	
							}
						}
						lnTest = xlSheetWriteNum(lpSheet, lpnRow, lpnCol, ldValue, lpFormat);
						//printf("Wrote Value %lf\n", ldValue);
					}
				else
					{
					lnTest = xlSheetWriteBlank(lpSheet, lpnRow, lpnCol, lpDefaultFormat); 
					//printf("wrote blank\n");
					}
					
				//printf("The lnTest %ld\n", lnTest);
				if (lnTest == 0)
					SetErrorCodes(-14);
				else
					lnReturn = 1;				
				break;
					
			case 'B':
				lpFormat = xlSheetCellFormat(lpSheet, lpnRow, lpnCol);
				if (lpFormat != NULL)
					lnTest = xlSheetWriteBlank(lpSheet, lpnRow, lpnCol, lpFormat);
				else
					lnTest = xlSheetWriteBlank(lpSheet, lpnRow, lpnCol, lpDefaultFormat);
				if (lnTest == 0)
					SetErrorCodes(-14);
				else
					lnReturn = 1;
				break;
				
			case 'F':
				lnTest = xlSheetWriteFormula(lpSheet, lpnRow, lpnCol, lpcValue, lpDefaultFormat);
				if (lnTest == 0)
					SetErrorCodes(-14);
				else
					lnReturn = 1;				
				break;
				
			default:
				strcpy(gcErrorMessage, "Bad Cell Type Code Passed");
				gnLastErrorNumber = -15;
				break;				
			}	
		if (lnReturn == 1) gbBookChanged = TRUE;	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return(lnReturn);
}

/* ********************************************************************************************** */
/* Gets a full row of data back from the spreadsheet.  Returns a buffer to a TAB delimited string
   with one item for each cell (empty cells have nothing between th TABs).  The type of each cell
   data returned is entered into a type string returned in the lpcTypeList passed by reference.
   Returns NULL on error.                                                                         */
/* Changed delimiter string from TAB ('\t') to an externally defined delimiter to allow for more
   creative selection of delims for special data which may include the tab character.             */                                                                     
DOEXPORT char* xlwGetRowValues(long lpnSheetPtr, long lnRow, long lnFromCol, long lnToCol, char *lpcTypeList)
{
	char* lpcPtrReturn = NULL;
	SheetHandle lpSheet = NULL;
	long lnType = 0;
	static char lcRetBuff[65000]; // Larger than we're likely to need.
	static char lcTypeList[500]; // Ditto
	long jj;
	long lnCellCnt = -1;
	long lbGoodCopy = TRUE;
	const char* lpPtr;
	char lcOneType[3];
	char lcNumBuff[20];
	
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	lpSheet = gpSheets[lpnSheetPtr];
	if (lpSheet != NULL)
		{
		lcRetBuff[0] = (char) 0;
		memset(lcTypeList, 0, (size_t) 500 * sizeof(char));
		lcNumBuff[0] = (char) 0;
		
		for (jj = lnFromCol; jj <= lnToCol; jj++)
			{
			lcOneType[1] = (char) 0;
			lpPtr = xlwGetCellValue(lpnSheetPtr, lnRow, jj, lcOneType);

			if (lpPtr != NULL)
				{
				strcat(lcRetBuff, lpPtr);
				lnCellCnt += 1;
				//if (jj < lnToCol) strcat(lcRetBuff, "\t");
				if (jj < lnToCol) strcat(lcRetBuff, gcCellDelimiter);
				lcTypeList[lnCellCnt] = lcOneType[0];
				}
			else
				{
				// Error condition.  The error codes are already set.
				strcat(gcErrorMessage, " - Col Number: ");
				sprintf(lcNumBuff, "%ld %s", jj, lcOneType);
				strcat(gcErrorMessage, lcNumBuff);
				gnLastErrorNumber = -19;
				lbGoodCopy = FALSE;
				break;	
				}
			}
		if (lbGoodCopy)
			{
			strcpy(lpcTypeList, lcTypeList);
			lpcPtrReturn = lcRetBuff;	
			}	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return(lpcPtrReturn);
}

/* ********************************************************************************************** */
/* Returns row of data back to the spreadsheet.  Returns a buffer to a TAB delimited string
   with one item for each cell (empty cells have nothing between the TABs).  The type of each cell
   data returned is entered into a type string returned in the lpcTypeList passed by reference.
   Returns NULL on error.                                                                         */
/* Note that now the delimiter is as defined by the xlwSetDelimiter() function.                   */
long DOEXPORT xlwWriteRowValues(long lpnSheetPtr, long lnRow, long lnFromCol, long lnToCol, char *lpcValueList, char *lpcTypeList)
{
	char* lpcPtrReturn = NULL;
	SheetHandle lpSheet = NULL;
	long lnResult;
	long jj;
	long lnCellCnt = -1;
	long lbGoodCopy = TRUE;
	char* lpPtr;
	char lcOneType[3];
	long lnReturn = FALSE;
	
	gcErrorMessage[0] = (char) 0;
	gnLastErrorNumber = 0;
	
	lpSheet = gpSheets[lpnSheetPtr];
	if (lpSheet != NULL)
		{
		lcOneType[0] = (char) 0;
		lcOneType[1] = (char) 0;
		for (jj = lnFromCol, lpPtr = strtokX(lpcValueList, gcCellDelimiter); jj <= lnToCol; jj++, lpPtr = strtokX(NULL, gcCellDelimiter))
			{
			if (lpPtr != NULL)
				{
				lnCellCnt += 1;
				lcOneType[0] = *(lpcTypeList + lnCellCnt);
				//printf("cnt %ld  type %s  value pointer %p \n", jj, lcOneType, lpPtr);
				lnResult = xlwWriteCellValue(lpnSheetPtr, lnRow, jj, lpPtr, lcOneType);
				if (lnResult == 0)
					{
					lbGoodCopy = FALSE;
					SetErrorCodes(-35);
					break;	
					} 
				}
			else
				{
				// Error condition.  The error codes are already set.
				lbGoodCopy = FALSE;
				strcpy(gcErrorMessage, "Ran Out of Values");
				gnLastErrorNumber = -101;
				break;	
				}
			}
		if (lbGoodCopy)
			{
			lnReturn = TRUE;
			gbBookChanged = TRUE;	
			}
//		else
//			{
//			strcpy(gcErrorMessage, "Row Write Failed");
//			gnLastErrorNumber = -18;	
//			}	
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Sets row heights for one or more rows.                                            */
long DOEXPORT xlwSetRowHeight(long lpnSheet, long lpnRow, double lpnHeight)
{
	FormatHandle lpFormat;
	SheetHandle lpSheet;
	long lnReturn = 0;
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		//lnReturn = xlSheetSetCol(lpSheet, (int) lpnFromCol, (int) lpnToCol, lpnWidth, NULL, 0);	
		lnReturn = xlSheetSetRow(lpSheet, (int) lpnRow, lpnHeight, NULL, 0);
		gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}		
	return (lnReturn);	
}

/* ************************************************************************************** */
/* Sets the print aspect ratio: Landscape or Portrait, for the specified worksheet.       */
/* Pass sheet number and 1 for Portrait and 2 for Landscape.                              */
long DOEXPORT xlwSetPrintAspect(long lpnSheet, long lpnAspect)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		xlSheetSetLandscape(lpSheet, (int) (lpnAspect == _PRINT_LANDSCAPE ? TRUE : FALSE));
		lnReturn = 1;
		gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Sets the print page margins for the specified worksheet.                               */
/* Pass the sheet number plus a code indicating which margin to set and then the width    */
/* of the margin in fractional Inches.                                                    */
/* Note margin selections:
   1 = LEFT
   2 = RIGHT
   3 = TOP
   4 = BOTTOM
   99 = ALL
   */
long DOEXPORT xlwSetPageMargins(long lpnSheet, long lpnMarginSelect, double lpnWidth)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		switch (lpnMarginSelect)
			{
			case _MARGIN_LEFT:
				xlSheetSetMarginLeft(lpSheet, lpnWidth);
				break;
				
			case _MARGIN_RIGHT:
				xlSheetSetMarginRight(lpSheet, lpnWidth);
				break;
				
			case _MARGIN_TOP:
				xlSheetSetMarginTop(lpSheet, lpnWidth);
				break;
				
			case _MARGIN_BOTTOM:
				xlSheetSetMarginBottom(lpSheet, lpnWidth);
				break;
				
			case _MARGIN_ALL:
				xlSheetSetMarginLeft(lpSheet, lpnWidth);
				xlSheetSetMarginRight(lpSheet, lpnWidth);
				xlSheetSetMarginTop(lpSheet, lpnWidth);
				xlSheetSetMarginBottom(lpSheet, lpnWidth);			
				break;
				
			default:	
				xlSheetSetMarginLeft(lpSheet, lpnWidth);			
			}
		lnReturn = 1;
		gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Sets the print aspect ratio: Landscape or Portrait, for the specified worksheet.       */
/* Pass sheet number and 1 for Portrait and 2 for Landscape.                              */
long DOEXPORT xlwSetPrintArea(long lpnSheet, long lpnFromRow, long lpnToRow, long lpnFromCol, long lpnToCol)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		xlSheetSetPrintArea(lpSheet, (int) lpnFromRow, (int) lpnToRow, (int) lpnFromCol, (int) lpnToCol);
		lnReturn = 1;
		gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Sets the paper size for determining printed appearance.  Currently 3 paper sizes are
   supported, but LibXL supports many more:
   1 = US Letter
   2 = US Legal
   3 = European A3 */
long DOEXPORT xlwSetPaperSize(long lpnSheet, long lpnSizeCode)
{
	SheetHandle lpSheet;
	long lnReturn = 0;
	
	lpSheet = gpSheets[lpnSheet];
	if (lpSheet != NULL)
		{
		switch(lpnSizeCode)
			{
			case _SIZE_PAPER_LETTER:
				xlSheetSetPaper(lpSheet, (int) PAPER_LETTER);
				break;
				
			case _SIZE_PAPER_LEGAL:
				xlSheetSetPaper(lpSheet, (int) PAPER_LEGAL);
				break;
				
			case _SIZE_PAPER_A3:
				xlSheetSetPaper(lpSheet, (int) PAPER_A3);
				break;
				
			default:	
				xlSheetSetPaper(lpSheet, (int) PAPER_LETTER);
			}	
		lnReturn = 1;
		gbBookChanged = TRUE;
		}
	else
		{
		strcpy(gcErrorMessage, gcMessageBadSheet);
		gnLastErrorNumber = gnMessageNumberBadSheet;	
		}
	return(lnReturn);	
}

/* ************************************************************************************** */
/* Sets the print aspect ratio: Landscape or Portrait, for the specified worksheet.       */
/* Pass sheet number and 1 for Portrait and 2 for Landscape.                              */
long DOEXPORT xlwSetPrintFit(long lpnSheet, long lpnWidthPages, long lpnHeightPages)
{
SheetHandle lpSheet;
long lnReturn = 0;

lpSheet = gpSheets[lpnSheet];
if (lpSheet != NULL)
	{
	xlSheetSetPrintFit(lpSheet, (int) lpnWidthPages, (int) lpnHeightPages);
	lnReturn = 1;
	gbBookChanged = TRUE;
	}
else
	{
	strcpy(gcErrorMessage, gcMessageBadSheet);
	gnLastErrorNumber = gnMessageNumberBadSheet;	
	}
return(lnReturn);	
}