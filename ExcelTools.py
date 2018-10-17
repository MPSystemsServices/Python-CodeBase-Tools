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

"""
This Module, ExcelTools.py, provides a Python wrapper around the MPSS, Inc., developed LibXLWrapper.DLL, which in turn
is a wrapper around the LIBXL.DLL 'C' Library from XLWare at www.LibXL.com.  LIBXL is a powerful alternative to using
MS Excel and controlling it with COM.  COM provides full access to all Excel features but at a significant cost in
time and licensing.

LIBXL is a commercial product, but has no run-time distribution fees and is considerably faster for common tasks
than Excel 2007 and higher.  It is available for non-Microsoft OS platforms as well.

Unfortunately for programming convenience LIBXL is highly granular and structured around the internal constructs of
the Excel file model.  These wrappers are designed to simplify the access and to allow Python modules to access it
without explicit use of ctypes -- although ctypes is the foundation on which this is built.

The work is performed by the object class xlWrap.  The methods in this module are documented fully in the docstrings
contained in it.  Also the test function at the bottom provides many examples of how to call the methods.

General Concepts:
    - This module only allows one XLS or XLSX to be worked on at one time.  If you need to access multiple spreadsheets
      simultaneously, you'll need to access LIBXL directly as it can handle multiple instances.
    - Spreadsheet contents is accessed by addressing the individual worksheets within the workbook.  Every workbook has at
      minimum one worksheet.  Worksheets are referenced by an integer index.  There is a method that returns the index
      if you pass the worksheet name.
    - Cells are accessed within a worksheet by referencing their row and column (0 based) offsets.
    - Each cell has three components which can be managed separately: Value, Format (including Font), and Formula.
    - Formats are applied by first defining the format and adding it to the workbook and saving the format index
      value returned.  Then formats can be applied to individual cells and/or rows by referring to the worksheet,
      the rows and columns and the format index value.
    - Either XLS or XLSX formats can be read and written, but this module doesn't allow for conversions since
      it doesn't permit two or more workbooks to be open at the same time.  You could, however read the contents
      of one workbook (returned as plain ASCII character strings) save them, then close the first workbook and open
      the target workbook.

2012 Jan 13 - Corrected documentation for GetRowValues(). JSH.
2012 May 30 JMc more accurate location method to find .dll
2013 Jan 04 JSH added ascii value protection with encoding to writecell() and writerow().
2018 Sep 28 JSH added xlwSetLicense function to store license for LibXL externally.
"""
from __future__ import print_function
from ctypes import *
import os
from time import time
import copy
import sys
import datetime
import MPSSCommon.MPSSBaseTools as mTools

# Defines that substitute for the C enums specified by LIBXL
# Cell Color Styles
xlColor = {"COLOR_BLACK":8, "COLOR_WHITE":9, "COLOR_RED":10, "COLOR_BRIGHTGREEN":11, "COLOR_BLUE":12, "COLOR_YELLOW":13, "COLOR_PINK":14, "COLOR_TURQUOISE":15, "COLOR_DARKRED":16,
           "COLOR_GREEN":17, "COLOR_DARKBLUE":18, "COLOR_DARKYELLOW":19, "COLOR_VIOLET":20, "COLOR_TEAL":21, "COLOR_GRAY25":22, "COLOR_GRAY50":23, "COLOR_PERIWINKLE_CF":24,
           "COLOR_PLUM_CF":25, "COLOR_IVORY_CF":26, "COLOR_LIGHTTURQUOISE_CF":27, "COLOR_DARKPURPLE_CF":28, "COLOR_CORAL_CF":29, "COLOR_OCEANBLUE_CF":30, "COLOR_ICEBLUE_CF":31,
           "COLOR_DARKBLUE_CL":32, "COLOR_PINK_CL":33, "COLOR_YELLOW_CL":34, "COLOR_TURQUOISE_CL":35, "COLOR_VIOLET_CL":36, "COLOR_DARKRED_CL":37, "COLOR_TEAL_CL":38,
           "COLOR_BLUE_CL":39, "COLOR_SKYBLUE":40, "COLOR_LIGHTTURQUOISE":41, "COLOR_LIGHTGREEN":42, "COLOR_LIGHTYELLOW":43, "COLOR_PALEBLUE":44, "COLOR_ROSE":45, "COLOR_LAVENDER":46,
           "COLOR_TAN":47, "COLOR_LIGHTBLUE":48, "COLOR_AQUA":49, "COLOR_LIME":50, "COLOR_GOLD":51, "COLOR_LIGHTORANGE":52, "COLOR_ORANGE":53, "COLOR_BLUEGRAY":54, "COLOR_GRAY40":55,
           "COLOR_DARKTEAL":56, "COLOR_SEAGREEN":57, "COLOR_DARKGREEN":58, "COLOR_OLIVEGREEN":59, "COLOR_BROWN":60, "COLOR_PLUM":61, "COLOR_INDIGO":62, "COLOR_GRAY80":63,
           "COLOR_DEFAULT_FOREGROUND": 0x0040, "COLOR_DEFAULT_BACKGROUND": 0x0041, "COLOR_TOOLTIP": 0x0051, "COLOR_AUTO": 0x7FFF}

# Cell Number Formats
xlNumFormat = {"NUMFORMAT_GENERAL":0, "NUMFORMAT_NUMBER":1, "NUMFORMAT_NUMBER_D2":2, "NUMFORMAT_NUMBER_SEP":3, "NUMFORMAT_NUMBER_SEP_D2":4,
               "NUMFORMAT_CURRENCY_NEGBRA":5, "NUMFORMAT_CURRENCY_NEGBRARED":6, "NUMFORMAT_CURRENCY_D2_NEGBRA":7, "NUMFORMAT_CURRENCY_D2_NEGBRARED":8,
               "NUMFORMAT_PERCENT":9, "NUMFORMAT_PERCENT_D2":10, "NUMFORMAT_SCIENTIFIC_D2":11, "NUMFORMAT_FRACTION_ONEDIG":12, "NUMFORMAT_FRACTION_TWODIG":13,
               "NUMFORMAT_DATE":14, "NUMFORMAT_CUSTOM_D_MON_YY":15, "NUMFORMAT_CUSTOM_D_MON":16, "NUMFORMAT_CUSTOM_MON_YY":17,
               "NUMFORMAT_CUSTOM_HMM_AM":18, "NUMFORMAT_CUSTOM_HMMSS_AM":19, "NUMFORMAT_CUSTOM_HMM":20, "NUMFORMAT_CUSTOM_HMMSS":21,
               "NUMFORMAT_CUSTOM_MDYYYY_HMM":22,
               "NUMFORMAT_NUMBER_SEP_NEGBRA":37, "NUMFORMAT_NUMBER_SEP_NEGBRARED":38,
               "NUMFORMAT_NUMBER_D2_SEP_NEGBRA":39, "NUMFORMAT_NUMBER_D2_SEP_NEGBRARED":40, "NUMFORMAT_ACCOUNT":41, "NUMFORMAT_ACCOUNTCUR":42,
               "NUMFORMAT_ACCOUNT_D2":43, "NUMFORMAT_ACCOUNT_D2_CUR":44, "NUMFORMAT_CUSTOM_MMSS":45, "NUMFORMAT_CUSTOM_H0MMSS":46,
               "NUMFORMAT_CUSTOM_MMSS0":47, "NUMFORMAT_CUSTOM_000P0E_PLUS0":48, "NUMFORMAT_TEXT":49}

# Horizontal Alignment Formats
xlAlignH = {"ALIGNH_GENERAL":0, "ALIGNH_LEFT":1, "ALIGNH_CENTER":2, "ALIGNH_RIGHT":3, "ALIGNH_FILL":4, "ALIGNH_JUSTIFY":5, "ALIGNH_MERGE":6, "ALIGNH_DISTRIBUTED":7}

# Vertical Alignment Formats
xlAlignV = {"ALIGNV_TOP":0, "ALIGNV_CENTER":1, "ALIGNV_BOTTOM":2, "ALIGNV_JUSTIFY":3, "ALIGNV_DISTRIBUTED":4}

# Cell Border Formats
xlBorder = {"BORDERSTYLE_NONE":0, "BORDERSTYLE_THIN":1, "BORDERSTYLE_MEDIUM":2, "BORDERSTYLE_DASHED":3, "BORDERSTYLE_DOTTED":4, "BORDERSTYLE_THICK":5,
            "BORDERSTYLE_DOUBLE":6, "BORDERSTYLE_HAIR":7}

# Cell Fill Pattern Formats -- This is a subset.  See the LIBXL documentation for additional options.
xlFill = {"FILLPATTERN_NONE":0, "FILLPATTERN_SOLID":1, "FILLPATTERN_GRAY50":2, "FILLPATTERN_GRAY75":3, "FILLPATTERN_GRAY25":4 }

# Text Underline Formats -- This is a subset.  See the LIBXL documentation for additional options.
xlUnderline = {"UNDERLINE_NONE":0, "UNDERLINE_SINGLE":1, "UNDERLINE_DOUBLE":2 }

# Print Margin Codes
xlMargins = {"MARGIN_LEFT": 1, "MARGIN_RIGHT": 2, "MARGIN_TOP": 3, "MARGIN_BOTTOM": 4, "MARGIN_ALL": 99}

# Print Aspect Format
xlAspect = {"PORTRAIT": 1, "LANDSCAPE": 2}

# Print Paper Sizes
xlPaperSize = {"LETTER": 1, "LEGAL": 2, "A3": 3}


class XLFORMAT(Structure):
    """
    A Data structure using the ctypes Structure construct.  This allows a single function call to pass all format information for one or more
    cells rather than a separate one for each characteristic.  The default values of -1 or "" are ignored when xlWrap sets cell formats.  Only
    members with other values are considered in format setting, so need only set the member properties that you need.
    """
    #noinspection PyTypeChecker
    _fields_ = [("nNumFormat", c_int),
                ("nAlignH", c_int),
                ("nAlignV", c_int),
                ("nWordWrap", c_int),
                ("nRotation", c_int),
                ("nIndentLevel", c_int),
                ("nBorderTop", c_int),
                ("nBorderLeft", c_int),
                ("nBorderRight", c_int),
                ("nBorderBottom", c_int),
                ("nBackColor", c_int),
                ("nFillPattern", c_int),
                ("bLocked", c_int),
                ("bHidden", c_int),
                ("cFontName", c_char * 45),
                ("bItalic", c_int),
                ("bBold", c_int),
                ("nColor", c_int),
                ("bUnderlined", c_int),
                ("nPointSize", c_int)]

    #noinspection PyMissingConstructor
    def __init__(self):
        self.nNumFormat = -1
        self.nAlignH = -1
        self.nAlignV = -1
        self.nWordWrap = -1
        self.nRotation = -1
        self.nIndentLevel = -1
        self.nBorderTop = -1
        self.nBorderLeft = -1
        self.nBorderRight = -1
        self.nBorderBottom = -1
        self.nFillPattern = -1
        self.nBackColor = -1
        self.bLocked = -1
        self.bHidden = -1
        self.cFontName = ""
        self.bItalic = -1
        self.bBold = -1
        self.nColor = -1
        self.bUnderlined = -1
        self.nPointSize = -1


class xlWrap(object):
    def __init__(self):
        """
        The DLL LibXLWrapper.dll must be in the same path as this module.  The LIBXL.DLL must be in the Python2x
        main directory.
        """
        thispath = os.path.abspath(sys.argv[0] if __name__=='__main__' else __file__)
        lcDllPath = os.path.join(os.path.split(thispath)[0], 'LibXLWrapper.dll')
        # The License codes for LIBXL must be found in a file with the path name below, consisting of
        # two lines of text, the first having the license user name, the second having the license key string.
        # Basically, the LIBXLWrapper.DLL and the LibXLLicenseInfo.txt files must be in the same directory.
        lcLicensePath = os.path.join(os.path.split(thispath)[0], "LibXLLicenseInfo.TXT")
        xFile = open(lcLicensePath, "r")
        cName = xFile.readline()
        cKey = xFile.readline()
        xFile.close()
        cLicenseName = cName.strip(" \n\r")
        cLicenseKey = cKey.strip(" \n\r")

        self.xlWrapx = CDLL(lcDllPath)

        self.xlWrapx.xlwSetLicense.argtypes = [c_char_p, c_char_p]
        self.xlWrapx.xlwSetLicense.restype = c_int

        lpName = c_char_p(cLicenseName)
        lpKey = c_char_p(cLicenseKey)
        nReturn = self.xlWrapx.xlwSetLicense(lpName, lpKey)  # This must be called first or the libxl
        # library will print garbage on all the output pages.

        self.xlWrapx.xlwOpenWorkbook.argtypes = [c_char_p]
        self.xlWrapx.xlwOpenWorkbook.restype = c_int

        self.xlWrapx.xlwCloseWorkbook.argtypes = [c_int]
        self.xlWrapx.xlwCloseWorkbook.restype = c_int

        self.xlWrapx.xlwCreateWorkbook.argtypes = [c_char_p, c_char_p, c_int]
        self.xlWrapx.xlwCreateWorkbook.restype = c_int

        self.xlWrapx.xlwAddSheet.argtypes = [c_char_p]
        self.xlWrapx.xlwAddSheet.restype = c_int

        self.xlWrapx.xlwGetCellValue.argtypes = [c_int, c_int, c_int, c_char_p]
        self.xlWrapx.xlwGetCellValue.restype = c_char_p

        self.xlWrapx.xlwGetErrorMessage.argtypes = []
        self.xlWrapx.xlwGetErrorMessage.restype = c_char_p

        self.xlWrapx.xlwGetErrorNumber.argtypes = []
        self.xlWrapx.xlwGetErrorNumber.restype = c_int

        self.xlWrapx.xlwGetSheetStats.argtypes = [c_int, POINTER(c_int), POINTER(c_int), POINTER(c_int), POINTER(c_int)]
        self.xlWrapx.xlwGetSheetStats.restype = c_char_p

        self.xlWrapx.xlwGetSheetFromName.argtypes = [c_char_p]
        self.xlWrapx.xlwGetSheetFromName.restype = c_int

        self.xlWrapx.xlwWriteCellValue.argtypes = [c_int, c_int, c_int, c_char_p, c_char_p]
        self.xlWrapx.xlwWriteCellValue.restype = c_int

        self.xlWrapx.xlwWriteRowValues.argtypes = [c_int, c_int, c_int, c_int, c_char_p, c_char_p]
        self.xlWrapx.xlwWriteRowValues.restype = c_int

        self.xlWrapx.xlwGetRowValues.argtypes = [c_int, c_int, c_int, c_int, c_char_p]
        self.xlWrapx.xlwGetRowValues.restype = c_char_p

        self.xlWrapx.xlwAddWrapFormat.argtypes = [POINTER(XLFORMAT)]
        self.xlWrapx.xlwAddWrapFormat.restype = c_int

        self.xlWrapx.xlwFormatCells.argtypes = [c_int, c_int, c_int, c_int, c_int, c_int]
        self.xlWrapx.xlwFormatCells.restype = c_int

        self.xlWrapx.xlwSetColWidths.argtypes = [c_int, c_int, c_int, c_double]
        self.xlWrapx.xlwSetColWidths.restype = c_int
        
        self.xlWrapx.xlwSetRowFormats.argtypes = [c_int, c_int, c_int, c_int]
        self.xlWrapx.xlwSetRowFormats.restype = c_int

        self.xlWrapx.xlwSetDelimiter.argtypes = [c_char_p]
        self.xlWrapx.xlwSetDelimiter.restype = c_int

        self.xlWrapx.xlwSetRowHeight.argtypes = [c_int, c_int, c_double]
        self.xlWrapx.xlwSetRowHeight.restype = c_int

        self.xlWrapx.xlwSplitSheet.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwSplitSheet.restype = c_int

        lcTab = "\t"
        self.cCellDelim = lcTab
        pDelim = c_char_p(lcTab)
        self.xlWrapx.xlwSetDelimiter(pDelim)  # Sets the delimiter as a tab unless set otherwise by the caller.

        # These print-related functions added 02/29/2012. JSH.
        self.xlWrapx.xlwSetPrintAspect.argtypes = [c_int, c_int]
        self.xlWrapx.xlwSetPrintAspect.restype = c_int

        self.xlWrapx.xlwSetPageMargins.argtypes = [c_int, c_int, c_double]
        self.xlWrapx.xlwSetPageMargins.restype = c_int

        self.xlWrapx.xlwSetPrintArea.argtypes = [c_int, c_int, c_int, c_int, c_int]
        self.xlWrapx.xlwSetPrintArea.restype = c_int

        self.xlWrapx.xlwSetPaperSize.argtypes = [c_int, c_int]
        self.xlWrapx.xlwSetPaperSize.restype = c_int

        self.xlWrapx.xlwSetPrintFit.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwSetPrintFit.restype = c_int

        self.xlWrapx.xlwIsRowHidden.argtypes = [c_int, c_int]
        self.xlWrapx.xlwIsRowHidden.restype = c_int

        self.xlWrapx.xlwIsColHidden.argtypes = [c_int, c_int]
        self.xlWrapx.xlwIsColHidden.restype = c_int

        self.xlWrapx.xlwSetRowHidden.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwSetRowHidden.restype = c_int

        self.xlWrapx.xlwSetColHidden.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwSetColHidden.restype = c_int

        self.xlWrapx.xlwSetRowGroupBottom.argtypes = [c_int, c_int]
        self.xlWrapx.xlwSetRowGroupBottom.restype = c_int

        self.xlWrapx.xlwSetColGroupRight.argtypes = [c_int, c_int]
        self.xlWrapx.xlwSetColGroupRight.restype = c_int

        self.xlWrapx.xlwGroupCols.argtypes = [c_int, c_int, c_int, c_int]
        self.xlWrapx.xlwGroupCols.restype = c_int

        self.xlWrapx.xlwGroupRows.argtypes = [c_int, c_int, c_int, c_int]
        self.xlWrapx.xlwGroupRows.restype = c_int

        self.xlWrapx.xlwSetCellColor.argtypes = [c_int, c_int, c_int, c_int, c_int, c_int]
        self.xlWrapx.xlwSetCellColor.restype = c_int

        self.xlWrapx.xlwInsertRows.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwInsertRows.restype = c_int

        self.xlWrapx.xlwInsertCols.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwInsertCols.restype = c_int

        self.xlWrapx.xlwDeleteRows.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwDeleteRows.restype = c_int

        self.xlWrapx.xlwDeleteCols.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwDeleteCols.restype = c_int

        self.xlWrapx.xlwImportSheetFrom.argtypes = [c_char_p, c_char_p]
        self.xlWrapx.xlwImportSheetFrom.restype = c_int

        self.nMainTitleFormat = -1
        self.nTitleFormat = -1
        self.nHeaderFormat = -1
        self.nBodyFormat = -1
        self.nBodyMoneyFormat = -1
        self.nBodyAccountsFormat = -1
        self.nBodyPercentFormat = -1
        self.nBodyNumberFormat = -1
        self.nBodyNumber2DFormat = -1
        self.nBodyDateFormat = -1  # Added 01/24/2012. JSH.
        self.nBodyNoWrapFormat = -1  # Ditto.
        self.nBodyBoldFormat = -1  # Added 01/25/2012. JSH
        self.nBodyBoldDateFormat = -1  # ditto.
        self.nTitleCenterFormat = -1
        self.nBodyHiddenFormat = -1
        self.xCustomFormats = dict()
        self.cErrorMessage = ""
        self.nErrorNumber = 0

    def SetPrintAspect(self, lnSheet, lpcHow="PORTRAIT"):
        """
        Sets the output printing format to either "PORTRAIT" or "LANDSCAPE".  Pass the string as the second
        parameter.  If unrecognized value is passed, defaults to PORTRAIT.  Returns 1 if the lnSheet value is
        recognized as a valid work sheet index, otherwise returns 0.
        """
        return self.xlWrapx.xlwSetPrintAspect(lnSheet, xlAspect.get(lpcHow, 0))

    def SetPageMargins(self, lnSheet, lcWhich, lnWidth):
        """
        Defines the printed output page margins for the specified worksheet.  The values to pass in lcWhich
        are TOP, BOTTOM, LEFT, RIGHT, or ALL.  Width of the margins is affected based on that selection and should
        be specified in inches as a decimal fraction.  Returns 1 if the lnSheet value is recognized as a valid
        work sheet index, otherwise returns 0.
        """
        return self.xlWrapx.xlwSetPageMargins(lnSheet, xlMargins.get(lcWhich, 99), lnWidth)

    def SetPrintArea(self, lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol):
        """
        Defines the Excel "print area" or portion of the worksheet that is "printable".  Returns 1 if the lnSheet
        value is recognized as a valid work sheet index, otherwise returns 0. Note that row and column counters
        are 0 (zero) based.
        """
        return self.xlWrapx.xlwSetPrintArea(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol)

    def SplitSheet(self, lnSheet, lnRow, lnCol):
        """
        The LibXL docs aren't very clear on what "Split" does, but it actually locks header rows and or left side
        label columns into position while the body scrolls.  In Excel-speak, this is "Freeze Panes", not "Split", which
        is something different.  To lock in place the first 4 rows, pass lnRow as 4 (the 5th physical row) and
        lnCol as 0.
        """
        return self.xlWrapx.xlwSplitSheet(lnSheet, lnRow, lnCol)

    def SetPaperSize(self, lnSheet, lpcType):
        """
        Determines the paper size for any printed output.  This overrides the printer setting based on Excel
        typical operation.  The lpcType value may be one of three text values: LETTER, LEGAL, and A3.  If the text
        string is unrecognized, the default will be set to LETTER.  Returns 1 if the lnSheet value is recognized as a
        valid work sheet index, otherwise returns 0.
        """
        return self.xlWrapx.xlwSetPaperSize(lnSheet, xlPaperSize.get(lpcType, 0))

    def SetPrintPageCountFit(self, lnSheet, lnPageWidth, lnPageHeight):
        """
        In the Excel print preview it is possible to specify that the print area to be output must fit in a
        specified number of pages wide and pages high.  This function provides that functionality.  Specify the
        work sheet index number and the number of pages of width for the output followed by the pages of height.
        The governing dimension is the width, typically set a one page (a value of 1 in the lnPageWidth
        parameter.  Height may be set to any non-zero value, but may be set to a value greater than the likely
        actual output height in pages so as to avoid unnecessary compression.  Again, note these values are set
        in numbers of pages.  Returns 1 if the lnSheet value is recognized as a valid work sheet index, otherwise
        returns 0.
        """
        return self.xlWrapx.xlwSetPrintFit(lnSheet, lnPageWidth, lnPageHeight)

    def SetRowGroupBottom(self, nSheet, bBottom):
        """
        When Rows are Grouped, typically the extra "Summary" row is assumed to be the one directly below the range
        of grouped rows.  Call this function with bBottom = True to ensure that.  Call it with False to force
        the "Summary" row to be the one immediately above the defined Group of Rows.
        """
        lnBottom = 1 if bBottom else 0
        return self.xlWrapx.xlwSetRowGroupBottom(nSheet, lnBottom)

    def SetColGroupRight(self, nSheet, bRight):
        """
        When Cols are Grouped, typically the extra "Summary" column is assumed to be the one immediately to the right
        of the specified group of Columns.  Call this function with bRight = True to ensure that.  Call it with False
        to force the "Summary" column to be the one immediately to the left of the Group of Rows.
        """
        lnRight = 1 if bRight else 0
        return self.xlWrapx.xlwSetColGroupRight(nSheet, lnRight)

    def GroupRows(self, nSheet, nFromRow, nToRow, bCollapse=False):
        """
        Sets up a group of rows to allow them to be opened and closed above or below a summary row.  Pass
        bCollapse as True to ensure that the Grouped rows are collapsed when the sheet is opened.
        """
        lnCollapse = 1 if bCollapse else 0
        return self.xlWrapx.xlwGroupRows(nSheet, nFromRow, nToRow, lnCollapse)

    def GroupCols(self, nSheet, nFromCol, nToCol, bCollapse=False):
        """
        Like Group Rows, but does it for Columns.  Once set, there is an expand/collapse button at the top of the
        spreadsheet for this Group.  Pass bCollapse as True to collapse the group prior to opening it.
        """
        return self.xlWrapx.xlwGroupCols(nSheet, nFromCol, nToCol, bCollapse)

    def InsertRows(self, lnSheet, lnFromRow, lnToRow):
        """
        Inserts rows into the specified worksheet.  Indicate the starting row where the insertion begins and
        the last row to be inserted below that.  Returns 1 on success, 0 on failure.
        """
        return self.xlWrapx.xlwInsertRows(lnSheet, lnFromRow, lnToRow)

    def InsertCols(self, lnSheet, lnFromCol, lnToCol):
        """
        Inserts columns into the specified worksheet.  Indicate the starting column (0 based count) where the
        insertion begins and the last column to the right of that.  Returns 1 on success, 0 on failure.
        """
        return self.xlWrapx.xlwInsertCols(lnSheet, lnFromCol, lnToCol)

    def DeleteRows(self, lnSheet, lnFromRow, lnToRow):
        """
        Removes the rows from the lnFromRow row to the lnToRow row.  Row numbers are 0-based.  Returns 1 on
        success, 0 on failure.
        """
        return self.xlWrapx.xlwDeleteRows(lnSheet, lnFromRow, lnToRow)

    def DeleteCols(self, lnSheet, lnFromCol, lnToCol):
        """
        Removes the specified columns from the specified worksheet.  Specify the from col and the to col (0-based
        column counts).  Returns 1 on success, 0 on failure.
        """
        return self.xlWrapx.xlsDeleteCols(lnSheet, lnFromCol, lnToCol)

    def SetRowFormats(self, lnSheet, lnFromRow, lnToRow, lnFmatIndex):
        """
        An alternative way to apply formats to all cells in one or more rows.  You'll need to test to see if the
        format characteristics you need are allowed to be set at the Row level.  If not, no error is raised.
        Returns 1 on success, 0 on failure.  Failure could be caused by not having any open spreadsheet workbooks.
        """
        return self.xlWrapx.xlwSetRowFormats(lnSheet, lnFromRow, lnToRow, lnFmatIndex)

    def SetColWidths(self, lnSheet, lnFromCol, lnToCol, ldWidth):
        """
        Sets the width (in Excel column width units passed as a double) of one or more adjacent columns.
        Returns 1 on success, 0 on failure.
        """
        return self.xlWrapx.xlwSetColWidths(lnSheet, lnFromCol, lnToCol, ldWidth)

    def SetRowHeight(self, lnSheet, lnRow, ldHeight):
        """
        Sets the height of one row in pixels.  Returns 1 on success, 0 on failure.
        """
        return self.xlWrapx.xlwSetRowHeight(lnSheet, lnRow, ldHeight)

    def AddFormat(self, lxFmat):
        """
        You define a format by creating an instance of the Format structure like:
        someFormat = XLFORMAT()
        Then set the properties (members) of the structure that you need to define for one or more cells.
        This method returns a Format index value that you'll use to refer to this Format.  If it returns
        -1, something went wrong.

        The first format you define with AddFormat is saved and becomes the Default Format.  This Format
        is automatically applied to any cells filled with WriteCellValue() or WriteRowValues() after it has
        been defined.  Defining a Default Format for the workbook prior to extensive writing can save
        considerable processing time when there are large numbers of identically formatted body cells
        to be loaded -- in comparison to writing the cell contents and adding the formatting in a second pass.
        """
        lpFM = pointer(lxFmat)
        return self.xlWrapx.xlwAddWrapFormat(lpFM)

    def FormatCells(self, lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, lnFmatIndex):
        """
        Allows you to apply a previously defined Format to a rectangular collection of cells.  Returns 1
        on success, 0 on failure.
        """
        return self.xlWrapx.xlwFormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, lnFmatIndex)

    def SetCellDelimiter(self, lcDelim):
        """
        Overrides the default value of the delimiter used to separate cell values both in the get and the
        write of multiple cells.  Should be a string value of up to 10 characters unlikely to be found in
        any text string cell value.  The default value, if this is not called is tab: '\t'.
        """
        lpDelim = c_char_p(lcDelim)
        self.cCellDelim = lcDelim
        return self.xlWrapx.xlwSetDelimiter(lpDelim)

    def OpenWorkbook(self, lcFileName):
        """
        Initializes a new spreadsheet, starting with an existing spreadsheet file on the disk.  The worksheet
        is immediately read into memory and information is captured on the number of worksheets it contains.
        This number is the return value.  If the return value is < 1, something is wrong.

        The Excel version assumed is determined by the file name extension: XLS or XLSX.

        You can access the existing worksheets by using sheet index values from 0 (the first sheet) to
        one less than the number of worksheets returned by this method.
        """
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        pName = c_char_p(lcFileName)
        nReturn = self.xlWrapx.xlwOpenWorkbook(pName)
        if nReturn <= 0:
            self.cErrorMessage = self.xlWrapx.xlwGetErrorMessage()
            self.nErrorNumber = self.xlWrapx.xlwGetErrorNumber()
        return nReturn

    def CloseWorkbook(self, lbSave):
        """
        Terminates work with the currently active spreadsheet.  If the spreadsheet has been changed during processing
        and you pass 1 for lbSave, then the workbook will be saved prior to the shutdown.  Once this method
        has executed, you can open or create another workbook (spreadsheet).
        """
        return self.xlWrapx.xlwCloseWorkbook(lbSave)

    def CreateWorkbook(self, lcFileName, lcFirstSheetName="", bNoSheet=False):
        """
        Initializes a new spreadsheet -- creating a new, empty one in memory.  The lcFileName parameter MUST
        be passed.  It must have an extension of XLS or XLSX, which tells the module what version of Excel to
        use.

        A Worksheet is immediately added to the workbook.  This worksheet will have the name as given in the
        second parameter lcFirstSheetName.  If you don't supply lcFirstSheetName, the sheet will be named
        "Sheet1".  Returns the sheet index value for that newly created sheet.

        01/19/2013 - Added bNoSheet parm.  If you set this to True, then no initial sheet will be added
        to the workbook, so you will need to start off by creating sheets with the AddSheet() or the
        ImportSheetFrom() methods.  Trying to save the Workbook without adding sheets will likely trigger
        an error.
        """
        pName = c_char_p(lcFileName)
        pSheet = c_char_p(lcFirstSheetName)
        nNoSheet = c_int(1 if bNoSheet else 0)
        return self.xlWrapx.xlwCreateWorkbook(pName, pSheet, nNoSheet)

    def ImportSheetFrom(self, cFile, cSheet):
        """
        Adds a copy of the sheet cSheet in XLS/XLSX file cFile as the next sheet in the current worksheet.
        You can create a workbook without sheets using CreateWorkbook() and specifying bNoSheet=True.
        Added 01/19/2013. JSH.
        Returns the index of the new sheet or -1 on failure.
        """
        pName = c_char_p(cFile)
        pSheet = c_char_p(cSheet)
        return self.xlWrapx.xlwImportSheetFrom(pName, pSheet)

    def AddSheet(self, lcSheetName):
        """
        Adds a new worksheet at the end of the set of worksheets.  The name lcSheetName is required.

        Returns the worksheet index value of the sheet you added.  This index will be required for all
        manipulation of that sheet.
        """
        pName = c_char_p(lcSheetName)
        return self.xlWrapx.xlwAddSheet(pName)

    def GetCellValue(self, lnSheet, lnRow, lnCol):
        """
        Pass this method a sheet index number and a row/column pair.  It will return the cell value as a tuple:
        The first element of the tuple will be a string representation of the value.  The second element of
        the tuple will be a string value indicating the type of the cell value.  Cell value types are:
        - "S" - character string
        - "N" - number (always a double)
        - "D" - date (including date, datetime, and time)
        - "L" - boolean or logical
        - "B" - Blank
        - "E" - Empty
        - "X" - Error
        Date values are returned in the form YYYYMMDDHHMMSSnnn where the year is YYYY, month MM, day DD, hour HH,
        minute MM, seconds SS, and milliseconds nnn.  Time values will have all 0 (zeros) in the date portion of the
        string.  Date values will have all 0 (zeros) in the time portion of the string.

        Boolean values are returned as "TRUE" or "FALSE".

        Returns None if there is no current spreadsheet.
        """
        lcType = "      "
        pType = c_char_p(lcType)
        lcStr = self.xlWrapx.xlwGetCellValue(lnSheet, lnRow, lnCol, pType)
        lxRet = (lcStr, pType.value)
        return lxRet

    def GetErrorMessage(self):
        """
        Returns the text string explaining the most recent error condition.
        """
        return self.xlWrapx.xlwGetErrorMessage()

    def GetErrorNumber(self):
        """
        Returns the last error number.
        """
        return self.xlWrapx.xlwGetErrorNumber()

    def __getattr__(self, cAttr):
        if cAttr == "cErrorMessage":
            return self.GetErrorMessage()
        elif cAttr == "nErrorNumber":
            return self.GetErrorNumber()
        else:
            return None

    def GetSheetStats(self, lnSheet):
        """
        Returns a tuple with 5 pieces of information about the worksheet specified in the lnSheet parameter:
        - Sheet Name
        - Topmost row with contents
        - Bottommost row with contents
        - Leftmost column with contents
        - Rightmost column with contents

        Returns an empty string for the Sheet Name if the Sheet Index is invalid or if there is no active workbook.
        """
        lnFirstCol = c_int()
        lnLastCol = c_int()
        lnFirstRow = c_int()
        lnLastRow = c_int()
        lpCol1 = pointer(lnFirstCol)
        lpColN = pointer(lnLastCol)
        lpRow1 = pointer(lnFirstRow)
        lpRowN = pointer(lnLastRow)
        lcName = self.xlWrapx.xlwGetSheetStats(lnSheet, lpCol1, lpColN, lpRow1, lpRowN)
        return lcName, lpRow1.contents.value, lpRowN.contents.value, lpCol1.contents.value, lpColN.contents.value

    def GetSheetFromName(self, lcName):
        """
        Returns the worksheet index value for the worksheet with the name passed to the function.  Returns
        -1 on error (no such sheet or no active workbook).
        """
        lpSH = c_char_p(lcName)
        return self.xlWrapx.xlwGetSheetFromName(lpSH)

    def WriteCellValue(self, lnSheet, lnRow, lnCol, lcStrValue, lcCellType):
        """
        Places a value or a formula into a specified cell.  There are five required parameters:
        - Worksheet index
        - Row number (0 based)
        - Column number (0 based)
        - Value represented as a text string.  See GetCellValue() for formatting date and logical values.
        - Desired cell type: "S", "N", "D", "L", "B", or "F".

        In the case of cell type "B" pass an empty string as the value (blank).

        In the case of "F" the string should contain a formula in the "letter number" type notation of cell
        reference.  In other words you might pass "=A1+B1" as the formula.

        Returns 1 on success, 0 on failure.
        Changed handling of unicode conversions to ascii for greater reliability.  Also added code to convert
        right and left quote signs to simple ascii quotes. jsh. 03/07/2013.
        """
        if mTools.isstr(lcStrValue):  # Added this 01/03/2013. JSH.
            bBadUnicode = False
            if isinstance(lcStrValue, unicode):
                try:
                    lcStrValue = lcStrValue.encode("cp1252", "ignore")
                except:
                    bBadUnicode = True
                if bBadUnicode:
                    lcStrValue = lcStrValue.encode("utf-8", "ignore")
            else:
                try:
                    lcStrValue = unicode(lcStrValue, "cp1252")
                except:
                    bBadUnicode = True
                if bBadUnicode:
                    bBadUnicode = False
                    try:
                        lcStrValue = unicode(lcStrValue, "utf-8")
                    except:
                        bBadUnicode = True
                if bBadUnicode:
                    raise TypeError("lcStrValue contains unrecognized unicode values")
                else:
                    lcStrValue = lcStrValue.replace(u"\x93", u'"')
                    lcStrValue = lcStrValue.replace(u"\x94", u'"')
                    lcStrValue = lcStrValue.encode("ascii", "ignore")
        else:
            raise TypeError("lcStrValue must be a string")

        lpVL = c_char_p(str(lcStrValue))  # prevent access violation
        lpTY = c_char_p(str(lcCellType))  # "
        return self.xlWrapx.xlwWriteCellValue(lnSheet, lnRow, lnCol, lpVL, lpTY)

    def WriteRowValues(self, lnSheet, lnRow, lnColFrom, lnColTo, lcStrValues, lcCellTypes):
        """
        Exactly like WriteCellValue() except that the lcStrValues parameter must be a TAB
        delimited string containing one value for each cell you specify to be filled.  The
        parameter lcCellTypes is a string with one character for each cell to be filled, containing
        the corresponding data type to be stored.  See WriteCellValue() for valid types.
        Changed handling of unicode conversions to ascii for greater reliability. JSH. 03/07/2013.
        """
        bUnicodeFailure = False
        if mTools.isstr(lcStrValues):  # Added this 01/03/2013. JSH.
            # lcStrValue = lcStrValues.encode("cp1252", "ignore")
            # Substituted "replace" for "ignore".  In Py 2.7, "ignore" still triggers an error condition
            # if a non-translatable character is found.  "replace" does not. 04/02/2015. JSH.
            # "replace" substitutes a '?' for the bad char.
            if isinstance(lcStrValues, unicode):
                try:
                    lcStrValues = lcStrValues.encode("cp1252", "replace")
                except:
                    bUnicodeFailure = True
                if bUnicodeFailure:
                    lcStrValues = lcStrValues.encode("utf-8", "replace")
            else:
                try:
                    lcStrValues = unicode(lcStrValues, "cp1252", "replace")
                except:
                    bUnicodeFailure = True
                if bUnicodeFailure:
                    try:
                        bUnicodeFailure = False
                        lcStrValues = unicode(lcStrValues, "utf-8", "replace")
                    except:
                        bUnicodeFailure = True
                if bUnicodeFailure:
                    raise TypeError("lcStrValues contains unrecognized unicode value")
                else:
                    lcStrValues = lcStrValues.encode("ascii", "replace")
        else:
            raise TypeError("lcStrValues must be a string")
            # lcStrValues = "" # In case things get past here...
        if ((self.cCellDelim + ".NULL.") in lcStrValues) or ((".NULL." + self.cCellDelim) in lcStrValues):
            # Yikes, there is a VFP data .NULL. value (= Python None) in the string, so we have to do
            # a type conversion.  Bummer, slows things down.  Someday handle this in the Excel wrapper engine.
            xVals = lcStrValues.split(self.cCellDelim)
            xTypes = list(lcCellTypes)
            for xI, xV in enumerate(xVals):
                if xV == ".NULL.":  # The signal that the DBF file has a .NULL. value in this field.
                    xVals[xI] = ".NULL."
                    xTypes[xI] = "S"  # change to type = "Blank"
            lcStrValues = self.cCellDelim.join(xVals)
            lcCellTypes = ("").join(xTypes)

        lpVL = c_char_p(str(lcStrValues))  # prevent access violation
        lpTY = c_char_p(str(lcCellTypes))  #
        return self.xlWrapx.xlwWriteRowValues(lnSheet, lnRow, lnColFrom, lnColTo, lpVL, lpTY)

    def ReadRowValues(self, lnSheet, lnRow, lnColFrom, lnColTo):
        """
        Like GetRowValues except returns a simple list of values (or None on error) containing
        the values of each cell specified in native Python format.  Not for use in high performance situations.
        """
        lcVals, lcTypes = self.GetRowValues(lnSheet, lnRow, lnColFrom, lnColTo)
        if lcVals is None:
            return None
        lxReturn = list()
        lxVals = lcVals.split("\t")
        for jj, xVal in enumerate(lxVals):
            lxReturn.append(self.TransformExcelValue(xVal, lcTypes[jj]))
        return lxReturn

    def TransformExcelValue(self, lpxVal, lpcType):
        """
        Takes the raw value from a cell plus the type lpcType and returns the native Python value which
        these represent:
        Allowed lpcType values...
        - "S" - character string - returns a string
        - "N" - number (always a double) - returns a double
        - "D" - date (including date, datetime, and time) - returns a datetime
          NOTE: If the date value in Excel is empty or invalid, (year 0), returns None.
        - "L" - boolean or logical - returns a logical
        - "B" - Blank returns ""
        - "E" - Empty returns ""
        - "X" - Error returns "ERROR"
        """
        if lpcType == "S":
            return str(lpxVal)
        elif lpcType == "N":
            if not "." in lpxVal:
                lpxVal = lpxVal + ".0"
            return float(lpxVal)
        elif lpcType == "D":  # YYYYMMDDHHMMSSnnn
            lnYear = int(lpxVal[0:4])
            lnMonth = int(lpxVal[4:6])
            lnDay = int(lpxVal[6:8])
            lnHour = int(lpxVal[8:10])
            lnMinute = int(lpxVal[10:12])
            lnSecond = int(lpxVal[12:14])
            lnMilli = int(lpxVal[14:17])
            if (not lnYear) or (not lnMonth) or (not lnDay):
                return None
            return datetime.datetime(lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond, lnMilli)
        elif lpcType == "L":
            return True if lpxVal == "TRUE" else False
        elif lpcType == "B":
            return ""
        elif lpcType == "E":
            return ""
        elif lpcType == "X":
            return 'ERROR'
        else:
            return "UNKNOWN"

    def GetRowValues(self, lnSheet, lnRow, lnColFrom, lnColTo):
        """
        Gets the values from a portion of a row from lnColFrom to lnColTo.  Exactly like GetCellValue()
        except that the tuple contains a TAB-delimited string of values and a value type string
        with one character per cell indicating the type found in the cell.

        On error returns an empty string as the first element of the tuple.
        CORRECTION: On error returns None as the first element of the tuple. Jan 13, 2012. JSH.
        """
        lcType = " " * 300
        lpTy = c_char_p(lcType)
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lcVals = self.xlWrapx.xlwGetRowValues(lnSheet, lnRow, lnColFrom, lnColTo, lpTy)
        if lcVals is None:
            self.cErrorMessage = self.xlWrapx.xlwGetErrorMessage()
            self.nErrorNumber = self.xlWrapx.xlwGetErrorNumber()
        return lcVals, lpTy.value

    def MergeCells(self, lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol):
        """
        Applies the Excel 'Cell Merge' feature to the specified cells.
        """
        lnReturn = self.xlWrapx.xlwMergeCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol)
        return lnReturn

    def SetColorByName(self, lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, lcColor):
        """
        Pass the name in the xlColor dict() rather than the number.
        """
        if lcColor in xlColor:
            return self.SetCellColor(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, xlColor[lcColor])
        else:
            return None

    def SetCellColor(self, lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, lnColor):
        """
        Applies the specified cell background color to the indicated range of cells.
        """
        lnFR = c_int(lnFromRow)
        lnTR = c_int(lnToRow)
        lnFC = c_int(lnFromCol)
        lnTC = c_int(lnToCol)
        lnSH = c_int(lnSheet)
        lnCO = c_int(lnColor)
        return self.xlWrapx.xlwSetCellColor(lnSH, lnFR, lnTR, lnFC, lnTC, lnCO)

    def IsRowHidden(self, lnSheet, lnRow):
        """
        Tests to see if the specified row in the specified sheet is hidden.
        Returns True if hidden, False otherwise.  Note that a False may also be returned
        if an error was encountered.  If you get False and were expecting True, then
        call the GetErrorMessage() method or GetErrorNumber() method to see if an
        error was reported.
        """
        return True if self.xlWrapx.xlwIsRowHidden(lnSheet, lnRow) == 1 else False

    def IsColHidden(self, lnSheet, lnCol):
        """
        Tests to see if the specified column in the specified sheet is hidden.
        Returns True if hidden, False otherwise.  Note that a False may also be returned
        if an error was encountered.  If you get False and were expecting True, then
        call the GetErrorMessage() method or GetErrorNumber() method to see if an
        error was reported.
        """
        return True if self.xlWrapx.xlwIsColHidden(lnSheet, lnCol) == 1 else False

    def SetRowHidden(self, lnSheet, lnRow, lbHide):
        """
        Sets the specified row in the specified sheet to hidden or NOT hidden,
        based on the True/False status of lbHide.  Returns 1 on success, 0 on failure.
        If 0 is returned, call GetErrorMessage() or GetErrorNumber() to determine the
        reason for the problem.  Does not report an error if the specified row's
        hidden status is already set to the value of lbHide.
        """
        lnHow = 1 if lbHide else 0
        return self.xlWrapx.xlwSetRowHidden(lnSheet, lnRow, lnHow)

    def SetColHidden(self, lnSheet, lnCol, lbHide):
        """
        Sets the specified column in the specified sheet to hidden or NOT hidden,
        based on the True/False status of lbHide.  Returns 1 on success, 0 on failure.
        If 0 is returned, call GetErrorMessage() or GetErrorNumber() to determine the
        reason for the problem.  Does not report an error if the specified column's
        hidden status is already set to the value of lbHide.
        """
        lnHow = 1 if lbHide else 0
        return self.xlWrapx.xlwSetColHidden(lnSheet, lnCol, lnHow)

    def makeHeaderFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.bBold = 1
        #lpFmat.cFontName = "Arial"
        lpFmat.nColor = xlColor["COLOR_BLACK"]
        lpFmat.nWordWrap = 1
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeTitleFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.bBold = 1
        lpFmat.cFontName = "Arial"
        lpFmat.nColor = xlColor["COLOR_BLACK"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 12
        return copy.copy(lpFmat)

    def makeBodyFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 1
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeBodyHiddenFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nAlignH = xlAlignH["ALIGNH_CENTER"]
        lpFmat.bBold = 1
        lpFmat.cFontName = "Arial"
        lpFmat.nColor = xlColor["COLOR_WHITE"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeTitleCenterFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nAlignH = xlAlignH["ALIGNH_CENTER"]
        lpFmat.bBold = 1
        lpFmat.cFontName = "Arial"
        lpFmat.nColor = xlColor["COLOR_BLACK"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 12
        return copy.copy(lpFmat)

    def makeMainTitleFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.bBold = 1
        lpFmat.cFontName = "Arial"
        lpFmat.nColor = xlColor["COLOR_BLACK"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 14
        return copy.copy(lpFmat)

    def makeBodyBoldFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.bBold = 1
        lpFmat.nWordWrap = 1
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeBodyBoldDateFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_DATE"]
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.bBold = 1
        lpFmat.nPointSize = 10
        lpFmat.nWordWrap = 0
        return copy.copy(lpFmat)

    def makeBodyNoWrapFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeBodyDateFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_DATE"]
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeBodyAccountsFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_CURRENCY_NEGBRA"]
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeBodyMoneyFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_CURRENCY_D2_NEGBRA"]
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeBodyPercentFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_PERCENT_D2"]
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeBodyNumberFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_NUMBER_SEP"]
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeBodyNumber2DFormat(self):
        lpFmat = XLFORMAT()
        lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_NUMBER_SEP_D2"]
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def setStandardFormat(self):
        self.nBodyFormat =         self.AddFormat(self.makeBodyFormat())
        self.nMainTitleFormat =    self.AddFormat(self.makeMainTitleFormat())
        self.nTitleFormat =        self.AddFormat(self.makeTitleFormat())
        self.nHeaderFormat =       self.AddFormat(self.makeHeaderFormat())
        self.nBodyNumberFormat =   self.AddFormat(self.makeBodyNumberFormat())
        self.nBodyMoneyFormat =    self.AddFormat(self.makeBodyMoneyFormat())
        self.nBodyNumber2DFormat = self.AddFormat(self.makeBodyNumber2DFormat())
        self.nTitleCenterFormat =  self.AddFormat(self.makeTitleCenterFormat())
        self.nBodyDateFormat =     self.AddFormat(self.makeBodyDateFormat())
        self.nBodyNoWrapFormat =   self.AddFormat(self.makeBodyNoWrapFormat())
        self.nBodyBoldFormat =     self.AddFormat(self.makeBodyBoldFormat())
        self.nBodyBoldDateFormat = self.AddFormat(self.makeBodyBoldDateFormat())
        self.nBodyPercentFormat  = self.AddFormat(self.makeBodyPercentFormat())
        self.nBodyAccountsFormat  = self.AddFormat(self.makeBodyAccountsFormat())
        self.nBodyHiddenFormat    = self.AddFormat(self.makeBodyHiddenFormat())
        return 1

    def applyStandardFormat(self, lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, lcType):
        if lcType.upper() == "BODY":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyFormat)
        elif lcType.upper() == "MAINTITLE":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nMainTitleFormat)
        elif lcType.upper() == "BODYNOWRAP":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyNoWrapFormat)
        elif lcType.upper() == "BODYBOLD":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyBoldFormat)
        elif lcType.upper() == "BODYBOLDDATE":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyBoldDateFormat)
        elif lcType.upper() == "TITLECENTER":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nTitleCenterFormat)
        elif lcType.upper() == "TITLE":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nTitleFormat)
        elif lcType.upper() == "HEADER":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nHeaderFormat)
        elif lcType.upper() == "ACCOUNTS":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyAccountsFormat)
        elif lcType.upper() == "MONEY":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyMoneyFormat)
        elif lcType.upper() == "PERCENT":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyPercentFormat)
        elif lcType.upper() == "NUMBER":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyNumberFormat)
        elif lcType.upper() == "NUMBER2D":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyNumber2DFormat)
        elif lcType.upper() == "DATE":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyDateFormat)
        elif lcType.upper() == "HIDDEN":
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyHiddenFormat)
        elif lcType in self.xCustomFormats:
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.xCustomFormats[lcType])
        else:
            self.FormatCells(lnSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyFormat)
        return 1

    def makeCustomFormat(self, lcBaseFormat, lcNewName, lcNewBackColor="", lcNewAlignH="", lcNewFont="",
                         lnNewPointSize=-1, lcNewFontColor=""):
        lcFmat = lcBaseFormat.upper()
        if lcNewBackColor in xlColor:
            lnNewBackColor = xlColor[lcNewBackColor]
        else:
            lnNewBackColor = -1
        if lcNewAlignH in xlAlignH:
            lnNewAlignH = xlAlignH[lcNewAlignH]
        else:
            lnNewAlignH = -1
        if lcNewFontColor in xlColor:
            lnNewFontColor = xlColor[lcNewFontColor]
        else:
            lnNewFontColor = -1

        lpFmat = None
        if lcFmat == "BODY":
            lpFmat = self.makeBodyFormat()
        elif lcFmat == "MAINTITLE":
            lpFmat = self.makeMainTitleFormat()
        elif lcFmat == "BODYNOWRAP":
            lpFmat = self.makeBodyNoWrapFormat()
        elif lcFmat == "TITLECENTER":
            lpFmat = self.makeTitleCenterFormat()
        elif lcFmat == "TITLE":
            lpFmat = self.makeTitleFormat()
        elif lcFmat == "HEADER":
            lpFmat = self.makeHeaderFormat()
        elif lcFmat == "MONEY":    # Money with decimal cents.
            lpFmat = self.makeBodyMoneyFormat()
        elif lcFmat == "ACCOUNTS":  # Money but drops the decimal cents from the $ amount
            lpFmat = self.makeBodyAccountsFormat()
        elif lcFmat == "PERCENT":
            lpFmat = self.makeBodyPercentFormat()
        elif lcFmat == "NUMBER":
            lpFmat = self.makeBodyNumberFormat()
        elif lcFmat == "NUMBER2D":
            lpFmat = self.makeBodyNumber2DFormat()
        elif lcFmat == "DATE":
            lpFmat = self.makeBodyDateFormat()
        elif lcFmat == "BODYBOLD":
            lpFmat = self.makeBodyBoldFormat()
        elif lcFmat == "BODYBOLDDATE":
            lpFmat = self.makeBodyBoldDateFormat()
        else:
            lpFmat = self.makeBodyFormat()

        if lnNewBackColor != -1:
            lpFmat.nBackColor = lnNewBackColor
            lpFmat.nFillPattern = xlFill["FILLPATTERN_SOLID"]
        if lnNewAlignH != -1:
            lpFmat.nAlignH = lnNewAlignH
        if lcNewFont != "":
            lpFmat.cFontName = lcNewFont
        if lnNewPointSize != -1:
            lpFmat.nPointSize = lnNewPointSize
        if lnNewFontColor != -1:
            lpFmat.nColor = lnNewFontColor
        self.xCustomFormats[lcNewName] = self.AddFormat(lpFmat)
        return 1

    # The values in WriteCellValue and WriteRowValues must be strings.  However,
    # date/datetime and logical values cannot be converted by str() to the required
    # format.  These functions provide the required conversions.  The values MUST
    # be either the required type or None.
    def prepDateTime(self, tDateTimeVal):
        if tDateTimeVal is None:
            lcReturn = ""
        else:
            lcReturn = mTools.TTOC(tDateTimeVal, 1) + "000"
        return lcReturn

    def prepDate(self, dDateVal):
        if dDateVal is None:
            lcReturn = ""
        else:
            xDT = mTools.DTOT(dDateVal)
            lcReturn = mTools.TTOC(xDT, 1) + "000"
        return lcReturn

    def prepLogical(self, bVal):
        lcReturn = "TRUE" if bVal else "FALSE"
        return lcReturn


def xlTest():
    lnStart = time()
    xlW = xlWrap()

    ln1 = xlW.CreateWorkbook("C:\\MPSSPythonScripts\\TestDataOutput\\testfile9999.xls", "myFirstSheet")
    # Open Source users should change this path to suit their installation.

    print("Create Result:", ln1)
    lnz = xlW.GetSheetFromName("myFirstSheet")
    print("Sheet from Name: ", lnz)
    xStats = xlW.GetSheetStats(0)
    print(xStats)

    fmt = XLFORMAT()
    fmt.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    fmt.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    fmt.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    fmt.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    fmt.nPointSize = 10
    fmt.nBackColor = xlColor["COLOR_LIGHTTURQUOISE"]
    fmt.nFillPattern = xlFill["FILLPATTERN_SOLID"]

    fmtHdr = XLFORMAT()
    fmtHdr.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    fmtHdr.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    fmtHdr.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    fmtHdr.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    fmtHdr.nAlignV = xlAlignV["ALIGNV_TOP"]
    fmtHdr.bBold = 1
    fmtHdr.cFontName = "Arial"
    fmtHdr.nColor = xlColor["COLOR_BLACK"]
    fmtHdr.nPointSize = 11
    fmtHdr.nWordWrap = 1

    #noinspection PyUnusedLocal
    lnFmt = xlW.AddFormat(fmt)
    lnFmtHdr = xlW.AddFormat(fmtHdr)
    xlW.SetCellDelimiter("\t")
    lcHdr = "First\tSecond\tThird\tFourth\tFifth\tSixth\tSeventh\tAnother1\tAnother two Now\tAnother3\tAnother4\tAnother5\tAnother6\tAnther7\tAnother8\tAnother9\tAnother10"
    lcHdrTypes = "SSSSSSSSSSSSSSSSS"
    lcRow = "12345\tTesting 123\tmoretesting\tFALSE\tTRUE\t235098.342\t23459087.13\t        \t20090402\t34343\t1234987\t234089\t23444.343\t343.33\t123.43\tCARRIER 1\tCARRIER 2"
    lcTypes = "NSSLLNNDDNNNNNNSS"
    print(lcTypes)
    
    xlW.WriteRowValues(lnz, 0, 0, 16, lcHdr, lcHdrTypes)
    for jj in range(1, 2400):
        xlW.WriteRowValues(lnz, jj, 0, 16, lcRow, lcTypes)

    lcGetRow, lcTypes = xlW.GetRowValues(lnz, 30, 0, 8)
    print(lcGetRow, lcTypes)
    print(xlW.GetErrorMessage())

    datalist = lcGetRow.split("\t")
    print(datalist)
    xlW.SetColWidths(lnz, 0, 17, 18.0)
    lcName, lnRow1, lnRowN, lnCol1, lnColN = xlW.GetSheetStats(lnz)
    # xlW.FormatCells(lnz, 1, lnRowN, 0, lnColN, lnFmt)
    # xlW.SetRowFormats(lnz, 2, lnRowN, lnFmt)
    xlW.FormatCells(lnz, 0, 1, 0, lnColN, lnFmtHdr)    

    print(lcName)
    print(lnCol1)
    print(lnColN)
    print(lnRow1)
    print(lnRowN)

    print("Testing HIDING")
    ln9 = xlW.SetRowHidden(lnz, 16, True)
    print("Set Hidden Row Result", ln9)
    ln9 = xlW.IsRowHidden(lnz, 16)
    print("Is Row Hidden Result:", ln9)
    ln9 = xlW.SetColHidden(lnz, 5, True)
    print("Set Hidden Col:", ln9)
    ln9 = xlW.IsColHidden(lnz, 5)
    print("Is Col Hidden:", ln9)

    ln1 = xlW.CloseWorkbook(1)
    lnEnd = time()
    print("DONE IN:", (lnEnd - lnStart))
    print(ln1)

    ln1 = xlW.CreateWorkbook("C:\\MPSSPythonScripts\\TestDataOutput\\GLOPGLOP.xls", "", True)
    # ln1a = xlW.AddSheet("myNewSheet")
    print("Empty Sheet:", ln1)
    ln2 = 999
    ln2 = xlW.ImportSheetFrom("C:\\MPSSPythonScripts\\TestDataOutput\\testfile9999.xls", "myFirstSheet")
    ln3 = xlW.CloseWorkbook(1)
    print("After Import:", ln2, ln3)

#    print("Opening LaneMaster")
#    ln2 = xlW.OpenWorkbook("\\mpsspythonscripts\\TestDataInput\\lanemaster.xlsx")
#    print(ln2)
#    if not ln2:
#        print(xlW.GetErrorMessage(), xlW.GetErrorNumber())
#    for jj in range(0, 35):
#        lcStr, lcType = xlW.GetCellValue(0, 1, jj)
#        print(lcStr, lcType)
#    xlW.CloseWorkbook(0)
#
#    print("Opening bidder list small")
#    ln3 = xlW.OpenWorkbook("\\MPSSPythonScripts\\TestDataInput\\bidderlist-small.xls")
#    print(ln3)
#    if not ln3:
#        print(xlW.GetErrorMessage())
#    for jj in range(0, 35):
#        lcStr, lcType = xlW.GetCellValue(0, 1, jj)
#        if jj < 5:
#            print(lcStr, " ", lcType)
#
#    xlW.CloseWorkbook(0)
    lnEnd = time()
    print(lnEnd - lnStart)
    print("done")
    return True


if __name__ == "__main__":
    print("***** Testing xlLibWrapper.py components")
    lbResult = xlTest()
    print(lbResult)
    print("***** Testing Complete")
