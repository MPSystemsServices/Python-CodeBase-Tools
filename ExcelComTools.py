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
ExcelComTools.py provides COM access for VFP applications to the functions in the ExcelTools.py module.
"""
from __future__ import print_function, absolute_import

from win32com.server.exception import COMException
from ExcelTools import *
from time import time
from WebServerTools.pycomsupport import COMprops, logging

if sys.version_info[0] <= 2:
    _ver3x = False
    xLongType = long
else:
    _ver3x = True
    xLongType = int

d2xconvert={"N": "N", "Y": "N", "C": "S", "D": "D", "T": "D", "I": "N", "L": "L", "M": "S", "B": "N", "F": "N",
            "G": "X", "Z": "S", "W": "X"}
x2dconvert={"N": "N", "S": "C", "L": "L", "D": "T", "X": "X", "B": "X", "E": "X"}
cErrorMessage = ""


def makeHeaderFormat():
    lpFmat = XLFORMAT()
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.bBold = 1
    lpFmat.cFontName = b"Arial"
    lpFmat.nColor = xlColor["COLOR_BLACK"]
    lpFmat.nWordWrap = 1
    return lpFmat


def makeTitleFormat():
    lpFmat = XLFORMAT()
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_NONE"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_NONE"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_NONE"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_NONE"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.bBold = 1
    lpFmat.cFontName = b"Arial"
    lpFmat.nColor = xlColor["COLOR_BLACK"]
    lpFmat.nWordWrap = 0
    lpFmat.nPointSize = 12
    return lpFmat


def makeTitleCenterFormat():
    lpFmat = XLFORMAT()
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.nAlignH = xlAlignH["ALIGNH_CENTER"]
    lpFmat.bBold = 1
    lpFmat.cFontName = b"Arial"
    lpFmat.nColor = xlColor["COLOR_BLACK"]
    lpFmat.nWordWrap = 0
    lpFmat.nPointSize = 12
    return lpFmat


def makeMainTitleFormat():
    lpFmat = XLFORMAT()
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_NONE"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_NONE"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_NONE"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_NONE"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.bBold = 1
    lpFmat.cFontName = b"Arial"
    lpFmat.nColor = xlColor["COLOR_BLACK"]
    lpFmat.nWordWrap = 0
    lpFmat.nPointSize = 14
    return lpFmat


def makeBodyFormat():
    lpFmat = XLFORMAT()
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.nWordWrap = 1
    return lpFmat


def makeBodyBoldFormat():
    lpFmat = XLFORMAT()
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.bBold = 1
    lpFmat.nWordWrap = 1
    return lpFmat


def makeBodyBoldDateFormat():
    lpFmat = XLFORMAT()
    lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_DATE"]
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.bBold = 1
    lpFmat.nWordWrap = 0
    return lpFmat


def makeBodyNoWrapFormat():
    lpFmat = XLFORMAT()
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.nWordWrap = 0
    return lpFmat


def makeBodyDateFormat():
    lpFmat = XLFORMAT()
    lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_DATE"]
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.nWordWrap = 0
    return lpFmat


def makeBodyMoneyFormat():
    lpFmat = XLFORMAT()
    lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_CURRENCY_D2_NEGBRA"]
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.nWordWrap = 0
    return lpFmat


def makeBodyNumberFormat():
    lpFmat = XLFORMAT()
    lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_NUMBER_SEP"]
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.nWordWrap = 0
    return lpFmat


def makeBodyNumber2DFormat():
    lpFmat = XLFORMAT()
    lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_NUMBER_SEP_D2"]
    lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
    lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
    lpFmat.nWordWrap = 0
    return lpFmat


def calcColumnWidth(lcType, lnWidth):
    lnWidthRet = 12.0
    if (lcType == "C") or (lcType == "N") or (lcType == "F"):
        lnWidthRet = lnWidth + 1
        if lnWidthRet > 90:
            lnWidthRet = 90

    elif lcType == "Y":
        lnWidthRet = 13.0

    elif lcType == "I":
        lnWidthRet = 12.0

    elif lcType == "T":
        lnWidthRet = 16.0

    elif lcType == "D":
        lnWidthRet = 12.0

    elif lcType == "L":
        lnWidthRet = 12.0

    elif lcType == "M":
        lnWidthRet = 40.0

    elif lcType == "B":
        lnWidthRet = 13.0

    if lnWidthRet < 12:
        lnWidthRet = 12

    return lnWidthRet


class excelTools (COMprops):
    """
    NOTE: This class is for COM applications ONLY.  Its methods are intended to be called by COM clients
    with the corresponding VARIANT type arguments.  Use from within a Python module may produce incorrect
    results.
    """
    # NEVER copy the following ID 
    # Use "print pythoncom.CreateGuid()" to make a new one    
    _public_methods_ = [ 'createworkbook', 'addsheet', 'closeworkbook', 'getsheetfromname',
                         'writecellvalue','getsheetstats', 'writerowvalues', 'setcurrentsheet',
                         'geterrormessage', 'setcolumnwidth', 'setstandardformat', 'applystandardformat',
                         'setprintaspect', 'setpagemargins', 'setprintarea', 'setpapersize', 'setprintpagecountfit',
                         'mergecells', 'setcellcolor', 'makecustomformat', 'setrowheight',
                         'openworkbook'] + COMprops._public_methods_
    if _ver3x:
        _reg_progid_ = "evosmodules.exceltools3"
        _reg_clsid_ = "{E26BED9E-391A-4BAC-9D5E-6E2B65A91B77}"  # Python V3.7
    else:
        _reg_clsid_ = "{230418F0-E086-4C49-ABF6-E115E948EA25}" # Python V2.7
        _reg_progid_ = "evosmodules.exceltools"
    
    def __init__(self):
        super(excelTools, self).__init__()
        self.cSheetName = ""
        self.bAppendFlag = False
        self.cTemplateDbf = ""
        self.lFieldList = None
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        self.cLastType = ""
        self.nCurrentSheet = -1
        self.oXL = xlWrap()
        self.nMainTitleFormat = -1
        self.nTitleFormat = -1
        self.nHeaderFormat = -1
        self.nBodyFormat = -1
        self.nBodyMoneyFormat = -1
        self.nBodyNumberFormat = -1
        self.nBodyNumber2DFormat = -1
        self.nBodyDateFormat = -1  # Added 01/24/2012. JSH.
        self.nBodyNoWrapFormat = -1  # Ditto.
        self.nBodyBoldFormat = -1  # Added 01/25/2012. JSH
        self.nBodyBoldDateFormat = -1  # ditto.
        self.nTitleCenterFormat = -1
        self.xCustomFormats = dict()

    #  For documentation on all of these methods, see the ExcelTools.py module.
    def createworkbook(self, lpcFileName, lpcFirstSheetName=""):
        """
        See documentation in ExcelTools for details
        :param lpcFileName: Fully qualified path name for the new XLS or XLSX workbook
        :param lpcFirstSheetName: Name for the first sheet.
        :return: >= 1 on OK -1 on error.
        """
        try:
            self.cErrorMessage = ""
            lnresult = self.oXL.CreateWorkbook(lpcFileName, lcFirstSheetName=lpcFirstSheetName)
            if lnresult < 0:
                self.cErrorMessage = self.oXL.GetErrorMessage()
            else:
                self.nCurrentSheet = 0

        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))        
        return lnresult
    
    def addsheet(self, lpcSheetName):
        """
        Appends a worksheet to the end of the spreadsheet's worksheet collection.
        :param lpcSheetName: The new sheet name
        :return: >=1 on OK, -1 on error.
        """
        lnResult = -1
        try:
            self.cErrorMessage = ""
            lnresult = self.oXL.AddSheet(lpcSheetName)
            if lnresult < 0:
                self.cErrorMessage = self.geterrormessage()
            else:
                self.nCurrentSheet = lnresult
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description=("Python Error Occurred: " + repr(lcType)))
        return lnresult

    def openworkbook(self, lpcFileName):
        """
        Opens an existing XLS or XLSX file and loads it into memory for further processing.  Returns the
        number of worksheets in the document or -1 on error.
        """
        lnResult = -1
        try:
            self.cErrorMessage = ""
            lnResult = self.oXL.OpenWorkbook(lpcFileName)
            if lnResult < 0:
                self.cErrorMessage = self.oXL.GetErrorMessage()
            else:
                self.nCurrentSheet = 0
        except:
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lnResult

    def closeworkbook(self, lpbDoSave):
        """
        Closes the currently open XLS or XLXS file.  Pass True to the parameter if you want to save any pending
        changes still only in memory.  If you pass False, then any unsaved changes will be discarded without
        warning.  Note that since these tools are geared to openly only one XLS or XLSX file at a time, you'll
        need to call this repeatedly to move between worksheets.
        :param lpbDoSave:
        :return:
        """
        lnresult = -1
        try:
            self.cErrorMessage = ""
            lnresult = self.oXL.CloseWorkbook(lpbDoSave)
            if lnresult < 0:
                self.cErrorMessage = self.oXL.GetErrorMessage()
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))                
        return lnresult

    def getsheetfromname(self, lpcSheetName):
        """
        Every work sheet has both a name and a number.  Most of the functions that address worksheets refer to
        them by number.  Get the number from the name using this function.
        :param lpcSheetName: Name to find.  Note that it is case sensitive.
        :return: The sheet number (0 based indexing) or -1 on error/not found.
        """
        lnResult = -1
        try:
            self.cErrorMessage = ""
            lnresult = self.oXL.GetSheetFromName(lpcSheetName)
            if lnresult < 0:
                self.cErrorMessage = self.oXL.GetErrorMessage()
            else:
                self.nCurrentSheet = lnresult
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))                
        return lnresult

    def writecellvalue(self, lpnrow, lpncol, lpxvalue, isformula=False):
        """
        Note that from the Visual FoxPro (or other COM client) side, the lpxvalue may be passed in native format.
        The conversions to a form required by the excel module are handled here.
        """
        try:
            if type(lpxvalue) == int:
                lxType = "N"
                lcTempVal = "%d"%lpxvalue

            elif type(lpxvalue) == float:
                lxType = "N"
                lcTempVal = "%f"%lpxvalue

            elif mTools.isunicode(lpxvalue):
                if isformula:
                    lxType = "F"
                    lcTempVal = lpxvalue
                else:
                    lxType = "S"
                    lcTempVal = lpxvalue

            elif type(lpxvalue).__name__ == "time":
                lxType = "D"
                # Note that the win32 module makes an automatic conversion to the PyTime type object, which
                # is found only here.  This solution comes from online wikis and is NOT documented in the Py Win32 book.
                lcTempVal = lpxvalue.Format("%Y") + lpxvalue.Format("%m") + lpxvalue.Format("%d") + \
                            lpxvalue.Format("%H") + lpxvalue.Format("%M") + lpxvalue.Format("%S")
                if lcTempVal.endswith("000000"):
                    lcTempVal = lcTempVal[0:8] # This is plain date, no time portion.
                
            elif type(lpxvalue) == bool:
                lxType = "L"
                lcTempVal = ("T" if lpxvalue else "F")
            else:
                lxType = "X"
                lcTempVal = ""
            self.cLastType = lxType
            lcTempVal = lcTempVal.strip("\x00 ")
            lnresult = self.oXL.WriteCellValue(self.nCurrentSheet, lpnrow, lpncol, lcTempVal, lxType)
            if lnresult == 0:
                self.cErrorMessage = self.oXL.GetErrorMessage()
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lnresult

    def writerowvalues(self, lpnrow, lpnstartcol, lpavalues):
        """
        Like writecellvalue, this makes the conversion. Pass lpavalues as a Visual FoxPro array of cell values.
        Writes all the values to the columns starting with lpnstartcol (0 based counting) for as many columns
        as there are elements in the array.

        Returns a string with one byte for each cell written, containing the type code for that cell.
        """
        try:
            lnCnt = len(lpavalues)
            lnResult = 0
            lcTypeList = ""
            for ix, la in enumerate(lpavalues):
                lnTest = self.writecellvalue(lpnrow, lpnstartcol + ix, la)
                if lnTest > 0:
                    lnResult = lnResult + lnTest
                    lcTypeList = lcTypeList + self.cLastType
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lcTypeList

    def getsheetstats(self):
        """
        Get row and column counts, etc. from the currently select worksheet.  See the ExcelTools documentation
        for all the details.
        :return: A tuple of values.
        """
        try:
           lcSheet, lnR0, lnRn, lnC0, lnCn = self.oXL.GetSheetStats(self.nCurrentSheet)
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return [lcSheet, lnR0, lnRn, lnC0, lnCn]

    def setcurrentsheet(self, lpnIndex):
        """
        Writing and reading cell contents refer to cells in the "current worksheet".  You set that using the
        worksheet index.  If you know the worksheet name, you can get the index with the method getsheetfromname().
        :param lpnIndex: 0-based index of the desired worksheet
        :return: Sheet name or "" on error.
        """
        try:
            laTest = self.oXL.GetSheetStats(lpnIndex)
            if laTest is None:
                lcReturn = ""
            else:
                self.nCurrentSheet = lpnIndex
                lcReturn = laTest[0]
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lcReturn

    def geterrormessage(self):
        return self.cErrorMessage

    def setcolumnwidth(self, lnCol, lnWidth):
        lnReturn = None
        try:
            lnReturn = self.oXL.SetColWidths(self.nCurrentSheet, lnCol, lnCol + 1, lnWidth)
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lnReturn

    def setrowheight(self, lnRow, lnHeight):
        lnReturn = None
        try:
            lnReturn = self.oXL.SetRowHeight(self.nCurrentSheet, lnRow, lnHeight)
        except:  # Another bare exception
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lnReturn

    def setstandardformat(self):
        try:
            self.nBodyFormat = self.oXL.AddFormat(makeBodyFormat())
            self.nMainTitleFormat = self.oXL.AddFormat(makeMainTitleFormat())
            self.nTitleFormat = self.oXL.AddFormat(makeTitleFormat())
            self.nHeaderFormat = self.oXL.AddFormat(makeHeaderFormat())
            self.nBodyNumberFormat = self.oXL.AddFormat(makeBodyNumberFormat())
            self.nBodyMoneyFormat = self.oXL.AddFormat(makeBodyMoneyFormat())
            self.nBodyNumber2DFormat = self.oXL.AddFormat(makeBodyNumber2DFormat())
            self.nTitleCenterFormat = self.oXL.AddFormat(makeTitleCenterFormat())
            self.nBodyDateFormat = self.oXL.AddFormat(makeBodyDateFormat())
            self.nBodyNoWrapFormat = self.oXL.AddFormat(makeBodyNoWrapFormat())
            self.nBodyBoldFormat = self.oXL.AddFormat(makeBodyBoldFormat())
            self.nBodyBoldDateFormat = self.oXL.AddFormat(makeBodyBoldDateFormat())
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return 1

    def applystandardformat(self, lnFromRow, lnToRow, lnFromCol, lnToCol, lcType):
        try:
            if lcType.upper() == "BODY":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyFormat)

            elif lcType.upper() == "MAINTITLE":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nMainTitleFormat)

            elif lcType.upper() == "BODYNOWRAP":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyNoWrapFormat)

            elif lcType.upper() == "BODYBOLD":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyBoldFormat)

            elif lcType.upper() == "BODYBOLDDATE":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyBoldDateFormat)

            elif lcType.upper() == "TITLECENTER":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nTitleCenterFormat)

            elif lcType.upper() == "TITLE":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nTitleFormat)

            elif lcType.upper() == "HEADER":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nHeaderFormat)

            elif lcType.upper() == "MONEY":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyMoneyFormat)

            elif lcType.upper() == "NUMBER":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyNumberFormat)

            elif lcType.upper() == "NUMBER2D":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyNumber2DFormat)

            elif lcType.upper() == "DATE":
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyDateFormat)

            elif lcType in self.xCustomFormats:
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.xCustomFormats[lcType])

            else:
                self.oXL.FormatCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol, self.nBodyFormat)

        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return 1

    def mergecells(self, lnFromRow, lnToRow, lnFromCol, lnToCol):
        try:
            lnReturn = self.oXL.MergeCells(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol)
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lnReturn

    def setcellcolor(self, lnFromRow, lnToRow, lnFromCol, lnToCol, lcColorCode):
        try:
            if lcColorCode in xlColor:
                lnReturn = self.oXL.SetCellColor(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol,
                                                 xlColor[lcColorCode])
            else:
                lnReturn = 0
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lnReturn

    def makecustomformat(self, lcBaseFormat, lcNewName, lcNewBackColor="", lcNewAlignH="", lcNewFont="", lnNewPointSize=-1, lcNewFontColor=""):
        try:
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
                lpFmat = makeBodyFormat()

            elif lcFmat == "MAINTITLE":
                lpFmat = makeMainTitleFormat()

            elif lcFmat == "BODYNOWRAP":
                lpFmat = makeBodyNoWrapFormat()

            elif lcFmat == "TITLECENTER":
                lpFmat = makeTitleCenterFormat()

            elif lcFmat == "TITLE":
                lpFmat = makeTitleFormat()

            elif lcFmat == "HEADER":
                lpFmat = makeHeaderFormat()

            elif lcFmat == "MONEY":
                lpFmat = makeBodyMoneyFormat()

            elif lcFmat == "NUMBER":
                lpFmat = makeBodyNumberFormat()

            elif lcFmat == "NUMBER2D":
                lpFmat = makeBodyNumber2DFormat()

            elif lcFmat == "DATE":
                lpFmat = makeBodyDateFormat()

            elif lcFmat == "BODYBOLD":
                lpFmat = makeBodyBoldFormat()

            elif lcFmat == "BODYBOLDDATE":
                lpFmat = makeBodyBoldDateFormat()

            else:
                lpFmat = makeBodyFormat()
                
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

            self.xCustomFormats[lcNewName] = self.oXL.AddFormat(lpFmat)
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description=("Python Error Occurred: " + repr(lcType)))
        return 1

    def setprintaspect(self, lpcHow="PORTRAIT"):
        """
        Sets the output printing format to either "PORTRAIT" or "LANDSCAPE".  Pass the string as the second
        parameter.  If unrecognized value is passed, defaults to PORTRAIT.  Returns 1 if the lnSheet value is
        recognized as a valid work sheet index, otherwise returns 0.
        """
        lnReturn = 0
        try:
            lnReturn = self.oXL.SetPrintAspect(self.nCurrentSheet, lpcHow)
        except:
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description=("Python Error Occurred: " + repr(lcType)))
        return lnReturn

    def setpagemargins(self, lcWhich, lnWidth):
        """
        Defines the printed output page margins for the specified worksheet.  The values to pass in lcWhich
        are TOP, BOTTOM, LEFT, RIGHT, or ALL.  Width of the margins is affected based on that selection and should
        be specified in inches as a decimal fraction.  Returns 1 if the lnSheet value is recognized as a valid
        work sheet index, otherwise returns 0.
        """
        lnReturn = 0
        try:
            lnReturn = self.oXL.SetPageMargins(self.nCurrentSheet, lcWhich, lnWidth)
        except:
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description=("Python Error Occurred: " + repr(lcType)))
        return lnReturn

    def setprintarea(self, lnFromRow, lnToRow, lnFromCol, lnToCol):
        """
        Defines the Excel "print area" or portion of the worksheet that is "printable".  Returns 1 if the lnSheet
        value is recognized as a valid work sheet index, otherwise returns 0. Note that row and column counters
        are 0 (zero) based.
        """
        lnReturn = 0
        try:
            lnReturn = self.oXL.SetPrintArea(self.nCurrentSheet, lnFromRow, lnToRow, lnFromCol, lnToCol)
        except:
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lnReturn

    def setpapersize(self, lpcType):
        """
        Determines the paper size for any printed output.  This overrides the printer setting based on Excel
        typical operation.  The lpcType value may be one of three text values: LETTER, LEGAL, and A3.  If the text
        string is unrecognized, the default will be set to LETTER.  Returns 1 if the lnSheet value is recognized as a
        valid work sheet index, otherwise returns 0.
        """
        lnReturn = 0
        try:
            lnReturn = self.oXL.SetPaperSize(self.nCurrentSheet, lpcType)
        except:
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return lnReturn

    def setprintpagecountfit(self, lnPageWidth, lnPageHeight):
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
        lnReturn = 0
        try:
            return self.oXL.SetPrintPageCountFit(self.nCurrentSheet, lnPageWidth, lnPageHeight)
        except:
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))


__all__ = ["makeBodyBoldDateFormat", "makeBodyBoldFormat", "makeBodyDateFormat", "makeBodyFormat", "makeBodyMoneyFormat",
           "makeBodyNoWrapFormat", "makeBodyNumber2DFormat", "makeBodyNumberFormat", "makeHeaderFormat",
           "makeMainTitleFormat", "makeTitleCenterFormat", "makeTitleFormat", "excelTools", "calcColumnWidth"]

def xlTest():
    global cErrorMessage
    xlTool = excelTools()
    lnresult = 1
    lnstart = time()
    lnend = time()
    print(lnend - lnstart)
    return lnresult


if __name__ == "__main__":
    print("***** Testing ExcelComTools.py components")
    lnResult = xlTest()
    print(lnResult)
    print("***** Testing Complete")
    if _ver3x:
        print("Registering COM server... evosmodules.exceltools3")
    else:
        print("Registering COM server... evosmodules.exceltools")
    print("")
    import win32com.server.register
    win32com.server.register.UseCommandLine(excelTools)
    print("Registration Completed")
