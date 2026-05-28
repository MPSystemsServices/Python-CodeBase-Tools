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

"""DBFtools.py created by M-P System Services, Inc. to provide access to the CodeBaseTools
and ExcelTools modules in an integrated package. """

# VERSION 2 that uses 'with' features for better exception handling
# Fix applied 07/16/2011. JSH: changed name of _setdbfstru to setdbfstru.
# 7/20/2011, JMc  default to .dbf if no extension passed as VFP filename
# 04/30/2014 - JSH - Added code to terminate reading an Excel sheet when a completely blank row is encountered.

from __future__ import print_function, absolute_import

from datetime import date, datetime
import copy
from time import time
import os, tempfile
from CodeBaseTools import VFPFIELD, cbTools
from ExcelTools import *
import MPSSBaseTools as mTools

__author__ = "J. S. Heuer"

d2xConvertFrom = "NYCDTILMBFGZW"
d2xConvertTo = "NNSDDNLSNNXSX"
x2dConvert = {"N": "N", "S": "C", "L": "L", "D": "T", "X": "X", "B": "X", "E": "X"}
cErrorMessage = ""  # Keeping some old code happy.

if sys.version_info[0] <= 2:
    _ver3x = False
    xLongType = long
else:
    _ver3x = True
    xLongType = int


def addExt(name, ext='dbf'):
    n, e = os.path.splitext(name)
    if e != "":
        return name
    return name + "." + ext


class DbfXlsEngine(object):
    """
    DbfXlsEngine provides a suite of methods to copy data from DBF tables to Excel files and the reverse.  It requires
    the installation of the ExcelTools module and the CodeBaseTools module.  Note that the ExcelTools module requires
    a licensed version of LibXL and the LibXLWrapper.dll.  LibXLWrapper.dll is an open source module that is part
    of the ExcelTools distribution, but LibXL is a commercial product requiring a user name and license code to
    be stored in a libXLLicenseInfo.txt file.  See the documentation for ExcelTools for more details.

    The normal usage for this object is to pass it an object reference to an instance of cbTools() from the
    CodeBaseTools module:

        oEngine = DbfXlsEngine(cbTools())

    But, if no cbTools instance is supplied, one will be created if possible, providing this is invoked via the
    "with" mechanism.  See the dtTest() function for an example of this usage.
    """
    def __init__(self, oVFP=None):
        global d2xConvertFrom
        global d2xConvertTo
        global x2dConvert

        self.d2xconvertFrom = d2xConvertFrom
        self.d2xconvertTo = d2xConvertTo
        self.cSheetName = ""
        self.x2dconvert = x2dConvert
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        self.cOutputFieldList = ""  # Set either externally or via the parm, to limit fields output.  Comma
        # delimited list of field names
        self.xlmgr = None
        self.dbfmgr = oVFP
        self.xWidthList = None  # Python list() of widths for each column.  bWidthPixels determines the width metric.
        self.bWidthPixels = False  # Set to False for values in xWidthList being in Excel column width units.  If
        # True, assumes the width units are in pixels.  Pixel values must be integers, but Excel width units may
        # be floating point values.
        # An Excel column width unit is approximately 8.4 pixels as of version 2007.
        self.cCellDelim = "\a\v\a"
        self.bFormatDecimals = False  # if set to True, then when converting from DBF to XLS will look for
        # number type fields with 2 decimal points (usually money fields) and output them with the special
        # format NUMFORMAT_NUMBER_SEP_D2.
        self.xDefaultValues = None  # Will be a dict() with fieldname:default value pairs.  Any fieldname
        # not found in the source XLS will be set to the default value in the output DBF.
        self.bForceDefault = False  # Set this to True if you want the default values to override anything in
        # the current table fields.
        self.lAppendSheet = False  # If True, then will attempt to append the sheet to the end of the sheets
        # collection if the file already exists.
        self.cHiddenCellValue = None  # Calls for writing something hidden into one row/col.  Set font color white.
        self.nHiddenCellRow = None
        self.nHiddenCellCol = None
        self.cLastExcelName = ""
        self.bXLSXmode = False  # XLS files are faster to output.

    def __enter__(self):
        if self.dbfmgr is None:
            self.dbfmgr = cbTools()
        self.xlmgr=xlWrap()
        self.xlmgr.SetCellDelimiter("\a\v\a")  # Bell, VTab, Bell
        self.cCellDelim = "\a\v\a"
        return self
        
    def makeHeaderFormat(self):
        """
        Format creator for the column header cells containing the field names.
        :return: Format Object
        """
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
        lpFmat.nPointSize = 10  # Changed to 10 points from previous 11, which is too big. 01/04/2013. JSH.
        return copy.copy(lpFmat)

    def makeHeaderFormatXLSX(self):
        """
        Format creator for the column header cells containing the field names.  Uses default fonts from
        XLSX type files.
        :return: Format Object
        """
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.bBold = 1
        lpFmat.cFontName = b"Calibri"
        lpFmat.nColor = xlColor["COLOR_BLACK"]
        lpFmat.nWordWrap = 1
        lpFmat.nPointSize = 10  # Changed to 10 points from previous 11, which is too big. 01/04/2013. JSH.
        return copy.copy(lpFmat)

    def makeBodyHiddenFormat(self):
        """
        Format creator for the cells in a hidden row.  The hidden row option allows the raw DBF field names to
        be included in the XLS file without actually displaying them -- the column headings can then display
        human readable names.
        :return: Format Object
        """
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_NONE"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.bBold = 1
        lpFmat.cFontName = b"Arial"
        lpFmat.nColor = xlColor["COLOR_WHITE"]
        lpFmat.nWordWrap = 0
        lpFmat.nPointSize = 10
        return copy.copy(lpFmat)

    def makeTitleFormat(self):
        """
        Format creator for the top-of-the-page title if one is to be displayed.
        :return: Format Object
        """
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
        return copy.copy(lpFmat)

    def makeBodyFormat(self):
        """
        Format creator for the standard body cells containing data displayed in text format (including dates)
        :return: Format Object
        """
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 1
        return copy.copy(lpFmat)

    def makeBodyDecimalFormat(self):
        """
        Format creator for the standard body cells containing data in numeric format.  Defaults to 2 decimal places
        and the standard thousands separator character.
        :return: Format Object
        """
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 1
        lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_NUMBER_SEP_D2"]
        return copy.copy(lpFmat)

    def makeBodyFormatXLSX(self):
        """
        Format creator for the standard body cells containing data displayed in text format (including dates).
        This version uses font types standard for XLSX files.
        :return: Format Object
        """
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 1
        lpFmat.bBold = 0
        lpFmat.cFontName = b"Calibri"
        lpFmat.nColor = xlColor["COLOR_BLACK"]
        lpFmat.nPointSize = 11
        return copy.copy(lpFmat)

    def makeBodyDecimalFormatXLSX(self):
        """
        Format creator for the standard body cells containing data in numeric format.  Defaults to 2 decimal places
        and the standard thousands separator character.  This version uses font types standard for XLSX files.
        :return: Format Object
        """
        lpFmat = XLFORMAT()
        lpFmat.nBorderTop = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderLeft = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderRight = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nBorderBottom = xlBorder["BORDERSTYLE_THIN"]
        lpFmat.nAlignV = xlAlignV["ALIGNV_TOP"]
        lpFmat.nWordWrap = 1
        lpFmat.bBold = 0
        lpFmat.cFontName = b"Calibri"
        lpFmat.nColor = xlColor["COLOR_BLACK"]
        lpFmat.nPointSize = 11
        lpFmat.nNumFormat = xlNumFormat["NUMFORMAT_NUMBER_SEP_D2"]
        return copy.copy(lpFmat)

    def calcColumnWidth(self, lcType, lnWidth, lnOverride=0.0):
        if lnOverride != 0.0:
            return lnOverride / (8.4 if self.bWidthPixels else 1.0)
        lnWidthRet = 12.0
        if (lcType == "C") or (lcType == "N") or (lcType == "F"):
            lnWidthRet = lnWidth + 2
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
        if lnWidthRet < 8:
            lnWidthRet = 8
        if not self.bXLSXmode:
            lnWidthRet += 1
        return lnWidthRet

    def bizobj2excel(self, oBiz=None, cExcelName="", xHeaderList=None, cTitleOne="", cTitleTwo="", bHiddenNames=False,
                     bNoClose=False, cFieldList="", cForExpr="", cSortOrder=""):
        """
        Similar to dbf2excel except that takes as input a fully configured biz object with a table attached.  This
        can be a Biz Object for either a DBF or a MongoDB table.  This pulls all the data out with a copytoarray()
        call, then stringifies the data for storing into the cells and outputs the specified Excel.  Outputs ONLY
        .XLSX files, although that may be slower than .XLS.
        :param oBiz: Object reference to the BizObject which will be the source of the records
        :param cExcelName: Fully qualified path name for the Excel output file.  Will be forced to ext of .XLSX
        :param xHeaderList: List() object with one text string for each column of output
        :param cTitleOne: Optional title for first row
        :param cTitleTwo: Optional title for second row
        :param bHiddenNames: If True, then we'll output a hidden row with the field names directly after the title row.
        :param bNoClose: Ignored
        :param cFieldList: Comma separated list of field names that will be output and in the order listed.
        :param cForExpr: If not all records are to be output, specify a legal "for" expression here.  This will be
               passed to the Biz Object to restrict the list of records in the copytoarray() call
        :param cSortOrder:  Pass the name of a DBF index TAG or the name of a MongoDB field to control the sort order.
        :return: True on Success, otherwise False, and set the error message.
        """
        def convString(cValue=""):
            return cValue

        def convNumber(nValue=0):
            return str(nValue).strip()

        def convDate(xValue=None):
            if xValue is None:
                return ""
            else:
                return ""

        def convBool(bValue=False):
            if bValue:
                return "TRUE"
            else:
                return "FALSE"

        def convOther(xValue=None):
            return str(xValue).strip()

        lbReturn = True
        if not oBiz or oBiz is None:
            raise ValueError("Bad BizObj value")
        if not oBiz.bTableOpen:
            bTest = oBiz.reattachtable()
            if not bTest:
                self.cErrorMessage = "Table is not open..."
                return False
        bMongoMode = oBiz.bMongoMode
        if oBiz.nMaxQueryCount > 100000:
            oBiz.nMaxQueryCount = 100000  # Won't output more than 100,000 Excel rows with this routine due to
            # time and memory limits in this version.
        nCurRow = 0
        nColWidthOverride = 0.0

        nColCount = 0
        aRawFlds = oBiz.xFieldList
        if cFieldList:
            self.cOutputFieldList = cFieldList.strip()
        else:
            self.cOutputFieldList.strip()
        if self.cOutputFieldList:
            # first produce a list of fields containing only fields which are in the table
            # and which are in the cOutputFieldList.
            aNewFlds = list()
            self.cOutputFieldList = self.cOutputFieldList.upper()
            xFlds = self.cOutputFieldList.split(",")
            xOrder = dict()
            nNewPos = 0
            for lf in aRawFlds:
                cTest = lf.cName.upper()
                if cTest in xFlds:
                    aNewFlds.append(lf)
                    xOrder[cTest] = nNewPos
                    nNewPos += 1
            nFldCnt = len(aNewFlds)

            if nFldCnt == 0:
                self.cErrorMessage = "Table fields don't match supplied field list"
                oBiz.detachtable()
                return False

            aFlds = list()
            for cN in xFlds:
                nPos = xOrder.get(cN, -1)
                if nPos != -1:
                    aFlds.append(aNewFlds[nPos])
            # now aFlds should be ordered in the same sequence as the fieldList.
        else:
            aFlds = copy.copy(aRawFlds)

        fList = list()
        cTypes = ""
        cHdr = ""
        cHdrTypes = ""
        cFldNames = ""
        cWorkExcelName = cExcelName
        nFldCnt = 0
        dList = list()  # Where we store the special formats for each column, if any.
        bFoundDecimals = False
        convList = list()
        xFldList = list()
        xColList = list()
        for lf in aFlds:
            cType = lf.cType
            if cType is None:
                cType = "_"  # Probably the _id field in a MongoDB table.
            if cType in self.d2xconvertFrom:
                nPt = self.d2xconvertFrom.index(cType)
                cTest = self.d2xconvertTo[nPt: nPt + 1]
            else:
                cTest = "X"
            if cTest != "X":
                if cType in "MC":
                    convList.append(convString)
                elif cType == "L":
                    convList.append(convBool)
                elif cType in "NBYIF":
                    convList.append(convNumber)
                elif cType in "DT":
                    convList.append(convDate)
                else:
                    convList.append(convOther)

                fList.append(lf.cName)
                xColList.append(lf)
                cTypes = cTypes + cTest
                cHdr = cHdr + lf.cName + self.cCellDelim
                cFldNames = cFldNames + lf.cName + self.cCellDelim  # In case we need to save a separate hidden
                # field name row.
                cHdrTypes = cHdrTypes + "S"
                nColCount += 1
            nFldCnt += 1

        cWorkFieldList = ",".join(fList)
        cWorkFieldList = cWorkFieldList.upper()
        if xHeaderList is not None:
            if isinstance(xHeaderList, list):
                cHdr = self.cCellDelim.join(xHeaderList)
            else:
                cHdr = cHdr.lower()
        if len(fList) > 0:
            # They have at least one valid field to output
            if bMongoMode:
                cWorkFieldList = cWorkFieldList.lower()
                if mTools.isstr(cSortOrder) and cSortOrder:
                    xSortRequested = [(cSortOrder, '1')]
                else:  # Hope for the best
                    xSortRequested = cSortOrder  # Must be a list of ordering tuples.  See the MongoDB biz object docs.
            else:
                xSortRequested = cSortOrder
            if not xSortRequested:
                xSortRequested = None

            xRecs = oBiz.copytoarray(maxcount=-1, cExpr=cForExpr, cFieldList=cWorkFieldList, xSortInfo=xSortRequested)
            if xRecs is None:
                self.cErrorMessage = "Unable to get table records: " + oBiz.cErrorMessage
                self.cLastExcelName = ""
            else:
                lnRecs = len(xRecs)
                bXLSXmode = False
                self.bXLSXmode = False
                if lnRecs < 64000:
                    xParts = os.path.splitext(cWorkExcelName)
                    cExt = xParts[1].upper()
                    if cExt == ".XLSX":  # Do what they tell us, if they specify XLSX
                        bXLSXmode = True
                        self.bXLSXmode = True
                    else:
                        cWorkExcelName = mTools.FORCEEXT(cWorkExcelName, "XLS")
                else:
                    # Too big for XLS format, so we force to XLSX
                    cWorkExcelName = mTools.FORCEEXT(cWorkExcelName, "XLSX")
                    bXLSXmode = True
                    self.bXLSXmode = True
                if self.cSheetName:
                    cSheetName = self.cSheetName
                else:
                    cSheetName = mTools.JUSTSTEM(oBiz.cTableName)
                if self.lAppendSheet and os.path.exists(cExcelName):
                    nSheetCount = self.xlmgr.OpenWorkbook(cExcelName)
                    nSheet = self.xlmgr.AddSheet(cSheetName)
                else:
                    self.xlmgr.CreateWorkbook(cExcelName, cSheetName)
                    nSheetCount = 1
                    nSheet = 0

                if bXLSXmode:
                    xFmatT = self.makeTitleFormat()
                    xFmatH = self.makeHeaderFormat()
                    xFmatT.cFontName = b"Calibri"
                    xFmatH.cFontName = b"Calibri"
                    xFmatB = self.makeBodyFormatXLSX()
                    xFmatD = self.makeBodyDecimalFormatXLSX()
                else:
                    xFmatB = self.makeBodyFormat()
                    xFmatT = self.makeTitleFormat()
                    xFmatH = self.makeHeaderFormat()
                    xFmatD = self.makeBodyDecimalFormat()

                nBodyFmat = self.xlmgr.AddFormat(xFmatB)
                nTitleFmat = self.xlmgr.AddFormat(xFmatT)
                nHdrFmat = self.xlmgr.AddFormat(xFmatH)
                nDecFmat = 0
                if self.bFormatDecimals:
                    nDecFmat = self.xlmgr.AddFormat(xFmatD)
                xFmatHide = self.makeBodyHiddenFormat()
                nHiddenFmat = self.xlmgr.AddFormat(xFmatHide)
                for jj, lf in enumerate(xColList):
                    nColWidthOverride = 0.0
                    if self.xWidthList is not None:
                        try:
                            nColWidthOverride = float(self.xWidthList[jj])
                        except:
                            nColWidthOverride = 0.0
                    if nColWidthOverride >= 0.0:
                        nTestWidth = lf.nWidth
                        nNameWidth = len(lf.cName)
                        if nNameWidth > nTestWidth:
                            nTestWidth = nNameWidth
                        self.xlmgr.SetColWidths(nSheet, jj, jj, self.calcColumnWidth(lf.cType, nTestWidth,
                                                lnOverride=nColWidthOverride))
                    else:
                        self.xlmgr.SetColHidden(nSheet, nColCount, True)
                    if self.bFormatDecimals and (cType == "N") and (lf.nDecimals == 2):
                        dList.append(nDecFmat)
                        bFoundDecimals = True
                    else:
                        dList.append(None)

                if cTitleOne != "":
                    self.xlmgr.WriteCellValue(nSheet, 0, 0, cTitleOne, "S")
                    self.xlmgr.FormatCells(nSheet, 0, 1, 0, 1, nTitleFmat)
                    nCurRow += 1

                if cTitleTwo != "":
                    self.xlmgr.WriteCellValue(nSheet, nCurRow, 0, cTitleTwo, "S")
                    self.xlmgr.FormatCells(nSheet, nCurRow, nCurRow + 1, 0, 1, nTitleFmat)
                    nCurRow += 1

                self.xlmgr.WriteRowValues(nSheet, nCurRow, 0, nColCount - 1, cHdr, cHdrTypes)
                self.xlmgr.FormatCells(nSheet, nCurRow, nCurRow + 1, 0, nColCount, nHdrFmat)
                nCurRow += 1
                if bHiddenNames:
                    cFldNames = cFldNames.upper()
                    self.xlmgr.WriteRowValues(nSheet, nCurRow, 0, nColCount - 1, cFldNames, cHdrTypes)
                    self.xlmgr.SetRowHidden(nSheet, nCurRow, True)
                    nCurRow += 1

                for valDict in xRecs:
                    xRow = list()
                    for ii, lv in enumerate(fList):
                        cItem = str(convList[ii](valDict[lv]))
                        xRow.append(cItem)
                    cRow = self.cCellDelim.join(xRow)

                    self.xlmgr.WriteRowValues(nSheet, nCurRow, 0, nColCount, cRow, cTypes)
                    if self.bFormatDecimals:
                        for nn in range(0, nColCount):
                            if not (dList[nn] is None):
                                self.xlmgr.FormatCells(nSheet, nCurRow, nCurRow + 1, nn, nn + 1, dList[nn])
                    nCurRow += 1

                if self.cHiddenCellValue is not None:
                    self.xlmgr.WriteCellValue(nSheet, self.nHiddenCellRow, self.nHiddenCellCol,
                                              self.cHiddenCellValue, "S")
                    self.xlmgr.FormatCells(nSheet, self.nHiddenCellRow, self.nHiddenCellRow + 1, self.nHiddenCellCol,
                                           self.nHiddenCellCol + 1, nHiddenFmat)
                self.xlmgr.CloseWorkbook(1)
                self.cLastExcelName = cWorkExcelName
        else:
            lbReturn = False
            self.cErrorMessage = "No fields found to output"
            oBiz.detachtable()
            self.xlmgr.CloseWorkbook(0)
            self.cLastExcelName = ""
        return lbReturn

    def dbf2excel(self, lcDbfName, lcExcelName, headerlist=None, titleone="", titletwo="", rawcopy=False,
                  hiddenNames=False, noclose=False, fieldList=""):
        """
        Copies the contents of the specific DBF file into the spreadsheet named lcExcelName.
        If the spreadsheet already exists, it is overwritten.  Otherwise it is created.  There
        will be a header row containing field names and one row following that for each table
        record (row).

        If headerlist is of type list(), it should contain one text value for each column in the table/sheet.
        These text values will replace the field names in the top row. (In that case, the output can't be used
        for import back into dbf via excel2dbf().

        If titleone and/or titletwo are non-empty strings, they will be placed in A1 and A2 respectively
        in 12 point Bold, and all content will be moved down accordingly.

        Data table fields will be properly converted to Excel value types.

        Fields such as General will not be converted.  Memo will be included.

        If you are confident you know that all fields in the table are convertable to Excel types (for example
        that there are no "General" or "Memo-Binary" fields in the table), then you can use the rawcopy option
        which does no type or content checking other than the standard conversions from VFP to Excel formats
        for dates, etc.  Pass the optional parameter rawcopy=True in that case.

        Table ordering is by record.

        If hiddenNames == True, then inserts a row directly below the title row (either containing the
        field names or the values in headerlist) which will be filled with the field names AND be set to
        hidden.

        Returns True on success.  False on Failure.
        """
        lbReturn = True
        lnResult = self.dbfmgr.use(addExt(lcDbfName), readOnly=True)
        lcLoadAlias = self.dbfmgr.alias()
        lnCurRow = 0
        lnColWidthOverride = 0.0
        if lnResult == 1:
            if self.cSheetName:
                lcSheetName = self.cSheetName
            else:
                lcSheetName = self.dbfmgr.alias()
            if self.lAppendSheet and os.path.exists(lcExcelName):
                lnSheetCount = self.xlmgr.OpenWorkbook(lcExcelName)
                lnSheet = self.xlmgr.AddSheet(lcSheetName)
            else:
                self.xlmgr.CreateWorkbook(lcExcelName, lcSheetName)
                lnSheetCount = 1
                lnSheet = 0
            lnColCount = 0
            laRawFlds = self.dbfmgr.afields()
            if fieldList:
                self.cOutputFieldList = fieldList.strip()
            else:
                self.cOutputFieldList.strip()
            if self.cOutputFieldList:
                # first produce a list of fields containing only fields which are in the table
                # and which are in the cOutputFieldList.
                aNewFlds = list()
                xFlds = self.cOutputFieldList.split(",")
                xOrder = dict()
                nNewPos = 0
                for lf in laRawFlds:
                    cTest = lf.cName.upper()
                    if cTest in xFlds:
                        aNewFlds.append(lf)
                        xOrder[cTest] = nNewPos
                        nNewPos += 1
                nFldCnt = len(aNewFlds)

                if nFldCnt == 0:
                    self.cErrorMessage = "Table fields don't match supplied field list"
                    self.dbfmgr.closetable(lcLoadAlias)
                    return False
                laFlds = list()
                for cN in xFlds:
                    nPos = xOrder.get(cN, -1)
                    if nPos != -1:
                        laFlds.append(aNewFlds[nPos])
                # now laFlds should be ordered in the same sequence as the fieldList.
            else:
                laFlds = copy.copy(laRawFlds)

            if ".XLSX" in lcExcelName.upper():
                xFmatT = self.makeTitleFormat()
                xFmatH = self.makeHeaderFormat()
                xFmatT.cFontName = b"Calibri"
                xFmatH.cFontName = b"Calibri"
                xFmatB = self.makeBodyFormatXLSX()
                xFmatD = self.makeBodyDecimalFormatXLSX()
            else:
                xFmatB = self.makeBodyFormat()
                xFmatT = self.makeTitleFormat()
                xFmatH = self.makeHeaderFormat()
                xFmatD = self.makeBodyDecimalFormat()
            lnBodyFmat = self.xlmgr.AddFormat(xFmatB)
            lnTitleFmat = self.xlmgr.AddFormat(xFmatT)
            lnHdrFmat = self.xlmgr.AddFormat(xFmatH)
            lnDecFmat = 0
            if self.bFormatDecimals:
                lnDecFmat = self.xlmgr.AddFormat(xFmatD)
            xFmatHide = self.makeBodyHiddenFormat()
            lnHiddenFmat = self.xlmgr.AddFormat(xFmatHide)
            lfList = list()
            ltList = list()
            lcTypes = ""
            lcHdr = ""
            lcHdrTypes = ""
            lcFldNames = ""

            lnFldCnt = 0
            ldList = list()  # Where we store the special formats for each column, if any.
            lbFoundDecimals = False
            for lf in laFlds:
                if lf.cType in self.d2xconvertFrom:
                    lnPt = self.d2xconvertFrom.index(lf.cType)
                    lcTest = self.d2xconvertTo[lnPt: lnPt + 1]
                else:
                    lcTest = "X"
                if lcTest != "X":
                    lfList.append(lf.cName)
                    lcTypes = lcTypes + lcTest
                    lcHdr = lcHdr + lf.cName + self.cCellDelim
                    lcFldNames = lcFldNames + lf.cName + self.cCellDelim  # In case we need to save a separate hidden
                    # field name row.
                    lcHdrTypes = lcHdrTypes + "S"
                    if self.xWidthList is not None:
                        try:
                            lnColWidthOverride = float(self.xWidthList[lnFldCnt])
                        except:
                            lnColWidthOverride = 0.0
                    if lnColWidthOverride >= 0.0:
                        self.xlmgr.SetColWidths(lnSheet, lnColCount, lnColCount,
                                                self.calcColumnWidth(lf.cType, lf.nWidth,
                                                                     lnOverride=lnColWidthOverride))
                    else:
                        self.xlmgr.SetColHidden(lnSheet, lnColCount, True)
                    if self.bFormatDecimals and (lf.cType == "N") and (lf.nDecimals == 2):
                        ldList.append(lnDecFmat)
                        lbFoundDecimals = True
                    else:
                        ldList.append(None)
                    lnColCount += 1
                lnFldCnt += 1

            if headerlist is not None:
                if isinstance(headerlist, list):
                    lcHdr = self.cCellDelim.join(headerlist)
                else:
                    lcHdr = lcHdr.lower()
            if len(lfList) > 0:
                # They have at least one valid field to output
                self.dbfmgr.goto("TOP")

                lnRecs = self.dbfmgr.reccount()
                if titleone != "":
                    self.xlmgr.WriteCellValue(lnSheet, 0, 0, titleone, "S")
                    self.xlmgr.FormatCells(lnSheet, 0, 1, 0, 1, lnTitleFmat)
                    lnCurRow += 1
                if titletwo != "":
                    self.xlmgr.WriteCellValue(lnSheet, lnCurRow, 0, titletwo, "S")
                    self.xlmgr.FormatCells(lnSheet, lnCurRow, lnCurRow + 1, 0, 1, lnTitleFmat)
                    lnCurRow += 1

                self.xlmgr.WriteRowValues(lnSheet, lnCurRow, 0, lnColCount - 1, lcHdr, lcHdrTypes)
                self.xlmgr.FormatCells(lnSheet, lnCurRow, lnCurRow + 1, 0, lnColCount, lnHdrFmat)
                lnCurRow += 1
                if hiddenNames:
                    lcFldNames = lcFldNames.upper()
                    self.xlmgr.WriteRowValues(lnSheet, lnCurRow, 0, lnColCount - 1, lcFldNames, lcHdrTypes)
                    self.xlmgr.SetRowHidden(lnSheet, lnCurRow, True)
                    lnCurRow += 1
                if not rawcopy:  # Apply type checking to eliminate fields that we can't convert.
                    self.dbfmgr.select(lcLoadAlias)
                    self.dbfmgr.goto("TOP")
                    for jj in range(0, lnRecs):
                        if (not self.dbfmgr.eof()) and (not self.dbfmgr.deleted()):
                            valDict = self.dbfmgr.scatter(alias=lcLoadAlias, converttypes=False, stripblanks=True)

                            if valDict is None:
                                lbReturn = False
                                self.cErrorMessage = self.dbfmgr.cErrorMessage
                                self.xlmgr.CloseWorkbook(0)
                                # Serious crash so close it all.
                                # self.dbfmgr.closedatabases()
                                self.dbfmgr.closetable(lcLoadAlias)
                                return lbReturn

                            xRow = list()
                            for lv in lfList:
                                xRow.append(valDict[lv].strip("\x00 "))
                            lcRow = self.cCellDelim.join(xRow)

                            self.xlmgr.WriteRowValues(lnSheet, lnCurRow, 0, lnColCount, lcRow, lcTypes)
                            if self.bFormatDecimals:
                                for nn in range(0, lnColCount):
                                    if not (ldList[nn] is None):
                                        self.xlmgr.FormatCells(lnSheet, lnCurRow, lnCurRow + 1, nn, nn + 1, ldList[nn])
                            lnCurRow += 1
                        self.dbfmgr.goto("NEXT")
                        valDict = None
                        if self.dbfmgr.eof():
                            break

                else:  # Copy the field data directly into the spreadsheet for maximum speed
                    for jj in range(0, lnRecs):
                        if not self.dbfmgr.deleted():
                            lcRow = self.dbfmgr.scatterraw(converttypes=False, stripblanks=True,
                                                           lcDelimiter=self.cCellDelim)
                            if lcRow is not None:
                                self.xlmgr.WriteRowValues(lnSheet, lnCurRow, 0, lnColCount, lcRow, lcTypes)
                                lnCurRow += 1
                        self.dbfmgr.goto("NEXT")
                        lcRow = None
            if not (self.cHiddenCellValue is None):
                self.xlmgr.WriteCellValue(lnSheet, self.nHiddenCellRow, self.nHiddenCellCol, self.cHiddenCellValue, "S")
                self.xlmgr.FormatCells(lnSheet, self.nHiddenCellRow, self.nHiddenCellRow + 1, self.nHiddenCellCol, self.nHiddenCellCol + 1, lnHiddenFmat)
            self.xlmgr.CloseWorkbook(1)
            self.dbfmgr.closetable(lcLoadAlias)
        else:
            lbReturn = False
            cErrorMessage = self.dbfmgr.cErrorMessage
            self.cErrorMessage = cErrorMessage
            self.xlmgr.CloseWorkbook(0)
            # Serious crash so close it all.
            self.dbfmgr.closetable(lcLoadAlias)
        return lbReturn

    def setdbfstru(self, lcFieldVals, lcFieldTypes, lxDbfTypes, lxDbfWidths, lxDbfDecimals):
        lxVals = lcFieldVals.split(self.cCellDelim)
        lcVal = ""
        lnCnt = len(lcFieldTypes)
        for jj in range(0, lnCnt):
            lcXType = lcFieldTypes[jj]
            if (lcXType == "X") or (lcXType == "E") or (lcXType == "B"):
                continue  # can't get any information from this cell.

            if lcXType == "S":  # string
                lnTestWidth = len(lxVals[jj])
                lcVal = lxVals[jj]
                if (lcVal == "TRUE") or (lcVal == "FALSE"):
                    if lxDbfTypes[jj] == "X":
                        lxDbfTypes[jj] = "L"
                        lxDbfWidths[jj] = 1
                else:
                    lxDbfTypes[jj] = "C"
                    if lnTestWidth == 0:
                        lnTestWidth = 1
                    if lnTestWidth > lxDbfWidths[jj]:
                        lxDbfWidths[jj] = lnTestWidth
                if lxDbfWidths[jj] > 254:
                    lxDbfWidths[jj] = 10
                    lxDbfTypes[jj] = "M"  # Memo for big strings.

            elif lcXType == "N":  # Number
                lcVal = lxVals[jj]
                lxParts = lcVal.split('.')
                lcBase = lxParts[0]
                lcDec = lxParts[1]
                lcDec = lcDec.rstrip('0')  # Get rid of useless trailing zeros.
                if len(lcDec) > lxDbfDecimals[jj]:
                    lxDbfDecimals[jj] = len(lcDec)
                if (lxDbfTypes[jj] == "X") and (lxDbfDecimals[jj] == 0):
                    lxDbfTypes[jj] = "I"
                    lxDbfWidths[jj] = 4
                elif (lxDbfTypes[jj] == "X") and (lxDbfDecimals[jj] > 0):
                    lxDbfTypes[jj] = "N"
                elif (lxDbfTypes[jj] == "I") and (lxDbfDecimals[jj] > 0):
                    lxDbfTypes[jj] = "N"
                elif lxDbfTypes[jj] in "CLDT":  # already something else so we make it 'C'
                    lxDbfTypes[jj] = 'C'
                    lxDbfDecimals[jj] = 0
                    lxDbfWidths[jj] = 20  # will hold something or another.
                    break
                lnTestWidth = len(lcDec) + lxDbfDecimals[jj] + 1

                if (lxDbfTypes[jj] == "N") and (lnTestWidth > lxDbfWidths[jj]):
                    lxDbfWidths[jj] = lnTestWidth
                    if lxDbfWidths[jj] < 10:
                        lxDbfWidths[jj] = 10  # nothing narrower than this for a number by default.
            elif lcXType == "L":
                if lxDbfTypes[jj] == "X":
                    lxDbfTypes[jj] = "L"
                    lxDbfWidths[jj] = 1
                    lxDbfDecimals[jj] = 0
                elif lxDbfTypes[jj] in "CNTDI":
                    lxDbfTypes[jj] = "C"
                    lxDbfWidths[jj] = 20
                    lxDbfDecimals[jj] = 0
            elif lcXType == "D":
                if lcVal == "00000000000000000":
                    # this is an empty datetime, so just break
                    break
                lcVal = lxVals[jj]
                if lcVal[0:8] != "00000000":
                    lxDbfTypes[jj] = "D"
                    lxDbfWidths[jj] = 8
                    lxDbfDecimals[jj] = 0

                if lcVal[8:99] != "000000000":
                    # this is a datetime cause it has a time component.
                    if lxDbfTypes[jj] == "D":
                        lxDbfTypes[jj] = "T"  # Change the type to a datetime
                    elif lxDbfTypes[jj] == "X":
                        lxDbfTypes[jj] = "C"  # Just a time, which we'll handle as an 8 character string
                        lxDbfWidths[jj] = 8
                        lxDbfDecimals[jj] = 0

            else:
                xyx = True  # nothing.

    def excel2dbf(self, lcExcelName, lcDbfName, sheetname="", appendflag=False, templatedbf="", fieldlist=None,
                  columnsonly=False, showErrors=False):
        """
        Writes the contents of the specified Excel file (xls or xlsx) into the target DBF type table.  The existence of the
        lcDbfName table and the appendflag value determine the action taken:

        lcDbfName exists, appendflag = False:
            - lcDbfName records are removed (zapped), all records are replaced with newly imported records.  Any value in templatedbf
              is ingnored.
        lcDbfName exists, appendflag = True:
            - Records from the spreadsheet are appended into the table after the existing records.  Any value in templatedbf
              is ignored.
        lcDbfName doesn't exist -- appendflag has no meaning:
            If templatedbf contains the name of a valid DBF table, then uses the structure of that table to create a new table
            and imports data into it.
            If templatedbf == "" or None, then inspects the contents of the spreadsheet and creates a table the best way it can
            from the cell information.

        Takes the data from the first worksheet in the workbook.  You can pass a worksheet name to use instead in the
        sheetname optional parameter.  Note that in Excel, worksheet names are case sensitive.

        If the templatedbf parameter contains the name of an existing DBF type table, then appendflag is ignored and
        a new table is created (overwriting any existing one by that name) using the identical structure of the template
        DBF table.

        You can supply a list() of field names to include in the output DBF table.  In that case, the fields in the output
        table consist of those that are BOTH in the spreadsheet AND in the fieldlist.  If appendflag == True, the fieldlist
        will be respected -- records will be appended to the lcDbfName table, but only the fields specified will be populated.

        The first row of the spreadsheet must contain text values that are valid field names.  No more than 255 are permitted.
        Field names may be upper or lower case but MUST start with a letter or underscore.  They may be no longer than 10 characters
        and may NOT include blanks.  No two field names may be the same.  Any violation of these rules will result in an error flag
        being returned.

        Returns the number of records stored or appended.  Returns -1 on error, however if 0 is returned, you likely should investigate
        the source of the problem.

        NOTE: The CodeBase utilities are fairly slow in creating new DBF tables.  Not sure why.  At least the first one they create
        seems to take around 0.45 seconds on a very fast machine.  The second and third are faster, however.  So... for simple things
        just let this method write into an existing table, for much better performance.
        """

        lbDBFexists = False
        lbSetDefaults = False
        cDelimString = "/a!~!~/a"  # Added 08/07/2013 so as to handle case where ~~ may be contained in a field,
        # which is the CBTools
        # default delimiter. JSH.
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        lnResult = self.xlmgr.OpenWorkbook(lcExcelName)
        if lnResult < 1:
            self.cErrorMessage = "Can't open Excel File: %s because - "%(lcExcelName,) + self.xlmgr.cErrorMessage
            self.nErrorNumber = self.xlmgr.nErrorNumber
            if showErrors:
                print(self.cErrorMessage)
            return -1

        if sheetname != "":
            lnSheet = self.xlmgr.GetSheetFromName(sheetname)
        else:
            lnSheet = 0

        if lnSheet == -1:
            self.cErrorMessage = "No Active Sheet found"
            self.xlmgr.CloseWorkbook(0)
            if showErrors:
                print(self.cErrorMessage)
            return -1

        lcSheetName, lnFromRow, lnToRow, lnFromCol, lnToCol = self.xlmgr.GetSheetStats(lnSheet)
        xHdrVals, xHdrTypes = self.xlmgr.GetRowValues(lnSheet, 0, lnFromCol, lnToCol)

        lbBadType = False
        lnColCnt = 0
        for jj in range(0, (lnToCol + 1)):
            if xHdrTypes[jj] != "S":
                break
            else:
                if xHdrVals[jj] == "":
                    break
            lnColCnt += 1
        if lnColCnt < 1:
            lbBadType = True
        else:
            lnToCol = lnColCnt - 1  # A zero based column numbering system.

        if lnColCnt > 254:
            self.cErrorMessage = "Too many fields for a DBF table"
            self.xlmgr.CloseWorkbook(0)
            if showErrors:
                print(self.cErrorMessage)
            return -1

        if lbBadType:
            self.cErrorMessage = "Bad Column Cell Types: " + xHdrTypes
            self.xlmgr.CloseWorkbook(0)
            if showErrors:
                print(self.cErrorMessage)
            return -1

        xHdrTypes = xHdrTypes[0:lnColCnt]

        lxFieldNames = xHdrVals.split(self.cCellDelim)  # This exotic string is the standard delimiter
        # for the xls tools module.
        lnLimit = lnColCnt
        if lnLimit > len(lxFieldNames):
            lnLimit = len(lxFieldNames)
        lxFieldNames = lxFieldNames[0: lnColCnt]

        if columnsonly:
            self.xlmgr.CloseWorkbook(0)
            return lxFieldNames

        if not (self.xDefaultValues is None):
            for xFld in lxFieldNames:
                cTest = xFld.upper()
                if not self.bForceDefault:
                    if cTest in self.xDefaultValues:
                        # No need for this in the default list...
                        del self.xDefaultValues[cTest]
            if len(self.xDefaultValues) > 0:
                lbSetDefaults = True

        lbMakeNewTable = True
        lbUseTemplate = False
        lcDBFtemplate = ""
        lcDbfName = addExt(lcDbfName)
        if os.path.exists(lcDbfName):
            lcDBFtemplate = lcDbfName
            lbDBFexists = True
            lbMakeNewTable = False
        if (templatedbf != "") and (templatedbf is not None) and (not lbDBFexists):
            if os.path.exists(templatedbf):
                lcDBFtemplate = templatedbf
                lbUseTemplate = True
                lbMakeNewTable = True
        laFlds = None
        if lbUseTemplate:
            lbTest = self.dbfmgr.use(lcDBFtemplate)
            if lbTest:
                laFlds = self.dbfmgr.afields()
                lcAlias = self.dbfmgr.alias()
                self.dbfmgr.closetable(lcAlias)
            else:
                self.cErrorMessage = "Could not open template table: " + lcDBFtemplate
                self.xlmgr.CloseWorkbook(0)
                if showErrors:
                    print(self.cErrorMessage)
                return -1

        if (not lbUseTemplate) and lbMakeNewTable:
            #  we have to figure out the structure from the data in the spreadsheet
            lxTypeList = ["X" for jj in range(0, (lnToCol + 1))]
            lxWidthList = [0 for jj in range(0, (lnToCol + 1))]
            lxDecimalList = [0 for jj in range(0, (lnToCol + 1))]
            laFlds = list()  # This will be our list of VFPFIELD() objects
            for jj in range(1, (lnToRow + 1)):
                xRow1, xTypes1 = self.xlmgr.GetRowValues(lnSheet, jj, lnFromCol, lnToCol)
                self.setdbfstru(xRow1, xTypes1, lxTypeList, lxWidthList, lxDecimalList)
            for jj in range(0, (lnToCol + 1)):
                if lxTypeList[jj] == "X":
                    lxTypeList[jj] = "C"
                    lxWidthList[jj] = 10
                    lxDecimalList[jj] = 0  # This is going to be an empty field but fill it with a Char(10)
                if lxTypeList[jj] == "C":
                    if lxWidthList[jj] == 0:
                        lxWidthList[jj] = 10  # Nothing less for empty field.

                lxFld = VFPFIELD()
                lxFld.cName = lxFieldNames[jj].upper()
                lxFld.cType = lxTypeList[jj]
                lxFld.nWidth = lxWidthList[jj]
                lxFld.nDecimals = lxDecimalList[jj]
                lxFld.bNulls = 0
                laFlds.append(lxFld)
                lxFld = None
        # Now we have a list of VFPFIELD() objects with which we can construct our new table.
        # We'll construct a new temp table
        # even if we are doing an append so as not to mess up the target table if we have a problem.

        # Set the field list to upper case.
        if fieldlist is not None:
            for ix, lx in enumerate(fieldlist):
                fieldlist[ix] = lx.upper()

        # Now determine the name of the table and memo file we'll be outputting.
        lcWorkTable = lcDbfName.lower()

        lcBaseTable, lcExtTable = os.path.splitext(lcWorkTable)
        if (lcExtTable == "") or (lcExtTable == "."):
            lcExtTable = ".dbf"
        lcExtMemo = ".fpt"
        lcWorkMemo = lcBaseTable + lcExtMemo
        lcWorkTable = lcBaseTable + lcExtTable
        laFlds2 = None
        if lbMakeNewTable:
            laFlds2 = list()
            ldFldPtrs = dict()
            # This will be a dictionary keyed by field name indicating the columns where each appears in the XLS
            if fieldlist is not None:
                for ix, lx in enumerate(laFlds):
                    if fieldlist.count(lx.cName.upper()) != 0:
                        laFlds2.append(lx)
                        ldFldPtrs[lx.cName.upper()] = ix
                laFlds = None
            else:
                laFlds2 = laFlds  # OK to have a simple copy of the object pointer
                laFlds = None
            if len(laFlds2) == 0:
                self.cErrorMessage = "No fields matching fieldlist found in spreadsheet"
                self.xlmgr.CloseWorkbook(0)
                if showErrors:
                    print(self.cErrorMessage)
                return -1
        lcWorkAlias = ""
        lxNameTest = None
        if lbMakeNewTable:
            lbResult = self.dbfmgr.createtable(lcWorkTable, laFlds2)
            if lbResult:
                lcWorkAlias = self.dbfmgr.alias()
            lxNameTest = dict()
            for lx in laFlds2:
                lxNameTest[lx.cName.upper()] = 1
        else:  # Just open the table
            lbResult = self.dbfmgr.use(lcWorkTable)
            if lbResult:
                lcWorkAlias = self.dbfmgr.alias()
                if not appendflag:
                    lbResult = self.dbfmgr.zap()
                laFlds = self.dbfmgr.afields()
                lxNameTest = dict()

            if laFlds is not None:
                for lx in laFlds:
                    lxNameTest[lx.cName.upper()] = 1
            else:
                lbResult = False

        if not lbResult:
            self.xlmgr.CloseWorkbook(0)
            self.dbfmgr.closetable(lcWorkAlias)
            if os.path.exists(lcWorkMemo):
                os.unlink(lcWorkMemo)
            if os.path.exists(lcWorkTable):
                os.unlink(lcWorkTable)
            self.cErrorMessage = self.dbfmgr.cErrorMessage
            if showErrors:
                print(self.cErrorMessage)
            return -1

        loDB = self.dbfmgr
        lcAlias = loDB.alias()
        #  Now to spin through the spreadsheet, grab the values and shove into the output table
        aFldList = loDB.afieldtypes()
        cEmptyTest = "E" * lnColCnt
        for jj in range(1, lnToRow):

            xRow1, xTypes1 = self.xlmgr.GetRowValues(lnSheet, jj, lnFromCol, lnToCol)
            if xTypes1 == cEmptyTest:
                # we have a case of a completely empty row.  We need to stop NOW!
                break
            if xRow1 is None:
                self.cErrorMessage = self.xlmgr.GetErrorMessage() + " Row Number %d" % (jj + 1)
                if showErrors:
                    print(self.cErrorMessage)
                lxDict = None
                return -1  # Bad data, best to stop and let user know.

            lbResult = loDB.appendblank()
            lxDict = dict()
            if fieldlist is None:
                lxVals = xRow1.split(self.cCellDelim)
                for ix, lx in enumerate(lxVals):
                    cCellType = xTypes1[ix]
                    cFldName = lxFieldNames[ix].upper()
                    cFldType = aFldList.get(cFldName, "X")
                    if cFldName not in lxNameTest:
                        continue

                    if cFldType != "M":
                        cFldValue = lx.strip()
                        if cFldType == "D":
                            cFldValue = cFldValue[0:8]
                        elif cFldValue == ".NULL." and cCellType == "S":
                            cFldValue = None
                    else:
                        cFldValue = lx

                    lxDict[cFldName] = cFldValue
            else:
                lcVals = ""
                lxVals = xRow1.split(self.cCellDelim)
                for lx in laFlds2:
                    cFldName = lx.cName.upper()
                    nPtr = ldFldPtrs[cFldName]
                    cFldType = aFldList.get(cFldName, "X")
                    if cFldName not in lxNameTest:
                        continue
                    cFldValue = lxVals[nPtr]
                    cCellType = xTypes1[nPtr]

                    if cFldType != "M":
                        cFldValue = cFldValue.strip()
                        if cFldType == "D":
                            cFldValue = cFldValue[0:8]
                        elif cCellType == "S" and cFldValue == ".NULL.":
                            cFldValue = None  # This is where we stored .NULL. as the value in the spreadsheet.
                    lxDict[cFldName] = cFldValue

            lbResult = loDB.gatherdict(cAlias=lcAlias, dData=lxDict)

            if lbSetDefaults:
                for xFld, xDef in self.xDefaultValues.items():
                    if xFld in aFldList:  # Added this to prevent a crash if the field is not found in the table.
                        #  05/14/2013 JSH
                        xTestVal = loDB.curval(xFld, convertType=True)
                        if not xTestVal:
                            loDB.replace(xFld, str(xDef))
            if not lbResult:
                self.cErrorMessage = loDB.cErrorMessage + " Row Number %d" % (jj + 1)
                if showErrors:
                    print(self.cErrorMessage)
                lxDict = None
                return -1  # Bad data, best to stop and let user know.
            lxDict = None
        self.dbfmgr.closetable(lcAlias)
        return 1

    def __exit__(self, lcClass, lcType, lcTrace):
        # self.dbfmgr.closedatabases()
        # Not necessary. Close tables when we're done with them.
        # self.dbfmgr.cb_shutdown() ## DO NOT DO THIS... MUCH TOO SLOW
        self.xlmgr.CloseWorkbook(0)

        if lcClass is not None:
            # Do something with the exception
            return False  # Forces the error to propagate and be trapped elsewhere.

    def __del__(self):
        # self.dbfmgr.cb_shutdown() ## EXTREMELY SLOW... Doesn't seem to hurt not to call this
        self.dbfmgr = None
        self.xlmgr = None


class dbfxlTools(mTools.baseService):
    """
    This is a COM wrapper around the DbfXlsEngine object above.  It is a bit of a hybrid, as it can be called stand-alone
    as a COM object, OR it can be launched by the PlanToolsPythonsService manager as a COM service object.

    The conversion tools in this module are somewhat simpler than for the main DbfXlsEngine class, but all the
    same functionality is available by setting properties of the class after creating the instance.  In Python
    this can be done directly.  In VFP via COM, it can be done by calling the set_property() method.

    Note that this class is no longer intended to be a stand alone COM object but is to be created and delivered
    to COM clients by an object factory dispatcher.  To make this a COM object, see the documentation in the
    book Python Programming on Win32 for details.
    """
    _public_methods_ = [ 'dbftoexcel', 'exceltodbf', 'getexcelcolumns', 'gettestdictionary', 'gettestcompound',
                         'gettestlist', 'setfielddefault', 'set_property', 'get_property', 'setwidthlist',
                         'setheaderlist', 'setoneproperty']
    _public_attrs_ = ['cSheetName', 'cWidthList', 'cErrorMessage', 'cTempFilePath', 'nErrorNumber', 'name',
                      'cHeaderList', 'cTitleOne', 'cTitleTwo', 'cFieldList']

    def __init__(self, oVFP=None, oCfg=None, bIsCOM=False):  # Parms fit signature of python services.
        bComTest = bIsCOM
        # super(dbfxlTools, self).__init__(bIsCOM=bComTest)
        self.cSheetName = ""
        self.bAppendFlag = False
        self.cTemplateDbf = ""
        self.lFieldList = None  # For Excel to DBF
        self.cFieldList = ""  # For DBF to Excel
        self.lAppendSheet = False  # If TRUE then if the file name passed exists, the new data will be added
        # as a new sheet in the sheets collection, leaving any existing sheets intact.
        self.cErrorMessage = ""
        self.nErrorNumber = 0
        self.cHeaderList = ""
        self.cTitleOne = ""
        self.cTitleTwo = ""
        self.name = "dbfxlTools COM Object"
        self.cWidthList = ""  # A tab delimited list of width numbers preceded by "PIXEL:" or "EXCEL:" which
        # indicates the metric associated with the width values.  Must be set by the caller prior to calling
        # dbftoexcel().
        self.xDefaultValues = dict()
        self.bForceDefault = False  # Passed to the underlying engine.
        self.cHiddenCellValue = None
        self.nHiddenCellRow = None
        self.nHiddenCellCol = None
        self.oVFP = (oVFP or None)
        self.oCfg = (oCfg or None)
        if self.oCfg:
            self.cErrorFilePath = self.oCfg.cErrorFilePath
        self.cTempFilePath = ""
        mTools.baseService.__init__(self, bIsCOM=bComTest)
        return

    def setoneproperty(self, cTheProp="", xTheValue=None):
        """
        Method used by a COM client to set the property value required to be changed.  Note that this value
        IS case sensitive
        :param cTheProp: String containing the name of the property to set
        :param xTheValue: Value to set it to
        :return: Nothing
        """
        if cTheProp:
            mTools.baseService.__setattr__(self, cTheProp, xTheValue)
        return

    def setheaderlist(self, cTheList=""):
        """
        To override the standard column header values (the field names from the DBF table), pass a comma separated
        list of column captions. The number of items in the list should equal the number of fields to be output.
        :param cTheList: The list string
        :return: True, always.
        """
        self.cHeaderList = cTheList
        return True

    def setwidthlist(self, cTheList=""):
        """
        The engine will attempt to set Excel column widths appropriate to the content of each field.  To override this
        behavior pass a comma separated list of number values specifying the widths of all the columns.  These values
        must be in Excel width units which are approximately 8.4 pixels each.  Widths may be specified by floating
        point decimal values like "5.5".
        :param cTheList: Comma separated list of width numbers.
        :return: True
        """
        self.cWidthList = cTheList
        return True

    def setfielddefault(self, lpcFieldName, lpcDefaultValue):
        """
        This function allows you to set default values for data table fields which may be empty or missing in the
        source XLS file, but for which there must be a value in the DBF table.
        :param lpcFieldName: Field name for the default value
        :param lpcDefaultValue: Value to supply if the XLS(X) source doesn't have a value.
        :return: 1 always
        """
        self.xDefaultValues[lpcFieldName.upper()] = lpcDefaultValue
        return 1

    def unpackWidths(self):
        """
        Internal Use
        :return: tuple of width information
        """
        loList = None
        lbPixels = False
        xParts = self.cWidthList.split(":")
        if len(xParts) != 2:
            return (loList, lbPixels)  # Nothing to do, they didn't provide the pixel indicator
        lbPixels = (True if (xParts[0].upper() == "PIXEL") else False)
        loList = xParts[1].split("\t")
        for lx in loList:
            try:
                lx = float(lx)
            except:
                lx = 0.0

        return (loList, lbPixels)

    def dbftoexcel(self, lcDbf, lcExcel, rawcopy=False, addHiddenNames=False):
        """
        Converts the DBF table specified into an Excel spreadsheet specified.  Both must be provided as fully
        qualified path names.  The file name extension for the Excel sheet tells the engine whether to produce
        an XLS ro an XLSX output.  XLS outputs are faster, but the result files are considerably larger than the
        XLSX version.  Also XLS sheets are limited to about 64,000 rows, while XLSX sheets are limited to a bit over
        1 million rows.

        The sheet name applied will be the alias name of the DBF table.  If you want a different sheet name, you'l
        need to set the value of the class property cSheetName prior to running this method.  By default, this
        process will create a new spreadsheet file, overwriting any existing file by that name.  If you set the
        property lAppendSheet to True, then if the Excel file exists, the engine will attempt to add the specified
        sheet to the end of the sheets collection in the existing file.

        :param lcDbf: Fully qualified path name of the DBF table.  Must be able to be opened in shared mode.
        :param lcExcel: Fully qualified path name of the Excel spreadsheet
        :param rawcopy: Set to True (default is False) to gain speed but give up some output error checking and data
        format control.
        This is relatively safe if all the data is standard string or simple numeric values.
        :param addHiddenNames: If True (default is False), then a hidden row will be added right after the column
        header row which will contain the actual field names in each column.  May be needed for loading an Excel
        sheet back into a DBF table when the column headings must contain human readable, meaningful captions.
        :return: Value >= 1 on success, 0 on failure.
        """
        if not _ver3x:
            if mTools.isunicode(lcDbf):
                lcDbf = lcDbf.encode('cp1252', 'ignore')
            if mTools.isunicode(lcExcel):
                lcExcel = lcExcel.encode('cp1252', 'ignore')
        else:
            pass
            # No else.  In the case of 3x, these will be passed by COM as unicode which is the native 3x string format.

        global cErrorMessage
        lxHdr = None
        lnReturn = 1
        if self.cHeaderList != "":
            # May be set by the caller.  Has to be tab delimited list.
            lxHdr = self.cHeaderList.split("\t")
        if self.cWidthList != "":
            loWidths, lbPixels = self.unpackWidths()
        else:
            loWidths = None
            lbPixels = False
        cTestHdr = str(lxHdr)

        xRaw = rawcopy
        xHiddenNames = addHiddenNames

        with DbfXlsEngine(self.oVFP) as dbcvrt:
            dbcvrt.xWidthList = loWidths
            dbcvrt.bWidthPixels = lbPixels
            if self.cSheetName:
                dbcvrt.cSheetName = self.cSheetName
            dbcvrt.lAppendSheet = self.lAppendSheet
            if not (self.cHiddenCellValue is None):
                dbcvrt.cHiddenCellValue = self.cHiddenCellValue
                dbcvrt.nHiddenCellRow = self.nHiddenCellRow - 1
                # We assume these col and row values are Excel 1-based values.
                dbcvrt.nHiddenCellCol = self.nHiddenCellCol - 1
            lbReturn = dbcvrt.dbf2excel(lcDbf, lcExcel, headerlist=lxHdr, titleone=self.cTitleOne,
                                        titletwo=self.cTitleTwo, rawcopy=xRaw, hiddenNames=xHiddenNames,
                                        fieldList=self.cFieldList)
            if not lbReturn:
                self.cErrorMessage = dbcvrt.cErrorMessage
                lnReturn = 0
        return lnReturn

    def exceltodbf(self, lcExcel, lcDbf, lcSheetName):
        """
        Writes the contents of the specified Excel spreadsheet into the specified table.  By default, the contents
        of the DBF file are deleted (Zapped) and replaced with the contents of the Excel spreadsheet.  If the
        lcSheetName is not provided, then the first sheet in the XLS(X) file will be used.  Each column heading
        must contain the name of a field found in the table.  Column headings with unrecognized field names will
        be ignored.

        If the 2nd row of the sheet is hidden, then an attempt will be made to read the field names from that row.

        :param lcExcel: Fully qualified path name to the spreadsheet file
        :param lcDbf: Fully qualified path name to the DBF table.
        :param lcSheetName: Optional sheet name from which to read the data.
        :return: Positive integer on success, negative integer on error.
        """
        if not _ver3x:
            if mTools.isunicode(lcDbf):
                lcDbf = lcDbf.encode('cp1252', 'ignore')
            if mTools.isunicode(lcExcel):
                lcExcel = lcExcel.encode('cp1252', 'ignore')
        else:
            pass
            # No else.  Python 3.x handles unicode from COM natively.

        global cErrorMessage
        if mTools.isstr(self.lFieldList):
            if not ("~~" in self.lFieldList):
                self.cErrorMessage = "Bad Field List, no ~~ delimiters"
                return -1
            lxFieldList = self.lFieldList.split("~~")
        else:
            lxFieldList = None
        with DbfXlsEngine(self.oVFP) as dbcvrt:
            if len(self.xDefaultValues) > 0:
                dbcvrt.xDefaultValues = self.xDefaultValues
            dbcvrt.bForceDefault = self.bForceDefault
            lnReturn = dbcvrt.excel2dbf(lcExcel, lcDbf, sheetname=lcSheetName, appendflag=self.bAppendFlag,
                                        templatedbf=self.cTemplateDbf, fieldlist=lxFieldList)
            if lnReturn < 0:
                self.cErrorMessage = dbcvrt.cErrorMessage
                self.nErrorNumber = dbcvrt.nErrorNumber

        return lnReturn

    def geterrormessage(self):
        """
        If this object is used by a Python module, all error information is available.  However COM clients will
        need to use this method to determine what when wrong when an error condition is reported by a method.
        :return: Text of the error message.
        """
        return self.cErrorMessage

    def getexcelcolumns(self, lcExcel, lcSheetName):
        """
        This method opens a spreadsheet and reads the first row of cells to get the column headers, which should
        normally be the field names defined for each column.  No data is read.
        :param lcExcel: Fully qualified path name of the XLS(X) file.
        :param lcSheetName: Name of the sheet to read if not the first one.
        :return: List of field names on success, and empty string on failure.
        """
        global cErrorMessage
        lcDbf = "sometable.dbf"
        with DbfXlsEngine(self.oVFP) as dbcvrt:
            lxReturn = dbcvrt.excel2dbf(lcExcel, lcDbf, sheetname=lcSheetName, columnsonly=True)
            if not isinstance(lxReturn, list):
                self.cErrorMessage = dbcvrt.cErrorMessage
                return ""
        return lxReturn


__all__ = ["DbfXlsEngine", "dbfxlTools"]

classToWrap = dbfxlTools


def dtTest():
    global cErrorMessage
    import random
    oDB = cbTools()
    random.seed()

    xFirstNames = ["Robert", "James", "Terry", "Janice", "Mary", "Elizabeth", "Bobbie"]
    xLastNames = ["Jones", "Wang", "Smith", "Carsky", "Johnson", "Goldman", "Thompson"]
    xCities = [("Rockford", "IL"), ("Minneapolis", "MN"), ("Portland", "OR"), ("Dallas", "TX")]

    xFlds = list()
    xFlds.append(oDB.newfield("FIRSTNAME", "C", 20, 0, False))
    xFlds.append(oDB.newfield("LASTNAME", "C", 20, 0, False))
    xFlds.append(oDB.newfield("CITY", "C", 15, 0, False))
    xFlds.append(oDB.newfield("STATE", "C", 2, 0, False))
    xFlds.append(oDB.newfield("BIRTHDATE", "D", 0, 0, False))
    xFlds.append(oDB.newfield("ACCOUNTNUM", "I", 4, 0, False))
    xFlds.append(oDB.newfield("BALANCE", "N", 8, 2, False))

    bDirOK = os.path.exists(r"c:\temp")
    if not bDirOK:
        print("SORRY, you need a C:\\TEMP directory to hold the outputs of this test.")
        return False

    cTestNameTable = r"c:\temp\customers.dbf"
    bTest = oDB.createtable(cTestNameTable, xFlds)
    if bTest:
        cAlias = oDB.alias()
        oDB.closetable(cAlias)
    else:
        print("Table Create Failed! " + oDB.cErrorMessage)
        return False

    bTest = oDB.use(cTestNameTable, alias="CUSTOMERS")
    if not bTest:
        print("Table re-open failed! " + oDB.cErrorMessage)
        return False

    for jj in range(0, 500):
        xData = dict()
        xData["FIRSTNAME"] = random.choice(xFirstNames)
        xData["LASTNAME"] = random.choice(xLastNames)
        xLoc = random.choice(xCities)
        xData["CITY"] = xLoc[0]
        xData["STATE"] = xLoc[1]
        dBase = date(1948, 1, 1)
        nDaysAge = random.randint(3000, 10000)
        xData["BIRTHDATE"] = dBase + datetime.timedelta(days=nDaysAge)
        xData["ACCOUNTNUM"] = random.randint(10000, 20000)
        xData["BALANCE"] = random.random() * 20000.0
        bTest = oDB.insertdict(xData, "CUSTOMERS")
        if not bTest:
            print("Data Insert Failed:", oDB.cErrorMessage)
            return False

    print("Records Created:", oDB.reccount())
    oDB.closetable("CUSTOMERS")

    xTools = dbfxlTools(oVFP=oDB)
    print(xTools.name)
    cExcelOutput = r"c:\temp\customers1.xlsx"
    nRet = xTools.dbftoexcel(cTestNameTable, cExcelOutput, rawcopy=False)
    print(nRet)

    if nRet > -1:
        cTargetNameTable = r"c:\temp\newcustomers.dbf"
        bTest = oDB.createtable(cTargetNameTable, xFlds)
        if not bTest:
            print("Unable to create test output DBF:", oDB.cErrorMessage)
            return False
        else:
            cAlias = oDB.alias()
            oDB.closetable(cAlias)
    else:
        print("Excel Output Failed: " + xTools.cErrorMessage)
        return False

    nRet = xTools.exceltodbf(cExcelOutput, cTargetNameTable, "")
    if nRet < 0:
        print("Copy to DBF table failed: ", xTools.cErrorMessage)
        return False

    bTest = oDB.use(cTargetNameTable, alias="NEWTABLE")
    print("Records copied to DBF:", oDB.reccount())
    oDB.closetable("NEWTABLE")
    return True


# def mongoTest():
#     print("Running mongoTest()")
#     global cErrorMessage
#     from LoadOptWeb.FreightDataManager.TLRateTable import TLRateTableBizObj
#     from MPSSCommon.LOConfigurator import LOConfig
#     from time import time
#     oCFG = LOConfig()
#     bTest = oCFG.init("live")
#     nStart = time()
#     if bTest:
#         with DbfXlsEngine() as oEng:
#             oBiz = TLRateTableBizObj("MASTER")
#             bTest = oBiz.initialize(oCFG=oCFG, oVFP=oCFG.VFP, oParent=oEng)
#             xTest = oBiz.attachtable()
#             oBiz.nMaxQueryCount = 75000
#             print(xTest)
#             if xTest is None:
#                 print(oBiz.cErrorMessage)
#             else:
#                 bTest = oEng.bizobj2excel(oBiz, cExcelName="c:\\temp\\myrates.xlsx", cForExpr="rts_key==11027",
#                                           cSortOrder="from_geo")
#                 print("RESULT: ", bTest)
#     nEnd = time()
#     print("TOTAL TIME:", (nEnd - nStart))
#     return bTest


if __name__ == "__main__":
    print("***** Testing DBFXLStools2.py components")
    bResult = dtTest()
    print(bResult)

    # Only uncomment this if MongoDB is available and then only if the table biz objects are defined.
    # bResult = mongoTest()
    # print(bResult)


