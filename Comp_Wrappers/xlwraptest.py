from ctypes import *
import os
from time import time
from datetime import date, datetime
import itertools
# from imp import find_module
import sys
from MPSSCommon.MPSSBaseTools import findrepopath

gcRepoPath = findrepopath()
# Cell Color Styles
xlColor = {"COLOR_BLACK":8, "COLOR_WHITE":9, "COLOR_RED":10, "COLOR_BRIGHTGREEN":11, "COLOR_BLUE":12, "COLOR_YELLOW":13, "COLOR_PINK":14, "COLOR_TURQUOISE":15, "COLOR_DARKRED":16, \
           "COLOR_GREEN":17, "COLOR_DARKBLUE":18, "COLOR_DARKYELLOW":19, "COLOR_VIOLET":20, "COLOR_TEAL":21, "COLOR_GRAY25":22, "COLOR_GRAY50":23, "COLOR_PERIWINKLE_CF":24, \
           "COLOR_PLUM_CF":25, "COLOR_IVORY_CF":26, "COLOR_LIGHTTURQUOISE_CF":27, "COLOR_DARKPURPLE_CF":28, "COLOR_CORAL_CF":29, "COLOR_OCEANBLUE_CF":30, "COLOR_ICEBLUE_CF":31, \
           "COLOR_DARKBLUE_CL":32, "COLOR_PINK_CL":33, "COLOR_YELLOW_CL":34, "COLOR_TURQUOISE_CL":35, "COLOR_VIOLET_CL":36, "COLOR_DARKRED_CL":37, "COLOR_TEAL_CL":38, \
           "COLOR_BLUE_CL":39, "COLOR_SKYBLUE":40, "COLOR_LIGHTTURQUOISE":41, "COLOR_LIGHTGREEN":42, "COLOR_LIGHTYELLOW":43, "COLOR_PALEBLUE":44, "COLOR_ROSE":45, "COLOR_LAVENDER":46, \
           "COLOR_TAN":47, "COLOR_LIGHTBLUE":48, "COLOR_AQUA":49, "COLOR_LIME":50, "COLOR_GOLD":51, "COLOR_LIGHTORANGE":52, "COLOR_ORANGE":53, "COLOR_BLUEGRAY":54, "COLOR_GRAY40":55, \
           "COLOR_DARKTEAL":56, "COLOR_SEAGREEN":57, "COLOR_DARKGREEN":58, "COLOR_OLIVEGREEN":59, "COLOR_BROWN":60, "COLOR_PLUM":61, "COLOR_INDIGO":62, "COLOR_GRAY80":63, \
           "COLOR_DEFAULT_FOREGROUND": 0x0040, "COLOR_DEFAULT_BACKGROUND": 0x0041, "COLOR_TOOLTIP": 0x0051, "COLOR_AUTO": 0x7FFF}

xlNumFormat = {"NUMFORMAT_GENERAL":0, "NUMFORMAT_NUMBER":1, "NUMFORMAT_NUMBER_D2":2, "NUMFORMAT_NUMBER_SEP":3, "NUMFORMAT_NUMBER_SEP_D2":4,\
               "NUMFORMAT_CURRENCY_NEGBRA": 5, "NUMFORMAT_CURRENCY_NEGBRARED": 6, "NUMFORMAT_CURRENCY_D2_NEGBRA": 7, "NUMFORMAT_CURRENCY_D2_NEGBRARED": 8,\
               "NUMFORMAT_PERCENT":9, "NUMFORMAT_PERCENT_D2":10, "NUMFORMAT_SCIENTIFIC_D2":11, "NUMFORMAT_FRACTION_ONEDIG":12, "NUMFORMAT_FRACTION_TWODIG":13,\
               "NUMFORMAT_DATE":14, "NUMFORMAT_CUSTOM_D_MON_YY":15, "NUMFORMAT_CUSTOM_D_MON":16, "NUMFORMAT_CUSTOM_MON_YY":17,\
               "NUMFORMAT_CUSTOM_HMM_AM":18, "NUMFORMAT_CUSTOM_HMMSS_AM":19, "NUMFORMAT_CUSTOM_HMM":20, "NUMFORMAT_CUSTOM_HMMSS":21,\
               "NUMFORMAT_CUSTOM_MDYYYY_HMM":22,\
               "NUMFORMAT_NUMBER_SEP_NEGBRA":37, "NUMFORMAT_NUMBER_SEP_NEGBRARED":38,\
               "NUMFORMAT_NUMBER_D2_SEP_NEGBRA":39, "NUMFORMAT_NUMBER_D2_SEP_NEGBRARED":40, "NUMFORMAT_ACCOUNT":41, "NUMFORMAT_ACCOUNTCUR":42,\
               "NUMFORMAT_ACCOUNT_D2":43, "NUMFORMAT_ACCOUNT_D2_CUR":44, "NUMFORMAT_CUSTOM_MMSS":45, "NUMFORMAT_CUSTOM_H0MMSS":46,\
               "NUMFORMAT_CUSTOM_MMSS0":47, "NUMFORMAT_CUSTOM_000P0E_PLUS0":48, "NUMFORMAT_TEXT":49};

xlAlignH = {"ALIGNH_GENERAL":0, "ALIGNH_LEFT":1, "ALIGNH_CENTER":2, "ALIGNH_RIGHT":3, "ALIGNH_FILL":4, "ALIGNH_JUSTIFY":5, "ALIGNH_MERGE":6, "ALIGNH_DISTRIBUTED":7}

xlAlignV = {"ALIGNV_TOP":0, "ALIGNV_CENTER":1, "ALIGNV_BOTTOM":2, "ALIGNV_JUSTIFY":3, "ALIGNV_DISTRIBUTED":4}

xlBorder = {"BORDERSTYLE_NONE":0, "BORDERSTYLE_THIN":1, "BORDERSTYLE_MEDIUM":2, "BORDERSTYLE_DASHED":3, "BORDERSTYLE_DOTTED":4, "BORDERSTYLE_THICK":5,\
            "BORDERSTYLE_DOUBLE":6, "BORDERSTYLE_HAIR": 7}

xlFill = {"FILLPATTERN_NONE":0, "FILLPATTERN_SOLID":1, "FILLPATTERN_GRAY50":2, "FILLPATTERN_GRAY75":3, "FILLPATTERN_GRAY25":4 }

xlUnderline = {"UNDERLINE_NONE":0, "UNDERLINE_SINGLE":1, "UNDERLINE_DOUBLE":2 }


class XLFORMAT(Structure):
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
        # lc1, lc2, lc3 = find_module('xlwraptest')  # Find the full path name of this module,
        # so we know where the dll is.
        lc2 = __file__
        print("THE FILE:", lc2)
        lcDllPath, lcName = os.path.split(lc2)
        lcDllPath = lcDllPath + "\\LibXLWrapper.dll"

        self.xlWrapx = CDLL(lcDllPath)

        self.xlWrapx.xlwOpenWorkbook.argtypes = [c_char_p]
        self.xlWrapx.xlwOpenWorkbook.restype = c_int

        self.xlWrapx.xlwCloseWorkbook.argtypes = [c_int]
        self.xlWrapx.xlwCloseWorkbook.restype = c_int

        self.xlWrapx.xlwCreateWorkbook.argtypes = [c_char_p, c_char_p]
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

        self.xlWrapx.xlwIsRowHidden.argtypes = [c_int, c_int]
        self.xlWrapx.xlwIsRowHidden.restype = c_int

        self.xlWrapx.xlwIsColHidden.argtypes = [c_int, c_int]
        self.xlWrapx.xlwIsColHidden.restype = c_int

        self.xlWrapx.xlwSetRowHidden.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwSetRowHidden.restype = c_int

        self.xlWrapx.xlwSetColHidden.argtypes = [c_int, c_int, c_int]
        self.xlWrapx.xlwSetColHidden.restype = c_int

        self.xlWrapx.xlwSetDelimiter.argtypes = [c_char_p]
        self.xlWrapx.xlwSetDelimiter.restype = c_int

        lcTab = "\t"
        lcTab = bytes(lcTab, encoding="ascii")
        pDelim = c_char_p(lcTab)
        self.xlWrapx.xlwSetDelimiter(pDelim)  # Sets the delimiter as a tab unless set otherwise by the caller.

    def IsRowHidden(self, lnSheet, lnRow):
        return True if self.xlWrapx.xlwIsRowHidden(lnSheet, lnRow) == 1 else False

    def IsColHidden(self, lnSheet, lnCol):
        return True if self.xlWrapx.xlwIsColHidden(lnSheet, lnCol) == 1 else False

    def SetRowHidden(self, lnSheet, lnRow, lbHide):
        lnHow = 1 if lbHide else 0
        return self.xlWrapx.xlwSetRowHidden(lnSheet, lnRow, lnHow)

    def SetColHidden(self, lnSheet, lnCol, lbHide):
        lnHow = 1 if lbHide else 0
        return self.xlWrapx.xlwSetColHidden(lnSheet, lnCol, lnHow)

    def OpenWorkbook(self, lcFileName):
        cFileName = bytes(lcFileName, encoding="utf-8")
        pName = c_char_p(cFileName)
        return self.xlWrapx.xlwOpenWorkbook(pName)

    def CloseWorkbook(self, lbSave):
        return self.xlWrapx.xlwCloseWorkbook(lbSave)

    def CreateWorkbook(self, lcFileName, lcFirstSheetName):
        cFileName = bytes(lcFileName, encoding="ascii")
        pName = c_char_p(cFileName)
        cFirstSheetName = bytes(lcFirstSheetName, encoding="ascii")
        pSheet = c_char_p(cFirstSheetName)
        return self.xlWrapx.xlwCreateWorkbook(pName, pSheet)

    def AddSheet(self, lcSheetName):
        pName = c_char_p(lcSheetName)
        return self.xlWrapx.xlwAddSheet(pName)

    def GetCellValue(self, lnSheet, lnRow, lnCol):
        lcType = b"      "
        pType = c_char_p(lcType)
        lcStr = self.xlWrapx.xlwGetCellValue(lnSheet, lnRow, lnCol, pType)
        lxRet = (lcStr, pType.value)
        return lxRet

    def GetErrorMessage(self):
        return self.xlWrapx.xlwGetErrorMessage()

    def GetErrorNumber(self):
        return self.xlWrapx.xlwGetErrorNumber()

    def GetSheetStats(self, lnSheet):
        lnFirstCol = c_int()
        lnLastCol = c_int()
        lnFirstRow = c_int()
        lnLastRow = c_int()
        lpCol1 = pointer(lnFirstCol)
        lpColN = pointer(lnLastCol)
        lpRow1 = pointer(lnFirstRow)
        lpRowN = pointer(lnLastRow)
        lcName = self.xlWrapx.xlwGetSheetStats(lnSheet, lpCol1, lpColN, lpRow1, lpRowN)
        return (lcName, lpCol1.contents.value, lpColN.contents.value, lpRow1.contents.value, lpRowN.contents.value)

    def GetSheetFromName(self, lcName):
        cName = bytes(lcName, encoding="ascii")
        lpSH = c_char_p(cName)
        return self.xlWrapx.xlwGetSheetFromName(lpSH)

    def WriteCellValue(self, lnSheet, lnRow, lnCol, lcStrValue, lcCellType):
        cStrValue = bytes(lcStrValue, encoding="utf-8")
        lpVL = c_char_p(cStrValue)
        cCellType = bytes(lcCellType, encoding="ascii")
        lpTY = c_char_p(lcCellType)
        return self.xlWrapx.xlwWriteCellValue(lnSheet, lnRow, lnCol, lpVL, lpTY)

    def WriteRowValues(self, lnSheet, lnRow, lnColFrom, lnColTo, lcStrValues, lcCellTypes):
        cStrValues = bytes(lcStrValues, encoding="utf-8")
        cCellTypes = bytes(lcCellTypes, encoding="ascii")
        lpVL = c_char_p(cStrValues)
        lpTY = c_char_p(cCellTypes)
        return self.xlWrapx.xlwWriteRowValues(lnSheet, lnRow, lnColFrom, lnColTo, lpVL, lpTY)

    def GetRowValues(self, lnSheet, lnRow, lnColFrom, lnColTo):
        lcType = b" " * 300
        lpTy = c_char_p(lcType)
        lcVals = self.xlWrapx.xlwGetRowValues(lnSheet, lnRow, lnColFrom, lnColTo, lpTy)
        return (lcVals, lpTy.value)


def xlTest():
    lnStart = time()
    xlW = xlWrap()

    ln1 = xlW.CreateWorkbook(gcRepoPath + "\\cmodules\\testfile.xls", "myFirstSheet")
    print(ln1)
    lnz = xlW.GetSheetFromName("myFirstSheet")
    print(lnz)
    lcRow = "12345\tTesting 123\tmoretesting\tFALSE\tTRUE\t235098.342\t23459087.13\t20090329\t20090402\t34343\t1234987\t234089\t23444.343\t343.33\t123.43\tCARRIER 1\tCARRIER 2"
    lcTypes = "NSSLLNNDDNNNNNNSS"
    print(lcTypes)

    for jj in range(0, 35):
        xlW.WriteRowValues(lnz, jj, 0, 16, lcRow, lcTypes)

    # lnx = xlW.WriteCellValue(lnz, 0, 0, "2340985.44", "N")
    # print "write result: ", lnx
    # lnx = xlW.WriteCellValue(lnz, 1, 0, "Testing 123", "S")
    # print lnx
    # lnx = xlW.WriteCellValue(lnz, 2, 0, "20070329", "D")
    # print lnx
    # lnx = xlW.WriteCellValue(lnz, 3, 0, "00000000093000", "D")
    # print "time only", lnx
    # lnx = xlW.WriteCellValue(lnz, 4, 0, "FALSE", "L")
    # print "logical ", lnx
    # lnx = xlW.WriteCellValue(lnz, 5, 0, "99434", "N")
    # print "integer ", lnx
    # lnx =  xlW.WriteCellValue(lnz, 6, 0, "59340", "N")
    # print "integer ",lnx
    # lnx = xlW.WriteCellValue(lnz, 7, 0, "=A6+A7", "F")
    # print lnx
    # lnx = xlWrap.WriteCellValue(lnz, 1, 12, "Way out here", "S")
    # print lnx

    lcGetRow, lcTypes = xlW.GetRowValues(lnz, 30, 0, 8)
    datalist = lcGetRow.split(b"\t")
    print(datalist)

    lcName, lnCol1, lnColN, lnRow1, lnRowN = xlW.GetSheetStats(lnz)
    print(lcName)
    print(lnCol1)
    print(lnColN)
    print(lnRow1)
    print(lnRowN)

    print("Testing HIDING")
    ln9 = xlW.SetRowHidden(lnz, 16, True)
    print("Set Hidden Row Result", ln9)
    print(xlW.GetErrorMessage())
    ln9 = xlW.IsRowHidden(lnz, 16)
    print("Is Row Hidden Result:", ln9)
    ln9 = xlW.SetColHidden(lnz, 5, True)
    print("Set Hidden Col:", ln9)
    ln9 = xlW.IsColHidden(lnz, 5)
    print("Is Col Hidden:", ln9)

    ln1 = xlW.CloseWorkbook(1)
    lnEnd = time()
    print (lnEnd - lnStart)
    print(ln1)

    print("\nopening truckload rate problem file")
    ln2 = xlW.OpenWorkbook(gcRepoPath + "\\cmodules\\QGTruckloadRateTEST.xlsx")
    print(ln2)
    xVal, xType = xlW.GetCellValue(0, 1, 0)
    print(xVal, xType)

    xVal, xType = xlW.GetCellValue(0, 1, 1)
    print(xVal, xType)

    xVal, xType = xlW.GetCellValue(0, 1, 2)
    print(xVal, xType)

    xVal, xType = xlW.GetCellValue(0, 1, 3)
    print(xVal, xType)
    xVal, xType = xlW.GetCellValue(0, 1, 4)
    print(xVal, xType)
    xVal, xType = xlW.GetCellValue(0, 1, 5)
    print(xVal, xType)
    xVal, xType = xlW.GetCellValue(0, 1, 6)
    print(xVal, xType)

    xlW.CloseWorkbook(0)
    print("Opening Baldwin")
    ln2 = xlW.OpenWorkbook("e:\\temp\\baldwinrated_RATED.xls")
    print(ln2)
    for jj in range(0, 6):
        lcStr, lcType = xlW.GetCellValue(0, 1, jj)
        print(lcStr, lcType)

    xlW.CloseWorkbook(0)
    #
    #  ln3 = xlW.OpenWorkbook(gcRepoPath + "\\cmodules\\bidderlist-small.xls")
    #  print ln3
    #  for jj in range(0, 35):
    #        lcStr, lcType = xlW.GetCellValue(0, 1, jj)
    #        print lcStr, " ", lcType
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
