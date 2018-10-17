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
# along with this program.  If not, see <https://www.gnu.org/licenses/>
"""
MPSS Basic Utility functions, many copied over from VFP MPSSLib.PRG, some based on VFP functions or commands
Created Nov. 8, 2010. JSH.
"""
#  8/30/2011 JMc  added FSacquire, thispath, getconfig
# 10/06/2011 JMc  added get
# 10/26/2011 JSH  mod to getconfig parameters.
# 03/18/2012 JSH  added MPlocal_except_handler
# 08/12/2012 JSH  added named parm to HexToString.
# 11/15/2012 JMc  1) fixed WEEK() so it would work.  2) added sanitizeHtml()
# 3/29/2016 JMc  Added NamedTuple (with defaults)
# 10/15/2018 JSH Modified for Python 3.x compatibility
from __future__ import print_function
import collections
import random
import os
import fnmatch
import traceback
import time
import datetime
from dateutil import parser

import sys
import math
import glob
import re
import json
import pywintypes
import calendar
import shutil

from win32com.server.exception import COMException
import win32con
import win32api
from functools import wraps
import tempfile
if sys.version_info[0] <= 2:
    _ver3x = False
    from BeautifulSoup import BeautifulSoup, Comment
    import _winreg as xWinReg
else:
    _ver3x = True
    from bs4 import BeautifulSoup, Comment
    import winreg as xWinReg
try:
    from validate import Validator
except ImportError:
    Validator = False
    print('### Unable to import configobj.  Please easy_install configobj')

gcNonPrintables = "\x01\x02\x03\x04\x05\x06\x07\x08\x09\x0A\x0B\x0C\x0D\x0E\x0F\x10\x11" + \
                  "\x12\x13\x14\x15\x16\x17\x18\x19\x1A\x1B\x1C\x1D\x1E\x1F"
gxNonPrintables = dict()
for c in gcNonPrintables:
    gxNonPrintables[ord(c)] = None

gxDateRangeTypes = {"LMONTH": "Previous Month",
                    "LWEEK": "Previous Week",
                    "LQUARTER": "Previous Quarter",
                    "LYEAR": "Previous Year",
                    "YTD": "Year to Date",
                    "MTD": "Month to Date",
                    "QTD": "Quarter to Date",
                    "DATERANGE": "Custom Date Range"}
gbIsRandomized = False
gcLastErrorMessage = ""
cMonths = "JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC"
gxMonthList = cMonths.split(",")
gxConversionFactors = {"LB2KG": 0.453592, "KG2LB": 2.20462, "CM2IN": 0.3937008, "IN2CM": 2.54, "IN2MT": 0.0254,
                       "FT2MT": 0.3048, "MT2FT": 3.28084, "CF2CM": 0.0283168, "MT2IN": 39.3701, "CM2CF": 35.3147,
                       "IN2FT": 0.083333, "FT2IN": 12.0, "MT2CM": 100.0, "CM2MT": 0.010, "LB2LB": 1.0, "KG2KG": 1.0,
                       "CM2CM": 1.0, "IN2IN": 1.0, "FT2FT": 1.0, "MT2MT": 1.0, "CF2CF": 1.0}

gxHTMLCodes = {
    ">": "&gt;",
    "<": "&lt;",
    "&": "&amp;",
    '"': "&quot;",
    "\\": "&#92;",
    "/": "&#47;",
    "~": "&tilde;",
    "^": "&circ;"}

gxHTMLSimpleCodes = {
    ">": "&gt;",
    "<": "&lt;",
    "&": "&amp;",
    '"': "&quot;",
    "\\": "&#92;"}

# MIME Types per http://filext.com/faq/office_mime_types.php
gxMIMEtypes = {"DOC":  "application/msword",
               "DOT":  "application/msword",
               "DOCX": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
               "DOTX": "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
               "DOCM": "application/vnd.ms-word.document.macroEnabled.12",
               "DOTM": "application/vnd.ms-word.template.macroEnabled.12",
               "XLS":  "application/vnd.ms-excel",
               "XLT":  "application/vnd.ms-excel",
               "XLA":  "application/vnd.ms-excel",
               "XLSX": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
               "XLTX": "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
               "XLSM": "application/vnd.ms-excel.sheet.macroEnabled.12",
               "XLTM": "application/vnd.ms-excel.template.macroEnabled.12",
               "XLAM": "application/vnd.ms-excel.addin.macroEnabled.12",
               "XLSB": "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
               "PPT":  "application/vnd.ms-powerpoint",
               "POT":  "application/vnd.ms-powerpoint",
               "PPS":  "application/vnd.ms-powerpoint",
               "PPA":  "application/vnd.ms-powerpoint",
               "PPTX": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
               "POTX": "application/vnd.openxmlformats-officedocument.presentationml.template",
               "PPSX": "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
               "PPAM": "application/vnd.ms-powerpoint.addin.macroEnabled.12",
               "PPTM": "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
               "POTM": "application/vnd.ms-powerpoint.template.macroEnabled.12",
               "PPSM": "application/vnd.ms-powerpoint.slideshow.macroEnabled.12",
               "TXT":  "text/plain",
               "XML":  "text/xml",
               "PDF":  "application/pdf",
               "CSV":  "application/vnd.ms_excel",
               "RFQ":  "text/xml"}

# More constants used by functions below.  Evaluate just once when this module is loaded.
allChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
passChars = "ABCDEFHJKLMNPQRSTUVWXYZ23456789"
gcErrorFilePath = "c:\\temp"

###############
# this copy&pasted from configtool.  Remove when getconfig in websession.py and purgeCmd.py use configtools
###############
if Validator:
    from configobj import ConfigObj, flatten_errors
    vdt = Validator()

    def error_list(self, copy=False):
        """
        Validate using vdt

        :copy: option to copy all defaults into current object

        :return: None if missing, False if verified, string of errors if not validated
        """
        if not self.configspec:
            return
        errors = list()
        ok = self.validate(vdt, copy=copy)    # converts to correct type ( integer, list)
        if ok is not True:
            if not ok:
                errors.append('no value present')
            for (section_list, key, _) in flatten_errors(self, ok):
                if key is not None:
                    errors.append('  ## The "%s" key in the section "%s" failed validation' % (key, ', '.join(section_list)))
                else:
                    errors.append('  ## The following section was missing:%s ' % ', '.join(section_list))
            return errors
        return False

    def add(self, other):
        """
        allow ConfigObj + ConfigObj structure to create new configs

        :param other: Other ConfigObj item
        :return: new (unless error, then old)
        """
        if not other:
            return self
        new = ConfigObj(self,configspec=self.configspec)
        custom = ConfigObj(other)  # allows either filename or dict
        new.merge( custom)
        errors = self.errors()
        if errors:
            print('###### %s in the below folder failed validation, ignored. #####' % other)
            for e in errors:
                print(e)
            return self
        return new
    ConfigObj.__add__ = add
    ConfigObj.errors = error_list

if _ver3x:
    def bytes2str(xbytes):
        """
        Transforms an ASCII byte string into unicode
        :param xbytes:
        :return:
        """
        return None
else:
    def bytes2str(xbytes):
        return xbytes


##########################################################################################
# Cleans up HTML to get rid of possible dangerous codes or disruptive codes when we are
# importing HTML into a formated text entry screen in a browser.
def sanitizeHtml(value):
    rjs = r'[\s]*(&#x.{1,7})?'.join(list('javascript:'))
    rvb = r'[\s]*(&#x.{1,7})?'.join(list('vbscript:'))
    re_scripts = re.compile('(%s)|(%s)' % (rjs, rvb), re.IGNORECASE)
    validTags = 'p i strong b u a h1 h2 h3 h4 h5 h6 blockquote br ul ol li span sup sub '\
                'strike div font hr'.split()
    validAttrs = 'href src width height style align face size color'.split()
    soup = BeautifulSoup(value)
    for comment in soup.findAll(text=lambda text: isinstance(text, Comment)):
        # Get rid of comments
        comment.extract()
    newoutput = soup.renderContents()
    while 1:
        # this will filter out tags in the text attempting to bypass normal filters
        oldoutput = newoutput
        soup = BeautifulSoup(newoutput)
        for tag in soup.findAll(True):
            if tag.name not in validTags:
                tag.hidden = True
            attrs = tag.attrs
            tag.attrs = []
            for attr, val in attrs:
                if attr in validAttrs:
                    val = re_scripts.sub('', val) # Remove scripts (vbs & js)
                    tag.attrs.append((attr, val))
        newoutput = soup.renderContents()
        if oldoutput == newoutput:
            break
    return soup.renderContents().decode('utf-8').encode('ascii', 'xmlcharrefreplace')


def findAllFiles(cDirPath, cSkel="", cExcl=""):
    """
    Finds all files matching the standard OS file wildcard pattern in cSkel and returns them in a list
    of fully qualified path names.
    :param cDirPath:
    :param cSkel: Include '?' or '*' to match multiple file names.  Do NOT include the path, just the file name.
    :param cExcl: Reverse of cSkel -- exclude files with matching this wild-card file-name string.
    :return: List of matching files or None on error.  List may be empty if there aren't any files there.
    """
    if not os.path.exists(cDirPath):
        return None
    xRet = list()
    xGen = os.walk(cDirPath)
    for xG in xGen:
        cPath = xG[0]
        for cF in xG[2]:
            if cExcl:
                if fnmatch.fnmatch(cF, cExcl):
                    continue
            if cSkel:
                if not fnmatch.fnmatch(cF, cSkel):
                    continue
            xRet.append(os.path.join(cPath, cF))
    return xRet


def getTempFileName(ext="tmp"):
    """
    Returns a fully qualified path name for a temporary file that can be used as needed.  The file will NOT
    be created.  That is the responsibility of the calling program.  It will NOT be deleted automatically either.
    :param ext: Pass an extension if that will be important, otherwise the extension .TMP will be used
    :return: The file name, or "" on error
    """
    cFileName = ""
    cPath = tempfile.gettempdir()
    cBase = strRand(10) + "." + ext
    if cPath:
        cFileName = os.path.join(cPath, cBase)
    return cFileName


def FSacquire(name, isdir=True, root=False, starting=None):
    '''Find a directory or file here or higher in the directory structure.
    Returns file/folder name or an empty string
    isdir: Must be a folder if True(default), otherwise must be a file
    root:  Allow it to be found at the root of the drive (default=False)
    starting: specify starting directory explicitly  # added JMc 9/12/12
    '''  # JMc 8/30/2011
    if isdir:
        test = os.path.isdir
    else:
        test = os.path.isfile
    fn = ''
    testpath = starting and os.path.abspath(starting) or os.getcwd()
    # print "IN FSACQUIRE:", testpath
    while True:
        testname = os.path.join(testpath, name)
        testpath, lastdir = os.path.split(testpath)
        atroot = not bool(lastdir)
        if not root and atroot:
            # print "FOUND AT ROOT: BAD lastdir=", lastdir, "testpath=", testpath
            break  # found at root not allowed
        # print "TESTING FOR: ", testname
        if test(testname):
            fn = testname
            break
        if atroot:
            break
    return fn


def isDigits(lpcStr, bAllowSpace=False):
    """
    examines every character in lpcStr and returns True if all are in the range from 0 to 9, else False
    :param lpcStr: String to test
    :param bAllowSpace: pass True if a ' ' char is OK.
    :return: See above
    """
    cValid = "0123456789"
    if bAllowSpace:
        cValid += " "

    for c in lpcStr:
        if c not in cValid:
            return False
    return True


def convertUnits(cConvCode="", nFromNumber=0.0):
    """
    Converts the nFromNumber (converted to float if necessary) into the specified output value with units conversion:
    :param cConvCode: Specifies the from and to units:
        LB2KG = Pounds to Kilograms: specify nFromNumber in pounds.
        KG2LB = Kilograms to Pounds: specify nFromNumber in kilograms
        CM2IN = Centimeters to Inches: specify nFromNumber in centimeters
        IN2CM = Inches to Centimeters: specify nFromNumber in inches
        IN2MT = Inches to Meters: specify from Number in inches
        FT2MT = Feet to Meters: specify nFromNumber in Feet
        MT2FT = Meters to Feet: specify nFromNumber in Meters
        CF2CM = Cubic Feet to Cubic Meters: specify nFromNumber in Cubic Feet
        CM2CF = Cubic Meters to Cubic Feet: specify nFromNumber in Cubic Meters
        MT2IN = Meters to Inches: specify nFromNumber in Meters
        IN2FT = Inches to Feet: specify nFromNumber in Inches
        FT2IN = Feet to Inches: specify nFromNumber in Feet
        MT2CM = Meters to Centimeters: specify nFromNumber in Meters
        CM2MT = Centimeters to Meters: specify nFromNumber in Centimeters
        Identities are also supported:
        CM2CM, IN2IN, FT2FT, CF2CF, MT2MT, LB2LB, KT2KG

    :param nFromNumber: The value you are converting FROM, see above.
    :return: the converted value or None on error.
    """
    global gxConversionFactors
    nFactor = gxConversionFactors.get(cConvCode, None)
    if nFactor is None:
        return None
    else:
        return nFromNumber * nFactor


def IsValidNumber(lpcStr, bForceInt=False):
    """
    Pass any string that purports to be a representation of a number (like a value passed back by
    a web page form).  Returns a float or an int consisting of the numeric value (or None if the value is not
    a valid number).  A non-string value will trigger an error, except for None, which returns None
    as of Oct. 10, 2013.
    """
    lnReturn = 0
    if lpcStr is None:
        lnReturn = None
    else:
        if lpcStr == "":
            lnReturn = 0
        else:
            if "." in lpcStr:
                # Assume they want a floating point value...
                try:
                    lnReturn = float(lpcStr)
                except:
                    lnReturn = None
                if bForceInt and lnReturn is not None:
                    lnReturn = int(lnReturn)
            else:
                try:
                    lnReturn = int(lpcStr)
                except:
                    lnReturn = None
    return lnReturn


def deleteByAge(cSkel, nDays):
    """
    Pass a file name skeleton (fully qualified path name) which may or may not match some files in the
    specified directory.  Any matching files older than nDays in age will be deleted.  Uses glob() for the
    file matching so ? and * are recognized in the skeleton.  Returns the number of files deleted.  Does
    not raise an error if no matching files are found, just returns 0.
    """
    lnReturn = 0
    xFiles = glob.glob(cSkel)
    dToday = datetime.date.today()
    dTest = dToday - datetime.timedelta(days=nDays)
    cPath, cFile = os.path.split(cSkel)
    if len(xFiles) > 0:
        for cF in xFiles:
            dDate = datetime.date.fromtimestamp(os.stat(cF).st_mtime)
            if dDate < dTest:
                try:
                    os.remove(cF)
                    lnReturn += 1
                except:
                    pass  # Nothing to be done
    return lnReturn


def isDateBetween(tTest, tFrom, tTo):
    """
    Simple test of whether a datetime or date value tTest is between the datetime or date values
    of tFrom and tTo, inclusive.
    """
    if isinstance(tTest, datetime.date):
        tTest = datetime.datetime(tTest.year, tTest.month, tTest.day)
    if isinstance(tFrom, datetime.date):
        tFrom = datetime.datetime(tFrom.year, tFrom.month, tFrom.day)
    if isinstance(tTo, datetime.date):
        tTo = datetime.datetime(tTo.year, tTo.month, tTo.day)

    dDiff1 = tTest - tFrom
    dDiff2 = tTo - tTest
    nDiff1 = dDiff1.days
    nDiff2 = dDiff2.days
    bReturn = True
    if nDiff1 < 0:
        bReturn = False
    if nDiff2 < 0:
        bReturn = False
    return bReturn


def CTOD(cDateString, cPattern="mm/dd/yy"):
    """
    Pass a simple date string, and this returns a Python datetime.date() value representing the string.  This is like
    the VFP CTOD() function.  The difference is that in VFP, the way the date is parsed is based on the setting
    of SET DATE TO, which has values like mdy or dmy, etc.  In this version, if you are starting with a non-US type
    date string, you'll need to pass the cPattern value where mm == the month number, yy = the year number, and
    dd = the day number separated by delimiters (required).

    Changed 08/26/2015 by addition of accepting the Evos/MPSS standard codes for dates: MDY, MDYY, DMY, DMYY, DMMY, and YYMD
    See FormatDateString() for details on their meaning. JSH.
    """
    global gxMonthList
    cDateString = cDateString.strip()
    cPattern = cPattern.strip()

    if len(cPattern) == 3 or len(cPattern) == 4:
        if cPattern == "MDY":
            cPattern = "mm-dd-yy"
        elif cPattern == "MDYY":
            cPattern = "mm-dd-yyyy"
        elif cPattern == "DMY":
            cPattern = "dd-mm-yy"
        elif cPattern == "DMYY":
            cPattern = "dd-mm-yyyy"
        elif cPattern == "DMMY":
            cPattern = "dd-mmm-yyyy"
        else:
            cPattern = "INVALID"
    cDateString = cDateString.replace("/", "-")
    cDateString = cDateString.replace("\\", "-")
    cDateString = cDateString.replace(".", "-")
    cDateString = cDateString.replace(" ", "-")

    cPattern = cPattern.replace("/", "-")
    cPattern = cPattern.replace("\\", "-")
    cPattern = cPattern.replace(".", "-")
    cPattern = cPattern.upper()

    if RIGHT(cPattern, 3) == "-YY":
        # handle the case where the pattern specifies a 2-digit year
        # but the actual date string has a 4-digit year
        cTestYear = RIGHT(cDateString, 4)
        try:
            nTestYear = int(cTestYear)
        except:
            nTestYear = 0
        if isBetween(nTestYear, 1800, 2500):
            cPattern += "YY"  # Add yy to support 4-digit year.

    nPatLen = len(cPattern)
    cDateString = cDateString[0: nPatLen]  # trim off any possible time info fields.

    cDelim = "-"
    if cDelim not in cDateString:
        return None  # Invalid format
    if cDelim not in cPattern:
        return None  # also invalid

    xPat = cPattern.split(cDelim)
    xParts = cDateString.split(cDelim)

    if len(xPat) != 3:
        return None  # also invalid
    if len(xParts) != 3:
        return None  # also invalid

    nYear = 0
    nMonth = 0
    nDay = 0

    for kk in range(0, 3):
        cPat = xPat[kk].upper()
        cElem = xParts[kk]
        cElem = cElem.strip()
        if cElem == "":
            return None  # Any empty spot in the date string results in a null date.
        if cPat == "MMM":
            try:
                nMonth = gxMonthList.index(cElem.upper())
            except:
                nMonth = -1
            nMonth += 1  # 0 if an error occurred.
        elif "M" in cPat:
            nMonth = int(cElem)
        elif "Y" in cPat:
            nYear = int(cElem)
            if cPat == "YY" and nYear < 100:
                nYear += 2000
            # Added this to correct the 2-digit year stuff. JSH. 09/25/2015.
        elif "D" in cPat:
            nDay = int(cElem)
    if (nYear == 0) or (nMonth == 0) or (nDay == 0):
        return None  # Again not valid
    return datetime.date(nYear, nMonth, nDay)


def cleanUnicodeString(cSource):
    """
    Applies transformations to strings where it may be possible to have higher-order unicode values embedded
    in a standard ascii or UTF8 string.
    :param cSource: Source string to be transformed.  Must be a string or a error will be triggered.
    :return: Clean-up string, or None on error condition
    NOTE: May not work reliably when utf-16 source is passed or 2 or 3 byte unicode characters are embedded
    in a UTF-8 string.
    """
    if isstr(cSource):  # Added this 01/03/2013. JSH.
        if isinstance(cSource, unicode):
            cSource = cSource.encode("utf-8", "replace")
        else:
            cSource = unicode(cSource, "utf-8", "replace")

        cReturn = cSource.encode("ascii", "replace")

        cReturn = cReturn.replace(u"\x93", u'"')
        cReturn = cReturn.replace(u"\x94", u'"')
    else:
        raise ValueError("Must be a string")
    return cReturn


def toAscii(cSrc):
    """
    Brute force removal of non-ascii characters from a string.  Returns a standard string, NOT unicode.
    :param cSrc:
    :return: standard ascii string.
    Always works, regardless of the input string, but will lose all upper range special characters.
    New 10/07/2016. JSH.
    """
    xList = list()
    for i, c in enumerate(cSrc):
        if ord(c) <= 128:
            xList.append(str(c))
        else:
            xList.append("")
    return "".join(xList)


def rawReplace(cSource, cOld, cNew, bMakeAscii=False):
    xList = list()
    for i, c in enumerate(cSource):
        if bMakeAscii:
            if ord(c) > 128:
                cNew = ""
        if c == cOld:
            xList.append(str(cNew))
        else:
            xList.append(c)
    return "".join(xList)


def safeHTMLcoding(lcSource, bFixNewLines=False, bSimple=False):
    """
    Replaces unsafe characters in a text string with their HTML codes to allow strings to be embedded
    more easily in javascript variables passed from the page engine.
    Fixed this function 04/20/2017. JSH.
    """
    cRet = toAscii(lcSource)
    xDict = gxHTMLSimpleCodes if bSimple else gxHTMLCodes

    for cChar, cCode in xDict.items():
        cTemp1 = cChar
        cTemp2 = cCode
        cRet = rawReplace(cRet, cTemp1, cTemp2, True)

    if bFixNewLines:
        cRet = cRet.replace("\r\n", "<br>")  # Convert the DOS EOL markers
        cRet = cRet.replace("\n", "<br>")  # If UNIX EOL markers are left, convert them.
    return cRet


def TTOD(tTheDT):
    """
    Like the VFP TTOD() function, converts a datetime type value to a date type value.
    """
    tReturn = None
    try:
        tReturn = datetime.date(tTheDT.year, tTheDT.month, tTheDT.day)
    except:
        tReturn = None
    return tReturn


def DTOT(dTheDate):
    """
    Quick conversion utility to convert a Python datetime.date() object to a datetime.datetime() object,
    like the corresponding VFP DTOT() function.  Returns the datetime value or None on error
    """
    tReturn = None
    try:
        tReturn = datetime.datetime(dTheDate.year, dTheDate.month, dTheDate.day, 0, 0, 0, 0)
    except:
        tReturn = None  # all we can do
    return tReturn


def fDiv(nDividend, nDivisor):
    """
    In Python 2.x the division of two integers results in an integer result.  If instead you want a floating result
    with the appropriate fractional part, you have to convert one of the values to float first.  This function simplifies
    that process and in addition protects against divide by zero errors.  In DBZ situations, it returns 0.0.  Always
    returns a float value.
    """
    if nDivisor == 0:
        return 0.0
    return float(nDividend) / float(nDivisor)


def CTOTX(cDateStr="", cDateFmat=None):
    """
    Addresses some of the shortcomings of CTOT for parsing datetimes in the form of 08/01/2018 10:30AM which
    CTOT struggles with.  Uses the Python dateutil object to do the parsing
    :param cDateStr: In one of many possible formats.  dateutil parser is clever about this
    :param cDateFmat: MDY, MDYY, YYDM, YYMD, DMY, DMYY, YDM, YMD
    :return: datetime.datetime object or None on error
    Added 03/01/2018. JSH.
    """
    if not cDateStr or not isstr(cDateStr):
        return None
    if not cDateFmat:
        cDateFmat = "MDY"  # US standard as default.
    bDayFirst = False
    bYearFirst = False
    tRet = None
    if cDateFmat[0: 1] == "Y":
        bYearFirst = True
    if "DM" in cDateFmat:
        bDayFirst = True
    try:
        tRet = parser.parse(cDateStr, dayfirst=bDayFirst, yearfirst=bYearFirst)
    except:
        tRet = None
    return tRet


def CTOT(cDateStr, cDateFmat=None):
    """
    The reverse of the TTOC() function below.  It auto-detects which of the datetime formats (not including
    the one which is "time only") has been passed, and interprets it accordingly.  If it doesn't recognize the
    format, or the date/time is otherwise not valid, returns None, otherwise returns a datetime.datetime value.

    08/21/2014 - Added a parser for MercuryGate XML date/time values of the form: 08/21/2014 10:23
    """
    if not isstr(cDateStr):
        return None
    cDateStr = cDateStr.upper()
    tDateTime = None
    cYr = ""
    cMo = ""
    cSe = ""
    cHr = ""
    cMi = ""
    cMo = ""
    cDa = ""
    cFrac = ""
    cDate = ""
    cTime = ""
    nHr = 0
    cPM = ""
    if ("T" in cDateStr) and ("-" in cDateStr) and (":" in cDateStr):
        # Type 3
        x = cDateStr.split("T")
        if len(x) == 2:
            cDate, cTime = x
        x = cDate.split("-")
        if len(x) == 3:
            cYr, cMo, cDa = x
        x = cTime.split(":")
        if len(x) == 3:
            cHr, cMi, cSe = x
        try:
            tDateTime = datetime.datetime(int(cYr), int(cMo), int(cDa), int(cHr), int(cMi), int(cSe))
        except:
            tDateTime = None

    elif (" " in cDateStr) and ("/" in cDateStr) and (("P" in cDateStr) or ("A" in cDateStr)):
        # Type 0
        x = cDateStr.split(" ")
        if len(x) == 3:
            cDate, cTime, cPM = x

        cDate = cDate.replace("-", "/")
        cDate = cDate.replace(".", "/")
        x = cDate.split("/")
        if len(x) == 3:
            cMo, cDa, cYr = x
        if len(cYr) == 2:
            if cYr.isdigit():
                if int(cYr) > 55:
                    cYr = "19" + cYr
                else:
                    cYr = "20" + cYr
        x = cTime.split(":")
        if len(x) == 3:
            cHr, cMi, cSe = x
        nHr = 0
        if cHr.isdigit():
            nHr = int(cHr)
            if ("P" in cPM) and (nHr < 12):
                nHr = nHr + 12
            elif ("A" in cPM) and (nHr == 12):
                nHr = 0
        try:
            tDateTime = datetime.datetime(int(cYr), int(cMo), int(cDa), nHr, int(cMi), int(cSe))
        except:
            tDateTime = None

    elif (len(cDateStr) == 16) and (":" in cDateStr) and ("/" in cDateStr) and (" " in cDateStr):
        # MercuryGate XML date/time representation
        x = cDateStr.split(" ")
        if len(x) == 2:
            cDate, cTime = x
        if cDate:
            x = cDate.split("/")
            if len(x) == 3:
                cMo, cDa, cYr = x
        if cTime:
            x = cTime.split(":")
            if len(x) == 2:
                cHr, cMi = x
        cSe = "0"
        try:
            tDateTime = datetime.datetime(int(cYr), int(cMo), int(cDa), int(cHr), int(cMi), int(cSe))
        except:
            tDateTime = None

    elif (len(cDateStr) == 20) and (":" in cDateStr):
        # Type 9
        cYr = cDateStr[0:4]
        cMo = cDateStr[4:6]
        cDa = cDateStr[6:8]
        cTime = cDateStr[8:]
        x = cTime.split(":")
        if len(x) == 4:
            cHr, cMi, cSe, cFrac = x
        try:
            tDateTime = datetime.datetime(int(cYr), int(cMo), int(cDa), int(cHr), int(cMi), int(cSe))
        except:
            tDateTime = None

    elif (len(cDateStr) == 14) and cDateStr.isdigit():
        # Type 1
        cYr = cDateStr[0:4]
        cMo = cDateStr[4:6]
        cDa = cDateStr[6:8]
        cHr = cDateStr[8:10]
        cMi = cDateStr[10:12]
        cSe = cDateStr[12:]
        try:
            tDateTime = datetime.datetime(int(cYr), int(cMo), int(cDa), int(cHr), int(cMi), int(cSe))
        except:
            tDateTime = None

    elif (len(cDateStr) == 26) and (":" in cDateStr) and (" " in cDateStr) and ("." in cDateStr) and ("-" in cDateStr):
        # Native Python str() of datetime.datetime
        x = cDateStr.split(" ")
        if len(x) == 2:
            cDate, cTime = x
        x = cDate.split("-")
        if len(x) == 3:
            cYr, cMo, cDa = x
        x = cTime.split(":")
        if len(x) == 3:
            cHr, cMi, cSe = x
        x = cSe.split(".")
        if len(x) == 2:
            cSe, cFrac = x
        try:
            tDateTime = datetime.datetime(int(cYr), int(cMo), int(cDa), int(cHr), int(cMi), int(cSe), int(cFrac))
        except:
            tDateTime = None
    else:
        tDateTime = None
        # Default return nothing.

    return tDateTime


def TTOC(tDateTime, nHow=0):
    """
    Pass a datetime value and an optional value of nHow and this creates a text string
    representation of the value in the same exact way as the VFP TTOC() functions.  This simplifies
    the creation of values to seek for when datetime values are included in VFP table indexes.
    - nHow = 0: yields MM/DD/YY HH:MM:SS xM (where xM is either AM or PM)
    - nHow = 1: yields YYYYMMDDHHMMSS with no punctuation.  This will match a VFP index value like:
       UPPER(LOGIN_ID) + TTOC(MYDATETIME, 1)
    - nHow = 2: Time only in form HH:MM:SS xM per the above.
    - nHow = 3: yields YYYY-MM-DDTHH:MM:SS (where time is expressed in 24 hour clock time) - used in some XML formats.
    - nHow = 4: yields YYYY_MM_DD_HH_MM_SS, which can be used in file names and read more easily than nHow = 1.
    - nHow = 9: yields YYYYMMDDHH:MM:SS:000 (the format needed by the CodeBaseTools seek() method when
      passing JUST the datetime value.  This is a format which is NOT supported by the VFP TTOC() function.
    - nHow set to any other number returns the built-in Python str() representation of a datetime.datetime type:
      YYYY-MM-DD HH:MM:SS.SSSSSS where the number of seconds is represented to 6 decimal places.

    For human readable date/time strings see FormatDateTimeString()
    """
    if not isinstance(tDateTime, datetime.datetime):
        lcReturn = ""
    elif nHow == 0:
        lcReturn = tDateTime.strftime("%m/%d/%y %I:%M:%S %p").replace(".", "")

    elif nHow == 1:
        lcReturn = tDateTime.strftime("%Y%m%d%H%M%S")

    elif nHow == 2:
        lcReturn = tDateTime.strftime("%I:%M:%S %p").replace(".", "")  # Just the time portion

    elif nHow == 3:
        lcReturn = tDateTime.strftime("%Y-%m-%dT%H:%M:%S")

    elif nHow == 4:
        lcReturn = tDateTime.strftime("%Y_%m_%d_%H_%M_%S")

    elif nHow == 9:
        lcReturn = tDateTime.strftime("%Y%m%d%H:%M:%S:000")
    else:
        lcReturn = str(tDateTime)
    return lcReturn


######################################################################
##
#  get - generic object keyed access with default
##
######################################################################
def get(obj, key, df=None):
    try:
        r = obj[key]
    except:
        r = df
    return r


######################################################################
##
# NamedTuple - create a named tuple with defaults.
##
######################################################################
# noinspection PyProtectedMember
class NamedTuple(object):

    def __init__(self, tupleClassName, tupleFieldnames, **defaults):
        """
        :param tupleClassName: namedtuple parameter
        :param tupleFieldnames: namedtuple parameter
        :param defaults: one or more default field assignments  I.E. age=0
        :returns: a namedtuple *template* instance with defaults
        example:
        > tmpl = NamedTuple('test', 'name state age', state='WA')
        > me = tmpl(name='Bob', age=99)
        test(name='Bob', state='WA', age=99)

        NOTE: Once you have created the actual named tuple from this template, you can convert it
        to an ordered dict and back again to the same kind of named tuple.  This makes the code for altering
        multiple elements much more straightforward:
        > meD = me._asdict()
        # the above meD is an ordered dict.
        # to create a copy of tmpl with no changes, do
        > me2 = tmpl(**meD)
        # Now me2 is an identical namedtuple.  You could have changed some of the values in the meD ordered
        # dict if you needed to.
        """
        emptyTuple = collections.namedtuple(tupleClassName, tupleFieldnames)
        setDict = dict()
        for field in emptyTuple._fields:
            setDict[field] = defaults.setdefault(field, "")
        self.namedTuple = emptyTuple(**setDict)

    def __call__(self, **fields):
        """
        create named tuple from this template.  All missing fields will use their default values
        """
        return self.namedTuple._replace(**fields)


######################################################################
##
# Bunch - generic object to store "things".  Access by dict or attribute
##
######################################################################
class Bunch(dict):
    """
    A Dictionary Object that can be accessed via attributes
    """
    def __getattr__(self, key):
        # noinspection PyTypeChecker
        return self.get(key, '')

    def __setattr__(self, key, value):
        self.__setitem__(key, value)


def thispath(name):
    """ returns full path name of the file 'name' in the same folder as the calling program
    """
    myPath = os.path.abspath(sys.argv[0] if __name__ == '__main__' else __file__)
    ret = os.path.join(os.path.split(myPath)[0], name)
    return ret


def getconfig( name=None, cspecFile='default.conf' ):
    """ return a default configobj or pass a different name to read that file
    Added the cspecFile parm so as to be able to override default.conf.  This
    really needs a .cspec type file for the validation anyway. JSH 10/26/2011.
    """
    if not ConfigObj:
        return    # cannot use if not imported
    if name:
        return ConfigObj(name)
    name = 'app.conf'
    # Note: if app.conf doesn't exist - you MUST run calling program from the app directory
    #      the first time so it is created in the correct folder
    # If it does exist anywhere up the directory structure, it will be found
    fullname = FSacquire(name, isdir=False) or name
    appdir = os.path.split(fullname)[0] or '.'
    config = ConfigObj(fullname,
                       configspec=thispath(cspecFile),
                       interpolation='template',   # allows use of $name inside CONF file
                       indent_type='    ',
                       )

    vdt = Validator()
    isNewfile = not bool(config)
    config['appdir'] = appdir
    # Added pgmdir to allow that to be templated in the conf file too. JSH. 10/26/2011
    pgmdir = os.path.split(__file__)[0]
    config['pgmdir'] = pgmdir
    xkeys = config.keys()
    results = config.validate(vdt, copy=True)
    if not results:
        print('###### %s failed validation. #####' % fullname)
        for (section_list, key, _) in flatten_errors(config, results):
            if key is not None:
                print('The "%s" key in the section "%s" failed validation' % (key, ', '.join(section_list)))
            else:
                print('The following section was missing:%s ' % ', '.join(section_list))
    if set(config.keys()) - set(xkeys):
        # this happens if a new section or attribute is added to the default.conf
        print('%s config file "%s"' % (('Update','Write new')[isNewfile], fullname))
        # don't write appdir to saved CONF
        del config['appdir']
        config.write()
        config['appdir'] = appdir
    return config


def getrandomseed(nSize=4):
    """
    Returns a seed that can be used by any randomizer, which is not dependant on the system clock.  Use
    if you think your processes may run concurrently and get the same seed for random() from the clock.
    """
    cRandStr = os.urandom(nSize)
    cRandStr = str(cRandStr)
    nSeed = 0
    for j in range(0, nSize):
        nVal = ord(cRandStr[j])
        if nVal > 1:
            nVal -= 1
        nSeed += (255**j) * nVal
    return nSeed


def strRand(StrLength, isPassword=0):
    """ This returns a string of random characters being either uppercase alpha letters
        or numbers.  If isPassword is FALSE or unassigned, then selects from all 36 possible values,
        otherwise omits 1,I,O,0 to avoid password confusion. """
    global gbIsRandomized
    if not gbIsRandomized:
        random.seed(getrandomseed())  # Ensures independence from the clock.
        gbIsRandomized = True
    lcResult = ""
    baseList = (allChars if isPassword == 0 else passChars)
    for i in range(0, StrLength):
        lcResult = lcResult + lcResult.join(random.choice(baseList))
    return lcResult


def isBetween(lpnTest, lpnFrom, lpnTo):
    """ Pass any three values of the same type in the parameters.  If the lpnTest
    value is >= lpnFrom AND <= lpnTo, this function returns 1, otherwise 0
    """
    lnReturn = 1
    if lpnTest < lpnFrom:
        lnReturn = 0
    elif lpnTest > lpnTo:
        lnReturn = 0
    return lnReturn


def getExcelColumn(lpnCol):
    """
    Converts a left-to-right columns number for a spreadsheet (starting with 1) to an alphabetic Excel Column
    ID equivalent.  For example, Column 1 is "A", and Column 26 is "Z".  After that, the letters repeat.
    Column 27 is "AA", 28 is "AB" and so on.
    """
    lcReturn = ""

    if isBetween(lpnCol, 1, 26) :
        lcReturn = allChars[lpnCol - 1]
    elif isBetween(lpnCol, 27, 52) :
        lcReturn = "A" + allChars[lpnCol - 26 - 1]
    elif isBetween(lpnCol, 53, 78):
        lcReturn = "B" + allChars[lpnCol - 52 - 1]
    elif isBetween(lpnCol, 79, 104):
        lcReturn = "C" + allChars[lpnCol - 78 - 1]
    elif isBetween(lpnCol, 105,130):
        lcReturn = "D" + allChars[lpnCol - 104 - 1]
    elif isBetween(lpnCol, 131,156):
        lcReturn = "E" + allChars[lpnCol - 130 - 1]
    elif isBetween(lpnCol, 157,182):
        lcReturn = "F" + allChars[lpnCol - 156 - 1]
    elif isBetween(lpnCol, 183,208):
        lcReturn = "G" + allChars[lpnCol - 182 - 1]
    elif isBetween(lpnCol, 209,234):
        lcReturn = "H" + allChars[lpnCol - 208 - 1]
    elif isBetween(lpnCol, 235,256):
        lcReturn = "I" + allChars[lpnCol - 234 - 1]
    else:
        lcReturn = ""

    return lcReturn


def confirmPath(thepath):
    global gcLastErrorMessage
    gcLastErrorMessage = ""
    if os.path.exists(thepath):
        return True
    else:
        try:
            os.mkdir(thepath)
        except IOError:
            gcLastErrorMessage = "IO Error"
            return False
        except WindowsError:
            gcLastErrorMessage = "Windows Error"
            return False
        except:
            gcLastErrorMessage = "Other Error"
            return False # Yes, this can mask other problems, but we want a simple OK, NOT OK result.
    return True


def ADDBS(lcPath):
    if not lcPath.endswith("\\"):
        lcPath += "\\"
    return lcPath


def ADDFS(lcPath):
    if not lcPath.endswith("/"):
        lcPath = lcPath.strip() + "/"
    return lcPath

# The following replaces the test isinstance(cX, basestring) or other tests against basestring which is not
# found in Python 3.x  isstr() works in either 2.7 or 3.x environment.
try:
    basestring  # attempt to evaluate basestring
    def isstr(s):
        return isinstance(s, basestring)

except NameError:
    def isstr(s):
        return isinstance(s, str)

# The following replaces the test isinstance(cX, unicode) or other tests against unicode, which is not found
# in Python 3.x.  isunicode() works in either 2.7 or 3.x environment
try:
    unicode  # attempt to evaluate unicode
    def isunicode(s):
        return isinstance(s, unicode)
except NameError:
        def isunicode(s):
            return isinstance(s, str)  # All strings are unicode.


def STRTOFILE(lcString, lcFile, lbAppend=False):
    """
    Added optional 3rd parm to cause the string contents to be appended to an existing file rather than creating a new file.
    Added trap for bad directory name, which causes this to return a False.
    """
    bReturn = True
    lxFile = None

    try:
        if not lbAppend:
            if _ver3x:
                if isstr(lcString):
                    lxFile = open(lcFile, "w", -1)
                else:
                    lxFile = open(lcFile, "wb", -1)
            else:
                lxFile = open(lcFile, 'wb', -1)
        else:
            if _ver3x:
                if isstr(lcString):
                    lxFile = open(lcFile, "a", -1)
                else:
                    lxFile = open(lcFile, "ab", -1)
            else:
                lxFile = open(lcFile, 'ab', -1)
    except:
        bReturn = False
    if bReturn:
        lxFile.write(lcString)
        lxFile.flush()
        os.fsync(lxFile.fileno())  # Added flushing. 07/19/2016. JSH.
        lxFile.close()

    return bReturn


def DTOR(nDegrees=0.0):
    """
    Converts a decimal float value in degrees to radians
    :param nDegrees: for example 45.3495 degrees
    :return: the angle in radians.
    NOTE: Differs from the result from Visual Foxpro DTOR() in the 7th or 8th decimal place.
    """
    try:
        nRet = (nDegrees * 3.1415926535897) / 180.0
    except:
        nRet = None
    return nRet


def RTOD(nRadians=0.0):
    """
    Converts an angle (or lat or lon) from radians to degrees)
    :param nRadians: must be a float
    :return: degrees as a float with a decimal component.  Returns None on error
    NOTE: Differs from the result from Visual FoxPro RTOD() in the 7th or 8th decimal place.
    """
    try:
        nRet = (nRadians * 180.0) / 3.1415926535897
    except:
        nRet = None
    return nRet


def DOW(ltDate, VFPstyle=True):
    """
    Pass this a datetime value and returns the Visual FoxPro Day of week value as returned by the DOW() function.
    By default, VFP treats Sunday as day 1, Monday as 2, etc. through Saturday as 7.  Python considers Monday
    to be 1.
    """
    try:
        if VFPstyle:
            lnDay = ltDate.isocalendar()[2] + 1
            nReturn = lnDay if (lnDay < 8) else 1
        else:
            nReturn = ltDate.isocalendar()[2]  # Native Python.
    except:
        nReturn = 0
    return nReturn  # Added error trapping 09/23/2013. JSH.


def WEEK(ltDate):
    """
    Works similar the VFP WEEK() function.  Pass it a datetime value and it returns the week number within the current
    year.  NOTE: in VFP the dividing line between weeks is between Saturday and Sunday (see DOW above).  In Python,
    the dividing line is between Sunday and Monday.  This function uses the native PY method of calculation.
    Week numbers can range from 1 to 53.
    """
    return ltDate.isocalendar()[1]


def StartOfWeek(ltDate):
    """
    Pass this any date/datetime value and it returns the date() value for the Monday of the week into which the
    ltDate value falls.
    """
    nDOW = DOW(ltDate, VFPstyle=False)
    nDiff = nDOW - 1
    xDiff = datetime.timedelta(nDiff)
    return ltDate - xDiff


def FILETOSTR(lcFile):
    """
    To emulate the VFP function this MUST be a type 'rb', since VFP does NO transformations in a
    FILETOSTR() function.  Note that in Python 3.x, this will return a bytes object, not a string.  In that
    case you'll need to convert to unicode with an encoding, based on what you expect is in the file.
    """
    try:
        lcReturn = open(lcFile, 'rb').read()
    except:  # yes, any error
        lcReturn = ''
    # if _ver3x:
    #     lcReturn = lcReturn.decode("UTF-8", "ignore")
    return lcReturn


def TestTimeEntry(cTestTime, bBlankOK=False):
    """
    Re-creation of the Evos/MPSS Time validator function that checks that a time is in the range from 0000 (midnight)
    to 2359 (one minute prior to midnight).  If lpbBlankOK is True, then allows a completely blank value to be passed.
    :param cTestTime: The text string to be evaluated
    :param bBlankOK: If True, blanks are OK
    :return: The corrected 4-character text time value or None on error.  Sets the gcLastErrorMessage global on error.
    If lpbBlankOK is True and a blank is received, then returns the empty string.
    """
    global gcLastErrorMessage
    cRetTime = None
    cTestTime = cTestTime.strip()
    if not cTestTime and bBlankOK:
        cRetTime = ""
    elif not cTestTime and not bBlankOK:
        gcLastErrorMessage = "Empty Time value is not valid1"
    elif len(cTestTime) > 4:
        gcLastErrorMessage = "Time is not in valid hhmm format2"
    elif len(cTestTime) < 3:
        gcLastErrorMessage = "Time is not in valid hhmm format3"
    else:
        if not isDigits(cTestTime, bAllowSpace=True):
            gcLastErrorMessage = "Time is not in valid hhmm format4"
        else:
            if len(cTestTime) == 3:
                cTestTime = "%04d" % int(cTestTime)
            if cTestTime == "2400":
                cTestTime = "0000"
            cTestHr = cTestTime[0:2]
            if int(cTestHr) > 23:
                gcLastErrorMessage = "Time must be in 24-hour clock time from 0000 to 23595"
            else:
                cTestMin = cTestTime[2:4]
                if int(cTestMin) > 59:
                    gcLastErrorMessage = "Time must be in 24-hour clock time from 0000 to 23596"
                else:
                    cRetTime = cTestTime
    return cRetTime


def FormatDateTimeString(lpdDate, fmatType="A", hourType="12", bOmitComma=False):
    """
    See the FormatDateString() function for details of formatting options.

    There are 4 variations of values for the hourType parm, governing the formatting of the time portion:
    12 >> 3:45pm
    12S >> 3:45:29 pm -- Shows the seconds
    24 >> 15:45 -- Military or 24 hour clock time
    24S >> 15:45:29 -- 24 hour time with seconds.

    See Also TTOC(), which outputs coded text versions of the datetime value.
    Added bOmitComma June 17, 2017. JSH.
    """

    lbDatetime = False
    lbTime = False
    lcStr = ""
    xx = type(lpdDate)
    lcType = repr(type(lpdDate))
    lxParts = lcType.split("'")
    lcType = lxParts[1]
    ltDate = None

    if lcType == 'float':
        ltDate = time.localtime(lpdDate)
        lbTime = True

    elif lcType == 'datetime':
        lbDatetime = True
        ltDate = lpdDate

    elif lcType == 'datetime.datetime':
        lbDatetime = True
        ltDate = lpdDate

    elif lcType == 'datetime.date':
        lbDatetime = True
        ltDate = lpdDate

    elif lcType == 'time.struct_time':
        ltDate = lpdDate
        lbTime = True
    else:
        ltDate = None

    if (lbTime or lbDatetime) and ltDate:
        lcStr = FormatDateString(ltDate, fmatType)
    lcTime = ""
    if lcStr:
        # We have a valid date, so now we can format the time component
        if hourType[0:2] == "12":
            lcHour = "00"
            lcMinute = "00"
            lcSeconds = "00"
            lcAmPm = ""
            if lbTime:
                lcHour = time.strftime("%I", ltDate)
                lcMinute = time.strftime("%M", ltDate)
                lcSeconds = time.strftime("%S", ltDate)
                lcAmPm = time.strftime("%p", ltDate)
            if lbDatetime:
                lcHour = ltDate.strftime("%I")
                lcMinute = ltDate.strftime("%M")
                lcSeconds = ltDate.strftime("%S")
                lcAmPm = ltDate.strftime("%p")
            lnHr = int(lcHour)
            lcHour = str(lnHr)
            if hourType.upper() == "12S":
                lcTime = "%s:%s:%s%s" % (lcHour, lcMinute, lcSeconds, lcAmPm)
            else:
                lcTime = "%s:%s%s" % (lcHour, lcMinute, lcAmPm)

        elif hourType[0:2] == "24":
            lcHour = "00"
            lcMinute = "00"
            lcSeconds = "00"
            lcAmPm = ""
            if lbTime:
                lcHour = time.strftime("%H", ltDate)
                lcMinute = time.strftime("%M", ltDate)
                lcSeconds = time.strftime("%S", ltDate)
            if lbDatetime:
                lcHour = ltDate.strftime("%H")
                lcMinute = ltDate.strftime("%M")
                lcSeconds = ltDate.strftime("%S")

            if hourType.upper() == "24S":
                lcTime = "%s:%s:%s" % (lcHour, lcMinute, lcSeconds)
            else:
                lcTime = "%s:%s" % (lcHour, lcMinute)
        if not bOmitComma:
            lcStr = lcStr + ", " + lcTime
        else:
            lcStr = lcStr + " " + lcTime
    return lcStr


def RIGHT(cStr, nLen):
    """
    Quick way to extract the nLen characters of cStr counting from the right-most end of the string.  So for example,
    mTools.RIGHT("1234567890", 4) returns "7890".  If nLen is greater than the actual length of the string, returns
    the entire string.  If nLen is 0, returns an empty string.  Note: we don't have a LEFT() because that is easy
    to do with one line of Python code using a slice.
    :param cStr: source string from which to extract a part
    :param nLen: the required length of the resulting string.
    :return: sub string as defined above
    """
    cRet = ""
    nTest = len(cStr)
    nStart = nTest - nLen
    if nStart < 0:
        nStart = 0
    if nStart < nTest:
        cRet = cStr[nStart:]
    return cRet


def getDateRangeTypeText(cRangeType):
    """
    Returns a text string which contains the explanatory text for the cRangeType values as required for
    the getCustomDateRange() function.
    """
    global gxDateRangeTypes
    cReturn = gxDateRangeTypes.get(cRangeType, "")
    return cReturn


def PADRIGHT(cStr, nLen, cPad=" "):
    """
    Returns a text string padded at the right with a specified character to a specified length.
    :param cStr: The source string.
    :param nLen: The target length.  If the source string exceeds the length of nLen, it will be truncated.  Normally
           this value should be a positive integer.  Passing a negative integer will remove the specified number of
           characters from the right end of the string -- which may or may not be what you want.
    :param cPad: The padding character to be used to extend the length of the string to nLen.  If omitted, an ascii
           space (32) will be used.
    :return: The resulting string
    This replicates Visual FoxPro behavior in the PADRIGHT() function.  Note that using a format code like "%-10s"
    is not sufficient because it does not truncate, which VFP does.
    """
    nCurLen = len(cStr)
    if nCurLen > nLen:
        cRet = cStr[0:nLen]
    elif nCurLen < nLen:
        cRet = cStr + (cPad * (nLen - nCurLen))
    else:
        cRet = cStr
    return cRet


def getCustomDateRange(cRangeType, tFromDate=None, tToDate=None):
    """
    Generic tool to get specific date ranges from user selections from typical drop-down selectors.
    cRangeType values recognized include:
    LMONTH = Last Month
    LWEEK = Last Week
    LYEAR = Last Year
    YTD = Year to Date
    MTD = Month to Date
    LQUARTER Last Quarter
    QTD = Quarter to Date
    DATERANGE = Date range specified by tFromDate, tToDate

    If cRangeType is any of the first 7, returns a tuple of tFromDate, tToDate (datetime objects) values
    calculated by the specification relative to TODAY.  If it is DATERANGE, then the tFromDate and tToDate values
    are simply returned.  On error, a None value is returned.  Error includes tFromDate or tToDate being empty or None
    or tToDate prior to tFromDate, if DATERANGE is specified, or if an unrecognized value of cRangeType is received.

    Added 09/27/2013. JSH.
    Bug fix 12/01/2015. JSH.
    """
    xReturn = None
    if cRangeType == "DATERANGE":
        if tFromDate is None or tToDate is None:
            return xReturn # bad.  we're out of here.
        else:
            dDiff = tToDate - tFromDate
            if dDiff.days < 0:
                return xReturn # From date after To date.  ERROR.
            else:
                return (tFromDate, tToDate)  # OK, so just return the specified range
    # now we have to calculate the range.
    tNow = datetime.datetime.now()
    nMonth = tNow.month
    nDay = tNow.day
    nYear = tNow.year

    if cRangeType == "LMONTH":
        nNewMonth = 12 if nMonth == 1 else (nMonth - 1)
        nNewYear = nYear if nMonth != 1 else (nYear - 1)
        tFrom = datetime.datetime(nNewYear, nNewMonth, 1, 0, 0, 0)
        nDOW, nDIM = calendar.monthrange(nNewYear, nNewMonth)
        tTo = datetime.datetime(nNewYear, nNewMonth, nDIM)

    elif cRangeType == "LYEAR":
        nNewYear = nYear - 1
        tFrom = datetime.datetime(nNewYear, 1, 1)
        tTo = datetime.datetime(nNewYear, 12, 31)

    elif cRangeType == "LWEEK":
        nDOW = DOW(tNow, VFPstyle=False)
        nLastMonday = nDOW + 6
        dDif = datetime.timedelta(days=(-1 * nLastMonday))
        tFrom = tNow + dDif
        tTo = tFrom + datetime.timedelta(days=6) # Takes you to Sunday.

    elif cRangeType == "YTD":
        tFrom = datetime.datetime(nYear, 1, 1, 0, 0, 0)
        tTo = tNow

    elif cRangeType == "MTD":
        tFrom = datetime.datetime(nYear, nMonth, 1, 0, 0, 0)
        tTo = tNow

    elif cRangeType == "LQUARTER":
        if isBetween(nMonth, 1, 3):
            nNewYear = nYear - 1
            nNewMonth = 10
        elif isBetween(nMonth, 4, 6):
            nNewYear = nYear
            nNewMonth = 1
        elif isBetween(nMonth, 7, 9):
            nNewYear = nYear
            nNewMonth = 4
        else:
            nNewYear = nYear
            nNewMonth = 7
        tFrom = datetime.datetime(nNewYear, nNewMonth, 1)
        nEndMonth = nNewMonth + 2
        nDOW, nDIM = calendar.monthrange(nNewYear, nEndMonth)
        tTo = datetime.datetime(nNewYear, nEndMonth, nDIM)

    elif cRangeType == "QTD":
        if isBetween(nMonth, 1, 3):
            nNewMonth = 1
        elif isBetween(nMonth, 4, 6):
            nNewMonth = 4
        elif isBetween(nMonth, 7, 9):
            nNewMonth = 7
        else:
            nNewMonth = 10
        tFrom = datetime.datetime(nYear, nNewMonth, 1)
        tTo = tNow

    else:
        return xReturn  # Still set to None 'cause we didn't recognized the cRangeType value.
    return (tFrom, tTo)


def getCSSCalDateFormat(fmatType="MDYY"):
    """
    Converts one of the 4 EVOS/MPSS standard date format codes used in our apps to the appropriate code used
    by the NewCSSCal widget for interactive calendar and date/time services (using the customized version from
    Dec., 2017.)
    :param fmatType: Must be: MDYY, MDY, DMY, DMMY, or YYMD
    :return: Corresponding code for the NewCssCal widget.
    """
    cTestType = fmatType.strip()
    if not cTestType:
        cTestType = "MDYY"  # A default
    if cTestType == "MDYY":
        cRetType = "MMDDYYYY"
    elif cTestType == "MDY":
        cRetType = "MMDDYY"
    elif cTestType == "DMY":
        cRetType = "DDMMYY"
    elif cTestType == "DMYY":
        cRetType = "DDMMYYYY"
    else:
        cRetType = "MMDDYYYY"
    return cRetType


def getJQueryDateFormat(fmatType="MDYY"):
    """
    Returns a format string as used by the jQuery datepicker widget which can then adapt to the current standard
    date format in the PlanTools/BidTools contexts.
    :param fmatType: Must be one of the following: MDY, DMY, MDYY, DMYY, YYMD, DMMY
    :return: format string.  raises a ValueError if a bad string is passed.
    """
    if fmatType == "MDYY":
        cRet = "mm/dd/yy" # Note the jQuery formatter sees yy as "4-digit year" and y as "2-digit year"
    elif fmatType == "MDY":
        cRet = "mm/dd/y"
    elif fmatType == "DMYY":
        cRet = "dd/mm/yy"
    elif fmatType == "DMY":
        cRet = "dd/mm/y"
    elif fmatType == "DMMY":
        cRet = "d M yy"
    elif fmatType == "YYMD":
        cRet = "yy/mm/dd"
    else:
        raise ValueError("unrecognized date format code")
    return cRet


def makeNewCSSDateFormat(cEvosFmat="MDYY"):
    """
    For most calendar lookups in data entry, we use the www.rainforest.com NewCSSCal() function in
    datetimepicker3_css.js as modified by Evos/MPSS.  This calendar widget uses different date format
    codes from PlanTools/BidTools standard.  Pass a PlanTools/BidTools date format string (from the Configurator)
    to this method to convert it to a form recognized by NewCSSCal().
    :return: The CSS calendar required date format string.
    """
    if not cEvosFmat:
        cRetFmat = "mmddyy"

    elif cEvosFmat == "MDY":
        cRetFmat = "mmddyy"

    elif cEvosFmat == "MDYY":
        cRetFmat = "mmddyyyy"

    elif cEvosFmat == "DMY":
        cRetFmat = "ddmmyyyy"

    elif cEvosFmat == "DMYY":
        cRetFmat = "ddmmyyyy"

    elif cEvosFmat == "DMMY":
        cRetFmat = "ddmmmyyyy"

    else:
        cRetFmat = "mmddyyyy"
    return cRetFmat


def FormatDateString(lpdDate, fmatType="A", cSep="/"):
    """
    Takes a time() type date value or a datetime() type date (date or datetime) value and transforms it into
    a text string in one of six forms as specified by the type parm:
    A >> "Mar. 15, 2012"
    F >> "March 15, 2012"
    C >> "03/15/2012" # Note the US date notation
    M >> "15 Mar 2012"
    G >> "031512" # Poor practice but required for some external applications.  Uses US MDY sequence. Added 01/09/2014. JSH.
    R >> "20100315" simple YYYYMMDD format.

    Now recognizes the following standard PlanTools/BidTools date format codes. Added 08/23/2015. JSH.
    MDY = This is the 'American' standard, e.g. 03/28/07 -- deprecated
    MDYY = American, but 4-digit year: 03/28/2007 -- same as 'C' above
    DMY = European and Canadian standard: 28/03/07 -- deprecated
    DMYY = European standard but 4-digit year: 28/03/2007
    DMMY = "Military Date": 28 Mar 2007 -- same as 'M' above.
    YYMD = ISO type date like 2007/12/31

    For outputs with a separator character like "/", you can specify an alternative character like "-" or "."
    with the cSep parameter.
    """

    lbDatetime = False
    lbTime = False
    lcStr = "UNKNOWN DATE"
    xx = type(lpdDate)
    lcType = repr(type(lpdDate))
    lxParts = lcType.split("'")
    lcType = lxParts[1]
    lcVal1 = ""
    lcVal2 = ""
    lcVal3 = ""
    if lcType == 'float':
        ltDate = time.localtime(lpdDate)
        lbTime = True

    elif lcType == 'datetime':
        lbDatetime = True
        ltDate = lpdDate

    elif lcType == 'datetime.datetime':
        lbDatetime = True
        ltDate = lpdDate

    elif lcType == 'datetime.date':
        lbDatetime = True
        ltDate = lpdDate

    elif lcType == 'time.struct_time':
        ltDate = lpdDate
        lbTime = True
    else:
        ltDate = None

    lcPattern = ""
    cSepPattern = "%s" + cSep + "%s" + cSep + "%s"
    if lbTime or lbDatetime:
        if fmatType == 'A':
            lcPattern = "%s. %s, %s"
            if lbTime:
                lcVal1 = time.strftime("%b", ltDate)
                lcVal2 = time.strftime("%d", ltDate)
                lcVal3 = time.strftime("%Y", ltDate)
            elif lbDatetime:
                lcVal1 = ltDate.strftime("%b")
                lcVal2 = ltDate.strftime("%d")
                lcVal3 = ltDate.strftime("%Y")
            lnDay = int(lcVal2)
            lcVal2 = str(lnDay) # Gets rid of any leading zeros

        elif fmatType == 'F':
            lcPattern = "%s %s, %s"
            if lbTime:
                lcVal1 = time.strftime("%B", ltDate)
                lcVal2 = time.strftime("%d", ltDate)
                lcVal3 = time.strftime("%Y", ltDate)
            elif lbDatetime:
                lcVal1 = ltDate.strftime("%B")
                lcVal2 = ltDate.strftime("%d")
                lcVal3 = ltDate.strftime("%Y")
            lnDay = int(lcVal2)
            lcVal2 = str(lnDay) # Gets rid of any leading zeros

        elif fmatType == 'MDY':
            lcPattern = cSepPattern
            if lbTime:
                lcVal1 = time.strftime("%m", ltDate)
                lcVal2 = time.strftime("%d", ltDate)
                lcVal3 = time.strftime("%y", ltDate) # Two digit year.
            elif lbDatetime:
                lcVal1 = ltDate.strftime("%m")
                lcVal2 = ltDate.strftime("%d")
                lcVal3 = ltDate.strftime("%y")

        elif (fmatType == 'C') or (fmatType == "MDYY"):
            lcPattern = cSepPattern
            if lbTime:
                lcVal1 = time.strftime("%m", ltDate)
                lcVal2 = time.strftime("%d", ltDate)
                lcVal3 = time.strftime("%Y", ltDate)
            elif lbDatetime:
                lcVal1 = ltDate.strftime("%m")
                lcVal2 = ltDate.strftime("%d")
                lcVal3 = ltDate.strftime("%Y")
            # Here we DO want the leading zeros.

        elif fmatType == 'DMY':
            lcPattern = cSepPattern
            if lbTime:
                lcVal1 = time.strftime("%d", ltDate)
                lcVal2 = time.strftime("%m", ltDate)
                lcVal3 = time.strftime("%y", ltDate) # Two digit year.
            elif lbDatetime:
                lcVal1 = ltDate.strftime("%d")
                lcVal2 = ltDate.strftime("%m")
                lcVal3 = ltDate.strftime("%y")

        elif fmatType == "DMYY":
            lcPattern = cSepPattern
            if lbTime:
                lcVal1 = time.strftime("%d", ltDate)
                lcVal2 = time.strftime("%m", ltDate)
                lcVal3 = time.strftime("%Y", ltDate)
            elif lbDatetime:
                lcVal1 = ltDate.strftime("%d")
                lcVal2 = ltDate.strftime("%m")
                lcVal3 = ltDate.strftime("%Y")
            # Here we DO want the leading zeros.

        elif fmatType == 'G':
            lcPattern = "%s%s%s"
            if lbTime:
                lcVal1 = time.strftime("%m", ltDate)
                lcVal2 = time.strftime("%d", ltDate)
                lcVal3 = time.strftime("%y", ltDate) # only two digit year.
            elif lbDatetime:
                lcVal1 = ltDate.strftime("%m")
                lcVal2 = ltDate.strftime("%d")
                lcVal3 = ltDate.strftime("%y")
                # Here we DO want the leading zeros, but the year is just two digits..

        elif fmatType == 'R':
            lcPattern = "%s%s%s"
            if lbTime:
                lcVal2 = time.strftime("%m", ltDate)
                lcVal3 = time.strftime("%d", ltDate)
                lcVal1 = time.strftime("%Y", ltDate)
            elif lbDatetime:
                lcVal2 = ltDate.strftime("%m")
                lcVal3 = ltDate.strftime("%d")
                lcVal1 = ltDate.strftime("%Y")
                # Here we DO want the leading zeros, but the year is just two digits..

        elif (fmatType == 'M') or (fmatType == 'DMMY'):
            lcPattern = "%s %s %s"
            if lbTime:
                lcVal1 = time.strftime("%d", ltDate)
                lcVal2 = time.strftime("%b", ltDate)
                lcVal3 = time.strftime("%Y", ltDate)
            elif lbDatetime:
                lcVal1 = ltDate.strftime("%d")
                lcVal2 = ltDate.strftime("%b")
                lcVal3 = ltDate.strftime("%Y")
            lnDay = int(lcVal1)
            lcVal1 = str(lnDay) # Gets rid of any leading zeros

        elif fmatType == "YYMD":
            lcPattern = cSepPattern
            if lbTime:
                lcVal2 = time.strftime("%m", ltDate)
                lcVal3 = time.strftime("%d", ltDate)
                lcVal1 = time.strftime("%Y", ltDate)
            elif lbDatetime:
                lcVal2 = ltDate.strftime("%m")
                lcVal3 = ltDate.strftime("%d")
                lcVal1 = ltDate.strftime("%Y")
    if lcPattern != "":
        lcStr = lcPattern % (lcVal1, lcVal2, lcVal3)

    return lcStr


def STR(lpnVal, lpnLen=10, lpnDec=0):
    """
    Function that works like the VFP STR() function.  It transforms any numeric value into a string
    with the value right-aligned (left justified).  Default size is 10 characters and no decimal places,
    but other sizes and decimal places are allowed.  As with VFP STR() a non-numeric value for the first
    parameter triggers an exception.  This makes it easy to construct string values which match the STR()
    function values in a CodeBaseTools index.

    Like the VFP STR() function, if the value won't fit in the string length provided, it is converted
    to exponential notation, which may not be what you want!
    """
    if (not isinstance(lpnVal, int) and not isinstance(lpnVal, float) and not isinstance(lpnVal, long)) or not isinstance(lpnLen, int) or not isinstance(lpnDec, int):
        raise TypeError
    if lpnLen < 1:
        lpnLen = 10
    if lpnDec >= lpnLen:
        raise ValueError

    cFmat = "%" + str(lpnLen) + "." + str(lpnDec) + "f"
    cRet = cFmat%lpnVal
    if len(cRet) > lpnLen: # Can happen if thing is too large...
        cFmat = "%" + str(lpnLen) + "." + str(lpnDec) + "e"
        cRet = cFmat % lpnVal
    return cRet


def StrToHex(lpcString, lpcKey = ""):
    """
    Function that converts any text string to a string of hex values.  Each
    character of the string is converted to two characters of Hex numbers.
    Optionally accepts an encryption key to encrypt the result string
    """
    lcResult = ""
    lnPtr = 0
    lnKeyLen = 0
    if not lpcKey == "":
        lbTestKey = True
        lnKeyLen = len(lpcKey)
        lnPtr = 0
    else:
        lbTestKey = False

    for cc in lpcString:
        lnValue = ord(cc)
        if lbTestKey:
            lnValue = lnValue + ord(lpcKey[lnPtr])
            lnPtr += 1
            if lnPtr >= lnKeyLen:
                lnPtr = 0

        lcResult = lcResult + "%02X" % (lnValue)
    return lcResult


def HexToStr(lpcHex, lpcKey=""):
    """
    Takes a string coded as hexadecimal (one 2 character hex expression per output string character)
    and returns the ASCII text string decoded from it.  Optionally pass an encryption key
    to decrypt the result.
    """
    lcResult = ""
    lnPtr = 0
    if not lpcKey == "":
        lbTestKey = True
        lnKeyLen = len(lpcKey)
    else:
        lbTestKey = False
        lnKeyLen = 0
    for nn in range(0, len(lpcHex), 2):
        lcStr = lpcHex[nn:nn+2]
        lnValue = int(lcStr, 16) - (ord(lpcKey[lnPtr]) if lbTestKey else 0)
        lnPtr += 1
        if lnPtr >= lnKeyLen:
            lnPtr = 0
        lcResult = lcResult + (chr(lnValue) if ((lnValue <= 255) and (lnValue >= 0)) else "")
    return lcResult


def DELETEFILE(lpcFileName):
    """
    Clean removal of a single file named by the fully qualified name.
    Returns True if removal was successful, else False.  Does NOT generate an
    error like the native Python stuff does, as often you just don't care if it fails.
    Works more like VFP DELETE FILE command.

    08/22/2015 - Add ability to delete multiple files based on a wild-card type skeleton using glob().
    """
    global gcLastErrorMessage
    lbReturn = True

    if ("*" not in lpcFileName) and ("?" not in lpcFileName):
        try:
            os.remove(lpcFileName)
            # they aren't passing any wildcards.  Just one file to be deleted.
        except:
            gcLastErrorMessage = "DELETE FAILED: " + str(sys.exc_info())
            lbReturn = False
    else:
        # case of wildcards, so we have to use glob() to get rid of multiple files.
        xFiles = glob.glob(lpcFileName)
        if len(xFiles) > 0:
            for cF in xFiles:
                try:
                    os.remove(cF)
                except:
                    gcLastErrorMessage = "DELETE FAILED: " + str(sys.exc_info())
                    lbReturn = False  # but we still keep going.
        else:
            lbReturn = False  # didn't find any like that.
    return lbReturn


def JUSTFNAME(lpcFile):
    """
    pass any fully qualified file or path name and this returns just the file name itself, including the extension
    and the dot separator, if any.  Like the VFP JUSTFNAME() function
    :param lpcFile: The file name
    :return: Just the base file name
    """
    cReturn = ""
    if not isstr(lpcFile):
        raise ValueError("Must be a string file name")
    xTest = os.path.split(lpcFile)
    if len(xTest) == 2:
        cReturn = xTest[1]
    return cReturn


def JUSTSTEM(lpcFile):
    """
    Pass any fully or partially qualified file name and this returns just the base name.  So for example, if
    you pass c:\temp\myfile.txt, this will return the text string 'myfile'.  If, there is not a valid file name
    (doesn't have to exist, just be in a valid format), then returns the empty string.
    :param lpcFile: Any file name.
    :return: string.  See above.
    """
    cReturn = ""
    if not isstr(lpcFile):
        raise ValueError("Must be a string file name")
    xTest = os.path.split(lpcFile)
    cName = xTest[1]
    xTest = os.path.splitext(cName)
    if len(xTest) == 2:
        cReturn = xTest[0]
    return cReturn


def FORCEEXT(lpcFile, lpcExt):
    """
    Pass any file name as lpcFile and an extension (less any '.') in lpcExt and this
    returns the full file name with the extension set as the lpcExt value.  Like the VFP FORCEEXT() function.
    """
    return os.path.splitext(lpcFile)[0] + "." + lpcExt.replace('.', '')


def STREXTRACT(lpcSource, lpcDelim1, lpcDelim2, lpnCnt=1):
    """
    function like VFP STREXTRACT() which returns the text string contained between two delimiters within a
    source string.  Returns the lpnCnt-th instance.  Returns "" on not found.  If lpcDelim1 == "", then
    returns a string starting with the beginning of lpcSource.  If lpcDelim2 == "", then returns from the
    lpnCnt-th occurrance of lpcDelim1 to the end of lpcSource.
    *********** Initial version returns only the first instance.
    """
    if lpcDelim1 > "":
        lnPt1 = lpcSource.find(lpcDelim1)
        lnLen1 = len(lpcDelim1)
        if lnPt1 == -1:
            return ""
    else:
        lnPt1 = 0
        lnLen1 = 0
    if lpcDelim2 > "":
        lnPt2 = lpcSource[lnPt1 + lnLen1:].find(lpcDelim2)
        lnLen2 = len(lpcDelim2)
        if lnPt2 == -1:
            return ""
    else:
        lnPt2 = len(lpcSource)
        lnLen2 = 0
    return lpcSource[lnPt1 + lnLen1:(lnPt1 + lnLen1 + lnPt2)]


def clearHTMLCode(cSource):
    """
    Takes any string with some HTML tags and removes them leaving only the base code. Very simple logic looking
    for "<" and ">" tag delimiters.
    :param cSource: String to clean
    :return: cSource less the HTML markeup
    """
    cTestSource = cSource
    while True:
        cNewSource = RemoveTextBlock(cTestSource, "<", ">")
        if cNewSource == cTestSource:
            break
        if "<" not in cNewSource and ">" not in cNewSource:
            break
        cTestSource = cNewSource
    return cNewSource


def RemoveTextBlock(cSource="", cDelim1="", cDelim2=""):
    """
    Like the same named tool in MPSS/Evos MPSSLib.PRG.  It removes all text found in the cSource block between
    two non-identical delimiters and returns the result.  The delimiters are removed too.
    :param cSource:
    :param cDelim:
    :return: The resulting string
    """
    cFront = STREXTRACT(cSource, "", cDelim1, 1)
    cBack = STREXTRACT(cSource, cDelim2, "", 2)
    return cFront + cBack

# ********************************************************************************
# * Utility function that removes a block of text between two identical delimiters
# * and returns the block after the removal.  Takes out the delimiters too.
# FUNCTION RemoveTextBlock
# 	LPARAMETERS lpcText, lpcDelimiter
# 	LOCAL lcFront, lcBack
# 	lcFront = STREXTRACT(lpcText, "", lpcDelimiter, 1)
# 	lcBack = STREXTRACT(lpcText, lpcDelimiter, "", 2, 2)
# 	RETURN lcFront + lcBack
# ENDFUNC

def MPlocal_except_handler(cType, xValue, xTraceback, cErrorFilePath):
    lcBaseErrFile = "PYCOMERR_" + strRand(8, True) + ".ERR"
    lcErrFile = os.path.join(cErrorFilePath, lcBaseErrFile)
    lbFileOK = False
    loHndl = None
    try:
        loHndl = open(lcErrFile, "w", -1)
        lbFileOK = True
    except IOError:
        lbFileOK = False
    if not lbFileOK:
        loHndl = open(lcBaseErrFile, "w")
        # If the above fails, it's BAD

    loHndl.write("MPSS Python Program Error Report: ")
    chgTimex = time.localtime()
    lcChgTime = time.strftime("%Y-%m-%d %H:%M:%S", chgTimex)
    loHndl.write(lcChgTime + "\n\n")
    loHndl.write(repr(cType) + "\n")
    loHndl.write(repr(xValue) + "\n\nEXCEPTION TRACE:\n")

    traceback.print_exception(cType, xValue, xTraceback, file=loHndl)

    loHndl.write("\n\nVARIABLES DUMP:\n")
    workTraceback = xTraceback
    if workTraceback is None:
        return False

    xFrame = workTraceback.tb_frame
    loHndl.write("GLOBALS:\n")
    loHndl.write(MPshow_variables(xFrame.f_globals,
        ['__builtins__','os','pywin','time','traceback','win32com']))

    while workTraceback is not None:
        xFrame = workTraceback.tb_frame
        lcName = xFrame.f_code.co_filename
        loHndl.write("\n\nMODULE NAME  : " + lcName + "\n")
        lcObjName = xFrame.f_code.co_name
        loHndl.write(  "FUNCTION NAME: " + lcObjName + "\n")
        loHndl.write("LOCALS:\n")
        loHndl.write(MPshow_variables(xFrame.f_locals))
        workTraceback = workTraceback.tb_next

    loHndl.close()
    loHndl = None
    return True


def MPshow_variables(obj,ignore=[]):
    "helper function to show variables in order"
    if hasattr(obj,'__dict__'): obj = obj.__dict__
    keys = obj.keys()
    keys.sort()
    return '\n'.join(['%s = %s' % (k, repr(obj[k])) for k in keys if k not in ignore])


# ***************************************************
# * Function that will convert a Lat/Lon...Lat/Lon pair to great circle miles
# * Lats and Lons must be in RADIANS.
# * Info from Gary Plazyk. Oct. 18, 1996. JSH.
# SEE ALSO identical version in GeoDataTools.py
def rad2mile(oLat, oLon, dLat, dLon):
    nReturn = 0.0
    try:
        nReturn = 4000 * math.acos(math.sin(dLat)*math.sin(oLat)+math.cos(dLat)*math.cos(oLat)*math.cos(abs(dLon - oLon)))
    except:
        nReturn = 0.0
        # Usually because dLat==oLat and/or oLon==dLon
    return nReturn


###########################################################################
# This function converts the PyTime objects which Win32 COM objects receive
# from Visual FoxPro when dates or date/time values are passed, into a native
# Python datetime.datetime type object.  The Win32 extensions were written before
# there was a datetime object in Python.
def PyTime2datetime(ptDate):
    global gcLastErrorMessage
    gcLastErrorMessage = ""
    try:
        now_datetime = datetime.datetime(
            year=ptDate.year,
            month=ptDate.month,
            day=ptDate.day,
            hour=ptDate.hour,
            minute=ptDate.minute,
            second=ptDate.second)
    except:
        now_datetime = None
        # If this fails, there's not much we can do but return None.
        gcLastErrorMessage = "Invalid PyTime object."
    return now_datetime


######################################################################################
# Some PlanTools table store date/time information as a DATE type field plus a
# separate C(4) field for the time.  This logic was implemented before there was a
# DATETIME field in the DBF format.  This routine transforms that combo into a Python
# datetime.datetime() type object.
def dtFromDateAndTime(xDate, cTime):
    if xDate is None:
        return None
    if cTime is None or cTime == "":
        cTime = "0800"
    if cTime > "2359" or len(cTime) != 4:
        cTime = "0800"
    nHr = int(cTime[0:2])
    nMin = int(cTime[2:4])
    nSec = 0
    nYear = xDate.year
    nMonth = xDate.month
    nDay = xDate.day
    return datetime.datetime(nYear, nMonth, nDay, nHr, nMin, nSec)


####################################################################################
# Calculates the difference between two datetime.datetime() type objects in terms
# of the number of minutes.  If tDateTime2 is greater (later) than tDateTime1, this
# returns a positive number of minutes.  Otherwise returns a negative number of
# minutes or 0.  Up to YOU to make sure the parms are OK.
def diffInMinutes(tDateTime1, tDateTime2):
    xDiff = tDateTime2 - tDateTime1
    nRet = xDiff.days * 1440 + xDiff.seconds / 60
    return nRet


def dict2lower(xDict):
    """
    Takes a dictionary keyed by text strings and converts all those text strings to lower case
    :param xDict: source dict
    :return: copy of the source dict
    """
    xRet = dict()
    for cKey, xVal in xDict.items():
        xRet[cKey.lower()] = xVal
    return xRet


# #####################################################################################
# def copydatatable(cTableName="", cSourceDir="", cTargetDir="", bByZap=False, oCBT=None):
# moved to Codebasetools
#     """ Function which copies all files for one DBF type table from one directory to another.
#      Takes 5 parameters:
#      - Name of the table, no path, extension optional unless NOT .DBF
#      - Name of the source directory - fully qualified
#      - Name of the target directory - fully qualified
#      - By Zap, which causes table content to be removed and replaced
#      - oCBT which is a handle to a CodeBaseTools object, if an external one is to be used.
#
#      Returns True on success, False on failure.  If there is an existing table in the
#      target directory, it and its related files are renamed to a random name prior
#      to the copy process and then copied to an ARCHIVE subdirectory.  If failure,
#      the partially copied files are removed and the originals renamed to their starting names.
#      On returning False, writes the error into the gcLastErrorMessage global.
#      In that case call getlasterrormessage() to retrieve the description of the error.
#      As of 12/29/2013, this method will attempt to copy by zapping the contents of the
#      target table and appending in the source, if the bByZap parameter is True.  If failing
#      that or if bByZap is False, then copies using the operating system copy functions.
#      Caller should provide a CodeBaseTools object as oCBT or this program will create its own.
#      Note that if the source table is in a VFP database container, and if the function
#      uses the OS to copy the file, the resulting copy will have a reference to the database
#      container in it, which may not be what you intended.
#     """
#     global gcLastErrorMessage
#     gcLastErrorMessage = ""
#     cRoot, cExt = os.path.splitext(cTableName)
#     if cExt and cExt.upper() != ".DBF":
#         cTable = cTableName
#     else:
#         cTable = cRoot + ".DBF"
#         cExt = ".DBF"
#     cSourceDir = cSourceDir.strip()
#     cTargetDir = cTargetDir.strip()
#
#     cSrc = os.path.join(cSourceDir, cTable)
#     cTarg = os.path.join(cTargetDir, cTable)
#
#     if not os.path.exists(cSrc):
#         gcLastErrorMessage = "Source Table does not exist"
#         return False
#     cSrcCdx = FORCEEXT(cSrc, "CDX")
#     cTargCdx = FORCEEXT(cTarg, "CDX")
#     if not os.path.exists(cSrcCdx):
#         cSrcCdx = ""
#         cTargCdx = ""
#     cSrcFpt = FORCEEXT(cSrc, "FPT")
#     cTargFpt = FORCEEXT(cTarg, "FPT")
#     if not os.path.exists(cSrcFpt):
#         cSrcFpt = ""
#         cTargFpt = ""
#
#     bRenameRequired = False
#     cTempDbf = ""
#     cTempCdx = ""
#     cTempFpt = ""
#     bRenameOK = False
#     if os.path.exists(cTarg):
#         # We have to rename it to something else...
#         # Or copy it to some separate table IF it is to be zapped in place.
#         cTimeStamp = TTOC(datetime.datetime.now(), 1)
#         bRenameRequired = True
#         cTemp = cRoot + "_" + cTimeStamp + "_" + strRand(4, isPassword=1)
#         cTempDbf = cTemp + cExt
#         cTempDbf = os.path.join(cTargetDir, cTempDbf)
#         if cTargCdx:
#             cTempCdx = FORCEEXT(cTempDbf, "CDX")
#         if cTargFpt:
#             cTempFpt = FORCEEXT(cTempDbf, "FPT")
#
#         bRenameOK = True
#         xProc = os.rename
#         if bByZap:
#             xProc = shutil.copy2
#         try:
#             xProc(cTarg, cTempDbf)
#         except:
#             bRenameOK = False
#         if bRenameOK and cTargCdx:
#             try:
#                 xProc(cTargCdx, cTempCdx)
#             except:
#                 bRenameOK = False
#             if not bRenameOK:
#                 try:
#                     xProc(cTempDbf, cTarg)
#                 except:
#                     bRenameOK = False  # Duhhh...
#         if bRenameOK and cTargFpt:
#             try:
#                 xProc(cTargFpt, cTempFpt)
#             except:
#                 bRenameOK = False
#             if not bRenameOK:
#                 try:
#                     xProc(cTempDbf, cTarg)
#                     if cTargCdx:
#                         xProc(cTempCdx, cTargCdx)
#                 except:
#                     bRenameOK = False
#     if bRenameRequired and not bRenameOK:
#         gcLastErrorMessage = "Existing Target Table could not be renamed or archived for backup"
#         return False
#
#     bCopyOK = True
#     if bByZap:
#         if oCBT is None:
#             oCBT = MPSSCommon.CodeBaseTools.cbTools()
#         oCBT.setdeleted(True)
#         bCopyOK = oCBT.use(cSrc, "THESOURCE", readOnly=True)
#         if bCopyOK:
#             bCopyOK = oCBT.use(cTarg, "THETARG", exclusive=True)
#         if bCopyOK:
#             bCopyOK = oCBT.select("THETARG")
#         if bCopyOK:
#             bCopyOK = oCBT.zap()
#         if bCopyOK:
#             oCBT.select("THESOURCE")
#             for xRec in oCBT.scan(noData=True):
#                 cData = oCBT.scatterraw(alias="THESOURCE", converttypes=False, stripblanks=False, lcDelimiter="~!~")
#                 bCopyOK = oCBT.insertintotable(cData, lcAlias="THETARG", lcDelimiter="~!~")
#                 if not bCopyOK:
#                     break
#         if not bCopyOK:
#             gcLastErrorMessage = "Codebase ERR: " + oCBT.cErrorMessage
#         oCBT.closetable("THETARG")
#         oCBT.closetable("THESOURCE")
#         oCBT = None
#     else:
#         try:
#             shutil.copy2(cSrc, cTarg)
#         except:
#             bCopyOK = False
#             # lcType, lcValue, lxTrace = sys.exc_info()
#             # gcLastErrorMessage = "Err: " + str(lcType) + " " + str(lcValue)
#             gcLastErrorMessage = "shutil() copy Failure from: " + cSrc + " To " + cTarg
#
#         if bCopyOK and cSrcCdx:
#             try:
#                 shutil.copy2(cSrcCdx, cTargCdx)
#             except:
#                 bCopyOK = False
#                 # lcType, lcValue, lxTrace = sys.exc_info()
#                 # gcLastErrorMessage = "Err: " + str(lcType) + " " + str(lcValue)
#                 gcLastErrorMessage = "shutil() copy Failure from: " + cSrcCdx + " to " + cTargCdx
#         if bCopyOK and cSrcFpt:
#             try:
#                 shutil.copy2(cSrcFpt, cTargFpt)
#             except:
#                 bCopyOK = False
#                 # lcType, lcValue, lxTrace = sys.exc_info()
#                 # gcLastErrorMessage = "Err: " + str(lcType) + " " + str(lcValue)
#                 gcLastErrorMessage = "shutil() copy Failure from: " + cSrcFpt + " to " + cTargFpt
#
#     if not bCopyOK:
#         DELETEFILE(cTarg)
#         DELETEFILE(cTargCdx)
#         DELETEFILE(cTargFpt)  # These are resilient and don't fail if the file isn't there.
#         if bRenameRequired:
#             try:
#                 os.rename(cTempDbf, cTarg)
#                 if cTargCdx:
#                     os.rename(cTempCdx, cTargCdx)
#                 if cTargFpt:
#                     os.rename(cTempFpt, cTargFpt)
#             except:
#                 bCopyOK = False  # of course
#         gcLastErrorMessage = "Copy Failed, Target Table restored: " + gcLastErrorMessage
#     else:
#         # we try archiving the replaced tables, if any, to an ARCHIVE directory under the specified target...
#         if bRenameRequired:
#             cArchiveDir = os.path.join(cTargetDir, "ARCHIVE")
#             bOkToArchive = True
#             if not os.path.exists(cArchiveDir):
#                 try:
#                     os.mkdir(cArchiveDir)
#                 except:
#                     bOkToArchive = False
#             cPath, cBaseTempDbf = os.path.split(cTempDbf)
#             cBackupDbf = os.path.join(cArchiveDir, cBaseTempDbf)
#             try:
#                 shutil.move(cTempDbf, cBackupDbf)
#                 if cTargCdx:
#                     cBackupCdx = FORCEEXT(cBackupDbf, "CDX")
#                     shutil.move(cTempCdx, cBackupCdx)
#                 if cTargFpt:
#                     cBackupFpt = FORCEEXT(cBackupDbf, "FPT")
#                     shutil.move(cTempFpt, cBackupFpt)
#             except:
#                 gcLastErrorMessage = "Moving Backup Files to Archive Failed"
#     return bCopyOK


def getlasterrormessage():
    global gcLastErrorMessage
    return gcLastErrorMessage


#
# bitfield manipulation -- From ActiveState
#
class bf(object):
    """
    The purpose of this class is to provide the bit twiddling functionality of the
    BITSET() functions in VFP.  It allows you to set individual bit values in an unsigned
    long integer and retrieve those values.  The result can be turned into an int at any time.
    You treat this like a list, setting the nth element of the list to 0 (0 bit) or any non-zero value
    (set the bit value as 1).  Standard slicing also works.
    Example:
    myBitField = bf() # Starts out with value of 0 in all 32 bits.
    myBitField[0] = 1 # Sets the 0th bit to 1
    myBitField[7] = 1 # Sets the 7th (actually, the physical 8th bit) to 1
    nVal = int(myBitField) # returns 129L (unsigned long integer)
    nVal2 = int(nVal) # returns signed integer 129

    """
    def __init__(self,value=0):
        self._d = value

    def __getitem__(self, index):
        return (self._d >> index) & 1

    def __setitem__(self, index, value):
        value = (value & 1) << index
        mask = 1 << index
        self._d = (self._d & ~mask) | value

    def __getslice__(self, start, end):
        mask = 2**(end - start) - 1
        return (self._d >> start) & mask

    def __setslice__(self, start, end, value):
        mask = 2**(end - start) - 1
        value = (value & mask) << start
        mask = mask << start
        self._d = (self._d & ~mask) | value
        return (self._d >> start) & mask

    def __int__(self):
        return self._d

    def reset(self, value=0):
        # Added this for extra convenience in updating the value as needed by the calling program.
        self._d = value
        return self._d

    def __str__(self):
        return str(self._d)


############################################################################
# This function returns the text string defining a MIME type based on the
# extension passed as the parameter.  This is used primarily in this system
# to populate the Content-Type field of the HTML document header for transmitting
# a file to be downloaded by a browser.  Setting the Content-Type correctly is
# helpful in telling the client's browser what application to use to open the file.
def getMIMEType(cExt=""):
    global gxMIMEtypes
    cTest = cExt.replace(".", "").upper()
    return gxMIMEtypes.get(cTest, "")


# ************************************************************************************8
# * Function that tests address information and returns .T. if it is in Canada
# * 11/13/2000 5:21PM - Added. JSH.
# Converted from VFP, March 5, 2015. JSH.
def IsCanada(cZip="", cProvince="", cCountry=""):
    if "CA" in cCountry.upper():
        return True
    if "US" in cCountry.upper():
        return False

    if len(cProvince.strip()) == 2:
        if ("," + cProvince.upper().strip() + ",") in ",ON,MB,QC,PQ,MB,SK,AB,BC,NS,PE,NF,NU,YK,YU,NB,":
            return True

    cTemp = cZip.upper().replace(" ", "").strip()
    if len(cTemp) >= 5:  # have some part or all of the string
        if cTemp[0].isalpha and cTemp[2].isalpha and cTemp[4].isalpha():
            return True

    return False


def setCommonErrorPath(cPath):
    global gcErrorFilePath
    gcErrorFilePath = cPath
    return os.path.exists(cPath)


def handleError(func):
    """decorator that will wrap error detection around a method"""
    @wraps(func)  # Essential to put back the __name__ and other vars
    # to make sure that COM objects processed with this match the _public_methods_ listings.
    def method(*args, **kwargs):
        global gcErrorFilePath
        # "handle error with local_except_handler"
        try:
            result = func(*args, **kwargs)
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            MPlocal_except_handler(lcType, lcValue, lxTrace, gcErrorFilePath)
            raise COMException(description=("Python Error Occurred: " + repr(lcType)))
        return result
    return method


def removeNonPrint(xStr):
    """
    Pass this any string or bytearray and will remove all the ascii characters below the space.
    """
    if _ver3x:
        if isinstance(xStr, bytes) or isinstance(xStr, bytearray):
            xRet = xStr.translate(None, bytes(gcNonPrintables, "utf-8"))
        if isinstance(xStr, str):
            xRet = xStr.translate(gxNonPrintables)
    else:
        xRet = str(xStr).translate(None, gcNonPrintables)
    return xRet


def cleanFileName(cName):
    """
    Expects as a parm a base file name (no path) and an optional extension.  This typically would be a file
    name as supplied by a user who will put just about anything into a file name.  This removes non-printing
    characters and replaces illegal characters with spaces.
    :param cName: File Name
    :return: Cleaned name
    """
    cNameRet = removeNonPrint(cName)
    cNameRet = cNameRet.replace(" ", "_")
    cNameRet = cNameRet.replace(",", "_")
    cNameRet = cNameRet.replace("*", "_")
    cNameRet = cNameRet.replace("^", "_")
    cNameRet = cNameRet.replace("&", "_")
    cNameRet = cNameRet.replace(":", "_")
    return cNameRet


def MPMsgBox(cMsg):
    ########################################################################################
    # Utility messagebox popup function to interrupt processing during debug.
    import easygui  # on demand.  should only be used in interactive mode
    easygui.msgbox(cMsg, "MPSS Development")
    return


def MPConfirm(cMsg):
    """
    Like MPMsgBox except expects a "Yes" or "No" answer
    :param cMsg:
    :return: Returns True for Yes, otherwise False
    """
    import easygui
    nResult = easygui.boolbox(cMsg, "MPSS Administration")
    if nResult == 1:
        return True
    return False


def multiItemFieldToList(cValue):
    """
    In DBF files where multiple values are stored in a Memo field either as json, comma delimited or line delmited
    values, this function is designed to extract that information into a Python list().
    :param cValue: Stored field data string
    :return: list of the values.  Returns an empty list of there is nothing there.  None on error.
    """
    xRet = None
    if cValue:
        try:
            if "]" in cValue:  # Probably JSON
                xRet = json.loads(cValue)
            elif "\r\n" in cValue:  # Line delimited
                xRet = cValue.split("\r\n")
            else:  # Assume comma delimited
                xRet = cValue.split(",")
        except:
            xRet = None
        return xRet
    else:
        return list()


def runPythonApp(cApp="", cParmString="", cHomeDir=""):
    """
    Determines where python.exe is and then runs it with the specified .py program and parms.  Looks first for
    a PYTHONPATH environment variable.  If that is empty, then determines the current Python version and tries
    either c:\python27 or c:\python36 .  If it can't find the app there, it will look at sys.executable to see
    if it contains "python.exe".  (Note that COM objects may report other .exe files.)  If that fails then
    raises an error.  Note we don't use os.environ because it doesn't reflect the most recent system environment
    variable updates.
    :param cApp: fully qualified path name to the .py file
    :param cParmString: any command line parms to be passed to the .py module
    :param cHomeDir: If not to run in the current directory, specify full path where it should be run.
    :return: True on OK
    """
    bRet = True
    cPath = ""
    cExe = ""
    rKey = None
    try:
        rPath = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        rKey = xWinReg.OpenKey(xWinReg.HKEY_LOCAL_MACHINE, rPath, 0, xWinReg.KEY_READ)
        xval = xWinReg.QueryValueEx(rKey, "PYTHONPATH")
        cPath = str(xval[0])
        xWinReg.CloseKey(rKey)
        rKey = None
    except WindowsError:
        cPath = ""
    except:
        cPath = None
    finally:
        if rKey is not None:
            xWinReg.CloseKey(rKey)
    if cPath:
        cExe = os.path.join(cPath, "python.exe")
        if not os.path.exists(cExe):
            cExe = ""
    if not cExe:
        cExe = sys.executable
        if "\\PYTHON.EXE" not in cExe.upper():
            cExe = ""
    if not cExe:
        if _ver3x:
            cExe = "c:\\python36\\python.exe"
        else:
            cExe = "c:\\python27\\python.exe"
        if not os.path.exists(cExe):
            raise ValueError("Python executable not found")
    cCommand = "%s %s" % (cApp, cParmString)
    win32api.ShellExecute(0, "open", cExe, cCommand, cHomeDir, win32con.SW_SHOWNORMAL)

    return bRet

#     ****************************************************************************************
# * Function that parses out the accessorial codes from the ACCESSREQD field in the
# * ShipMstr and related tables.  This field can have either a JSON, comma delimited, or
# * cr/lf delimited list of standard Evos Accessorial codes.  Returns the number of
# * items found.  If non-zero count, then lpaRetArray on return will refer to an
# * array passed by reference.
# FUNCTION accessListToArray
# LPARAMETERS lpcList, lpaRetArray
# 	EXTERNAL ARRAY lpaRetArray
# 	LOCAL lnCount, cTest
# 	LOCAL ARRAY aAccess[39]
# 	lnCount = 0
# 	IF !EMPTY(lpcList)
# 		DO CASE
# 			CASE "]" $ lpcList
# 				&& This is JSON
# 				cTest = STRTRAN(lpcList, "[", "")
# 				cTest = STRTRAN(cTest, "]", "")
# 				cTest = STRTRAN(cTest, '"', '')
# 				cTest = STRTRAN(cTest, " ", "")
# 				lnCount = ALINES(aAccess, cTest, 1 + 4, ",")
#
# 			CASE (CHR(13) + CHR(10)) $ lpcList
# 				&& This is CR/LF delimited
# 				cTest = STRTRAN(lpcList, " ", "")
# 				cTest = STRTRAN(cTest, '"', '')
# 				lnCount = ALINES(aAccess, cTest, 1 + 4, CHR(13) + CHR(10))
#
# 			OTHERWISE
# 				&& Probably just comma delimited
# 				cTest = STRTRAN(lpcList, " ", "")
# 				cTest = STRTRAN(cTest, '"', '')
# 				lnCount = ALINES(aAccess, cTest, 1 + 4, ",")
#
# 		ENDCASE
# 		IF lnCount > 0
# 			DIMENSION lpaRetArray[lnCount]
# 			ACOPY(aAccess, lpaRetArray)
# 		ENDIF
# 	ENDIF
# 	RETURN lnCount
# ENDFUNC

def getfilechanged(cFile=""):
    """
    Simple utility that returns the datetime value of the last modification date for any file.
    :param cFile: fully qualified path name for the file in question
    :return: datetime.datetime value of the last modification date or None if file is not found or otherwise
    not accessible.
    """
    global gcLastErrorMessage
    tRet = None
    if cFile:
        try:
            xStat = os.stat(cFile)
        except WindowsError:
            gcLastErrorMessage = "FILE NOT FOUND"
            xStat = None
        except:
            gcLastErrorMessage = "FILE " + cFile + " not accessible"
            xStat = None
        if xStat is not None:
            tRet = datetime.datetime.fromtimestamp(xStat.st_mtime)
    return tRet


class baseService(object):
    def __init__(self, bIsCOM=False):
        # This wrapper stuff is ONLY for when the module is used as a COM-type service object
        # and not a direct call by a Python client, which normally will provide its own error
        # handling capability.
        if bIsCOM:
            cCode = ""
            xStuff = dir(self)
            # This wraps every non-private class method in the handleError
            # decorator.
            for xc in xStuff:
                xThing = eval('self.' + xc)
                if callable(xThing) and not xc.startswith("_"):
                    cCode = "self." + xc + "=handleError(self." + xc + ")"
                    exec(cCode)

    def resetConfig(self, oNewCfg=None, oNewVfp=None):
        return True


__all__ = ["baseService", "bf", "Bunch", "ADDBS", "ADDFS", "bytes2str", "cleanFileName", "clearHTMLCode",
           "cleanUnicodeString", "confirmPath", "convertUnits", "CTOD", "CTOT", "CTOTX", "deleteByAge",
           "DELETEFILE", "dict2lower", "diffInMinutes", "DOW", "dtFromDateAndTime", "DTOR", "DTOT", "fDiv",
           "FILETOSTR", "findAllFiles", "FORCEEXT", "FormatDateString", "FormatDateTimeString", "FSacquire",
           "get", "getconfig", "getCSSCalDateFormat", "getCustomDateRange", "getDateRangeTypeText", "getExcelColumn",
           "getfilechanged", "getJQueryDateFormat", "getlasterrormessage", "getMIMEType", "getrandomseed",
           "getTempFileName", "handleError", "HexToStr", "isBetween", "IsCanada", "isDateBetween", "isDigits",
           "isstr", "isunicode", "IsValidNumber", "JUSTFNAME", "JUSTSTEM", "makeNewCSSDateFormat", "MPConfirm",
           "MPlocal_except_handler", "MPMsgBox", "MPshow_variables", "multiItemFieldToList", "NamedTuple",
           "PADRIGHT", "PyTime2datetime", "rad2mile", "rawReplace", "removeNonPrint", "RemoveTextBlock", "RIGHT",
           "RTOD", "runPythonApp", "safeHTMLcoding", "sanitizeHtml", "setCommonErrorPath", "StartOfWeek", "STR",
           "STREXTRACT", "strRand", "STRTOFILE", "StrToHex", "TestTimeEntry", "thispath", "toAscii", "TTOC",
           "TTOD", "WEEK"]

###########################################################################
# Testing section...


if __name__ == "__main__":
    print("strRand in Normal Mode - 25 chars")
    print(strRand(25, 0))
    print("strRand in Password Mode - 12 chars")
    print(strRand(12, 1))
    print("Testing isBetween for 99, 66, 135")
    print(isBetween(99, 66, 135))
    print("Testing isBetween for 44, 66, 135")
    print(isBetween(44, 66, 135))
    print("Testing getExcelColumn for 12")
    print(getExcelColumn(12))
    print("Testing getExcelColumn with 39")
    print(getExcelColumn(39))
    print("Testing getExcelColumn with 69")
    print(getExcelColumn(69))
    print("Testing getExcelColumn with 193")
    print(getExcelColumn(193))
    print("Testing confirmPath with c:\\temp")
    bIsPath = confirmPath("c:\\temp")
    print(bIsPath)
    if not bIsPath:
        print("Creating c:\\temp path...")
        os.mkdir("c:\\temp")
    print("Testing confirmPath with x:\\glop")
    print(confirmPath("x:\\glop"))
    print("Testing ADDBS")
    print(ADDBS("c:\somedirectory"))
    print("Testing STRTOFILE and FILETOSTR")
    print(STRTOFILE("This is a test", "c:\\temp\\testtextfile.txt"))
    cResult = FILETOSTR("c:\\temp\\testtextfile.txt")
    print("RAW RESULT:", cResult)  # Should return a bytes object
    print("UNICODE RESULT", cResult.decode("cp1252", "ignore"))  # transform the Windows Code Page 1252 string
    # into a unicode string.

    ltTime = time.time()
    print(ltTime)
    lc = FormatDateString(ltTime)
    print(lc)
    lc = FormatDateString(ltTime, 'M')
    print(lc)
    lc = FormatDateString(ltTime, "F")
    print(lc)
    lc = FormatDateString(ltTime, "C")
    print(lc)
    lc = FormatDateTimeString(ltTime, "M", "12")
    print(lc)
    lc = FormatDateTimeString(ltTime, "F", "24")
    print(lc)
    lc = FormatDateTimeString(ltTime, "C", "24S")
    print(lc)
    lx = StrToHex("Vfp34234981", "wewewe")
    print("RESULT: ", lx, HexToStr(lx, "wewewe"))
    lx = FORCEEXT("c:\\temp\\myfile.html", ".dbf")
    print(lx)
    print("NOW", TTOC(datetime.datetime.now()))
    print("NOW 1", TTOC(datetime.datetime.now(), 1))
    print("NOW 2", TTOC(datetime.datetime.now(), 2))
    print("NOW 3", TTOC(datetime.datetime.now(), 3))
    print("NOW 4", TTOC(datetime.datetime.now(), 9))
    xx = datetime.datetime.now()
    yy = datetime.datetime(2014, 1, 1, 0, 0, 0)
    zz = datetime.datetime(2011, 1, 1, 0, 0, 0)
    print(isDateBetween(xx, zz, yy))
    print(isDateBetween(yy, zz, xx))
    print(isDateBetween(zz, xx, yy))
    print(strRand(8))
    print(STR(3))
    print(STR(993.53, 12, 5))
    print(STR(123409883984))
    print(STR(14098309884, 15, 3))
    print(FormatDateString(datetime.date.today(), "R"))
    print(IsCanada("H3N 4K3"))
    print(CTOD("8 Mar 1943", "DMMY"))
    print(CTOD("08/19/2014", "mm-dd-yy"))
    print(CTOD("09/12/2014", "mm/dd/yyyy"))
    print(getJQueryDateFormat("DMMY"))
    print(JUSTSTEM(r"c:\temp\myfile.txt"))
    print(JUSTSTEM(r"qewrqwerqwer"))
    print(JUSTSTEM(r"blather.34i"))
    # try:
    #     x = 2
    #     y = 0
    #     z = x / y
    # except:
    #     var1, var2, var3 = sys.exc_info()
    #     MPlocal_except_handler(var1, var2, var3, "c:\\temp")
    #     print("Error Occured")
    # testList = findAllFiles("c:\\mpsspythonscripts\\loadoptweb", cSkel="*.py", cExcl="__init__.*")
    # if testList is not None:
    #     for tL in testList:
    #         print(tL)
    # else:
    #     print("Bad file path skeleton")

    print("Test Time Entry for 2400 %s" % TestTimeEntry("2400"))
    print("Test Time Entry for 800 %s" % TestTimeEntry("800"))
    print("Test Time Entry for qwer %s" % str(TestTimeEntry("qwer")))
    print("Test Time Entry for 0800 %s" % TestTimeEntry("0800"))
    print(CTOT("08/01/2018 10:55AM", "MDY"))  # Note produces None -- not what is wanted.
    print(CTOD("08/01/2018 10:55AM", "MDYY"))
    print(CTOTX("08/01/2018 10:55AM", "MDY"))  # Use CTOTX() instead.
