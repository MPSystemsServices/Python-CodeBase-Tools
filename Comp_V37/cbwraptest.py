from __future__ import print_function, absolute_import

from ctypes import *
import os
from time import time
from datetime import date, datetime
import itertools
import MPSSCommon.MPSSBaseTools as mTools


class VFPFIELD(Structure):
    _fields_ = [("cName", c_char * 132),
                ("cType", c_char * 4),
                ("nWidth", c_int),
                ("nDecimals", c_int),
                ("bNulls", c_int)]


class VFPINDEXTAG(Structure):
    _fields_ = [("cTagName", c_char * 130),
                ("cTagExpr", c_char * 256),
                ("cTagFilt", c_char * 256),
                ("nDirection", c_int),
                ("nUnique", c_int)]

cbWrap = CDLL("CodeBaseWrapper.dll")

cbWrap.cbwInit.argtypes = []
cbWrap.cbwInit.restype = c_int

cbWrap.cbwShutDown.argtypes = []
cbWrap.cbwShutDown.restype = c_int

cbWrap.cbwUSE.restype = c_int
cbWrap.cbwUSE.argtypes = [c_char_p, c_char_p, c_int]

cbWrap.cbwUSEEXCL.restype = c_int
cbWrap.cbwUSEEXCL.argtypes = [c_char_p, c_char_p]

cbWrap.cbwCLOSEDATABASES.restype = c_int
cbWrap.cbwCLOSEDATABASES.argtypes = []

cbWrap.cbwALIAS.restype = c_char_p
cbWrap.cbwALIAS.argtypes = []

cbWrap.cbwERRORNUM.restype = c_int
cbWrap.cbwERRORNUM.argtypes = []

cbWrap.cbwSELECT.restype = c_int
cbWrap.cbwSELECT.argtypes = [c_char_p]

cbWrap.cbwCLOSETABLE.restype = c_int
cbWrap.cbwCLOSETABLE.argtypes = [c_char_p]

cbWrap.cbwERRORMSG.restype = c_char_p
cbWrap.cbwERRORMSG.argtypes = []

cbWrap.cbwDBF.restype = c_char_p
cbWrap.cbwDBF.argtypes = [c_char_p]

cbWrap.cbwAFIELDS.restype = POINTER(VFPFIELD)
cbWrap.cbwAFIELDS.argtypes = [POINTER(c_int)]

cbWrap.cbwATAGINFO.restype = POINTER(VFPINDEXTAG)
cbWrap.cbwATAGINFO.argtypes = [POINTER(c_int)]

cbWrap.cbwREINDEX.restype = c_int
cbWrap.cbwREINDEX.argtypes = []

cbWrap.cbwISREADONLY.restype = c_int
cbWrap.cbwISREADONLY.argtypes = []

cbWrap.cbwPACK.restype = c_int
cbWrap.cbwPACK.argtypes = []

cbWrap.cbwRECCOUNT.restype = c_int
cbWrap.cbwRECCOUNT.argtypes = []

cbWrap.cbwRECNO.restype = c_int
cbWrap.cbwRECNO.argtypes = []

cbWrap.cbwZAP.restype = c_int
cbWrap.cbwZAP.argtypes = []

cbWrap.cbwEOF.restype = c_int
cbWrap.cbwEOF.argtypes = []

cbWrap.cbwBOF.restype = c_int
cbWrap.cbwBOF.argtypes = []

cbWrap.cbwRECALLALL.restype = c_int
cbWrap.cbwRECALLALL.argtypes = []

cbWrap.cbwRECALL.restype = c_int
cbWrap.cbwRECALL.argtypes = []

cbWrap.cbwDELETED.restype = c_int
cbWrap.cbwDELETED.argtypes = []

cbWrap.cbwTALLY.restype = c_int
cbWrap.cbwTALLY.argtypes = []

cbWrap.cbwLOCK.restype = c_int
cbWrap.cbwLOCK.argtypes = []

cbWrap.cbwDELETE.restype = c_int
cbWrap.cbwDELETE.argtypes = []

cbWrap.cbwSETORDERTO.restype = c_int
cbWrap.cbwSETORDERTO.argtypes = [c_char_p]

cbWrap.cbwSCATTER.argtypes = [c_char_p]
cbWrap.cbwSCATTER.restype = c_char_p

cbWrap.cbwSCATTERFIELD.argtypes = [c_char_p]
cbWrap.cbwSCATTERFIELD.restype  = c_char_p

cbWrap.cbwGOTO.argtypes = [c_char_p, c_int]
cbWrap.cbwGOTO.restype = c_int

cbWrap.cbwSEEK.restype = c_int
cbWrap.cbwSEEK.argtypes = [c_char_p, c_char_p, c_char_p]

cbWrap.cbwSCATTERFIELDLONG.argtypes = [c_char_p, POINTER(c_int)]
cbWrap.cbwSCATTERFIELDLONG.restype = c_int

cbWrap.cbwSCATTERFIELDDOUBLE.argtypes = [c_char_p, POINTER(c_double)]
cbWrap.cbwSCATTERFIELDDOUBLE.restype = c_int

cbWrap.cbwSCATTERFIELDDATE.argtypes = [c_char_p, POINTER(c_int), POINTER(c_int), POINTER(c_int)]
cbWrap.cbwSCATTERFIELDDATE.restype = c_int

cbWrap.cbwSCATTERFIELDDATETIME.restype = c_int
cbWrap.cbwSCATTERFIELDDATETIME.argtypes = [c_char_p, POINTER(c_int), POINTER(c_int), POINTER(c_int), POINTER(c_int), POINTER(c_int), POINTER(c_int)]

cbWrap.cbwSCATTERFIELDLOGICAL.argtypes = [c_char_p]
cbWrap.cbwSCATTERFIELDLOGICAL.restype = c_int

cbWrap.cbwUSED.restype = c_int
cbWrap.cbwUSED.argtypes = [c_char_p]

cbWrap.cbwEngineActive.restype = c_int
cbWrap.cbwEngineActive.argtypes = []

cbWrap.cbwFLUSH.restype = c_int
cbWrap.cbwFLUSH.argtypes = []

cbWrap.cbwFLUSHALL.restype = c_int
cbWrap.cbwFLUSHALL.argtypes = []

cbWrap.cbwFCOUNT.restype = c_int
cbWrap.cbwFCOUNT.argtypes = []

cbWrap.cbwAPPENDBLANK.restype = c_int
cbWrap.cbwAPPENDBLANK.argtypes = []

cbWrap.cbwGATHERMEMVAR.restype = c_int
cbWrap.cbwGATHERMEMVAR.argtypes = [c_char_p, c_char_p, c_char_p]

cbWrap.cbwINSERT.restype = c_int
cbWrap.cbwINSERT.argtypes = [c_char_p, c_char_p]

cbWrap.cbwTAGCOUNT.restype = c_int
cbWrap.cbwTAGCOUNT.argtypes = []

cbWrap.cbwORDER.restype = c_int
cbWrap.cbwORDER.argtypes = []

cbWrap.cbwINDEXON.restype = c_int
cbWrap.cbwINDEXON.argtypes = [c_char_p, c_char_p, c_char_p, c_int, c_int]

cbWrap.cbwREPLACE_DATETIMEC.restype = c_int
cbWrap.cbwREPLACE_DATETIMEC.argtypes = [c_char_p, c_char_p]

cbWrap.cbwREPLACE_DATETIMEN.restype = c_int
cbWrap.cbwREPLACE_DATETIMEN.argtypes = [c_char_p, c_int, c_int, c_int, c_int, c_int, c_int]

cbWrap.cbwREPLACE_DATEC.restype = c_int
cbWrap.cbwREPLACE_DATEC.argtypes = [c_char_p, c_char_p]

cbWrap.cbwREPLACE_DATEN.restype = c_int
cbWrap.cbwREPLACE_DATEN.argtypes = [c_char_p, c_int]

cbWrap.cbwCREATETABLE.restype = c_int
cbWrap.cbwCREATETABLE.argtypes = [c_char_p, c_char_p]

lnResult = cbWrap.cbwInit()
print(lnResult)
print("Creating employees")

lcFldSpec = \
"""firstname,C,25,0,FALSE
lastname,C,25,0,FALSE
city,C,35,0,FALSE
state,C,2,0,FALSE
zipcode,C,10,0,FALSE
hire_date,D,8,0,FALSE
login_date,T,8,0,TRUE
salary,N,10,2,FALSE
ssn,C,10,0,FALSE
married,L,1,0,FALSE
dob,D,8,0,TRUE
"""
lpFields = c_char_p(lcFldSpec)
lnStart = time()
lnResult = cbWrap.cbwCREATETABLE("employees", lpFields)
print(lnResult)
lnEnd = time()
print("time:", lnEnd - lnStart)
print(cbWrap.cbwERRORMSG())
lnResult = cbWrap.cbwINDEXON("firstname", "firstname", "", 0, 0)
print(lnResult)
lnResult = cbWrap.cbwAPPENDBLANK()
print(lnResult)
lnResult = cbWrap.cbwREPLACE_CHARACTER("firstname", "charlie")
print("Replacing: ", lnResult)

lnResult = cbWrap.cbwUSEEXCL("data2.dbf", "DATA2")
print("DATA2 Exclusive Open: ", lnResult)

##print "testing zippairs"
##lnResult = cbWrap.cbwUSE("e:\\loadbuilder2\\appdbfs\\zippairs.dbf", "zippairs", 0)
##lnResult = cbWrap.cbwSEEK("62305     19975     ", "zippairs", "main")
##lnResult = cbWrap.cbwSELECT("zippdirs")
##lnResult = cbWrap.cbwRECNO()
##print lnResult

lnResult = cbWrap.cbwUSE("E:\\LoadBuilder2\\AppDBFS\\FACLOGPR\\shipmstr.dbf", "shipmstr", 0)
print(lnResult)
lnResult = cbWrap.cbwGOTO("RECORD", 55)
print(lnResult)
lnResult = cbWrap.cbwTAGCOUNT()
print("shipmaster tags", lnResult)

lnCount = c_int()
lpPtr = pointer(lnCount)
laTagList = cbWrap.cbwATAGINFO(lpPtr)
print("Shipmstr TagCount", lnCount.value)
for jj in range(0, lnCount.value):
    print(laTagList[jj].cTagName, laTagList[jj].cTagExpr, jj)

xField = VFPFIELD()
xFldList = [VFPFIELD(), VFPFIELD(), VFPFIELD()]
print("Type checking: ", isinstance(xField, VFPFIELD))
print("List checking: ", isinstance(xFldList, list))
print("list type chk: ", isinstance(xFldList[0], VFPFIELD))
xStr = "This is a test string"
# print("string Typeck: ", isinstance(xStr, basestring))
print("string Typeck: ", mTools.isstr(xStr))

##lnResult = cbWrap.cbwSELECT("DATA2")
##lnResult = cbWrap.cbwRECNO()
##print "DATA 2 records: ", lnResult
##
##lnResult = cbWrap.cbwTAGCOUNT()
##print "data2 TAGS ", lnResult
##
##lcPath = cbWrap.cbwDBF("DATA2")
##print "Data2 Path: ", lcPath + "<<"
##
##print "read only test: "
##print cbWrap.cbwISREADONLY()
##
##lcAlias = cbWrap.cbwALIAS()
##print lcAlias, " data2 alias?"
##
##lnResult = cbWrap.cbwSELECT("DATA2")
##lnFldCnt = cbWrap.cbwFCOUNT()
##print lnFldCnt, " data2 field count"
##
##lpName = c_char_p("L_NAME")
##lpExpr = c_char_p("UPPER(L_NAME+F_NAME)")
##lpFilt = c_char_p("")
##
##lnResult = cbWrap.cbwINDEXON(lpName, lpExpr, lpFilt, 0, 0)
##print lnResult
##print cbWrap.cbwERRORMSG(), " << ERR message"
##print cbWrap.cbwERRORNUM(), " << ERR number"
##
##lnResult = cbWrap.cbwSELECT("DATA2")
##lnResult = cbWrap.cbwGOTO("RECORD", 1)

#lnResult = cbWrap.cbwREPLACE_DATETIMEC("loaddt", "20110525173502")
#lnResult = cbWrap.cbwREPLACE_DATETIMEN("loaddt", 1998, 4, 2, 17, 3, 23)
##lnResult = cbWrap.cbwREPLACE_DATEN("birth_date", 2450000)
##print lnResult
##print cbWrap.cbwERRORMSG()
##
##print lnResult, " record move"
##lnxYear = c_int()
##lnxMonth= c_int()
##lnxDay = c_int()
##lnpYear = pointer(lnxYear)
##lnpMonth = pointer(lnxMonth)
##lnpDay = pointer(lnxDay)
##lnxHour = c_int()
##lnxMinute = c_int()
##lnxSecond = c_int()
##lnpHour = pointer(lnxHour)
##lnpMinute = pointer(lnxMinute)
##lnpSecond = pointer(lnxSecond)
##lnResult = cbWrap.cbwSCATTERFIELDDATE("loaddt", lnpYear, lnpMonth, lnpDay)
##print lnResult
##lnYear = lnpYear.contents.value
##lnMonth = lnpMonth.contents.value
##lnDay = lnpDay.contents.value
##print "converting the dates"
##print lnYear, lnMonth, lnDay
##lnTheDate = date(lnYear, lnMonth, lnDay)
##print lnTheDate
##lnTheDateTime = datetime(lnYear, lnMonth, lnDay, 0, 0, 0)
##print lnTheDateTime
##lnResult = cbWrap.cbwSCATTERFIELDDATETIME("loaddt", lnpYear, lnpMonth, lnpDay, lnpHour, lnpMinute, lnpSecond)
##lnYear = lnpYear.contents.value
##lnMonth = lnpMonth.contents.value
##lnDay = lnpDay.contents.value
##lnHour = lnpHour.contents.value
##lnMinute = lnpMinute.contents.value
##lnSecond = lnpSecond.contents.value
##lnTheDateTime = datetime(lnYear, lnMonth, lnDay, lnHour, lnMinute, lnSecond)
##print lnTheDateTime

##lcValues = "George~~Smith~~9821 SE Burnside St~~95~~19151203~~F~~59.66~~NO COMMENT~~20050428093022~~5983408~~354.33~~34.123~~395.45"
##lcDelim = "~~"
##lpVal = c_char_p(lcValues)
##lpDel = c_char_p(lcDelim)
##lnResult = cbWrap.cbwINSERT(lpVal, lpDel)
##print lnResult
lnCount = c_int(0)
lpPtr = pointer(lnCount)
laFields = cbWrap.cbwAFIELDS(lpPtr)
# print lnCount.value
# print "fields done"
# print laFields[0].cName
# lxList = list()
# for jj in range(0, lnCount.value):
#     lxList.append(laFields[jj])
# print type(lxList)
# print type(lxList[0])

#
# lnResult = cbWrap.cbwGOTO("RECORD", 2)
# print lnResult
# if lnResult > 0 :
#     lcFldStr = cbWrap.cbwSCATTER("&$&")
#     print lcFldStr
#     laFldsx = lcFldStr.split("&$&")
#     print laFldsx[2]

# cbWrap.cbwSELECT("zippairs")
# cbWrap.cbwSETORDERTO("")
# cbWrap.cbwGOTO("TOP", 0)
# cbWrap.cbwGOTO("SKIP", 1000)
#
# lcStr = cbWrap.cbwSCATTERFIELD("zippairs.from_zip")
# print lcStr + ">>"
# lcStr = cbWrap.cbwSCATTERFIELD("data2.l_name")
# print lcStr + ">>"
# lcStr = cbWrap.cbwSCATTERFIELD("to_zip")
# print lcStr + ">>"
#
# lnMilesX = c_int()
# lpPtr = pointer(lnMilesX)
# lnResult = cbWrap.cbwSCATTERFIELDLONG("zippairs.pc_miles", lpPtr)
# lnMiles = lpPtr.contents.value
# print lnMiles
#
# lnMilesX = c_double()
# lpPtr = pointer(lnMilesX)
# lnResult = cbWrap.cbwSCATTERFIELDDOUBLE("zippairs.geo_miles", lpPtr)
# lnMiles = lpPtr.contents.value
# lcMsg = cbWrap.cbwERRORMSG()
# print lcMsg
# print lnResult

# print lnMiles

# lnStart = time()
# for kk in range(1, 1000000):
#     cbWrap.cbwGOTO("SKIP", 1)
#     lcFldStr = cbWrap.cbwSCATTER("main")
#     laFldsx = lcFldStr.split("~~")
# lnEnd = time()
# print lnEnd - lnStart
# print lnResult

lnResult = cbWrap.cbwCLOSEDATABASES()
print(lnResult)

lnResult = cbWrap.cbwShutDown()
print(lnResult)

print("done")




