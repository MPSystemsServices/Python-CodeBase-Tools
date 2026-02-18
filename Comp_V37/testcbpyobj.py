import CodeBaseTools as cbx
import MPSSCommon.MPSSBaseTools as mTools
import datetime
from time import time
vfp = cbx.cbTools()

gcRepoPath = mTools.findrepopath()
xx = vfp.use(gcRepoPath + "\\TestDataOutput\\clients.dbf", "TESTCLIENT")
print(xx)

aInfo = vfp.ataginfo()
print(vfp.tally)
for aI in aInfo:
    print(aI.cTagName, aI.cTagExpr)

print(vfp.setorderto("CLIENTNAME"))
print(vfp.order())

print(vfp.curvalstr("CLIENTNAME"))
print(vfp.curvalstr("CLIENTS.CLIENTNAME"))
print(vfp.cErrorMessage, "<< SECOND")
print(vfp.curvalstrEX("NOFIELDNAME"))
print(vfp.cErrorMessage)

print(vfp.goto("TOP"))
print(vfp.seek("GP"))
print(vfp.alias())
print(vfp.cErrorMessage, "<< THIRD")
print(vfp.curvalstrEX("CLIENTNAME"))
fInfo = vfp.afields()
for f in fInfo:
    print(f.cName, f.cType, f.nWidth)

fDict = vfp.afieldtypes()
print(fDict)

eDict = vfp.scatterblank()
print("EMPTY")
print(eDict)

vfp.goto("RECORD", 5)
dDict = vfp.scatter(converttypes=False, stripblanks=True, fieldList="CLIENT_NUM,CLIENTNAME,CREATED,MONTHRATE,XTODAY")
print(dDict)
print("MONTHRATE", float(dDict["MONTHRATE"]))

lDict = vfp.scattertoarray(stripblanks=False, fieldList="CLIENT_NUM,CLIENTNAME,CREATED,MONTHRATE,XTODAY")
print("<<< LIST OUTPUT >>>")
print(lDict)
if lDict is None:
    print(vfp.cErrorMessage)

bAdd = vfp.curvallogical("TESTCLIENT.ADDLTL")
print("ADD VALUE:", bAdd)
dCreate =vfp.curvaldate("DATECREATE")
tCreate =vfp.curvaldatetime("DATECREATE")
d2Create =vfp.curvaldate("XTODAY")
t2Create =vfp.curvaldatetime("XTODAY")
print("Four Dates:")
print(dCreate, tCreate, d2Create, t2Create)
  
nTest = vfp.calcfor("AVG", "CLIENT_NUM+0", "CLIENT_NUM>11000")
print(nTest)

bTest = vfp.goto("RECORD", 25)
print(vfp.recno())

print(vfp.curval("client_num", convertType=True))
print(vfp.curval("glop", convertType=True))
print(vfp.curval("client_num", convertType=False))
print(vfp.curval("monthrate", convertType=True))
print(vfp.curval("testclient.monthrate"))

print("FLOAT VALUES....")
print(vfp.curvalfloat("MBL_FROM"))
print(vfp.curvalfloat("RETAINDAYS"))
print(vfp.curvalfloat("MONTHRATE"))
print(vfp.curvalfloat("TESTDECIM"))

print("DECIMAL VALUES....")
print(vfp.curvaldecimal("MBL_FROM"))
print(repr(vfp.curvaldecimal("MBL_FROM")))
print(vfp.curvaldecimal("RETAINDAYS"))
print(vfp.curvaldecimal("MONTHRATE"))
print(vfp.curvaldecimal("TESTDECIM"))

print("CLIENT_NUM", vfp.curvallong("CLIENT_NUM"))
cRawData = vfp.scatterraw()
print(">>>> RAW >>>>")
print(cRawData)
print(">>>> DONE >>>>")

dDict = vfp.scatter(converttypes=True, stripblanks=True)
dDict["RETAINDAYS"] = 55
dDict["MONTHRATE"] = dDict["MONTHRATE"] + 2
dDict["BILL_NAME"] = "TRANZACT TECH " + mTools.strRand(8)
# nReturn = vfp.gatherfromarray(dDict)
nReturn = vfp.gatherdict(dData=dDict)
print(nReturn)
print(vfp.curval("RETAINDAYS", convertType=True))
print(vfp.curval("MONTHRATE", convertType=True))
print(vfp.curval("BILL_NAME", convertType=True))

print(type(dDict))
dDict["RETAINDAYS"] = 95
dDict["MONTHRATE"] = dDict["MONTHRATE"] + 6
dDict["BILL_NAME"] = "TRANZACT TECH " + mTools.strRand(8)
dDict["XTODAY"] = "20140101"
nReturn = vfp.gatherdict(dData=dDict)
print("GatherDict returned:", nReturn)
print(vfp.curval("RETAINDAYS", convertType=True))
print(vfp.curval("MONTHRATE", convertType=True))
print(vfp.curval("BILL_NAME", convertType=True))
print(vfp.curvaldate("XTODAY"))
print("Now for copy to array")
xList = vfp.copytoarray(fieldtomatch="CLIENTNAME", matchvalue="GP")
print(len(xList))
print(xList[2])

xToday = datetime.datetime.today()
print(xToday)
print(vfp.goto("RECORD", 13))
print(vfp.replace("CREATED", xToday))
print(vfp.cErrorMessage)
print(vfp.flush())
print(vfp.curvaldatetime("CREATED"), "CREATED!")

print(vfp.goto("NEXT", 1))
xNow = datetime.date.today()
print(xNow)
print(vfp.replace("CREATED", xNow))
print(vfp.flush())
print(vfp.curvaldatetime("CREATED"), "C DATE")
xStructure = vfp.afields()
bRet = vfp.createtable(gcRepoPath + "\\TestDataOutput\\clients2.dbf", xStructure)
cAlias2 = vfp.alias()
print(cAlias2, "alias 2")
print(bRet, vfp.cErrorMessage)

cStru = """FIELD1,C,25,0,FALSE
FIELD2,N,14,3,FALSE
TESTING,D,8,0,FALSE
MYDATES,T,8,0,FALSE
ISITTRUE,L,1,0,FALSE"""
bRet = vfp.createtable(gcRepoPath + "\\TestDataOutput\\testnew.dbf", cStru)
print(bRet, vfp.cErrorMessage)

nRet = vfp.copyto("TESTCLIENT", gcRepoPath + "\\TestDataOutput\\clientscopy.dbf", "DBF")
print("Copy To DBF result:", nRet)

nRet = vfp.copyto("TESTCLIENT", gcRepoPath + "\\TestDataOutput\\clientscopy.sdf", "SDF")
print("Copy To SDF result:", nRet)

nRet = vfp.copyto("TESTCLIENT", gcRepoPath + "\\TestDataOutput\\clientscopy.txt", "CSV", bHeader=True, bStripBlanks=True)
print("Copy To CSV result:", nRet)

nRet = vfp.copyto("TESTCLIENT", gcRepoPath + "\\TestDataOutput\\clientscopytab.txt", "TAB", bHeader=True, bStripBlanks=True)
print("Copy To TAB result:", nRet)

nRet = vfp.copyto("TESTCLIENT", gcRepoPath + "\\TestDataOutput\\clientscopy2.txt", "CSV", bHeader=True, cTestExpr='CLIENTNAME="GP"', bStripBlanks=True)
print("Copy To CSV result:", nRet)

nRet = vfp.appendfrom(cAlias=cAlias2, cSource=gcRepoPath + "\\TestDataOutput\\clientscopy2.txt", cType="CSV")
print(nRet, "appended count")
vfp.select(cAlias2)
vfp.goto("RECORD", 5)
cName = vfp.curval("CLIENTNAME", convertType=True)
print(cName)
print(">>>> SCANNING")
for xRec in vfp.scan():
    print(xRec["CLIENTNAME"], xRec["CLIENT_NUM"])

nRet = vfp.appendfrom(cSource=gcRepoPath + "\\TestDataOutput\\clientscopy.sdf")
vfp.goto("RECORD", 35)
cName = vfp.curval("CLIENTNAME", convertType=True)
print(cName)

nStart = time()
for kk in range(0, 500):
    xx = vfp.getNewKey(gcRepoPath + "\\TestDataOutput", "CLIENTS")
nEnd = time()
print((nEnd - nStart) / 500.0, "key time")

vfp.use(gcRepoPath + "\\TestDataOutput\\nextkey.dbf", "NEXTKEY", readOnly=False, noBuffering=True)
nStart = time()
for kk in range(0, 500):
    xx = vfp.getNewKey(gcRepoPath + "\\TestDataOutput", "CLIENTS")
nEnd = time()
print((nEnd - nStart) / 500.0, "key time, table open")
vfp.closetable("NEXTKEY")

bTest = vfp.use("e:\\loadbuilder2\\appdbfs\\geo.dbf")
print(vfp.alias())
nStart = time()
for xRec in vfp.scan(getList=False, forExpr='COUNTRY="USA"'):
    pass
nEnd = time()
print("SCAN TIME WITH DICT:", nEnd - nStart)
print("Used TESTCLIENT", vfp.used("TESTCLIENT"))
vfp.select("TESTCLIENT")
vfp.goto("TOP")
vfp.select("GEO")
xRec = vfp.scatter("TESTCLIENT")
print(vfp.cErrorMessage)
for xKey, xVal in xRec.items():
    print(xKey, xVal)
nStart = time()
nCounter = 0
vfp.closetable("GEO")
print(nStart, "START TIME")

cUnicode = mTools.FILETOSTR(gcRepoPath + "\\testdataoutput\\UnicodeXMLTest.xml")
print(len(cUnicode), "TEST XML LENGTH")
bOK = vfp.select("TESTCLIENT")
print(bOK, "SELECTED")
print(vfp.dbf())
bOK = vfp.goto("RECORD", 109)
print(bOK, "GOTO")
bOK = vfp.replace("TEST_MEMO", cUnicode)
print(bOK, "REPLACED")


##for nRec in vfp.scan(noData=True, forExpr='COUNTRY="USA"'):
##    nCounter += 1
##nEnd = time()
##print("SCAN TIME WITH REC:", nEnd - nStart, nCounter
##
##vfp.goto("TOP")
##nStart = time()
##while not vfp.eof():
##    nTest = vfp.recno()
##    vfp.skip(1)
##nEnd = time()
##print("LOOP TIME WITH REC:", nEnd - nStart
##
##vfp.closetable("GEO")
##
##oldVFP = cbOld.cbTools()
##bTest = oldVFP.use("e:\\loadbuilder2\\appdbfs\\geo.dbf")
##print(oldVFP.alias()
##nStart = time()
##for xRec in oldVFP.scan():
##    pass
##nEnd = time()
##print("OLD SCAN TIME WITH DICT:", nEnd - nStart
##nStart = time()
##for nRec in oldVFP.scan(noData=True):
##    pass
##nEnd = time()
##print("OLD SCAN TIME WITH REC:", nEnd - nStart
##oldVFP.closetable("GEO")

del vfp
##del oldVFP

#vfpL = cbx._cbTools(True)
#print(vfpL.nDataSession

#vfpX = cbx._cbTools(False)
#print(vfpX.nDataSession

#print(vfpX.deleted()
#print(vfpX.cErrorMessage
