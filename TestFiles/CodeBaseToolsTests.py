

"""
Testing tool for all versions of CodeBaseTools: Version Specific for PY 2.7, 3.7, 3.8, 3.9, Generic Library for versions 3.10+, and 64-bit versions.

Copyright M-P System Services, Inc., May, 2026, and released under the Lesser GPL License to Open Source users.  No claims are made for
the completeness or applicability of these tests to your conditions of use of the CodeBaseTools for Python.

"""
import sys
import os
import datetime
from time import time
import glob
import decimal
import easygui  # Available from PPI via a pip call from your Python scripts directory.
from CodeBaseTools import cbTools, copydatatable

class cbtTester(object):

    def __init__(self):
        self.cName = "CodeBaseToolsTester"
        self.cTestFileBasePath = easygui.diropenbox(msg="Select CodeBaseTools Test File Base Directory")
        if not self.cTestFileBasePath:
            print("No CodeBaseTools Base Test File Directory Selected")
            sys.exit(1)
            return

        bOK = os.chdir(self.cTestFileBasePath)
        self.cOutputPath = os.path.join(self.cTestFileBasePath, "TestDataOutput")
        self.oDB = cbTools()
        print("CodeBaseTools Test File Directory Selected")
        print("Selected Path:" + str(self.cTestFileBasePath))
        return

    def cbtWork(self):
        # Create Table Example #1
        tmpdir = self.cTestFileBasePath
        xFlds = list()
        print("starting cbtWork")
        xFlds.append("firstname,C,25,0,FALSE")
        xFlds.append("lastname,C,25,0,FALSE")
        xFlds.append("city,C,35,0,FALSE")
        xFlds.append("state,C,2,0,FALSE")
        xFlds.append("zipcode,C,10,0,FALSE")
        xFlds.append("hire_date,D,8,0,FALSE")
        xFlds.append("login_date,T,8,0,TRUE")
        xFlds.append("num_kids,I,4,0,FALSE")
        xFlds.append("salary,N,10,2,FALSE")     # Note Number type.  Exact decimals in the table but will be
        #  converted in Python to a double.
        xFlds.append("wageperhr,Y,10,3,FALSE")  # Note currency type with exact decimals
        xFlds.append("ssn,C,10,0,FALSE")
        xFlds.append("married,L,1,0,FALSE")
        xFlds.append("dob,D,8,0,TRUE")
        xFlds.append("comments,M,10,0,FALSE")
        xFlds.append("coverage,B,12,4,FALSE")  # This is a DOUBLE field, which functions as a floating point
        lcFld = "\n".join(xFlds)
        # The init() of the cbTools class initializes the engine for you.
        vfp = self.oDB
        lnStart = time()

        testname = os.path.join(self.cOutputPath, "employee1.dbf")
        DELETEFILE(testname)
        print(testname)
        bResult = vfp.createtable(testname, lcFld)
        lnEnd = time()
        print("Create Time: %d" % (lnEnd - lnStart))
        print("Create Work Result: %s" % str(bResult))
        print("ERR: ", vfp.cErrorMessage)
        cAlias = vfp.alias()
        print("Alias>%s" % cAlias)
        xFlds = vfp.afields()
        print("Fields:")
        for xF in xFlds:
            print(xF)
        if not bResult:
            print("ERROR", vfp.cErrorMessage)
            exit(1)
        print("Alias Name: %s" % vfp.alias())
        print(vfp.refreshbuffers(), "buffers refreshed")
        vfp.closetable("employee1")
        bexcl = vfp.use(testname, alias="EMPLOYEE1", exclusive=True)
        print("bxl", bexcl)
        # Index new table
        lnStart = time()
        bResult = vfp.indexon("fullname", "UPPER(lastname+firstname)")
        bResult = vfp.indexon("geo", "state+city")
        print("index result", bResult)
        print("indx err:", vfp.cErrorMessage)
        lnEnd = time()
        print("Index time: ", lnEnd - lnStart)

        # Append record from two lists (Value and Field Name)
        xVal = list()
        xFld = list()
        xFld.append("firstname")  # Note that field names may be specified in lower case, but will be UPPER in table.
        xVal.append("George")
        xFld.append("lastname")
        xVal.append("Weibladson")
        xFld.append("city")
        xVal.append("Chicago")
        xFld.append("state")
        xVal.append("IL")
        xFld.append("zipcode")
        xVal.append("60607")
        xFld.append("hire_date")
        xVal.append(datetime.date(2009, 6, 12))
        xFld.append("login_date")
        xVal.append(datetime.datetime(2011, 1, 24, 9, 30, 00))
        xFld.append("salary")
        xVal.append(decimal.Decimal("35535.50"))
        xFld.append("num_kids")
        xVal.append(2)
        xFld.append("wageperhr")
        xVal.append(23.45)
        xFld.append("ssn")
        xVal.append("123456789")
        xFld.append("married")
        xVal.append(True)
        xFld.append("dob")
        xVal.append(datetime.date(1955, 10, 3))
        xFld.append("comments")
        xVal.append("hello to all")
        xFld.append("coverage")
        xVal.append(2345.19003)

        lnStart = time()
        bResult = vfp.appendblank()
        print("Append Result: ", bResult)
        print("append err ", vfp.cErrorMessage)
        bResult = vfp.gatherfromarray(xVal, xFld)
        print("Gather Result: ", bResult)
        lnEnd = time()
        print("Tot Append Time: ", lnEnd - lnStart)
        print("Gather Error: ", vfp.cErrorMessage)

        # Append record from string of values (all fields updated) -- 10X faster
        lnStart = time()
        bResult = vfp.appendblank()
        lcVals = "Pete<~!~>Smith<~!~>Boston<~!~>MA<~!~>03493<~!~>19980224<~!~>20110103140203<~!~>2<~!~>66039.39<~!~>45.37<~!~>987734959<~!~>FALSE<~!~>19591203<~!~>Testing Again<~!~>0.5343"
        bResult = vfp.gatherfromarray(lcVals, None, "<~!~>")
        lnEnd = time()
        print("String Gather Time: ", lnEnd - lnStart)
        print("Gather From String Result: ", bResult)
        # Append record from a dictionary
        lxD = dict()
        lnStart = time()
        lxD["firstname"] = "William"
        lxD["lastname"] = "Jones"
        lxD["city"] = "Los Angeles"
        lxD["state"] = "CA"
        lxD["zipcode"] = "90272"
        lxD["hire_date"] = datetime.date(2005, 7, 29)
        lxD["login_date"] = datetime.datetime(2011, 2, 12, 14, 34, 21)
        lxD["salary"] = 42000.00
        lxD["num_kids"] = 2
        lxD["wageperhr"] = 22.35
        lxD["ssn"] = "292929292"
        lxD["married"] = False
        lxD["dob"] = None
        lxD["comments"] = "asdfwe kkkkkkkkkkkkkkkkkk     iiiiiiiiiiiiiiio \n iiiiiiiiiiiii"
        lxD["coverage"] = 0.35239
        bResult = vfp.insertdict(lxD)
        lnEnd = time()
        print("gather from dict result: ", bResult)
        print("gather from dict time: ", lnEnd - lnStart)

        # Go back to record 1 and get several field values
        bResult = vfp.goto("RECORD", 1)
        print("record 1?", vfp.recno(), "count", vfp.reccount())
        lxDict = vfp.scatter()
        if lxDict is None:
            print(vfp.cErrorMessage, 1)
        print(lxDict["FIRSTNAME"])  # Note upper case field names are REQUIRED.
        print(lxDict["LASTNAME"])
        print(lxDict["LOGIN_DATE"])  # Note conversion to Python datetime value
        print(lxDict["SALARY"])  # Converted to a double
        print(lxDict["WAGEPERHR"])  # Converted to currency value with exact decimals
        print("NAME", vfp.curvalstr("EMPLOYEE1.FIRSTNAME"), "ENDNAME")
        print("NUMKIDS", vfp.curvallong("EMPLOYEE1.NUM_KIDS"))
        #
        bResult = vfp.goto("RECORD", 2)
        vfp.replace("employee1.firstname", "Charles")
        vfp.replace("employee1.wageperhr", 37.50)
        vfp.replace("employee1.dob", datetime.date(1945, 6, 13))
        vfp.replace("employee1.married", True)
        lxDict = vfp.scatter()
        if lxDict is None:
            print(vfp.cErrorMessage, 2)
        print(lxDict["FIRSTNAME"])
        print(lxDict["DOB"])
        print(lxDict["MARRIED"])
        print("READY TO SCAtteR")
        bResult = vfp.goto("RECORD", 1)
        print("lbResult", bResult)
        lxDict = vfp.scatter(converttypes=False, fieldList="firstname,LASTNAME,SALARY")
        print(lxDict)
        print("that was dict")
        print("READY TO SCATTER RECORD")
        lxRec = vfp.scattertorecord()
        print(lxRec)
        print("FIRSTNAME:", lxRec.FIRSTNAME)
        print("LASTNAME:", lxRec.LASTNAME)
        cOutputXML = os.path.join(self.cOutputPath, "myxml.xml")
        cXML = vfp.cursortoxml(cFileName=cOutputXML)
        print(vfp.dbf())
        vfp.select("employee1")
        cNewXML = os.path.join(self.cOutputPath, "myxml.xml")
        vfp.xmltocursor(cFileName=cNewXML)
        xArray = vfp.copytoarray()
        print("LISTING XML CONVERTED TABLE DATA.....")
        for xA in xArray:
            print(xA)
        vfp.closedatabases()
        return True


    def cbt_test(self):
        cDest = os.path.join(self.cOutputPath, "SMCC2.DBF")
        vfp = self.oDB
        cShipMstrTable = os.path.join(self.cTestFileBasePath, "shipmstr.dbf")
        bTestUse = vfp.use(cShipMstrTable, alias="SHIPMSTR")
        print("USING SHIPMSTR:", cShipMstrTable, bTestUse)
        nStart = time()
        nResult = vfp.copyto(cAlias="SHIPMSTR", cOutput=cDest, cType="DBF", cTestExpr='', bHeader=False, bStripBlanks=False)
        nEnd = time()
        vfp.closetable("SHIPMSTR")
        print("ELAPSED TIME: ", (nEnd - nStart))
        if nResult < 0:
            print("ERROR RESULT WAS:", nResult)
            print("THE ERROR:", vfp.cErrorMessage)
            print("THE ERR NUMBER:", vfp.nErrorNumber)
        assert nResult > 0, "CopyTo function failed..."

        print("NRESULT: ", nResult)
        # return True
        nNewKey = vfp.getNewKey(self.cTestFileBasePath, filename="CLIENTS", readOnly=False)
        print("THE NEW KEY WAS:", nNewKey)
        #  Create Table Example #1
        assert nNewKey > 0, "getNewKey function failed..."

        lcFld = "firstname,C,25,0,FALSE\n"
        lcFld += "lastname,C,25,0,FALSE\n"
        lcFld += "city,C,35,0,FALSE\n"
        lcFld += "state,C,2,0,FALSE\n"
        lcFld += "zipcode,C,10,0,FALSE\n"
        lcFld += "hire_date,D,8,0,FALSE\n"
        lcFld += "login_date,T,8,0,TRUE\n"
        lcFld += "num_kids,I,4,0,FALSE\n"
        lcFld += "salary,N,10,2,FALSE\n"   # Note Number type.  Exact decimals in the table but will be
        # converted in Python to a double.
        lcFld += "wageperhr,Y,10,3,FALSE\n"  # Note currency type with exact decimals
        lcFld += "ssn,C,10,0,FALSE\n"
        lcFld += "married,L,1,0,FALSE\n"
        lcFld += "dob,D,8,0,TRUE\n"
        lcFld += "comments,M,10,0,FALSE\n"
        lcFld += "coverage,B,12,4,FALSE\n"  # This is a DOUBLE field, which functions a floating point
        # The init() of the cbTools class initializes the engine for you.
        vfp = cbTools()
        lnStart = time()
        testname = os.path.join(self.cOutputPath, "employee1.dbf")
        print(testname)
        bResult = vfp.createtable(testname, lcFld)
        assert bResult, "createtable function failed..."

        lnEnd = time()
        print("Create Time: ", lnEnd - lnStart)
        print("Create Result: ", bResult)
        if not bResult:
            print("ERROR", vfp.cErrorMessage)
            exit(1)
        print("Alias Name X: ", vfp.alias())
        zFlds = vfp.afields()
        print("zFlds below")
        print(zFlds)
        print("zFlds above")
        for zFld in zFlds:
            print(str(zFld))
        print("zFlds List Done")
        vfp.closetable("employee1")
        print("REOPENING:", vfp.use(testname, alias="employee1"))
        print("GOTO BOTTOM:", vfp.goto("BOTTOM"))
        # Append record from a dictionary
        lxD = dict()
        lnStart = time()
        lxD["firstname"] = "William"
        lxD["lastname"] = "Jones"
        lxD["city"] = "Los Angeles"
        lxD["state"] = "CA"
        lxD["zipcode"] = "90272"
        lxD["hire_date"] = datetime.date(2005, 7, 29)
        lxD["login_date"] = datetime.datetime(2011, 2, 12, 14, 34, 21)
        lxD["salary"] = 42000.00
        lxD["num_kids"] = 2
        lxD["wageperhr"] = 22.35
        lxD["ssn"] = "292929292"
        lxD["married"] = False
        lxD["dob"] = None
        lxD["comments"] = "This is a test of the memo field, this is only a test.  The lazy fox, etc."
        lxD["coverage"] = 0.35239

        bInsertTest = vfp.insertdict(lxD)
        assert bInsertTest, "insert function failed..."

        xRec = vfp.scatter()
        assert xRec is not None, "scatter function failed..."

        print("RECORD\n", xRec)
        print(vfp.cErrorMessage)
        lnEnd = time()
        print("gather from dict result: ", bResult)
        print("gather from dict time: ", lnEnd - lnStart)
        vfp.closetable(vfp.alias())

       # Go back to record 1 and get several field values
        bOK = vfp.use(testname, alias="employee1")
        assert bOK, "use function failed..."

        lbResult = vfp.goto("RECORD", 1)
        assert lbResult, "goto function failed..."

        lxDict = vfp.scatter()
        print(lxDict["FIRSTNAME"]) ## Note upper case field names are REQUIRED.
        print(lxDict["LASTNAME"])
        print(lxDict["LOGIN_DATE"]) ## Note conversion to Python datetime value
        print(lxDict["SALARY"]) ## Converted to a double
        print(lxDict["WAGEPERHR"]) ## Converted to currency value with exact decimals
        print("NAME", vfp.curvalstr("EMPLOYEE1.FIRSTNAME"), "ENDNAME")
        print("NUMKIDS", vfp.curvallong("EMPLOYEE1.NUM_KIDS"))

        lbResult = vfp.goto("RECORD", 2)
        assert lbResult == False, "Failed to detect end of file error..."

        vfp.replace("employee1.firstname", "Charles")
        vfp.replace("employee1.wageperhr", 37.50)
        vfp.replace("employee1.dob", datetime.date(1945, 6, 13))
        vfp.replace("employee1.married", True)
        lxDict = vfp.scatter()
        print(lxDict["FIRSTNAME"])
        print(lxDict["DOB"])
        print(lxDict["MARRIED"])

        lxaData = vfp.copytoarray(fieldtomatch="dob", matchvalue=None, matchtype="<>")
        print("LXADATA:", lxaData)
        print(lxaData[0]["FIRSTNAME"], lxaData[0]["LASTNAME"])

        lcTableName = vfp.dbf()
        print("Table Name: ", lcTableName)
        lcTableName = vfp.dbf("NOSUCHTABL")
        print("Bad Table Name: >" + lcTableName + "<")
        print("Error Message: " + vfp.cErrorMessage)
        print("Error Number: ", vfp.nErrorNumber)

        ## Create a new table as a copy of another by capturing a list of the first
        ## table's fields and then pass that list to the createtable function.
        vfp.closedatabases()
    #     lnStart = time()
    #     lbResult = vfp.use(testname, "emp1")
    #     lxTempList, lxTempFields = vfp.fielddicttolist(lxDict, "emp1")
    #     print lxTempList
    #     print lxTempFields
    #     for jj in range(0, 20):
    #         vfp.insertintotable(lxTempList, lxTempFields, lcAlias="emp1")
    #     print vfp.alias(), vfp.dbf()
    #     print vfp.reccount()
    #
    #     vfp.goto("RECORD", 19)
    #     vfp.delete()
    #     vfp.pack()
    #     print "After pack: ", vfp.reccount()
    #     lnCnt = 0
    #     lxFlds = vfp.afields()
    #
    #     vfp.closedatabases()
    #     lnStartC = time()
    #     assert True == vfp.createtable(os.path.join(tmpdir,"employee2.dbf"), lxFlds)
    #     print "### test TableObj iterator ###"
    #     cnt=0
    #     for this in vfp['employee2']:
    #         cnt += 1
    #     assert 0 == cnt , "Iterator not working"
    #     assert True == vfp.appendblank()
    #     cnt=0
    #     for this in vfp['employee2']:
    #         cnt += 1
    #     assert 1 == cnt , "Iterator not working"
    #     assert True == vfp.closetable("employee2")
    #     assert True == vfp.use(os.path.join(tmpdir,"employee2.dbf"))
    #     xflds = vfp.scatter(converttypes=False)
    #     print xflds
    #     lnEndC = time()
    #     print "Create Time 2: ", lnEndC - lnStartC
    #     vfp.dispstru()
    #     lnEnd = time()
    #     print "Open and Copy Time: ", lnEnd - lnStart
    #     if (lbResult == False):
    #         print vfp.cErrorMessage, vfp.nErrorNumber
    #
    #     ## test table from TableObj
    #     print 'TableObj from vfp',
    #     table = vfp.TableObj( testname)
    #     if not table.open: print 'Open error:',table.errormessage
    #     print table.name
    #     print 'Fields: ',table._fields
    #     del table
    #     print 'TableObj direct',
    #     table = TableObj( vfp, testname)
    #     print 'kids =',table.num_kids,  # test getattr
    #     print ', salary =',table['salary'],  # test getitem
    #     table.num_kids = 1  # test setattr
    #     table['salary'] = 20000.5  # test setitem
    #     print
    #     fields = 'firstname hire_date married num_kids salary login_date coverage wageperhr'.split()
    #     fsize = 10
    #     ff = '%%-%ss' % fsize
    #     print '  ',
    #     for n in fields:
    #         print ff%n,
    #     print
    #     while not vfp.eof():
    #         print;print '%2d' % vfp.recno(),
    #         for n in fields:
    #             print ff % str(table[n])[:fsize],
    #         table.next()
    #     print
    #     lastAlias = vfp.alias()
    #     print 'open file=',vfp.alias()
    #
    #     ## test getNewKey()
    #     print 'testing getNewKey()'
    #     testLocation = 'e:/loadbuilder2/appdbfs'
    #     assert vfp.getNewKey(testLocation,'bob',readOnly=True) == 11000  # new file
    #     assert vfp.alias() == lastAlias  #doesn't change current file
    #     vfp.getNewKey('c:/','xyz', readOnly=True)
    #     assert (vfp.cErrorMessage[:14] == "Unable to open"), "ErrorMessage = '%s'. curFile = '%s'" % (
    #         vfp.cErrorMessage, vfp.alias()
    #         )
    #     assert vfp.alias() == lastAlias  #doesn't change current file on error either
    #     if vfp.getNewKey(testLocation,'bob',readOnly=True, stayOpen=True) != 11000:
    #         # delete item 'bob'
    #         vfp.select('NextKey')
    #         x = vfp.curval('table_name',True) == 'BOB' and vfp.delete()
    #     assert vfp.getNewKey(testLocation,readOnly=True) == 11000  # `testname` is new file also
    #     assert vfp.getNewKey(testLocation,'SHIPMSTR',readOnly=True) > int(1e8)  # bigger than 100 million
    #     assert vfp.alias() == lastAlias  # make sure we haven't changed the active file
    #     assert vfp.getNewKey(testLocation,'bob', stayOpen=True) == 11000  # new file
    #     assert vfp.getNewKey(testLocation,'bob', stayOpen=True) == 11001  # added one
    #     assert vfp.getNewKey(testLocation,'bob', stayOpen=True) == 11002  # added one
    #     assert vfp.select('NextKey') is True
    #     assert vfp.curval('table_name',True) == 'BOB' and vfp.delete()
    #     assert vfp.getNewKey(testLocation,'bob',readOnly=True) == 11000  # new file
    #     assert vfp.getNewKey(testLocation,'bob',readOnly=True) == 11000  # still new file
    #     # test rlock() changes
    #     vfp.select(lastAlias)
    #     assert vfp.rlock(1) is True
    #     assert vfp.unlock() is True
    #
    #     vfp.closedatabases()
    #     lnTest = vfp.use(r"e:\loadbuilder2\appdbfs\geo.dbf", "GEO")
    #     print "GEO: ", lnTest
    #     print vfp.alias()
    #     lbFound = vfp.locate("ST_PROV = 'OR' .AND. UPPER(CITY) = 'PORTLAND'")
    #     print lbFound
    #     print vfp.cErrorMessage
    #     lcTest = vfp.curval("POSTALCODE")
    #     print lcTest
    #     lbFound = vfp.locatecontinue()
    #     print lbFound
    #     lcTest = vfp.curval("POSTALCODE")
    #     print lcTest
    #     vfp.locateclear()
    #
    #     vfp.setorderto("POSTALCODE")
    #     print "IN CHICAGO"
    #     lbFound = vfp.locate("ST_PROV = 'IL' .AND. CITY = 'CHICAGO'")
    #     while lbFound:
    #         print vfp.curvalstr("POSTALCODE")
    #         lbFound = vfp.locatecontinue()
    #     vfp.locateclear()
    #
    #     ## Shut down the DBF engine.
        vfp.closedatabases()
        vfp = None
        return lbResult

    def TestCopyTables(self) -> bool:
        cTargetPath = self.cOutputPath
        bTest = copydatatable(cTableName="DOMAINS.DBF", cSourceDir=self.cTestFileBasePath, cTargetDir=cTargetPath,
                              oCBT=self.oDB, bByZap=True)
        print("Result:", bTest)
        print("Error?:", self.oDB.getErrorMessage())
        return bTest

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

if __name__ == "__main__":
    print("***** Testing CodeBaseTools.py components")
    oTest = cbtTester()
    if oTest.cTestFileBasePath:
        print("************ cbt test starting *************")
        oTest.cbt_test()
        print("************ cbt test done **************")
        print("************ testing Copy Tables Starting ************")
        lbResult = oTest.TestCopyTables()
        print("************ Test Copy Tables Done ************")
        print("RESULT:", lbResult)
        print("************ cbtWork starting **************")
        lbResult = oTest.cbtWork() # More Extensive Tests
        print("************* cbtWork Done **************")
        print("RESULT:", lbResult)
        print("***** Testing Complete")
        if 'stop' in sys.argv:
            # if not _ver3x:
            #     raw_input('press <enter>')
            # else:

            input('press <enter>')
        else:
            print("DONE")
