import CodeBaseTools
from MPSSCommon.MPSSBaseTools import findrepopath
import easygui

cbt = CodeBaseTools.cbTools()
gcRepoPath = findrepopath()

print("RUNNING")
bTest = cbt.use(gcRepoPath + "\\testdataoutput\\clients.dbf", "CLIENTS")

for jj in range(0, 250000):
    # xData = cbt.scatter(converttypes=False)
    # print len(xData)
    cbt.select("CLIENTS")
    xflds = cbt.afields()
    cTable = gcRepoPath + "\\testdataoutput\\clients99.dbf"
    bReturn = cbt.createtable(cTable, xflds)
    if bReturn:
        cAlias = cbt.alias()
        # cbt.closetable(cAlias)
        # cbt.use(cTable, "CLIENTSX9X")
        # print cbt.appendblank()
        cbt.closetable(cAlias)
    del xflds
    # print xflds[0].cName
    # xTags = cbt.ataginfo()
    # print len(xTags)
    # xTags = None
    # cName = cbt.curvallong("CLIENT_NUM")
    # bTest = cbt.seek("GP", "CLIENTS", "CLIENTNAME")
    # xTypes = cbt.afieldtypes()
    # print xTypes["CLIENTNAME"]
    # del xTypes
    # xBlank = cbt.scatterblank()
    # xBlank["CLIENT_NUM"] = 0
    # del xBlank
    
cbt.closetable("CLIENTS")
