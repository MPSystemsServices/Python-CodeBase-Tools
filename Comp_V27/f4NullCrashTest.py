import CodeBaseTools
import MPSSCommon.MPSSBaseTools as mTools

gcRepoPath = mTools.findrepopath()

xp = CodeBaseTools.cbTools()
print(xp.use("e:\\loadbuilder2\\appdbfs\\clients.dbf"), "Using clients")
print(xp.alias())
print(xp.use("e:\\loadbuilder2\\appdbfs\\clients.dbf", "XCLIENTSLTL"), "Using Clients again")
print(xp.cErrorMessage, xp.nErrorNumber)
print(xp.alias())
xp.use(gcRepoPath + "\\loadoptwebfiles\\vfptabletemplates\\tablelist.dbf")
xp.use("e:\\loadbuilder2\\appdbfs\\ltloptions.dbf")
xp.use("e:\\loadbuilder2\\appdbfs\\clients.dbf")
print(xp.alias())
print("READY TO SEEK")
bResult = xp.seek("11458", "CLIENTS", "CLIENT_NUM")
xRec = xp.scatter()
print(xRec["CLIENTNAME"])
print("SEEK DONE", bResult)
xp.closedatabases()
print("DONE DONE")
