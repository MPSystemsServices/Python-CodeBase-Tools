import os
from CodeBaseTools import cbTools
oDB = cbTools()
os.chdir(r"E:\GIT_Repositories\Python-CodeBase-Tools\TestFiles")
# Open both tables
oDB.use("PRODCATG")
oDB.use("PRODUCTS")

xProdList = oDB.copytoarray(alias="PRODUCTS", maxcount=20000)
xPORProdList = list()

oDB.select("PRODCATG")
for xProd in xProdList:  # examine each one in the big list
        # Iterate through each PRODCATG record (currently selected table) that is stored in 'POR'
        cCatg = xProd["PROD_CATG"]
        bFound = oDB.seek(cCatg, "PRODCATG", "PROD_CATG")
        if bFound:
                cWhse = oDB.curvalstr("PRODCATG.CATG_WHSE")
                if cWhse.strip() == "POR":
                        xPORProdList.append(xProd)

# Print out the results.
for xProd in xPORProdList:
		print(xProd["PROD_NAME"], xProd["PROD_COUNT"])

# close the tables
oDB.closetable("PRODUCTS")
oDB.closetable("PRODCATG")
