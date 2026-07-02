import os
from CodeBaseTools import cbTools
oDB = cbTools()
os.chdir(r"E:\GIT_Repositories\Python-CodeBase-Tools\TestFiles")
# Open both tables
oDB.use("PRODUCTS")
oDB.use("PRODCATG")

xPORProdList = list() # Define the target list to store the results

for xRec in oDB.scan(forExpr="CATG_WHSE='POR'"):
	# Iterate through each PRODCATG record (currently selected table) that is stored in 'POR'
	cCatg = xRec["PROD_CATG"]
	# Look for all products in that category
	oDB.select("PRODUCTS")
	if oDB.locate("PROD_CATG='" + cCatg + "'"):
		xItem = oDB.scatter()
		xPORProdList.append(xItem)
		while oDB.locatecontinue():
			xItem = oDB.scatter()
			xPORProdList.append(xItem)
		oDB.locateclear()
		oDB.select("PRODCATG")

# Print out the results.
for xProd in xPORProdList:
		print(xProd["PROD_NAME"], xProd["PROD_COUNT"])

# close the tables
oDB.closetable("PRODUCTS")
oDB.closetable("PRODCATG")
