import CodeBasePYWrapper312 as cbp
xx = cbp.initdatasession(0)
print("Returned Session", xx)
print("Retrieved Session", cbp.getcurrentsession())
print("Large Mode Status:", cbp.getlargemode())
yy = cbp.setdateformat("DMY")
print(yy)

cf = cbp.getdateformat()
print("Date Format", cf)

cbp.setdebugmode(1)
bOpened = cbp.use("e:\\loadbuilder2\\appdbfs\\clients.dbf", "CLIENTS", 0, 0)
print(cbp.alias())
bTestLocate = cbp.locate("CLIENTNAME = 'CA'")
print("Is Found", bTestLocate)
cbp.locateclear()

x2 = cbp.initdatasession(0)
print("Second Session", x2)
print("Retrieved Session", cbp.getcurrentsession())

cf2 = cbp.getdateformat()
print("Second data format", cf2)

zz = cbp.closedatasession(xx)
print("Closed session", zz)
print(cbp.geterrormessage())

sw = cbp.switchdatasession(x2)
print("switch return", sw)

zz2 = cbp.closedatasession(x2)
print("Closed Session 2", zz2)
print(cbp.geterrormessage())

print("Setting Debug Mode TRUE")
zz3 = cbp.setdebugmode(True)
print("SETTING RESULT", zz3)

print("Retrieving Debug Mode")
zz4 = cbp.setdebugmode(-1)
print("Retrieved Result:", zz4)
zz5 = cbp.setdebugmode(1)
print("AGAIN", zz5)

