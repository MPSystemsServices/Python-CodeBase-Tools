import MPSSCommon.CodeBasePYWrapper as cbp
xx = cbp.initdatasession(0)
print "Returned Session", xx
print "Retrieved Session", cbp.getcurrentsession()
print "Large Mode Status:", cbp.getlargemode()
yy = cbp.setdateformat("DMY")
print yy

cf = cbp.getdateformat()
print "Date Format", cf

x2 = cbp.initdatasession(0)
print "Second Session", x2
print "Retrieved Session", cbp.getcurrentsession()

cf2 = cbp.getdateformat()
print "Second data format", cf2

zz = cbp.closedatasession(xx)
print "Closed session", zz
print cbp.geterrormessage()

sw = cbp.switchdatasession(x2)
print "switch return", sw

zz2 = cbp.closedatasession(x2)
print "Closed Session 2", zz2
print cbp.geterrormessage()

