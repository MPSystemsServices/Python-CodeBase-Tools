from dbf import Table
from time import time
zz = time()
xx = Table("c:\\temp\\mytest2.dbf", "name C(45);age N(10,0);ssn C(9);emptest L;hired D;clocked T;xcounts I", dbf_type="vfp")
zzend = time()
print zzend - zz
print xx
