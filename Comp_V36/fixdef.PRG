SET SAFETY OFF
lcText = FILETOSTR("deflist.txt")
lnCnt = ALINES(laLines, lcText)
lcOut = ""
lcOut = laLines[1] + CHR(13) + CHR(10)
FOR jj = 2 TO lnCnt
	IF LEN(laLines[jj]) > 20
		laLines[jj] = "   " + SUBSTR(laLines[jj], 27)
	ENDIF
	lcOut = lcOut + laLines[jj] + CHR(13) + CHR(10)
ENDFOR

STRTOFILE(lcOut, "deflistout.txt")
SET SAFETY ON
RETURN