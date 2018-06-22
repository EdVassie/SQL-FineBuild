@ECHO OFF
REM Copyright FineBuild Team © 2008 - 2016.  Distributed under Ms-Pl License
REM
CALL "SqlFineBuild.bat" %*     /Type:Fix                     /IAcceptLicenseTerms          ^
 /SetupSP:Yes /SetupSPCU:Yes   /SetupSPCUSNAC:Yes /SetupBOL:Yes     ^
 /SPLevel:SP2 /SPCULevel:CU3