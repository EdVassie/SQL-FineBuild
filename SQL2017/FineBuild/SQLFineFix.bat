@ECHO OFF
REM Copyright FineBuild Team © 2015 - 2016.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Fix                     /IAcceptLicenseTerms          ^
 /SetupSP:Yes /SetupSPCU:Yes   /SetupSPCUSNAC:Yes            /SetupBOL:Yes                 ^
 /SPLevel:SP1 /SPCULevel:CU1