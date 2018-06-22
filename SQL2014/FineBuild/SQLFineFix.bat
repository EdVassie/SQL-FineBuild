@ECHO OFF
REM Copyright FineBuild Team © 2012 - 2015.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Fix                     /IAcceptLicenseTerms          ^
 /SetupSP:Yes /SetupSPCU:Yes   /SetupSPCUSNAC:Yes            /SetupBOL:Yes                 ^
 /SPLevel:SP1 /SPCULevel:CU4