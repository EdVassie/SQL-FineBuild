@ECHO OFF
REM Copyright FineBuild Team � 2015 - 2016.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Client                  /IAcceptLicenseTerms          ^
 /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01"
