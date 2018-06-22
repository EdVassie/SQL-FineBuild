@ECHO OFF
REM Copyright FineBuild Team © 2012 - 2015.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Client                  /IAcceptLicenseTerms          ^
 /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01"
