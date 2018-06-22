@ECHO OFF
REM Copyright FineBuild Team © 2009 - 2016.  Distributed under Ms-Pl License
REM
CALL "SqlFineBuild.bat" %*     /Type:Workstation             /IAcceptLicenseTerms         ^
 /Edition:Express                       ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01" 