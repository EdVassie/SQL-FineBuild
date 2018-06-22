@ECHO OFF
REM Copyright FineBuild Team © 2012 - 2017.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms             ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01"      ^
 /SetupSQLDB:YES                             ^
 /SetupSQLAS:YES                             ^
 /SetupSQLRS:YES                             ^
 /SetupSQLIS:YES                             ^
 /SQLSVCAccount:"ROOT\ServGB_SQLDB_1$"       ^
 /AGTSVCACCOUNT:"ROOT\ServGB_SQLAG_1$"       ^
 /ASSVCACCOUNT:"ROOT\ServGB_SQLAS_1$"        ^
 /FTSVCACCOUNT:"ROOT\ServGB_SQLFT_1$"        ^
 /ISSVCACCOUNT:"ROOT\ServGB_SQLIS_1$"        ^
 /RSSVCACCOUNT:"ROOT\ServGB_SQLRS_1$"        ^
 /BROWSERSVCACCOUNT:"ROOT\ServGB_SQLBR_1$"   ^
 /VolProg:C          ^
 /VolBackup:I /VolData:JF /VolDataFT:F /VolLog:KG /VolTemp:T                                  ^
 /SETUPCMDSHELL:YES                          ^
 /CMDSHELLACCOUNT:"ROOT\APPGB_SQLCS_0001"    /CMDSHELLPASSWORD:"j25Fb*ef$36ySIyBW7hZ"         ^
 /SETUPRSEXEC:YES                            ^
 /RSEXECACCOUNT:"ROOT\APPGB_SQLRS_0001"      /RSEXECPASSWORD:"Prf53g#fdf$Efbv8QGH3"           ^
 /SETUPBPE:YES /VOLBPE:T /BPEFile:30GB
