@ECHO OFF
REM Copyright FineBuild Team © 2015 - 2017.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms             ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01"      ^
 /SetupSQLDBCluster:YES                      ^
 /SetupSQLASCluster:YES                      ^
 /SetupSQLRSCluster:YES                      ^
 /SetupSQLIS:YES                             ^
 /SetupAlwaysOn:Yes                          ^
 /SQLSVCAccount:"ROOT\ServGB_SQLDB_1$"       ^
 /AGTSVCACCOUNT:"ROOT\ServGB_SQLAG_1$"       ^
 /ASSVCACCOUNT:"ROOT\ServGB_SQLAS_1$"        ^
 /FTSVCACCOUNT:"ROOT\ServGB_SQLFT_1$"        ^
 /ISSVCACCOUNT:"ROOT\ServGB_SQLIS_1$"        ^
 /RSSVCACCOUNT:"ROOT\ServGB_SQLRS_1$"        ^
 /BROWSERSVCACCOUNT:"ROOT\ServGB_SQLBR_1$"   ^
 /VolProg:C /VolTempWin:C /VolDTC:M          ^
 /VolBackup:C:\ClusterStorage\Volume4        ^
 /VolData:"C:\ClusterStorage\Volume1,C:\ClusterStorage\Volume2"                               ^
 /VolLog:C:\ClusterStorage\Volume3           ^
 /VolDataFS:H                                ^
 /VolBackupAS:G /VolDataAS:F /VolLogAS:G /VolTempAS:F                                         ^
 /SetupCmdShell:Yes                          ^
 /CmdshellAccount:"ROOT\AppGB_SQLCS_0001"    /CmdshellPassword:"j25Fb*ef$36ySIyBW7hZ"         ^
 /SetupRSExec:Yes                            ^
 /RSEXECACCOUNT:"ROOT\APPGB_SQLRS_0001"      /RSEXECPASSWORD:"Prf53g#fdf$Efbv8QGH3"           ^
 /SetupPolyBase:Yes                          ^
 /PBDMSSvcAccount:"ROOT\ServGB_SQLPB_1$"     ^
 /SetupAnalytics:Yes                         ^
 /ExtSvcAccount:"ROOT\ServGB_SQLES_1$"       