@ECHO OFF
REM Copyright FineBuild Team © 2008 - 2017.  Distributed under Ms-Pl License
REM
CALL "SqlFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms        ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01" ^
 /SetupSQLDBCluster:YES                    ^
 /SetupSQLASCluster:YES                    ^
 /SetupSQLRSCluster:YES                    ^
 /SetupSQLIS:YES                           ^
 /SQLAccount:"ROOT\ServGB_SQLDB_0001"      /SQLPassword:"Argyt$6hsGGWMP894s4Gw2b73GS2o0" ^
 /AGTACCOUNT:"ROOT\ServGB_SQLAG_0001"      /AGTPASSWORD:"F6tbmd*nf!dfGFrcQnm84g4K7zwq2j" ^
 /ASACCOUNT:"ROOT\ServGB_SQLAS_0001"       /ASPASSWORD:"kE44bmutFGS579*bssJW84f=Rb6ehj"  ^
 /ISACCOUNT:"ROOT\ServGB_SQLIS_0001"       /ISPASSWORD:"bSHG5iuf9DFF#dw2!F5sKSIw43tnb7"  ^
 /RSACCOUNT:"ROOT\ServGB_SQLRS_0001"       /RSPASSWORD:"Orfd450!#DTWjn63hw45JDD873hk84"  ^
 /SQLBROWSERACCOUNT:"ROOT\ServGB_SQLBR_0001"  /SQLBROWSERPASSWORD:"w#d6gh*ge$dvnHHq1knbtd$Wd68Zj9" ^
 /VolProg:C /VolTempWin:C /VolDTC:M        ^
 /VolBackup:J /VolData:J /VolDataFT:J /VolLog:K /VolLogTemp:K /VolSysDB:J /VolTemp:J     ^
 /VolBackupAS:G /VolDataAS:F /VolLogAS:G /VolTempAS:F                                    ^
 /SETUPCMDSHELL:YES ^
 /CmdshellAccount:"ROOT\AppGB_SQLCS_0001"  /CmdshellPassword:"j25Fb*ef$36ySIyBW7hZ"      ^
 /SetupRSExec:Yes                          ^
 /RSEXECACCOUNT:"ROOT\APPGB_SQLRS_0001"    /RSEXECPASSWORD:"Prf53g#fdf$Efbv8QGH3"        ^
 /SQLCLUSTERGROUP:"ROOT\GBGGSQLC01DB"      /AGTCLUSTERGROUP:"ROOT\GBGGSQLC01AGT"         ^
 /ASCLUSTERGROUP:"ROOT\GBGGSQLC01AS"       /FTSCLUSTERGROUP:"ROOT\GBGGSQLC01FTS"
REM
REM /AdminPassword: must be supplied and contain the password for the account running SQL FineBuild
REM For details see http://sqlserverfinebuild.codeplex.com/wikipage?title=SQL%20Server%20Cluster%20Install