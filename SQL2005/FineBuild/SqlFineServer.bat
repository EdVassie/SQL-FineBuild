@ECHO OFF
REM Copyright FineBuild Team © 2008 - 2016.  Distributed under Ms-Pl License
REM
CALL "SqlFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms        ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01" ^
 /SetupSQLDB:YES                             ^
 /SetupSQLAS:YES                             ^
 /SetupSQLRS:YES                             ^
 /SetupSQLIS:YES                             ^
 /SQLAccount:"ROOT\ServGB_SQLDB_0001"      /SQLPassword:"Argyt$6hsGGWMP894s4Gw2b73GS2o0" ^
 /AGTACCOUNT:"ROOT\ServGB_SQLAG_0001"      /AGTPASSWORD:"F6tbmd*nf!dfGFrcQnm84g4K7zwq2j" ^
 /ASACCOUNT:"ROOT\ServGB_SQLAS_0001"       /ASPASSWORD:"kE44bmutFGS579*bssJW84f=Rb6ehj"  ^
 /ISACCOUNT:"ROOT\ServGB_SQLIS_0001"       /ISPASSWORD:"bSHG5iuf9DFF#dw2!F5sKSIw43tnb7"  ^
 /SQLBROWSERACCOUNT:"ROOT\ServGB_SQLBR_0001"  /SQLBROWSERPASSWORD:"w#d6gh*ge$dvnHHq1knbtd$Wd68Zj9" ^
 /VolBackup:I /VolData:JF /VolDataFT:F     /VolLog:KG /VolProg:C /VolTemp:T              ^
 /SETUPCMDSHELL:YES ^
 /CMDSHELLACCOUNT:"ROOT\APPGB_SQLCS_0001"  /CMDSHELLPASSWORD:"j25Fb*ef$36ySIyBW7hZ" 