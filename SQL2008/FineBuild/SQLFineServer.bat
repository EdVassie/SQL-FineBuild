@ECHO OFF
REM Copyright FineBuild Team © 2008 - 2016.  Distributed under Ms-Pl License
REM
CALL "SqlFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms             ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01"      ^
 /SetupSQLDB:YES                             ^
 /SetupSQLAS:YES                             ^
 /SetupSQLRS:YES                             ^
 /SetupSQLIS:YES                             ^
 /SQLSVCAccount:"ROOT\ServGB_SQLDB_0001"     /SQLSVCPassword:"Argyt$6hsGGWMP894s4Gw2b73GS2o0" ^
 /AGTSVCACCOUNT:"ROOT\ServGB_SQLAG_0001"     /AGTSVCPASSWORD:"F6tbmd*nf!dfGFrcQnm84g4K7zwq2j" ^
 /ASSVCACCOUNT:"ROOT\ServGB_SQLAS_0001"      /ASSVCPASSWORD:"kE44bmutFGS579*bssJW84f=Rb6ehj"  ^
 /FTSVCACCOUNT:"ROOT\ServGB_SQLFT_0001"      /FTSVCPASSWORD:"w$Yhfb84nmkl5r*hsdFR7yNs2$ynd6"  ^
 /ISSVCACCOUNT:"ROOT\ServGB_SQLIS_0001"      /ISSVCPASSWORD:"bSHG5iuf9DFF#dw2!F5sKSIw43tnb7"  ^
 /RSSVCACCOUNT:"ROOT\ServGB_SQLRS_0001"      /RSSVCPASSWORD:"Orfd450!#DTWjn63hw45JDD873hk84"  ^
 /BROWSERSVCACCOUNT:"ROOT\ServGB_SQLBR_0001" /BROWSERSVCPASSWORD:"w#d6gh*ge$dvnHHq1knbtd$Wd68Zj9" ^
 /VolProg:C          ^
 /VolBackup:I /VolData:JF /VolDataFT:F /VolLog:KG /VolTemp:T                                  ^
 /SETUPCMDSHELL:YES  ^
 /CMDSHELLACCOUNT:"ROOT\APPGB_SQLCS_0001"    /CMDSHELLPASSWORD:"j25Fb*ef$36ySIyBW7hZ"         ^
 /SETUPRSEXEC:YES    ^
 /RSEXECACCOUNT:"ROOT\APPGB_SQLRS_0001"      /RSEXECPASSWORD:"Prf53g#fdf$Efbv8QGH3"
