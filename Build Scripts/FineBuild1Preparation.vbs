''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FineBuild1Preparation.vbs  
'  Copyright FineBuild Team © 2008 - 2018.  Distributed under Ms-Pl License
'  Code to clear IndexingEnabled flag adapted from "Windows Server Cookbook" by Robbie Allen, ISBN 0-596-00633-0
'
'  Purpose:      Builds directory structure and shares for use in a standard
'                SQL Server build as defined in the FineBuild Reference document.
'
'  Author:       Ed Vassie, based on work for SQL 2000 by Mark Allison
'
'  Date:         December 2007
'
'  Change History
'  Version  Author        Date         Description
'  2.3      Ed Vassie     18 Jun 2010  Initial SQL Server R2 version
'  2.2.2    Ed Vassie     29 Oct 2009  Added extra drives to support clustering
'  2.2.1    Ed vassie     28 Jun 2009  Added support for Express Edition
'  2.2      Ed Vassie      8 Oct 2008  Major rewrite for FineBuild V2.0.0
'  2.1      Ed Vassie     20 Feb 2008  Bypass create of SQLAS files for Workgroup Edition
'                                      Add configure of local groups
'  2.0      Ed Vassie     02 Feb 2008  Initial version for FineBuild v1.0.0
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim arrProfFolders, arrUserSid
Dim colSysEnvVars, colUsrEnvVars, colVol
Dim intErrSave, intIdx
Dim objADOCmd, objADOConn, objCluster, objDrive, objFile, objFolder, objFSO, objOS, objNetwork, objShell, objVol, objWMI, objWMIReg
Dim strAction, strActionDTC, strActionSQLAS, strActionSQLDB, strAgtAccount, strAlphabet, strAnyKey, strAsAccount, strAVCmd
Dim strClusIPVersion, strClusIPV4Network, strClusIPV6Network, strClusStorage, strClusterGroupAO, strClusterHost, strClusterAction, strClusterName, strClusterNameAS, strClusterNameRS, strClusterNameSQL, strClusterNode, strClusterRoot, strCmd, strCmdshellAccount, strCSVRoot
Dim strDomain, strDomainSID, strDriveList, strDirDBA, strDirProg, strDirProgX86, strDirSQL, strDirSys, strDirSystemDataBackup, strDirSystemDataShared, strDirProgSys, strDirProgSysX86, strDRUCtlrAccount, strDRUCltAccount
Dim strEdition, strExtSvcAccount, strFirewallStatus, strFolderName, strFSLevel, strFSShareName, strGroupAdmin, strGroupDBA, strGroupDBANonSA, strGroupDistComUsers, strGroupMSA, strGroupPerfLogUsers, strGroupPerfMonUsers, strGroupRDUsers, strGroupUsers, strHKLMFB, strHKU, strHTTP
Dim strFTAccount, strIsAccount, strIsMasterPort, strLocalAdmin, strLocalDomain, strNTAuthAccount, strNTService, strRsAccount, strRSAlias, strSqlAccount
Dim strInstance, strInstAgent, strInstLog, strInstNode, strInstNodeAS, strInstNodeIS, strInstRS, strIsMasterAccount, strIsWorkerAccount
Dim strMenuSSMS, strNetworkGUID, strOSName, strOSType, strOSVersion
Dim strPath, strPathFB, strPathFBScripts, strPathNew, strPathTemp, strPathTempUser, strPBPortRange, strPrepareFolderPath, strProcArc, strProfDir, strProgCacls, strProgNtrights, strProgSetSPN, strProgReg, strReboot
Dim strRSInstallMode, strSchedLevel, strSecDBA, strSecMain, strSecNull, strSecTemp, strServer, strServerSID, strServInst, strSetupAlwaysOn, strSetupNetBind, strSetupNetName, strSetupNoDefrag, strSetupNoDriveIndex, strSetupNoTCPNetBios, strSetupNoTCPOffload
Dim strInstRSURL, strInstSQL, strSetupPowerCfg, strSetupPolyBase, strSetupSPN, strSetupSQLAS, strSetupSQLASCluster, strSetupSQLDB, strSetupSQLDBCluster, strSetupSQLDBAG, strSetupSQLDBFS, strSetupSQLDBFT, strSetupSQLIS, strSetupSQLRS, strSetupSQLRSCluster, strSetupSQLTools
Dim strSetupWinAudit, strSetupBPE, strSetupCmdshell, strSetupISMaster, strSetupTempWin, strSetupNoWinGlobal, strSetupDRUClt, strDTCClusterRes, strDTCMultiInstance, strSetupDTCCluster, strSetupDTCClusterStatus, strSetupDTCNetAccess, strSetupFirewall, strSetupSP, strSetupSSISCluster
Dim strSetupMyDocs, strSetupShares, strSIDDistComUsers, strSPLevel, strSQLVersion, strSQLVersionNum, strSQLVersionNet, strSqlBrowserAccount
Dim strTCPPort, strTCPPortDTC, strTCPPortISMaster, strType, strUserAccount, strUserName, strUserDNSDomain
Dim strLabBackup, strLabBackupAS, strLabBPE, strLabData, strLabDataAS, strLabDataFS, strLabDataFT, strLabDTC, strLabLog, strLabLogAS, strLabLogTemp, strLabPrefix, strLabProg, strLabSysDB, strLabSystem, strLabTemp, strLabTempAS, strLabTempWin, strLabDBA
Dim strSpace
Dim strVol, strVolType, strVolUsed, strVolBackup, strVolBackupAS, strVolBPE, strVolData, strVolDataAS, strVolDataFS, strVolDataFT, strVolDTC, strVolLog, strVolLogAS, strVolSysDB, strVolLogTemp, strVolTemp, strVolTempAS, strVolTempWin, strVolProg, strVolSys, strVolDBA
Dim strWaitLong, strWaitShort

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strProcessId >= "1TZ"
      ' Nothing
    Case Else
      Call PreparationTasks()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1UZ" ' 1UA to 1UZ reserved for User Preparation processing
      ' Nothing
    Case Else
      Call UserPreparation()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination
  
  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case strProcessId > "1TZ"
      ' Nothing
    Case err.Number = 0 
      Call objShell.Popup("SQL Server install preparation complete", 2, "Preparation processing" ,64)
    Case Else
      Call FBLog("***** Error has occurred *****")
      If strProcessIdLabel <> "" Then
        Call FBLog(" Process    : " & strProcessIdLabel & ": " & strProcessIdDesc)
      End If
      If err.Number <> "" Then
        Call FBLog(" Error code : " & err.Number)
      End If
      If err.Source <> "" Then
        Call FBLog(" Source     : " & err.Source)
      End If
      If err.Description <> "" Then
        Call FBLog(" Description: " & err.Description)
      End If
      If strDebugDesc <> "" And strDebugDesc <> err.Description Then
        Call FBLog(" Last Action: " & strDebugDesc)
      End If
      If strDebugMsg1 <> "" Then
        Call FBLog(" " & strDebugMsg1)
      End If
      If strDebugMsg2 <> "" Then
        Call FBLog(" " & strDebugMsg2)
      End If
      Call FBLog(" SQL Server install preparation failed")
  End Select

  Wscript.Quit(err.Number)

End Sub


Sub Initialisation()
' Perform initialisation procesing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"
  Include "FBManageInstall.vbs"
  Include "FBUtils.vbs"
  Call SetProcessIdCode("FB1P")

  Set objADOConn    = CreateObject("ADODB.Connection")
  Set objADOCmd     = CreateObject("ADODB.Command")
  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objNetwork    = CreateObject("Wscript.Network")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colSysEnvVars = objShell.Environment("System")
  Set colUsrEnvVars = objShell.Environment("User")

  objADOConn.Provider            = "ADsDSOObject"
  objADOConn.Open "ADs Provider"
  Set objADOCmd.ActiveConnection = objADOConn

  strHKLMFB         = GetBuildfileValue("HKLMFB")
  strHKU            = &H80000003
  strSpace          = Space(20)
  strAction         = GetBuildfileValue("Action")
  strActionDTC      = GetBuildfileValue("ActionDTC")
  strActionSQLAS    = GetBuildfileValue("ActionSQLAS")
  strActionSQLDB    = GetBuildfileValue("ActionSQLDB")
  strAgtAccount     = GetBuildfileValue("AgtAccount")
  strAlphabet       = GetBuildfileValue("Alphabet")
  strAnyKey         = GetBuildfileValue("AnyKey")
  strAsAccount      = GetBuildfileValue("AsAccount")
  strAVCmd          = GetBuildfileValue("AVCmd")
  strSqlBrowserAccount = GetBuildfileValue("SqlBrowserAccount")
  strClusterHost    = GetBuildfileValue("ClusterHost")
  strClusterAction  = GetBuildfileValue("ClusterAction")
  strClusterGroupAO = GetBuildfileValue("ClusterGroupAO")
  strClusterName    = GetBuildfileValue("ClusterName")
  strClusterNameAS  = GetBuildfileValue("ClusterNameAS")
  strClusterNameRS  = GetBuildfileValue("ClusterNameRS")
  strClusterNameSQL = GetBuildfileValue("ClusterNameSQL")
  strClusterNode    = GetBuildfileValue("ClusterNode")
  strClusterRoot    = GetBuildfileValue("ClusterRoot")
  strClusIPVersion  = GetBuildfileValue("ClusIPVersion")
  strClusIPV4Network  = GetBuildfileValue("ClusIPV4Network")
  strClusIPV6Network  = GetBuildfileValue("ClusIPV6Network")
  strClusStorage    = GetBuildfileValue("ClusStorage")
  strCmdshellAccount  = GetBuildfileValue("CmdshellAccount")
  strCSVRoot        = GetBuildfileValue("CSVRoot")
  strDirDBA         = GetBuildfileValue("DirDBA")
  strDirProg        = GetBuildfileValue("DirProg")
  strDirProgX86     = GetBuildfileValue("DirProgX86")
  strDirProgSys     = GetBuildfileValue("DirProgSys")
  strDirProgSysX86  = GetBuildfileValue("DirProgSysX86")
  strDirSQL         = GetBuildfileValue("DirSQL")
  strDirSys         = GetBuildfileValue("DirSys")
  strDirSystemDataBackup = GetBuildfileValue("DirSystemDataBackup")
  strDirSystemDataShared = GetBuildfileValue("DirSystemDataShared")
  strDomain         = GetBuildfileValue("Domain")
  strDomainSID      = GetBuildfileValue("DomainSID")
  strPath           = Mid(strHKLMFB, 6)
  strDRUCtlrAccount = GetBuildfileValue("DRUCtlrAccount")
  strDRUCltAccount  = GetBuildfileValue("DRUCltAccount")
  objWMIReg.GetStringValue strHKLM,strPath,"DTCClusterRes",strDTCClusterRes
  strDTCMultiInstance = GetBuildfileValue("DTCMultiInstance")
  strEdition        = GetBuildfileValue("AuditEdition")
  strExtSvcAccount  = GetBuildfileValue("ExtSvcAccount")
  strFirewallStatus = GetBuildfileValue("FirewallStatus")
  strFSShareName    = GetBuildfileValue("FSShareName")
  strFSLevel        = GetBuildfileValue("FSLevel")
  strFTAccount      = GetBuildfileValue("FtAccount")
  strGroupAdmin     = GetBuildfileValue("GroupAdmin")
  strGroupDBA       = GetBuildfileValue("GroupDBA")
  strGroupDBANonSA  = GetBuildfileValue("GroupDBANonSA")
  strGroupDistComUsers = GetBuildfileValue("GroupDistComUsers")
  strGroupMSA       = GetBuildfileValue("GroupMSA")
  strGroupPerfLogUsers = GetBuildfileValue("GroupPerfLogUsers")
  strGroupPerfMonUsers = GetBuildfileValue("GroupPerfMonUsers")
  strGroupRDUsers   = GetBuildfileValue("GroupRDUsers")
  strGroupUsers     = GetBuildfileValue("GroupUsers")
  strHTTP           = GetBuildfileValue("HTTP")
  strInstance       = GetBuildfileValue("Instance")
  strInstLog        = GetBuildfileValue("InstLog")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstNodeAS     = GetBuildfileValue("InstNodeAS")
  strInstNodeIS     = GetBuildfileValue("InstNodeIS")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strIsAccount      = GetBuildfileValue("IsAccount")
  strIsMasterAccount  = GetBuildfileValue("IsMasterAccount")
  strIsMasterPort   = GetBuildfileValue("IsMasterPort")
  strIsWorkerAccount  = GetBuildfileValue("IsWorkerAccount")
  strLabBackup      = GetBuildfileValue("LabBackup")
  strLabBackupAS    = GetBuildfileValue("LabBackupAS")
  strLabBPE         = GetBuildfileValue("LabBPE")
  strLabData        = GetBuildfileValue("LabData")
  strLabDataAS      = GetBuildfileValue("LabDataAS")
  strLabDataFS      = GetBuildfileValue("LabDataFS")
  strLabDataFT      = GetBuildfileValue("LabDataFT")
  strLabDBA         = GetBuildfileValue("LabDBA")
  strLabDTC         = GetBuildfileValue("LabDTC")
  strLabLog         = GetBuildfileValue("LabLog")
  strLabLogAS       = GetBuildfileValue("LabLogAS")
  strLabLogTemp     = GetBuildfileValue("LabLogTemp")
  strLabSysDB       = GetBuildfileValue("LabSysDB")
  strLabPrefix      = GetBuildfileValue("LabPrefix")
  strLabProg        = GetBuildfileValue("LabProg")
  strLabSystem      = GetBuildfileValue("LabSystem")
  strLabTemp        = GetBuildfileValue("LabTemp")
  strLabTempAS      = GetBuildfileValue("LabTempAS")
  strLabTempWin     = GetBuildfileValue("LabTempWin")
  strLocalAdmin     = GetBuildfileValue("LocalAdmin")
  strLocalDomain    = GetBuildfileValue("LocalDomain")
  strNetworkGUID    = GetBuildfileValue("NetworkGUID")
  strNTAuthAccount  = GetBuildfileValue("NTAuthAccount")
  strNTService      = GetBuildfileValue("NTService")
  strOSName         = GetBuildfileValue("OSName")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPathFBScripts  = FormatFolder("PathFBScripts")
  strPathTemp       = GetBuildfileValue("PathTemp")
  strPathTempUser   = GetBuildfileValue("PathTempUser")
  strPBPortRange    = GetBuildfileValue("PBPortRange")
  strPrepareFolderPath = ""
  strProcArc        = GetBuildfileValue("ProcArc")
  strProfDir        = GetBuildfileValue("ProfDir")
  strProgCacls      = GetBuildfileValue("ProgCacls")
  strProgNtrights   = GetBuildfileValue("ProgNTRights")
  strProgSetSPN     = GetBuildfileValue("ProgSetSPN")
  strProgReg        = GetBuildfileValue("ProgReg")
  strReboot         = GetBuildfileValue("RebootStatus")
  strRsAccount      = GetBuildfileValue("RsAccount")
  strRSAlias        = GetBuildfileValue("RSAlias")
  strRSInstallMode  = GetBuildfileValue("RSInstallMode")
  strSchedLevel     = GetBuildfileValue("SchedLevel")
  strSecDBA         = GetBuildfileValue("SecDBA")
  strSecMain        = GetBuildfileValue("SecMain")
  strSecNull        = ""
  strSecTemp        = GetBuildfileValue("SecTemp")
  strServerSID      = GetBuildfileValue("ServerSID")
  strServInst       = GetBuildfileValue("ServInst")
  strSetupAlwaysOn  = GetBuildfileValue("SetupAlwaysOn")
  strSetupBPE       = GetBuildfileValue("SetupBPE")
  strSetupCmdshell  = GetBuildfileValue("SetupCmdshell")
  strSetupDRUClt    = GetBuildfileValue("SetupDRUClt")
  strSetupDTCCluster   = GetBuildfileValue("SetupDTCCluster")
  strSetupDTCNetAccess = GetBuildfileValue("SetupDTCNetAccess")
  strSetupFirewall  = GetBuildfileValue("SetupFirewall")
  strSetupISMaster  = GetBuildfileValue("SetupISMaster")
  strSetupMyDocs    = GetBuildfileValue("SetupMyDocs")
  strSetupNetBind   = GetBuildfileValue("SetupNetBind")
  strSetupNetName   = GetBuildfileValue("SetupNetName")
  strSetupNoDefrag  = GetBuildfileValue("SetupNoDefrag")
  strSetupNoDriveIndex = GetBuildfileValue("SetupNoDriveIndex")
  strSetupNoTCPNetBios = GetBuildfileValue("SetupNoTCPNetBios")
  strSetupNoTCPOffload = GetBuildfileValue("SetupNoTCPOffload")
  strSetupNoWinGlobal  = GetBuildfileValue("SetupNoWinGlobal")
  strSetupPolyBase  = GetBuildfileValue("SetupPolyBase")
  strSetupPowerCfg  = GetBuildfileValue("SetupPowerCfg")
  strSetupShares    = GetBuildfileValue("SetupShares")
  strSetupSPN       = GetBuildfileValue("SetupSPN")
  strSetupSQLASCluster = GetBuildfileValue("SetupSQLASCluster")
  strSetupSQLAS     = GetBuildfileValue("SetupSQLAS")
  strSetupSQLDB     = GetBuildfileValue("SetupSQLDB")
  strSetupSQLDBCluster = GetBuildfileValue("SetupSQLDBCluster")
  strSetupSQLDBAG   = GetBuildfileValue("SetupSQLDBAG")
  strSetupSQLDBFS   = GetBuildfileValue("SetupSQLDBFS")
  strSetupSQLDBFT   = GetBuildfileValue("SetupSQLDBFT")
  strSetupSQLIS     = GetBuildfileValue("SetupSQLIS")
  strSetupSQLTools  = GetBuildfileValue("SetupSQLTools")
  strSetupSQLRS     = GetBuildfileValue("SetupSQLRS")
  strSetupSQLRSCluster = GetBuildfileValue("SetupSQLRSCluster")
  strSetupSP        = GetBuildfileValue("SetupSP")
  strSetupSSISCluster = GetBuildfileValue("SetupSSISCluster")
  strSetupTempWin   = GetBuildfileValue("SetupTempWin")
  strSetupWinAudit  = GetBuildfileValue("SetupWinAudit")
  strServer         = GetBuildfileValue("AuditServer")
  strSPLevel        = GetBuildfileValue("SPLevel")
  strSqlAccount     = GetBuildfileValue("SqlAccount")
  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strSQLVersionNet  = GetBuildfileValue("SQLVersionNet")
  strSQLVersionNum  = GetBuildfileValue("SQLVersionNum")
  strTCPPort        = GetBuildfileValue("TCPPort")
  strTCPPortDTC     = GetBuildfileValue("TCPPortDTC")
  strType           = GetBuildfileValue("Type")
  strUserAccount    = GetBuildfileValue("UserAccount")
  strUserDnsDomain  = GetBuildfileValue("UserDNSDomain")
  strUserName       = GetBuildfileValue("AuditUser")
  strDriveList      = GetBuildfileValue("DriveList")
  strVolProg        = GetBuildfileValue("VolProg")
  strVolBackup      = GetBuildfileValue("VolBackup")
  strVolBackupAS    = GetBuildfileValue("VolBackupAS")
  strVolData        = GetBuildfileValue("VolData")
  strVolDataAS      = GetBuildfileValue("VolDataAS")
  strVolDataFS      = GetBuildfileValue("VolDataFS")
  strVolDataFT      = GetBuildfileValue("VolDataFT")
  strVolDBA         = GetBuildfileValue("VolDBA")
  strVolDTC         = GetBuildfileValue("VolDTC")
  strVolLog         = GetBuildfileValue("VolLog")
  strVolLogAS       = GetBuildfileValue("VolLogAS")
  strVolLogTemp     = GetBuildfileValue("VolLogTemp")
  strVolSys         = GetBuildfileValue("VolSys")
  strVolSysDB       = GetBuildfileValue("VolSysDB")
  strVolTemp        = GetBuildfileValue("VolTemp")
  strVolTempAS      = GetBuildfileValue("VolTempAS")
  strVolBPE         = GetBuildfileValue("VolBPE")
  strVolTempWin     = GetBuildfileValue("VolTempWin")
  strVolUsed        = ""
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitShort      = GetBuildfileValue("WaitShort")
  Set arrProfFolders  = objFSO.GetFolder(strProfDir).SubFolders

End Sub


Sub PreparationTasks()
  Call SetProcessId("1", strSQLVersion & " Preparation processing (FineBuild1Preparation.vbs)")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case Else
      Call SetupFineBuild()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1BZ"
      ' Nothing
    Case Else
      Call SetupServer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CZ"
      ' Nothing
    Case Else
      Call SetupWindows()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DZ"
      ' Nothing
    Case Else
      Call SetupNetwork()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EZ"
      ' Nothing
    Case Else
      Call SetupAccounts()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1FZ"
      ' Nothing
    Case Else
      Call SetupDrives()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GZ"
      ' Nothing
    Case Else
      Call SetupFolders()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1HZ"
      ' Nothing
    Case Else
      Call PostPreparation()
  End Select

  Call SetProcessId("1TZ", " Preparation processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupFineBuild()
  Call SetProcessId("1A", "Setup FineBuild")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case Else
      Call GetClusterDetails()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1AB"
      ' Nothing
    Case strClusterAction = "" 
      ' Nothing
    Case Else
      Call SetupClusterStorageGroup()
  End Select

  Call SetProcessId("1AZ", " Setup FineBuild" & strStatusComplete)
  Call ProcessEnd("")

End Sub 


Sub GetClusterDetails()
  Call SetProcessId("1AA", "Get Cluster details")
  On Error Resume Next

  If strClusterHost = "YES" Then
    Set objCluster  = CreateObject("MSCluster.Cluster")
    objCluster.Open ""
  End If

End Sub


Sub SetupClusterStorageGroup()
  Call SetProcessId("1AB", "Setup Cluster Storage Group")

  strCmd            = "CLUSTER """ & strClusterName & """ GROUP """ & strClusStorage & """ /CREATE"
  Call Util_RunExec(strCmd, "", strResponseYes, 5010)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupServer()
  Call SetProcessId("1B", "Setup Server")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1BA"
      ' Nothing
    Case Else
      Call SetupServerName()
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub 


Sub SetupServerName()
  Call SetProcessId("1BA", "Setup Server Name")
  Dim colHostname
  Dim objHostname
  Dim strHostname

  strHostname       = objShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Hostname")
  If StrComp(strHostname, UCase(strHostname), vbBinaryCompare) <> 0 Then
    Call DebugLog("Change Hostname to upper case")
    Set colHostname = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each objHostname In colHostname
      Call DebugLog("Set server name to upper case " & objHostname.Name)
      Call objHostname.Rename(Ucase(objHostname.Name))
    Next
    Call SetBuildfileValue("RebootStatus", "Pending") 
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupWindows()
  Call SetProcessId("1C", "Setup Windows")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CA"
      ' Nothing
    Case Else
      Call SetupServiceTimeout()
  End Select

  ' ProcessId 1CB available for reuse

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CC"
      ' Nothing
    Case strSetupPowerCfg <> "YES"
      ' Nothing
    Case Else
      Call SetupPowerCfg()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CD"
      ' Nothing
    Case strSetupNoDefrag <> "YES"
      ' Nothing
    Case Else
      Call SetupNoDefrag()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CE"
      ' Nothing
    Case strSetupWinAudit <> "YES"
      ' Nothing
    Case strOSVersion < "6.0"
      Call SetBuildfileValue("SetupWinAuditStatus", strStatusManual)
    Case Else
      Call SetupWinAudit()
  End Select

  Call SetProcessId("1CZ", " Setup Windows" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupServiceTimeout()
  Call SetProcessId("1CA", "Setup Service Timeout")
  Dim intTime, intSpeedTest

  intSpeedTest      = GetBuildfileValue("SpeedTest")
  intTime           = GetBuildfileValue("BuildFileTime") 
  Select Case True 
    Case CDbl(intTime) <= CDbl(intSpeedTest)
      ' Nothing
    Case Else 
      intTime       = Cstr((Int(intTime) + 1) * 10000) ' Increase service startup time 1/10 second for every second of Buildfile time.
      Call SetServicePipeTimeOut(intTime, "Slow system detected, service start time allowance increased.")
  End Select

  Select Case True
    Case strSetupSQLRS <> "YES"
      ' Nothing
   Case strSQLVersion < "SQL2008"
      ' Nothing
    Case strSQLVersion > "SQL2008R2"
      ' Nothing
    Case strOSVersion >= "6.2"
      ' Nothing
    Case Else
      Call SetServicePipeTimeOut("60000", "Service start time allowance increased to 1 minute for Reporting Services.")   
  End Select

  Select Case True
    Case strSetupPolybase <> "YES"
      ' Nothing
    Case Else
      Call SetServicePipeTimeOut("60000", "Service start time allowance increased to 1 minute for PolyBase.")   
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetServicePipeTimeOut(intTime, strMsg)
  Call DebugLog("SetServicePipeTimeOut:")
  Dim intTimeout

  strPath           = "SYSTEM\CurrentControlSet\Control"
  objWMIReg.GetDwordValue strHKLM, strPath, "ServicesPipeTimeout", intTimeout
  If IsNull(intTimeout) Then
    intTimeout      = 30000
  End If

  If CLng(intTimeout) < CLng(intTime) Then
    strPath         = "HKLM\" & strPath & "\ServicesPipeTimeout"
    Call SetBuildMessage(strMsgInfo, strMsg)
    Call DebugLog("Adjusting " & strPath & " from " & Cstr(intTimeout) & " to " & CStr(intTime) & " milliseconds")
    Call Util_RegWrite(strPath, intTime, "REG_DWORD") 
    Call SetBuildfileValue("RebootStatus", "Pending")   ' Reboot needed so new Timeout can take effect 
  End If

End Sub


Sub SetupCertificateLog()
  Call DebugLog("SetupCertificateLog:")
' Described in KB2661254, SQL Self-Signs with 1024 bit Certificates and this change allows them to be accepted by Windows

  strPath           = strDirSys & "\Logs\CertLog"
  If Not objFSO.FolderExists(strPath) Then
    objFSO.CreateFolder(strPath)
    WScript.Sleep strWaitShort
  End If
  strCmd            = "CERTUTIL -SETREG chain\WeakSignatureLogDir """ & strPath & "\Under1024Key.Log"""
  Call Util_RunExec(strCmd, "", strResponseYes, 0)
  strCmd            = "CERTUTIL -SETREG chain\EnableWeakSignatureFlags 8"
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

End Sub


Sub SetupPowerCfg()
  Call SetProcessId("1CC", "Setup Windows Power Configuration")
  Dim strPowerScheme

  strPath           = "SOFTWARE\Policies\Microsoft\Power\PowerSettings"
  objWMIReg.GetStringValue strHKLM,strPath,"ActivePowerScheme",strPowerScheme

  Select Case True
    Case strPowerScheme = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
      Call SetBuildfileValue("SetupPowerCfgStatus", strStatusComplete)
    Case Else
      strPowerScheme = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
      strPath        = "HKLM\" & strPath & "\ActivePowerScheme"
      Call Util_RegWrite(strPath, strPowerScheme, "REG_SZ") 
      Call SetBuildfileValue("SetupPowerCfgStatus", strStatusComplete)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNoDefrag()
  Call SetProcessId("1CD", "Setup No Disk Defragmentation")

  strCmd            = "SCHTASKS /Change /tn ""Microsoft/Windows/Defrag/ScheduledDefrag"" /DISABLE"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)

  Call SetBuildfileValue("SetupNoDefragStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupWinAudit()
  Call SetProcessId("1CE", "Setup Windows Audit")

  Call Util_RunExec("AUDITPOL /set /Category:""Account Logon""      /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Account Management"" /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""DS Access""          /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Logon/Logoff""       /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Object Access""      /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Policy Change""      /success:enable",                   "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Privilege Use""      /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Detailed Tracking""  /success:disable /failure:disable", "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""System""             /success:enable",                   "", "", 0) 

  Call SetBuildfileValue("SetupWinAuditStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNetwork()
  Call SetProcessId("1D", "Setup SQL Server Network")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DA"
      ' Nothing
    Case strSetupFirewall <> "YES"
      ' Nothing
    Case Else
      Call SetupFireWall()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DB"
      ' Nothing
    Case strSetupNetName <> "YES"
      ' Nothing
    Case strClusterHost <> "YES"
      Call SetBuildfileValue("SetupNetNameStatus", strStatusBypassed)
    Case Else
      Call SetupNetName()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DC"
      ' Nothing
    Case strSetupNetBind <> "YES"
      ' Nothing
    Case strClusterHost <> "YES"
      Call SetBuildfileValue("SetupNetBindStatus", strStatusBypassed)
    Case Else
      Call SetupNetBind()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DD"
      ' Nothing
    Case Else
      Call SetupAdapter()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DE"
      ' Nothing
    Case strSetupNoTCPNetBios <> "YES"
      ' Nothing
    Case Else
      Call SetupNoTCPNetBios()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DF"
      ' Nothing
    Case strSetupNoTCPOffload <> "YES"
      ' Nothing
    Case Else
      Call SetupNoTCPOffload()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DG"
      ' Nothing
    Case GetBuildfileValue("SetupTLS12") <> "YES"
      ' Nothing
    Case Else
      Call SetupTLS12()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DH"
      ' Nothing
    Case GetBuildfileValue("SetupNoSSL3") <> "YES"
      ' Nothing
    Case Else
      Call SetupNoSSL3()
  End Select

  Call SetProcessId("1DZ", " Setup SQL Server Network" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupFirewall()
  Call SetProcessId("1DA", "Setup Firewall")

  Select Case True
    Case strClusterAction = ""
      ' Nothing
    Case Else
      Call OpenPort("Failover Clusters (UDP-In)",     "3343", "UDP", "IN", "")
  End Select

  Call OpenPort("RPC Endpoint Mapper",                "135",  "TCP", "IN", "")

  If strSetupISMaster = "YES" Then
    Call OpenPort("IS Master",                        strIsMasterPort, "TCP", "IN", "")
  End If

  If strSetupSQLDB = "YES" Then
    Call OpenPort("SQL Server (" & strInstance & ")", strTCPPort,    "TCP", "IN", "")
    Call OpenPort("SQL DAC",                          GetBuildfileValue("TCPPortDAC"), "TCP", "IN", "")
    Call OpenPort("SQL Service Broker",               "4022", "TCP", "IN", "")
    Call OpenPort("SQL DB Mirroring",                 GetBuildfileValue("TCPPortAO"),  "TCP", "IN", "")
    If strInstance <> "MSSQLSERVER" Then
      Call OpenPort("SQL Browser",                    "1434", "UDP", "IN", "")
    End If
  End If

  If strSetupPolyBase = "YES" Then
    Call OpenPort("PolyBase",                         strPBPortRange, "TCP", "IN", "")
  End If

  If strSetupSQLAS = "YES" Then
    Call OpenPort("SQL Analysis Server",              GetBuildfileValue("TCPPortAS"),  "TCP", "IN", "")
    If strInstance <> "MSSQLSERVER" Then
      Call OpenPort("SQL Browser",                    "2382", "TCP", "IN", "")
    End If
  End If

  Select Case True
    Case strSetupSQLDBFS <> "YES"
      ' Nothing
    Case strFSLevel < "2"
      ' Nothing
    Case Else
      Call OpenPort("SQL Filestream",                 "139",  "TCP", "IN", "")
      Call OpenPort("SQL Filestream",                 "145",  "TCP", "IN", "")
  End Select

  If strSetupSQLRS = "YES" Then
    Call OpenPort("HTTP",                             "80",   "TCP", "IN", "")
  End If

  Select Case True
    Case strSQLVersion > "SQL2005"
      ' Nothing
    Case strClusterAction = "" 
      ' Nothing
    Case Else
      Call OpenPort("SQL Server Setup",               "",    "TCP", "IN", strVolSys & ":\Program Files\Microsoft SQL Server\90\Setup Bootstrap\setup.exe")
      Call OpenPort("SQL Server Setup",               "",    "UDP", "IN", strVolSys & ":\Program Files\Microsoft SQL Server\90\Setup Bootstrap\setup.exe")
  End Select

  Call DebugLog("Add Firewall Exception for DTC")
  Select Case True
    Case (strSetupDTCNetAccess <> "YES") And (strClusterHost <> "YES")
      ' Nothing
    Case Left(strOSVersion, 1) >= "6"
      strCmd        = "NETSH ADVFIREWALL FIREWALL SET RULE NAME=""Distributed Transaction Coordinator (RPC)"" "
      strCmd        = strCmd & "NEW PROFILE=DOMAIN ENABLE=YES"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      strCmd        = "NETSH ADVFIREWALL FIREWALL SET RULE NAME=""Distributed Transaction Coordinator (RPC-EPMAP)"" "
      strCmd        = strCmd & "NEW PROFILE=DOMAIN ENABLE=YES"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      strCmd        = "NETSH ADVFIREWALL FIREWALL SET RULE NAME=""Distributed Transaction Coordinator (TCP-In)"" "
      strCmd        = strCmd & "NEW PROFILE=DOMAIN ENABLE=YES"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      strCmd        = "NETSH ADVFIREWALL FIREWALL SET RULE NAME=""Distributed Transaction Coordinator (TCP-Out)"" "
      strCmd        = strCmd & "NEW PROFILE=DOMAIN ENABLE=YES"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
    Case Else
      strCmd        = "NETSH FIREWALL ADD ALLOWEDPROGRAM NAME=""MSDTC"" "
      strCmd        = strCmd & "PROGRAM=""" & strDirSys & "\system32\msdtc.exe"" "
      strCmd        = strCmd & "MODE=ENABLE SCOPE=ALL PROFILE=DOMAIN"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

  Call SetBuildfileValue("SetupFirewallStatus", strStatusComplete)  
  Call ProcessEnd(strStatusComplete)

End Sub


Sub OpenPort (strFWName, strFWPort, strFWType, strFWDir, strFWProgram)
  Call DebugLog("OpenPort: " & strFWName & " for " & strFWPort)

  Select Case True
    Case strFirewallStatus <> "1"
      ' Nothing
    Case Left(strOSVersion, 1) < "6"
      strCmd        = "NETSH FIREWALL ADD PORTOPENING NAME=""" & strFWName & """ "
      strCmd        = strCmd & "PROTOCOL=" & strFWType & " MODE=ENABLE SCOPE=ALL PROFILE=DOMAIN "
      If strFWPort <> "" Then
        strCmd      = strCmd & "PORT=" & strFWPort & " "
      End If
      If strFWProgram <> "" Then
        strCmd      = strCmd & "PROGRAM=""" & strFWProgram & """ "
      End If
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
    Case Else
      strCmd        = "NETSH ADVFIREWALL FIREWALL ADD RULE NAME=""" & strFWName & """ "
      strCmd        = strCmd & "PROTOCOL=" & strFWType & " ACTION=ALLOW PROFILE=DOMAIN DIR=" & strFWDir & " "
      If strFWPort <> "" Then
        strCmd      = strCmd & "LOCALPORT=" & strFWPort & " "
      End If
      If strFWProgram <> "" Then
        strCmd      = strCmd & "PROGRAM=""" & strFWProgram & """ "
      End If
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

End Sub


Sub SetupNetName()
  Call SetProcessId("1DB", "Setup Network Names")
  Dim arrInterfaces, arrNetworks
  Dim colInterface, colNetwork
  Dim strInterfaceName, strNetNameSource, strNetworkName, strPathAdapter, strPathNetworks, strPathInterface, strPathInterfaces
  Dim intIdx, intIdxNew

  strNetNameSource  = GetBuildfileValue("NetNameSource")
  strPathNetworks   = "HKLM\Cluster\Networks\"
  objWMIReg.EnumKey strHKLM, Mid(strPathNetworks, 6), arrNetworks
  strPathInterfaces = "HKLM\Cluster\NetworkInterfaces\"
  objWMIReg.EnumKey strHKLM, Mid(strPathInterfaces, 6), arrInterfaces
  
  For Each colNetwork In arrNetworks
    strPath         = strPathNetworks & colNetwork & "\Name"
    strNetworkName  = objShell.RegRead(strPath)
    Call DebugLog("Processing Network " & strNetworkName)
    For Each colInterface In arrInterfaces
      strPathInterface = strPathInterfaces & colInterface
      Select Case True
        Case objShell.RegRead(strPathInterface & "\Network") <> colNetwork
          ' Nothing
        Case objShell.RegRead(strPathInterface & "\Node") <> strClusterNode
          ' Nothing
        Case Else
          strPathAdapter   = objShell.RegRead(strPathInterface & "\AdapterId")
          strInterfaceName = objShell.RegRead("HKLM\System\CurrentControlSet\Control\Network\{" & strNetworkGUID &"}\{" & strPathAdapter & "}\Connection\Name")
          Select Case True
            Case strInterfaceName = strNetworkName
              ' Nothing
            Case (strNetNameSource = "CLUSTER") Or (strClusterAction = "ADDNODE")
              strCmd = "NETSH INTERFACE SET INTERFACE NAME=""" & strInterfaceName & """ NEWNAME=""" & strNetworkName & """ "
              Call Util_RunExec(strCmd, "", strResponseYes, 0)
              Call DebugLog(" Network Adapter '" & strInterfaceName & "' renamed to '" & strNetworkName & "'")
              Wscript.Sleep strWaitShort 
            Case Else
              strCmd = "CLUSTER " & strClusterName & " NETWORK """ & strNetworkName & """ /RENAME:""" & strInterfaceName & """ "
              Call Util_RunExec(strCmd, "", strResponseYes, 0)
              Call DebugLog(" Cluster Network '" & strNetworkName & "' renamed to '" & strInterfaceName & "'")
              If strClusIPV4Network = strNetworkName Then
                strClusIPV4Network = strInterfaceName
                Call SetBuildfileValue("ClusIPV4Network", strClusIPV4Network)
              End If
              If strClusIPV6Network = strNetworkName Then
                strClusIPV6Network = strInterfaceName
                Call SetBuildfileValue("ClusIPV6Network", strClusIPV6Network)
              End If
              Wscript.Sleep strWaitShort 
          End Select
      End Select
    Next 
  Next

  Call SetBuildfileValue("SetupNetNameStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNetBind()
  Call SetProcessId("1DC", "Setup Network Bindings")
  Dim arrInterfaces
  Dim colInterface
  Dim strAdapter, strNetwork, strNetworkRole, strNode, strPathInterface, strPathInterfaces, strPathNetwork

  strPathInterfaces = "HKLM\Cluster\NetworkInterfaces\"
  objWMIReg.EnumKey strHKLM, Mid(strPathInterfaces, 6), arrInterfaces
  For Each colInterface In arrInterfaces
    strPathInterface = strPathInterfaces & colInterface
    strNetwork      = objShell.RegRead(strPathInterface & "\Network")
    strNode         = objShell.RegRead(strPathInterface & "\Node")
    strPathNetwork  = "HKLM\Cluster\Networks\" & strNetwork
    strNetworkRole  = objShell.RegRead(strPathNetwork & "\Role")
    Select Case True
      Case strNode <> strClusterNode
        ' Nothing
      Case CStr(strNetworkRole) < "2"
        ' Nothing
      Case Else
        strAdapter  = objShell.RegRead(strPathInterface & "\AdapterId")
        Call SetBindingOrder("IPv4", strAdapter)
        Call SetBindingOrder("IPv6", strAdapter)
    End Select
  Next 

  Call SetBuildfileValue("SetupNetBindStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetBindingOrder(strTCPVersion, strAdapter)
  Call DebugLog("SetBindingOrder: " & strTCPVersion & " for " & strAdapter)
  Dim arrBindings
  Dim bFound
  Dim intBind, intIdx, intIdxNew
  Dim strAdapterBind

  Select Case True
    Case strTCPVersion = "IPv6"
      strPath       = "HKLM\SYSTEM\CurrentControlSet\Services\TCPIP6\Linkage\"
    Case Else
      strPath       = "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Linkage\"
  End Select
  objWMIReg.GetMultiStringValue strHKLM, Mid(strPath, 6), "Bind", arrBindings
  If IsNull(arrBindings) Then
    Exit Sub
  End If

  intBind           = Ubound(arrBindings)
  bFound            = False
  strAdapterBind    = "\Device\{" & strAdapter & "}"
  For intIdx = 0 To intBind
    If arrBindings(intIdx) = strAdapterBind Then
      bFound        = True
    End If
  Next
  If Not bFound Then
    Exit Sub
  End If

  Call DebugLog("Checking Bindings")
  ReDim arrBindingsNew(intBind)
  arrBindingsNew(0) = strAdapterBind
  intIdxNew         = 1
  strDebugMsg2      = "IdxNew: " & cStr(intIdxNew)
  For intIdx = 0 To intBind
    strDebugMsg1    = "Idx: " & cStr(intIdx)
    Select Case True
      Case arrBindings(intIdx) = strAdapterBind
        If intIdx <> 0 Then
          Call FBLog(" TCP " & strTCPVersion & " Network bind order corrected")
        End If
      Case Else
        arrBindingsNew(intIdxNew) = arrBindings(intIdx)
        intIdxNew    = intIdxNew + 1
        strDebugMsg2 = "IdxNew: " & cStr(intIdxNew)
    End Select    
  Next
  objWMIReg.SetMultiStringValue strHKLM, Mid(strPath, 6), "Bind", arrBindingsNew
 
End Sub


Sub SetupAdapter()
  Call SetProcessId("1DD", "Setup Network Adapter Parameters")
  Dim arrInterfaces
  Dim intIdx
  Dim strNameServer, strDomain, strPathV4

  strPathV4         =  "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\"
  objWMIReg.EnumKey strHKLM, strPathV4, arrInterfaces

  For intIdx = 0 To Ubound(arrInterfaces)
    objWMIReg.GetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "NameServer", strNameServer
    objWMIReg.GetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "Domain",     strDomain
    If (strNameServer > "") And (Not (strDomain > "")) Then
      Call DebugLog("Processing Adapter " & arrInterfaces(intIdx))
      objWMIReg.SetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "Domain", strUserDNSDomain
    End If
  Next

  If strOSVersion >= "6.0" Then
    strPath         =  "SYSTEM\CurrentControlSet\Services\TCPIP6\Parameters\Interfaces\"
    objWMIReg.EnumKey strHKLM, strPath, arrInterfaces

    For intIdx = 0 To Ubound(arrInterfaces)
      objWMIReg.GetStringValue strHKLM, strPath   & arrInterfaces(intIdx), "NameServer", strNameServer
      objWMIReg.GetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "Domain",     strDomain
      If (strNameServer > "") And (Not (strDomain > "")) Then
        Call DebugLog("Processing Adapter " & arrInterfaces(intIdx))
        objWMIReg.SetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "Domain", strUserDNSDomain
      End If
    Next
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNoTCPNetBios()
  Call SetProcessId("1DE", "Setup No TCP NetBios access")
' Based on code published by Mark Harris http://lifeofageekadmin.com/disable-netbios-over-tcpip-with-vbscript
  Dim arrInterfaces
  Dim intIdx

  strPath           =  "SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces\"
  objWMIReg.EnumKey strHKLM, strPath, arrInterfaces

  For intIdx = 0 To Ubound(arrInterfaces)
    Call DebugLog("Processing Adapter " & arrInterfaces(intIdx))
    objWMIReg.SetDWORDValue  strHKLM, strPath & arrInterfaces(intIdx), "NetBIOSOptions", Hex(2)
  Next

  Call SetBuildfileValue("SetupNoTCPNetBiosStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNoTCPOffload()
  Call SetProcessId("1DF", "Setup No TCP Offload")
' Process described in KB976640
  Dim arrAdapters
  Dim colAdapter
  Dim intFound
  Dim strOffload, strPathAdapters

  intFound          = 0
  Call DebugLog("Turn of TCP Offload in Network Adapters")
  strPathAdapters   = "System\CurrentControlSet\Control\Class\{" & strNetworkGUID & "}\"
  objWMIReg.EnumKey strHKLM, strPathAdapters, arrAdapters
  For Each colAdapter In arrAdapters
    Select Case True
      Case colAdapter = "Properties"
        ' Nothing
      Case Else
        strPath     = strPathAdapters & colAdapter
        intFound    = intFound + AdapterOffloadDisable(strPath)
    End Select
  Next

  Call DebugLog("Turn of TCP Offload in Windows")
  strPath           = "System\CurrentControlSet\Services\TCPIP\Parameters"
  objWMIReg.GetDWordValue strHKLM,strPath,"DisableTaskOffload",strOffload
  Select Case True
    Case strOffload = 1
      ' Nothing
    Case Else
      intFound       = 1
      strPath        = "HKLM\" & strPath & "\DisableTaskOffload"
      Call Util_RegWrite(strPath, "1", "REG_DWORD") 
  End Select

  If intFound > 0 Then
    Call DebugLog(" TCP Offload Disabled")  
  End If

  Call SetBuildfileValue("SetupNoTCPOffloadStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Function AdapterOffloadDisable(strPathAdapter)
  Call DebugLog("AdapterOffloadDisable: " & strPathAdapter)
  Dim arrValueNames, arrValueTypes
  Dim intFound, intIdx
  Dim strValueName, strValueType

  intFound          = 0
  objWMIReg.EnumValues strHKLM, strPathAdapter, arrValueNames, arrValueTypes
  Select Case True
    Case Not IsArray(arrValueNames)
      ' Nothing
    Case Else
      For intIdx = 0 To UBound(arrValueNames)
        strValueName = arrValueNames(intIdx)
        strValueType = arrValueTypes(intIdx)
        Select Case True
          Case Left(strValueName, 1) <> "*"
            ' Nothing
          Case Instr(strValueName, "Offload") > 0
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*FlowControl"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*LsoV1IPv4"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*LsoV2IPv4"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*LsoV2IPv6"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*RSS"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
        End Select
      Next
  End Select

  AdapterOffloadDisable = intFound

End Function


Function OptionOffloadDisable(strPathAdapter, strOption, strType)
  Call DebugLog("OptionOffloadDisable: " & strOption)
  Dim intFound
  Dim strOffload, strRegType, strRegValue

  intFound          = 0
  Select Case True
    Case strType = 4
      strRegType    = "REG_DWORD"
      strRegValue   = 0
      objWMIReg.GetDWordValue strHKLM,strPathAdapter,strOption,strOffload
    Case Else
      strRegType    = "REG_SZ"
      strRegValue   = "0"
      objWMIReg.GetStringValue strHKLM,strPathAdapter,strOption,strOffload
  End Select

  Select Case True
    Case IsNull(strOffload)
      ' Nothing
    Case strOffload = strRegValue
      ' Nothing
    Case Else
      intFound      = 1
      strPath       = "HKLM\" & strPathAdapter & "\" & strOption
      Call Util_RegWrite(strPath, strRegValue, strRegType) 
  End Select

  OptionOffloadDisable = intFound

End Function


Sub SetupTLS12()
  Call SetProcessId("1DG", "Setup TLS 1.2 Support")
' More information given in KB3135244
  Dim intProcess, intRegValue

  intProcess        = 0
  strPath           = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"
  objWMIReg.GetDWordValue strHKLM,strPath,"DisabledByDefault",intRegValue
  Select Case True
    Case intRegValue = 0
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\DisabledByDefault", "0", "REG_DWORD")
      intProcess    = 1
  End Select

  objWMIReg.GetDWordValue strHKLM,strPath,"Enabled",intRegValue
  Select Case True
    Case intRegValue = 1
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\Enabled",           "1", "REG_DWORD")
      intProcess    = 1
  End Select

  strPath           = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"
  objWMIReg.GetDWordValue strHKLM,strPath,"DisabledByDefault",intRegValue
  Select Case True
    Case intRegValue = 0
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\DisabledByDefault", "0", "REG_DWORD")
      intProcess    = 1
  End Select

  objWMIReg.GetDWordValue strHKLM,strPath,"Enabled",intRegValue
  Select Case True
    Case intRegValue = 1
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\Enabled",           "1", "REG_DWORD")
      intProcess    = 1
  End Select
  
  Select Case True
    Case intProcess = 0
      Call SetBuildfileValue("SetupTLS12Status", strStatusPreConfig)
    Case Else
      Call SetBuildfileValue("RebootStatus", "Pending")
      Call SetBuildfileValue("SetupTLS12Status", strStatusComplete)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNoSSL3()
  Call SetProcessId("1DH", "Setup Disable SSL3")
' More information given in KB3009008
  Dim intProcess, intRegValue

  intProcess        = 0
  strPath           = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Client"
  objWMIReg.GetDWordValue strHKLM,strPath,"Enabled",intRegValue
  Select Case True
    Case intRegValue = 1
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\Enabled",           "0", "REG_DWORD")
      intProcess    = 1
  End Select

  strPath           = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server"
  objWMIReg.GetDWordValue strHKLM,strPath,"Enabled",intRegValue
  Select Case True
    Case intRegValue = 1
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\Enabled",           "0", "REG_DWORD")
      intProcess    = 1
  End Select

  Select Case True
    Case intProcess = 0
      Call SetBuildfileValue("SetupNoSSL3Status", strStatusPreConfig)
    Case Else
      Call SetBuildfileValue("RebootStatus", "Pending")
      Call SetBuildfileValue("SetupNoSSL3Status", strStatusComplete)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupAccounts()
  Call SetProcessId("1E", "Setup Accounts")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EA"
      ' Nothing
    Case Else
      Call SetupLocalGroups()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EB"
      ' Nothing
    Case Else
      Call SetupGroupRights()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EC"
      ' Nothing
    Case Else
      Call SetupAccountRights()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1ED"
      ' Nothing
    Case strSetupSPN <> "YES"
      ' Nothing
    Case Else
      Call SetupSPN()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EE"
      ' Nothing
    Case strSetupNoWinGlobal <> "YES"
      ' Nothing
    Case Else
      Call SetupNoWinGlobal()
  End Select

  Call SetProcessId("1EZ", " Setup Accounts" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupLocalGroups()
  Call SetProcessId("1EA", "Setup Local Groups")

  Call ProcessAccounts("AssignUserGroups", "")

  Call DebugLog("Process Computer accounts")
  If strClusterHost = "YES" Then
    strCmd          = "NET LOCALGROUP """ & strGroupUsers & """ """ & strDomain & "\" & strClusterName & "$" & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End If

  If strGroupMSA <> "" Then
    strCmd          = "NET LOCALGROUP """ & strGroupUsers & """ """ & strDomain & "\" & strGroupMSA & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupGroupRights()
  Call SetProcessId("1EB", "Setup Group Rights")

  Call RunNTRights("""" & strGroupUsers & """ +r SeNetworkLogonRight")
  Call RunNTRights("""" & strGroupUsers & """ +r SeInteractiveLogonRight")
  Call RunNTRights("""" & strGroupUsers & """ +r SeChangeNotifyPrivilege")

  Call RunNTRights("""" & strGroupAdmin & """ +r SeInteractiveLogonRight")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeRemoteInteractiveLogonRight")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeRemoteShutdownPrivilege")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeManageVolumePrivilege")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeProfileSingleProcessPrivilege")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeSystemProfilePrivilege")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeShutdownPrivilege")

  If strGroupRDUsers <> "" Then
    Call RunNTRights("""" & strGroupRDUsers & """ +r SeInteractiveLogonRight")
    Call RunNTRights("""" & strGroupRDUsers & """ +r SeRemoteInteractiveLogonRight")
  End If

  If (strSetupCmdshell = "YES") And (strCmdshellAccount <> "") Then
    Call RunNTRights("""" & strCmdshellAccount & """ +r SeBatchLogonRight")
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupAccountRights()
  Call SetProcessId("1EC", "Setup Account Rights")
  Dim arrShareList
  Dim strShareList
  Dim intIdx

  Call ProcessAccounts("AssignAccountRights", "")

  strShareList      = GetBuildfileValue("ShareList")
  If strShareList <> "" Then
    strShareList      = Left(strShareList, Len(strShareList) - 1)
    arrShareList      = Split(LTrim(RTrim(Replace(strShareList, ",", " "))))
    For intIdx = 0 To Ubound(arrShareList)
      If strSetupSQLDB = "YES" Then
        Call SetupRemoteShareRights(arrShareList(intIdx), strSQLAccount, "SqlAccount")
      End If
      If strSetupSQLDBAG = "YES" Then
        Call SetupRemoteShareRights(arrShareList(intIdx), strAgtAccount, "AgtAccount")
      End If
    Next
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRemoteShareRights(strShareName, strAccount, strAccountParm)
  Call DebugLog("SetupRemoteShareRights: " & strShareName & " for " & strSQLAccount)
  Dim arrACEs
  Dim objACE, objACEAccount, objSecDesc, objShareSec, objWMIRemote
  Dim strRemoteServer, strShare
  Dim intIdx, intRC

  intIdx            = Instr(strShareName, "\")
  strRemoteServer   = Left(strShareName, intIdx - 1)
  strShare          = Mid(strShareName, intIdx + 1)

  On Error Resume Next
  Set objWMIRemote  = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strRemoteServer & "\root\cimv2")
  Wscript.Sleep strWaitShort
  Set objShareSec   = objWMIRemote.Get("Win32_LogicalShareSecuritySetting.Name=""" & strShare & """")
  If IsEmpty(objShareSec) Then
    Err.Clear
    Exit Sub
  End If

  On Error GoTo 0
  intRC             = objShareSec.GetSecurityDescriptor(objSecDesc)
  arrACEs           = objSecDesc.DACL
  Set objACEAccount = GetShareDACL(strSQLAccount, "Full", strAccountParm)
  For Each objACE In arrACEs
    If objACEAccount.Trustee.Name = objACE.Trustee.Name Then
      objACEAccount.Trustee.Name = ""
    End If
  Next

  Select Case True
    Case objACEAccount.Trustee.Name = ""
      ' Nothing
    Case Else
      intIdx        = UBound(arrACEs) + 1
      ReDim Preserve arrACEs(intIdx)
      Set arrACEs(intIdx) = objACEAccount
  End Select

  objSecDesc.DACL   = arrACEs
  intRC             = objShareSec.SetSecurityDescriptor(objSecDesc)

  Set objWMIRemote  = Nothing
  Call SetBuildfileValue("SetupSharesStatus", strStatusProgress) 

End Sub


Sub SetupSPN()
  Call SetProcessId("1ED", "Setup Service Principal Names")
  Dim objInstParm, objSPNFile
  Dim strSPNFile, strSPNPath

  strSPNFile        = strInstance
  If strSetupSQLAS = "YES" Then
    strSPNFile      = strSPNFile & "AS"
  End If
  If strSetupSQLDB = "YES" Then
    strSPNFile      = strSPNFile & "DB"
  End If
  If strSetupSQLRS = "YES" Then
    strSPNFile      = strSPNFile & "RS"
  End If
  strSPNPath        = objShell.ExpandEnvironmentStrings("%Temp%")
  strSPNPath        = objFSO.GetAbsolutePathName(strSPNPath)
  strSPNFile        = "SetupSPNCmd" & strSPNFile & ".bat"
  strDebugMsg1      = "SPN commands: " & strSPNPath & "\" & strSPNFile
  Set objSPNFile    = objFSO.OpenTextFile(strSPNPath & "\" & strSPNFile, 2, True)
  objSPNFile.WriteLine "@ECHO OFF"
  objSPNFile.WriteLine "ECHO SPN Command File: " & strSPNPath & "\" & strSPNFile
  objSPNFile.WriteLine "SET CMDRC=0"
  objSPNFile.WriteLine "SET MAXRC=0"

  Select Case True
    Case strOSVersion >= "6.0"
      Call WriteSPNFile(objSPNFile, "SETSPN -S ")
    Case Else
      Call WriteSPNFile(objSPNFile, "SETSPN -D ")
      Call WriteSPNFile(objSPNFile, "SETSPN -A ")
  End Select
  
  objSPNFile.WriteLine "EXIT /B %MAXRC%"
  objSPNFile.Close

  Call SetXMLParm(objInstParm, "PathMain",     strSPNPath)
  Call RunInstall("SPN", strSPNFile, objInstParm)

  If GetBuildfileValue("SetupSPNStatus") <> strStatusComplete Then
    Call SetBuildMessage(strMsgWarning, "Unable to create SPNs. " &strSPNPath & "\" &  strSPNFile & " must be run by a Domain Administrator")
  End If

  Call ProcessEnd("")

End Sub


Sub WriteSPNFile(objSPNFile, strSPNCmd)
  Call DebugLog("WriteSPNFile: " & strSPNCMD)
  Dim strUserDomain

  strUserDomain     = ""
  If strUserDNSDomain <> "" Then
    strUserDomain   = "." & strUserDNSDomain
  End If

  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case strSetupSQLASCluster = "YES"
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPSvc.3/" & strClusterNameAS & " " & GetSPNAccount(strASAccount, strClusterNameAS))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPSvc.3/" & strClusterNameAS & strUserDomain & " " & GetSPNAccount(strASAccount, strClusterNameAS))
    Case strInstance <> "MSSQLSERVER"
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPSvc.3/" & strServer & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPSvc.3/" & strServer & strUserDomain & " " & GetSPNAccount(strASAccount, strServer))
    Case Else
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPSvc.3/" & strServer & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPSvc.3/" & strServer & ":" & strInstance & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPSvc.3/" & strServer & strUserDomain & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPSvc.3/" & strServer & strUserDomain & ":" & strInstance & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPDisco.3/" & strServer & " " & GetSPNAccount(strSqlBrowserAccount, strServer))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSOLAPDisco.3/" & strServer & strUserDomain & " " & GetSPNAccount(strSqlBrowserAccount, strServer))
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster = "YES"
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strClusterNameSQL & GetSPNInstance(strInstance) & " " & GetSPNAccount(strSQLAccount, strClusterNameSQL))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strClusterNameSQL & strUserDomain & " " & GetSPNAccount(strSQLAccount, strClusterNameSQL))
    Case Else
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strServer & GetSPNInstance(strInstance) & " " & GetSPNAccount(strSQLAccount, strServer))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strServer & ":" & strTCPPort & " " & GetSPNAccount(strSQLAccount, strServer))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strServer & strUserDomain & GetSPNInstance(strInstance) & " " & GetSPNAccount(strSQLAccount, strServer))
      Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strServer & strUserDomain & ":" & strTCPPort & " " & GetSPNAccount(strSQLAccount, strServer))
  End Select

  If strSetupAlwaysOn = "YES" Then
    Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strClusterGroupAO & " " & GetSPNAccount(strSQLAccount, strServer))
    Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strClusterGroupAO & ":" & strTCPPort & " " & GetSPNAccount(strSQLAccount, strServer))
    Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strClusterGroupAO & strUserDomain & " " & GetSPNAccount(strSQLAccount, strServer))
    Call WriteSPNCmd(objSPNFile, strSPNCMD & "MSSQLSvc/" & strClusterGroupAO & strUserDomain & ":" & strTCPPort & " " & GetSPNAccount(strSQLAccount, strServer))
  End If

  If (strSetupSQLRS = "YES") Or (strSetupISMaster = "YES") Then
    Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strServer & " " & strServer)
    Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strServer & strUserDomain & " " & strServer)
  End If

  intIdx            = Instr(strRSAccount, "\")
  Select Case True
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case Left(strRSAccount, intIdx - 1) <> strUserDNSDomain
      ' Nothing
    Case Else
      If strSetupSQLRSCluster = "YES" Then
        Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strClusterNameRS & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strClusterNameRS & ":80" & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strClusterNameRS & strUserDomain & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strClusterNameRS & strUserDomain & ":80" & " " & GetSPNAccount(strRSAccount, strServer))
      End If
      If strRSAlias <> "" Then
        Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strRSAlias & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strRSAlias & ":80" & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strRSAlias & strUserDomain & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteSPNCmd(objSPNFile, strSPNCMD & "HTTP/" & strRSAlias & strUserDomain & ":80" & " " & GetSPNAccount(strRSAccount, strServer))
      End If
  End Select

End Sub


Sub WriteSPNCmd(objSPNFile, strSPNCmd)
  Call DebugLog("WriteSPNCmd:")

  objSPNFile.WriteLine "ECHO " & strSPNCmd
  objSPNFile.WriteLine strSPNCmd
  objSPNFile.WriteLine "SET CMDRC=%ERRORLEVEL%"
  objSPNFile.WriteLine "IF %CMDRC% == 1 SET CMDRC=0"
  objSPNFile.WriteLine "IF %CMDRC% LSS 0 SET /A CMDRC=0 - %CMDRC%"
  objSPNFile.WriteLine "IF %CMDRC% GTR %MAXRC% SET MAXRC=%CMDRC%"

End Sub


Function GetSPNAccount(strAccount, strHost)
  Call DebugLog("GetSPNAccount: " & strAccount)
  Dim intIdx
  Dim strSPNAccount

  intIdx            = Instr(strAccount, "\")
  Select Case True
    Case intIdx = 0
      strSPNAccount = strHost
    Case Left(strAccount, intIdx - 1) <> strDomain
      strSPNAccount = strHost
    Case Else
      strSPNAccount = strAccount
  End Select

  GetSPNAccount     = UCase(strSPNAccount)

End Function


Function GetSPNInstance(strInstance)
  Call DebugLog("GetSPNInstance: " & strInstance)
  Dim strSPNInstance

  strSPNInstance    = ""
  If strInstance <> "MSSQLSERVER" Then
    strSPNInstance  = ":" & strInstance
  End If

  GetSPNInstance    = strSPNInstance

End Function


Sub SetupNoWinGlobal()
  Call SetProcessId("1EE", "Disble Windows Guest Access")
  Dim objAccount, objUser
  Dim strAccount, strAccountSID
' Do not remove 'Authenticated Users', it is needed for Kerberos

  If strType <> "WORKSTATION" Then
    Call DebugLog("Disable Domain Non-Specific Access")
    Call RemoveUser(strGroupUsers, "S-1-1-0",  "L") ' Everyone
    Call RemoveUser(strGroupUsers, "S-1-5-4",  "L") ' NT AUTHORITY\INTERACTIVE
    Call RemoveUser(strGroupUsers, "S-1-5-7",  "L") ' NT AUTHORITY\Anonymous
    Call RemoveUser(strGroupUsers, "S-1-5-13", "L") ' NT AUTHORITY\Terminal Service Users
    If strDomainSID <> "" Then
      Call RemoveUser(strGroupUsers, "S-1-5-21-" & strDomainSID & "-501", "D") ' domain\Guest
      Call RemoveUser(strGroupUsers, "S-1-5-21-" & strDomainSID & "-513", "D") ' domain\Domain Users
      Call RemoveUser(strGroupUsers, "S-1-5-21-" & strDomainSID & "-514", "D") ' domain\Domain Guests
    End If
  End If

  Call DebugLog("Disable Local Guest Account")
  Call GetGuestAccount(strAccount, strAccountSID)
  If strAccount <> "" Then
    Call RemoveUser(strGroupUsers, strAccountSID, "L") '  Local Guest
    strDebugMsg1    = "Disabling local Guest account: " & strAccount
    Set objUser     = GetObject("WinNT://./" & strAccount)
    objUser.AccountDisabled = True
    objUser.SetInfo
  End If

  Call SetBuildfileValue("SetupNoWinGlobalStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub RemoveUser(strGroup, strSID, strType)
  Call DebugLog("RemoveUser: " & strGroup & " for " & strSID)
  Dim objAccount
  Dim strAccount

  Set objAccount    = objWMI.Get("Win32_SID.SID='" & strSid & "'") 
  strAccount        = objAccount.AccountName
  Select Case True
    Case strAccount = ""
      ' Nothing
    Case strType = "L"
      strCmd        = "NET LOCALGROUP """ & strGroup & """ """ & strAccount & """ /DELETE"
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
    Case Else
      strCmd        = "NET LOCALGROUP """ & strGroup & """ """ & objAccount.ReferencedDomainName & "\" & strAccount & """ /DELETE"
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End Select

End Sub


Sub GetGuestAccount(strAccount, strAccountSID)
  Call DebugLog("GetGuestAccount:")
  Dim colUsers
  Dim objUser

  strAccount        = ""
  Set colUsers      = objWMI.ExecQuery("SELECT * FROM Win32_UserAccount WHERE LocalAccount=True") 
  For Each objUser In colUsers
    If Mid(objUser.SID, InstrRev(objUser.SID, "-") + 1) = "501" Then
      strAccount    = objUser.Name
      strAccountSID = objUser.SID
    End If
  Next

End Sub


Sub SetupDrives()
  Call SetProcessId("1F", "Setup Drives")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1FA"
      ' Nothing
    Case Else
      Call SetupDriveLabels()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1FB"
      ' Nothing
    Case Else
      Call SetupDriveShares()
  End Select

  Call SetProcessId("1FZ", " Setup Drives" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupDriveLabels()
  Call SetProcessId("1FA", "Setup Drive Labels")

  Call DebugLog("Setup System drives")
  Call SetupVolumes("VolSys",           strVolSys,      strLabSystem)
  Call SetupVolumes("VolProg",          strVolProg,     strLabProg)
  Call SetupVolumes("VolDBA",           strVolDBA,      strLabDBA)
  If strSetupTempWin = "YES" Then
    Call SetupVolumes("VolTempWin",     strVolTempWin,  strLabTempWin)
  End If

  Select Case True
    Case strSetupDTCCluster <> "YES"
      ' Nothing
    Case strActionDTC = "ADDNODE"
      ' Nothing
    Case (strOSVersion < "6.0") And (strDTCClusterRes > "")
      ' Nothing
    Case (strOSVersion >= "6.0") And (strDTCClusterRes > "") And (strDTCMultiInstance <> "YES")
      ' Nothing
    Case Else  
      Call DebugLog("Setup MSDTC drive")
      Call SetupVolumes("VolDTC",       strVolDTC,      strLabDTC)
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call DebugLog("Setup SQL drives")
      Call SetupVolumes("VolData",        strVolData,     strLabData)
      Call SetupVolumes("VolLog",         strVolLog,      strLabLog)
      Call SetupVolumes("VolSysDB",       strVolSysDB,    strLabSysDB)
      Call SetupVolumes("VolBackup",      strVolBackup,   strLabBackup)
      Select Case True
        Case strSQLVersion >= "SQL2012"
          Call SetupVolumes("VolTemp",    strVolTemp,     strLabTemp)
          Call SetupVolumes("VolLogTemp", strVolLogTemp,  strLabLog)
        Case Else
          Call SetupVolumes("VolTemp",    strVolTemp,     strLabTemp)
          Call SetupVolumes("VolLogTemp", strVolLogTemp,  strLabLog)
        End Select
      If strSetupBPE = "YES" Then
        Call SetupVolumes("VolBPE",       strVolBPE,      strLabBPE)
      End If
      If strSetupSQLDBFS = "YES" Then
        Call SetupVolumes("VolDataFS",    strVolDataFS,   strLabDataFS)
      End If
      If strSetupSQLDBFT = "YES" Then
        Call SetupVolumes("VolDataFT",    strVolDataFT,   strLabDataFT)
      End If
  End Select

  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case strActionSQLAS = "ADDNODE"
      ' Nothing
    Case Else
      Call DebugLog("Setup SQL AS drives")
      Call SetupVolumes("VolDataAS",      strVolDataAS,   strLabDataAS)
      Call SetupVolumes("VolLogAS",       strVolLogAS,    strLabLogAS)
      Call SetupVolumes("VolBackupAS",    strVolBackupAS, strLabBackupAS)
      Call SetupVolumes("VolTempAS",      strVolTempAS,   strLabTempAS)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupVolumes(strVolParam, strVolList, strVolLabel)
  Call DebugLog("SetupVolumes: " & strVolParam)
  Dim arrItems
  Dim strVol, strVolReq, strVolSource

  strVolReq         = GetBuildfileValue(strVolParam & "Req")
  strVolSource      = GetBuildfileValue(strVolParam & "Source")
  Select Case True
    Case strVolSource = "C"
      arrItems      = Split(Replace(strVolList, ",", " "))
      For intIdx = 0 To UBound(arrItems)
        strVol      = arrItems(intIdx)
        strVol      = UCase(Mid(strVol, Len(strCSVRoot) + 1))
        Call SetupThisCSV(strVolParam, strVol)
      Next
    Case strVolSource = "D"
      For intIdx = 1 To Len(strVolList)
        strVol      = Mid(strVolList, intIdx, 1)
        Call SetupThisDrive(strVolParam, strVol, strVolLabel, strVolList, strVolReq)
      Next
  End Select

End Sub


Sub SetupThisCSV(strVolParam, strVol)
  Call DebugLog("SetupThisCSV: " & strVol)
  Dim strName

  strName           = GetBuildfileValue("Vol_" & strVol & "Name")
  Select Case True
    Case UCase(strName) = UCase(strVol)
      ' Nothing
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE """ & strName & """ /RENAME:""" & strVol & """"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      Call SetBuildfileValue("Vol_" & strVol & "Name", strVol)
  End Select

End Sub


Sub SetupThisDrive(strVolParam, strVol, strVolLabel, strVolList, strVolReq)
  Call DebugLog("SetupThisDrive: " & strVol)
  Dim strNewLabel

  strNewLabel       = GetDriveLabel(strVolParam, strVol, strVolList, strVolLabel)

  Select Case True
    Case (Instr(strDriveList, strVol) = 0) And (strVol <> Left(strDriveList, Len(strVol)))
      ' No action, not a valid drive
    Case (Not objFSO.FolderExists(strVol & ":\")) And (strVol <> Left(strDriveList, Len(strVol)))
      Call FBLog(" Setup " & strVol & ": for " & strNewLabel & strStatusBypassed)
    Case Instr(strVolUsed, strVol) > 0
      ' Nothing
    Case Else
      Call FBLog(" Setup " & strVol & ": drive for " & strNewLabel)
      strVolUsed    = strVolUsed & " " & strVol
      If strClusterAction <> "" Then
        Call LabelThisClusterDrive(strVol, strNewLabel, strVolReq)
      End If
      Call LabelThisDrive(strVolParam,strVol, strNewLabel, strVolReq)
      Call CreateThisShare(strVol, strNewLabel)
  End Select

End Sub


Function GetDriveLabel(strVolParam, strVol, strVolList, strVolLabel)
  Call DebugLog("GetDriveLabel: " & strVolParam)
  Dim strVolNewLabel

  Select Case True
    Case Len(strVolList) = 1 
      strVolNewLabel = strVolLabel
    Case Else
      strVolNewLabel = strVolLabel & strVol
  End Select

  Select Case True
    Case strVol = strVolSys
      ' Nothing
    Case strVol = strVolProg
      ' Nothing
    Case strVol = strVolDTC
      ' Nothing
    Case strInstance = "MSSQLSERVER"
      ' Nothing
    Case strInstance = "SQLEXPRESS"
      ' Nothing
    Case Else
      strVolNewLabel = strVolNewLabel & "-" & strInstance
  End Select 

  If strLabPrefix <> "" Then
    strVolNewLabel  = Left(strLabPrefix & "-" & Replace(strVolNewLabel, " ", ""), 32)
    Call SetBuildfileValue("Lab" & Mid(strVolParam, 4), strVolNewLabel)
  End If

  strVolNewLabel    = Left(strVolNewLabel, 32)
  GetDriveLabel     = strVolNewLabel

End Function


Sub LabelThisClusterDrive(strVol, strVolLabel, strVolReq)
  Call DebugLog("LabelThisClusterDrive: " & strVol)
  Dim colClusGroups, colClusPartitions, colClusResources
  Dim objClusDisk, objClusGroup, objClusPartition, objClusResource
  Dim intCount, intFound

  strCmd            = "CLUSTER """ & strClusterName & """ GROUP """ & strClusStorage & """ /MOVETO:""" & strServer & """" 
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

  intFound          = 0
  Set colClusGroups = objCluster.ResourceGroups
  For Each objClusGroup In colClusGroups                   
    Set colClusResources = objClusGroup.Resources
    intCount        = 99
    For Each objClusResource In colClusResources
      If objClusResource.TypeName = "Physical Disk" Then
        Set objClusDisk           = objClusResource.Disk
        Set colClusPartitions     = objClusDisk.Partitions
        For Each objClusPartition In colClusPartitions
          Select Case True
            Case intFound <> 0
              ' Nothing
            Case strVolReq = "L" And Left(objClusPartition.DeviceName, 1) = strVol 
              ' Nothing 
            Case Left(objClusPartition.DeviceName, 1) = strVol 
              intCount = colClusResources.Count
              intFound = 1
              strCmd   = "CLUSTER """ & strClusterName & """ GROUP """ & objClusGroup.Name & """ /MOVETO:""" & strServer & """" 
              Call Util_RunExec(strCmd, "", strResponseYes, 0)
              strCmd   = "CLUSTER """ & strClusterName & """ RESOURCE """ & objClusResource.Name & """ /MOVE:""" & strClusStorage & """"
              Call Util_RunExec(strCmd, "", strResponseYes, 183)
              strCmd   = "CLUSTER """ & strClusterName & """ RESOURCE """ & objClusResource.Name & """ /RENAME:""" & strVolLabel & """"
              Call Util_RunExec(strCmd, "", strResponseYes, 0) 
          End Select
        Next
      End If
    Next

    Select Case True
      Case intCount > 1
        ' Nothing
      Case UCase(objClusGroup.Name) = UCase(strClusStorage)
        ' Nothing
      Case Else
        Call DebugLog("Delete empty group " & objClusGroup.Name)
        strCmd      = "CLUSTER """ & strClusterName & """ GROUP """ & objClusGroup.Name & """ /DELETE"
        Call Util_RunExec(strCmd, "", strResponseYes, 5010)
    End Select
  Next 

End Sub


Sub LabelThisDrive(strVolParam, strVol, strVolLabel, strVolReq)
  Call DebugLog("LabelThisDrive: " & strVol)
' Code to clear IndexingEnabled flag adapted from "Windows Server Cookbook" by Robbie Allen, ISBN 0-596-00633-0
  Dim strVolType

  If Not objFSO.FolderExists(strVol & ":\") Then
    Call SetBuildMessage(strMsgError, "Volume not found: " & strVol & ":\")
  End If

  Select Case True
    Case strOSVersion > "5.1"
      strCmd        = "SELECT * FROM Win32_Volume WHERE DriveLetter='" & strVol & ":'"
      Set colVol    = objWMI.ExecQuery(strCmd)
      For Each objVol In colVol
        If strSetupNoDriveIndex = "YES" Then
          objVol.IndexingEnabled = 0
        End If
        objVol.Label = strVolLabel
        objVol.Put_
      Next
    Case Else
      strCmd        = "SELECT * FROM Win32_LogicalDisk WHERE DeviceID='" & strVol & ":'"
      Set colVol    = objWMI.ExecQuery(strCmd)
      For Each objVol In colVol
        objVol.VolumeName = strVolLabel
        objVol.Put_
      Next
  End Select

  If strSetupNoDriveIndex = "YES" Then
    Call DebugLog("Clearing index attribute from drive " & strVol)
    strCmd            = "ATTRIB +I " & strVol & ":\*.* /D /S"
    Call Util_RunCmdAsync(strCmd, 0)
    Call SetBuildfileValue("SetupNoDriveIndexStatus", strStatusComplete)
  End If

  Call SetBuildfileValue("Vol" & strVol & "Label", strVolLabel)

  strVolType        = GetBuildfileValue("Vol" & strVol & "Type")
  If strVolType = "" Then
    Call SetBuildfileValue("Vol" & strVol & "Type", "L")
  End If
  Select Case True
    Case strVolReq = "C" And Instr("CX", strVolType) = 0
      Call SetBuildMessage(strMsgError, strVolParam & ": " & strVol & ": must be a Cluster Drive")
    Case strVolReq = "L" And strVolType = "C"
      Call SetBuildMessage(strMsgError, strVolParam & ": " & strVol & ": must NOT be a Cluster Drive")
  End Select

End Sub


Sub SetupDriveShares()
  Call SetProcessId("1FB", "Setup Drive Shares")

' KB245117 fix for share visibility
  Select Case True
    Case strSetupShares <> "YES"
      ' Nothing
    Case Left(strOSVersion, 1) >= "6"
      ' Nothing
    Case Else
      strPath       = "HKLM\System\CurrentControlSet\Services\LanmanServer\Parameters\AutoShareServer"
      Call Util_RegWrite(strPath, 1, "REG_DWORD")
      strPath       = "HKLM\System\CurrentControlSet\Services\LanmanServer\Parameters\AutoShareWks"
      Call Util_RegWrite(strPath, 1, "REG_DWORD")
      Call SetBuildfileValue("RebootStatus", "Pending")    
  End Select

  Call SetBuildfileValue("SetupSharesStatus", strStatusComplete) 
  Call ProcessEnd(strStatusComplete)

End Sub


Sub CreateThisShare(strVol, strVolLabel)
  Call DebugLog("CreateThisShare: " & strVol)
  Dim strShareName

  strVolType        = GetBuildfileValue("Vol" & strVol & "Type")
  strShareName      = "(" & strVol & ") " & strVolLabel
  Select Case True
    Case strSetupShares <> "YES" 
      ' Nothing
    Case strVolType = "L"
      Call SetupLocalShare(strVol & ":\", strShareName)
  End Select

  Call SetBuildfileValue("Vol" & strVol & "Share", strShareName)
  Call SetBuildfileValue("SetupSharesStatus", strStatusProgress)

End Sub


Sub SetupLocalShare(strVol, strShareName)
  Call DebugLog("SetupLocalShare: " & strVol)
  Dim objACEAdmin, objACEUser, objSecDesc, objShare, objShareParm

  Set objSecDesc    = objWMI.Get("Win32_SecurityDescriptor").SpawnInstance_
  Set objShare      = objWMI.Get("Win32_Share")
  Set objACEAdmin   = GetShareDACL(strGroupAdmin, "Full",   "")
  Set objACEUser    = GetShareDACL(strGroupUsers, "Change", "")
  Set objShareParm  = objShare.Methods_("Create").InParameters.SpawnInstance_ 

  objSecDesc.DACL          = Array(objACEAdmin, objACEUser)
  objShareParm.Access      = objSecDesc
  objShareParm.Description = strShareName & " Share"
  objShareParm.Name = strShareName
  objShareParm.Path = strVol
  objShareParm.Type = 0
  objShare.ExecMethod_ "Create",  objShareParm

End Sub


Function GetShareDACL(strAccount, strAccess, strAccountParm)
  Call DebugLog("GetShareDACL: " & strAccount)
  Dim objACE, objTrustee

  Set objTrustee    = SetTrustee(strAccount, strAccountParm)
  Set objACE        = objWMI.Get("Win32_Ace").SpawnInstance_

  objACE.AceFlags   = 3
  objACE.AceType    = 0
  objACE.Trustee    = objTrustee
  Select Case True
    Case strAccess = "Full"
      objACE.AccessMask = 2032127
    Case Else
      objACE.AccessMask = 1245631 ' Change
  End Select

  Set GetShareDACL  = objAce

End Function


Function SetTrustee(strAccount, strAccountParm) 
  Call DebugLog("SetTrustee: " & strAccount & " for " & strAccountParm)
  Dim objRecordSet, objTrustee
  Dim strAttrObject, strDNSDomain, strLocal, strSID, strSIDBinary, strQueryDomain, strQueryAccount
  Dim intIdx

  strLocal          = ""
  intIdx            = InStr(strAccount, "\")
  Select Case True
    Case intIdx = 0
      strDNSDomain     = strServer
      strQueryDomain   = strServer
      strQueryAccount  = strAccount
    Case Left(strAccount, intIdx - 1) = strDomain
      strDNSDomain     = strUserDNSDomain
      strQueryDomain   = strDomain
      strQueryAccount  = Mid(strAccount, intIdx + 1)
    Case Else
      strDNSDomain     = strServer
      strQueryDomain   = strServer
      strQueryAccount  = Mid(strAccount, intIdx + 1)
      strLocal         = ",LocalAccount=True"
  End Select
  strDebugMsg1      = "QueryDomain=" & strQueryDomain & " QueryAccount=" & strQueryAccount

  Select Case True
    Case strAccountParm = ""
      strDebugMsg2  = "Group Account"
      strSID        = objWMI.Get("Win32_Group.Domain='" & strQueryDomain & "',Name='" & strQueryAccount & "'" & strLocal).SID
      strSIDBinary  = objWMI.Get("Win32_SID.SID='" & strSID &"'").BinaryRepresentation 
    Case GetBuildfileValue(strAccountParm & "Type") = "M"
      strDebugMsg2  = "MSA Account"
      strAttrObject = "objectClass=msDS-GroupManagedServiceAccount"
      objADOCmd.CommandText = "<LDAP://DC=" & Replace(strDNSDomain, ".", ",DC=") & ">;(&(" & strAttrObject & ")(CN=" & strQueryAccount & "));CN,objectSID"
      Set objRecordSet  = objADOCmd.Execute
      objRecordset.MoveFirst
      strSIDBinary = objRecordset.Fields(1).Value
    Case Else
      strDebugMsg2  = "User Account"
      strAttrObject = "objectClass=user"
      objADOCmd.CommandText = "<LDAP://DC=" & Replace(strDNSDomain, ".", ",DC=") & ">;(&(" & strAttrObject & ")(CN=" & strQueryAccount & "));CN,SID"
      Set objRecordSet  = objADOCmd.Execute
      objRecordset.MoveFirst
      strSIDBinary = objRecordset.Fields(1).Value
  End Select

  Set objTrustee    = objWMI.Get("Win32_Trustee").Spawninstance_ 
  objTrustee.Domain = strQueryDomain 
  objTrustee.Name   = strQueryAccount
  objTrustee.SID    = strSIDBinary
 
  Set SetTrustee    = objTrustee 

End Function


Sub SetupFolders()
  Call SetProcessId("1G", "Setup SQL Server folders")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GA"
      ' Nothing
    Case Else
      Call SetupSystemFolders()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GB"
      ' Nothing
    Case Else
      Call SetupStdDrives()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GC"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call SetupSQLServer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GD"
      ' Nothing
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case strActionSQLAS = "ADDNODE"
      ' Nothing
    Case Else
      Call SetupSQLASDrives()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GE"
      ' Nothing
    Case strSetupTempWin <> "YES"
      ' Nothing
    Case Else
      Call SetupTempDrive()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GF"
      ' Nothing
    Case strSetupTempWin <> "YES"
      ' Nothing
    Case Else
      Call SetupAllUsersTemp()
  End Select

  Call SetProcessId("1GZ", " Setup SQL Server Folders" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupSystemFolders()
  Call SetProcessId("1GA", "Setup System folders")
  Dim objFolderParm

  Call SetXMLParm(objFolderParm, "Account1",   strUserAccount)
  Call SetXMLParm(objFolderParm, "Access",     strSecMain)
  Call PrepareFolder("PROG",     objFolderParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupStdDrives()
  Call SetProcessId("1GB", "Setup Standard drives")
  Dim objFolderParm

  Call SetXMLParm(objFolderParm,   "Folder1",    "\Scripts")
  Call SetXMLParm(objFolderParm,   "Folder2",    "\Servers")
  If strSetupSQLTools = "YES" Then
    Call SetXMLParm(objFolderParm, "Folder3",  "\SQL Server Management Studio\Custom reports")
  End If
  Call PrepareFolder("DBA",        objFolderParm)

  If strSetupDRUClt = "YES" Then
    Call SetXMLParm(objFolderParm, "Account1", strSqlAccount)
    Call SetXMLParm(objFolderParm, "Account2", strAgtAccount)
    Call SetXMLParm(objFolderParm, "Folder1",  "\DRU.Work")
    Call SetXMLParm(objFolderParm, "Folder2",  "\DRU.Result")
    Call PrepareFolder("DRU",      objFolderParm)
  End If

  If GetBuildfileValue("SetupManagementDW") = "YES" Then
    Call SetXMLParm(objFolderParm, "Account1", GetBuildfileValue("MDWAccount"))
    Call SetXMLParm(objFolderParm, "Account2", strAgtAccount)
    Call PrepareFolder("MDW",      objFolderParm)
  End If

  Select Case True
    Case strSetupSQLIS <> "YES"
      ' Nothing
    Case strClusterAction = ""
      ' Nothing
    Case strSetupSSISCluster <> "YES"
      ' Nothing
    Case Else
      Call SetXMLParm(objFolderParm, "Account1",  strIsAccount)
      Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
      Call SetXMLParm(objFolderParm, "Folder1",   "\Packages")
      Call PrepareFolder("DataIS",   objFolderParm)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSQLServer()
  Call SetProcessId("1GC", "Setup SQL Server drives")
  Dim objFolderParm

  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("Data",     objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("Log",      objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("SysDB",    objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call SetXMLParm(objFolderParm, "Folder1",   "\AdHocBackup")
  Call SetXMLParm(objFolderParm, "Folder2",   "\Reports")
  Call PrepareFolder("Backup",   objFolderParm)
  Call PrepareFolderPath("Backup", strAction, strDirSystemDataBackup,      strSecNull, "", "")
  Call PrepareFolderPath("Backup", strAction, strDirSystemDataShared,      strSecNull, "", "")

  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("Temp",     objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("LogTemp",  objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call PrepareFolder("BPE",      objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call PrepareFolder("DataFS",   objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call PrepareFolder("DataFT",   objFolderParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSQLASDrives()
  Call SetProcessId("1GD", "Setup AS Service drives")
  Dim objFolderParm

  Call SetXMLParm(objFolderParm, "Account1",  strAsAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("DataAS",   objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strAsAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("LogAS",    objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strAsAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call SetXMLParm(objFolderParm, "Folder1",   "\Data")
  Call SetXMLParm(objFolderParm, "Folder2",   "\AdHocBackup")
  Call PrepareFolder("BackupAS", objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",  strAsAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("TempAS",   objFolderParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupTempDrive()
  Call SetProcessId("1GE", "Setup Temp folder drive")

  Call PrepareFolderPath("TempWin", strAction, strPathTemp, strSecTemp, "", "")

  colSysEnvVars("TEMP") = strPathTemp
  colSysEnvVars("TMP")  = strPathTemp

  colUsrEnvVars("TEMP") = strPathTemp
  colUsrEnvVars("TMP")  = strPathTemp

  Call SetBuildFileValue("SetupTempWinStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupAllUsersTemp()
  Call SetProcessId("1GF", "Temp Folder for All Users Account")
  Dim strUserSid

  For Each objFolder In arrProfFolders
    Select Case True
      Case Not objFSO.FileExists (objFolder.Path & "\NTUSER.DAT")
        ' Nothing
      Case Else
        Call DebugLog("Account path: " & objFolder.Path)
        strCmd        = strProgReg & " LOAD ""HKLM\FBTempProf"" """ & objFolder.Path & "\NTUSER.DAT"""
        Call Util_RunExec(strCmd, "", strResponseYes, -1)
        Select Case True
          Case intErrSave = 0
            Call SetTempLoc(strHKLM, "HKLM", "FBTempProf")
            strCmd      = strProgReg & " UNLOAD  ""HKLM\FBTempProf"""
            Call Util_RunExec(strCmd, "", strResponseYes, 1)
          Case intErrSave = 1
            ' Nothing
          Case intErrSave = 1332
            ' Nothing
          Case Else
            Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
        End Select
    End Select
  Next

  objWMIReg.EnumKey strHKU, "", arrUserSid
  For Each strUserSid In arrUserSid
    Select Case True
      Case Right(strUserSid, 8) = "_Classes"
        ' Nothing
      Case Else
        Call DebugLog("Account SID: " & strUserSid)
        Call SetTempLoc(strHKU, "HKEY_USERS", strUserSid)
    End Select
  Next

  Call SetBuildFileValue("SetupTempWinStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetTempLoc(strKeyValue, strKey, strSid)
  Call DebugLog("SetTempLoc:")
  Dim strTempVar

  strPath           = strSid & "\Environment"
  objWMIReg.GetStringValue strKeyValue,strPath,"TEMP",strTempVar

  If Not IsNull(strtempVar) Then
    strCmd          = strKey & "\" & strPath & "\TEMP" 
    Call Util_RegWrite(strCmd, strPathTempUser, "REG_SZ")
    strCmd          = strKey & "\" & strPath & "\TMP" 
    Call Util_RegWrite(strCmd, strPathTempUser, "REG_SZ")
  End If

End Sub


Sub PrepareFolder(strFType, objFolderParm)
  Call DebugLog("PrepareFolder: " & strFType)
  Dim arrVolumes
  Dim strDirBase, strDirName
  Dim strFAction, strFAccount1, strFAccount2, strFAccess, strFFolder1, strFFolder2, strFFolder3
  Dim strVolume, strVolList, strVolSource
  Dim intIdx, intVol
' VolSource: C=CSV, D=Disk, M=Mount Point, N=Mapped Network Drive, S=Share
' VolType:   C=Clustered, L=Local, X=Either

  strFAction        = GetXMLParm(objFolderParm, "Action",   "strAction")
  strFAccount1      = GetXMLParm(objFolderParm, "Account1", "")
  strFAccount2      = GetXMLParm(objFolderParm, "Account2", "")
  strFAccess        = GetXMLParm(objFolderParm, "Access",   strSecNull)
  strFFolder1       = GetXMLParm(objFolderParm, "Folder1",   "")
  strFFolder2       = GetXMLParm(objFolderParm, "Folder2",   "")
  strFFolder3       = GetXMLParm(objFolderParm, "Folder3",   "")

  strDirName        = GetBuildfileValue("Dir" & strFType)
  strDirBase        = GetBuildfileValue("Dir" & strFType & "Base")
  strVolList        = GetBuildfileValue("Vol" & strFType)
  strVolSource      = GetBuildfileValue("Vol" & strFType & "Source")

  arrVolumes        = Split(Replace(strVolList, ",", " "))
  For intVol = 0 To UBound(arrVolumes)
    strVolume       = Trim(arrVolumes(intVol))
    Select Case True
      Case strDirName = ""
        ' Nothing
      Case strVolSource <> "D"
        strPath     = strVolume & strDirBase
        Call PrepareFolderPath(strFType, strFAction, strPath, strSecDBA, strFAccount1, strFAccount2)
        Call SetupAVExclude(strFType, strPath)
      Case Else
        For intIdx = 1 To Len(strVolume)
          strVol    = Mid(strVolume, intIdx, 1)
          strPath   = strVol & strDirBase
          Call PrepareFolderPath(strFType, strFAction, strPath, strSecDBA, strFAccount1, strFAccount2)
          Call SetupAVExclude(strFType, strPath)
        Next
    End Select
  Next

  If strFFolder1 <> "" Then
    Call PrepareFolderPath(strFType, strFAction, strDirName & strFFolder1, strFAccess, "", "")
  End If

  If strFFolder2 <> "" Then
    Call PrepareFolderPath(strFType, strFAction, strDirName & strFFolder2, strFAccess, "", "")
  End If

  If strFFolder3 <> "" Then
    Call PrepareFolderPath(strFType, strFAction, strDirName & strFFolder3, strFAccess, "", "")
  End If

  objFolderParm     = ""

End Sub


Sub PrepareFolderPath(strType, strAction, strPath, strSec, strAccount1, strAccount2)
  Call DebugLog("PrepareFolderPath: " & strPath)
  Dim strPathFolder

  strPrepareFolderPath = strPath
  Select Case True
    Case strAction <> "ADDNODE"
      ' Nothing
    Case GetBuildfileValue("Vol" & strType & "Source") <> "D"
      ' Nothing
    Case GetBuildfileValue("Vol" & strType & "Type") <> "L"
      Exit Sub
  End Select

  strPathFolder     =  SetupFolder(strPath, strSec)
  If strAccount1 <> "" Then
    strCmd          = """" & strPathFolder & """ /T /C /E /G """ & FormatAccount(strAccount1) & """:F"
    Call RunCacls(strCmd)
  End If
  If strAccount2 <> "" Then
    strCmd          = """" & strPathFolder & """ /T /C /E /G """ & FormatAccount(strAccount2) & """:F"
    Call RunCacls(strCmd)
  End If

End Sub


Function SetupFolder(strPath, strSec)
  Call DebugLog("SetupFolder: " & strPath)
  Dim strNull, strPathAlt, strPathAltParent, strPathFolder, strPathParent, strPathRoot

  If Right(strPath, 1) = "\" Then
    strPath         = Left(strPath, Len(strPath) - 1)
  End If

  strPathAlt        = strPath
  Select Case True
    Case Left(strPath, 2) <> "\\"
      ' Nothing
    Case Instr(3, strPath, "\") = 0
      SetupFolder   = strPath
      Exit Function
    Case Else
      strPathRoot   = Left(strPathAlt, Instr(3, strPathAlt, "\") - 1)
      strPathAlt    = strPathRoot & Mid(strPathRoot, 2) & Mid(strPathAlt, Instr(3, strPathAlt, "\")) ' For SOFS
  End Select
  strPathParent     = Left(strPath, InstrRev(strPath, "\") - 1)
  strPathAltParent  = Left(strPathAlt, InstrRev(strPathAlt, "\") - 1)

  strDebugMsg1      = "PathParent: " & strPathParent
  strPathFolder     = ""
  Select Case True
    Case objFSO.FolderExists(strPath & "\")
      strPathFolder = strPath
    Case objFSO.FolderExists(strPathAlt & "\")
      strPathFolder = strPathAlt
    Case objFSO.FolderExists(strPathParent & "\")
      strPathFolder = strPath
      Call CreateThisFolder(strPathFolder, strSec)
    Case objFSO.FolderExists(strPathAltParent & "\")
      strPathFolder = strPathAlt
      Call CreateThisFolder(strPathFolder, strSec)
    Case Else
      strNull       = SetupFolder(strPathParent, strSec)
      Select Case True
        Case objFSO.FolderExists(strPathParent & "\")
          strPathFolder = strPath
          Call CreateThisFolder(strPathFolder, strSec)
        Case objFSO.FolderExists(strPathAltParent & "\")
          strPathFolder = strPathAlt
          Call CreateThisFolder(strPathFolder, strSec)
      End Select
  End Select

  SetupFolder       = strPathFolder

End Function


Sub CreateThisFolder(strFolder, strSec)
  Call DebugLog("CreateThisFolder: " & strFolder)
  Dim strCreate

  strCreate         = "N"
  Select Case True
    Case objFSO.FolderExists(strFolder)
      ' Nothing
    Case Else
      objFSO.CreateFolder(strFolder)
      Wscript.Sleep strWaitShort
      strCreate     = "Y"
  End Select

  Select Case True
    Case strSec = ""
      ' Nothing
    Case (strSec = strSecDBA) And (strCreate = "N")
      ' Nothing
    Case strSec = strSecDBA
      strCmd        = """" & strFolder & """ /T /C /G " & strSec 
      Select Case True
        Case strGroupDBANonSA = ""
          ' Nothing
        Case strFolder = strDirDBA
          strCmd    = strCmd & """" & FormatAccount(strGroupDBANonSA) & """:F "
        Case strFolder = strPathTemp
          strCmd    = strCmd & """" & FormatAccount(strGroupDBANonSA) & """:F "
        Case Else
          strCmd    = strCmd & """" & FormatAccount(strGroupDBANonSA) & """:R "
      End Select
      Call RunCacls(strCmd)
    Case strSec = strSecTemp
      strCmd        = """" & strFolder & """ /T /C /G " & strSec 
      Call RunCacls(strCmd)     
    Case Else
      Call ProcessAccounts("AssignFolderRights", strFolder)
  End Select

  strPrepareFolderPath = ""

End Sub


Sub SetupAVExclude(strType, strPath)
  Call DebugLog("SetupAVExclude: " & strType & " for " & strPath)

  strCmd            = strAVCmd & """" & strPath & """"
  Call Util_RunExec(strCmd, "", "", -1)

End Sub


Sub ProcessAccounts(strProcess, strParameter)
  Call DebugLog("ProcessAccounts: " & strProcess)
  Dim intDomIdx, strParm

  intDomIdx         = InStr(strNTAuthAccount, "\")

  If strParameter <> "" Then
    strParm         = ",""" & strParameter & """"
  End If

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSQLAccount = ""
      ' Nothing
    Case Else
      strCmd        = strProcess & "(""" & strSQLAccount & """" & strParm & ")"
      Execute "Call " & strCmd
  End Select

  Call ProcessServiceAccount(strProcess, strSetupSQLDBAG,                     strAgtAccount,                        strParm)
  Call ProcessServiceAccount(strProcess, strSetupSQLAS,                       strASAccount,                         strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupDRUCtlr"),   strDRUCtlrAccount,                    strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupDRUClt"),    strDRUCltAccount,                     strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupSQLDB"),     strSQLBrowserAccount,                 strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupAnalytics"), GetBuildfileValue("ExtSvcAccount"),   strParm)
  Call ProcessServiceAccount(strProcess, strSetupSQLDBFT,                     strFTAccount,                         strParm)
  Call ProcessServiceAccount(strProcess, strSetupSQLIS,                       strIsAccount,                         strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupIsMaster"),  GetBuildfileValue("IsMasterAccount"), strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupISWorker"),  GetBuildfileValue("IsWorkerAccount"), strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupPolyBase"),  GetBuildfileValue("PBDMSSvcAccount"), strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupPolyBase"),  GetBuildfileValue("PBEngSvcAccount"), strParm)
  Call ProcessServiceAccount(strProcess, strSetupSQLRS,                       strRSAccount,                         strParm)

  Select Case True
    Case strGroupDBA = ""
      ' Nothing
    Case Else
      strCmd        = strProcess & "(""" & FormatAccount(strGroupDBA) & """" & strParm & ")"
      Execute "Call " & strCmd
  End Select

  Select Case True
    Case strGroupDBANonSA = ""
      ' Nothing
    Case Else
      strCmd        = strProcess & "(""" & FormatAccount(strGroupDBANonSA) & """" & strParm & ")"
      Execute "Call " & strCmd
  End Select

End Sub


Sub ProcessServiceAccount(strProcess, strSetup, strAccount, strParm)
  Call DebugLog("ProcessServiceAccount: " & strAccount)

  Select Case True
    Case strSetup <> "YES"
      ' Nothing
    Case strAccount = ""
      ' Nothing
    Case strAccount = strSqlAccount
      ' Nothing
    Case Else
      strCmd        = strProcess & "(""" & strAccount & """" & strParm & ")"
      Execute "Call " & strCmd
  End Select

End Sub


Sub AssignUserGroups(strAccount)
  Call DebugLog("AssignUserGroups: " & strAccount)
  Dim objAccount
  Dim intServerLen

  intServerLen      = Len(strServer) + 1
  If strGroupDistComUsers = "" Then
    strGroupDistComUsers = "Distributed COM Users"
    strCmd             = "NET LOCALGROUP """ & strGroupDistComUsers & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, 2)
    strCmd             = "Win32_Group.Domain='" & strLocalDomain & "',Name='" & strGroupDistComUsers & "'"
    Set objAccount     = objWMI.Get(strCmd) 
    strSIDDistComUsers = objAccount.SID
    Call SetBuildfileValue("SIDDistComUsers",    strSIDDistComUsers)
    Call SetBuildfileValue("GroupDistComUsers",  strGroupDistComUsers)
  End If

  Select Case True
    Case Left(strGroupDBA, intServerLen) = strServer & "\"
      ' Nothing
    Case strGroupDBA = strLocalAdmin
      ' Nothing
    Case Ucase(strAccount) = strGroupDBA
      strCmd        = "NET LOCALGROUP """ & strGroupRDUsers & """ """ & strAccount & """ /ADD"
      Call Util_RunExec(strCmd, "", strResponseYes, 2)
      Call AssignAccountGroups(strAccount)
  End Select

  Select Case True
    Case strGroupDBANonSA = ""
      ' Nothing
    Case Left(strGroupDBANonSA, intServerLen) = strServer & "\"
      ' Nothing
    Case Ucase(strAccount) = strGroupDBANonSA
      strCmd        = "NET LOCALGROUP """ & strGroupRDUsers & """ """ & strAccount & """ /ADD"
      Call Util_RunExec(strCmd, "", strResponseYes, 2)
      Call AssignAccountGroups(strAccount)
  End Select

  Select Case True
    Case Ucase(strAccount) = strGroupDBA
      ' Nothing
    Case Ucase(strAccount) = strGroupDBANonSA
      ' Nothing
    Case Left(strAccount, intServerLen) = strServer & "\"
      ' Nothing
    Case Ucase(strAccount) = strNTAuthAccount 
      ' Nothing
    Case Left(strAccount, Len(strNTService) + 1) = strNTService & "\"
      ' Nothing
    Case Else
      Call AssignAccountGroups(strAccount)
  End Select

End Sub


Sub AssignAccountGroups(strAccount)
  Call DebugLog("AssignAccountGroups: " & strAccount)

  strCmd            = "NET LOCALGROUP """ & strGroupUsers & """ """ & strAccount & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, 2)

  strCmd            = "NET LOCALGROUP """ & strGroupDistComUsers & """ """ & strAccount & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, 2)

  If strGroupPerfLogUsers <> "" Then
    strCmd          = "NET LOCALGROUP """ & strGroupPerfLogUsers & """ """ & strAccount & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, 2)
  End If

  If strGroupPerfMonUsers <> "" Then
    strCmd          = "NET LOCALGROUP """ & strGroupPerfMonUsers & """ """ & strAccount & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, 2)
  End If

End Sub


Sub AssignAccountRights(strAccount)
  Call DebugLog("AssignAccountRights: " & strAccount)

  Select Case True
    Case Ucase(strAccount) = Ucase(strGroupDBA)
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeManageVolumePrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeProfileSingleProcessPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeRemoteShutdownPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeShutdownPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeSystemProfilePrivilege")
    Case (Ucase(strAccount) = Ucase(strGroupDBANonSA)) And (strGroupDBANonSA <> "")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeProfileSingleProcessPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeSystemProfilePrivilege")
    Case strAccount = strSqlAccount
      Call RunNTRights("""" & strAccount & """ +r SeAssignPrimaryTokenPrivilege")    ' Replace a process-level token
      Call RunNTRights("""" & strAccount & """ +r SeBatchLogonRight")                ' Log on as a Batch Job
      Call RunNTRights("""" & strAccount & """ +r SeCreateGlobalPrivilege")          ' Create Global objects
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")          ' Bypass traverse checking
      Call RunNTRights("""" & strAccount & """ +r SeImpersonatePrivilege")           ' Impersonate a client after Authentication
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseBasePriorityPrivilege")  ' Adjust scheduling priority
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseQuotaPrivilege")         ' Adjust memory quotas
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseWorkingSetPrivilege")    ' Adjust Working Set
      Call RunNTRights("""" & strAccount & """ +r SeLockMemoryPrivilege")            ' Lock pages in memory
      Call RunNTRights("""" & strAccount & """ +r SeManageVolumePrivilege")          ' Manage files on a volume
      Call RunNTRights("""" & strAccount & """ +r SeProfileSingleProcessPrivilege")  ' Profile a process
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")              ' Log on as a Service
      Call RunNTRights("""" & strAccount & """ +r SeSystemProfilePrivilege")         ' Profile System performance
      Call RunNTRights("""" & strAccount & """ +r SeTcbPrivilege")                   ' Act as part of the Operating System
    Case strAccount = strAgtAccount
      Call RunNTRights("""" & strAccount & """ +r SeAssignPrimaryTokenPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeBatchLogonRight")
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeImpersonatePrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
    Case strAccount = strFTAccount
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseQuotaPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
    Case strAccount = strAsAccount
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeLockMemoryPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseBasePriorityPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseQuotaPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseWorkingSetPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeLockMemoryPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
    Case strAccount = strIsAccount
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeImpersonatePrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
    Case strAccount = strExtSvcAccount
      Call RunNTRights("""" & strAccount & """ +r SeAssignPrimaryTokenPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseQuotaPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
    Case Left(strAccount, Len(strNTService) + 1) = strNTService & "\"
      ' Nothing
    Case Else
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
  End Select

End Sub


Sub AssignFolderRights(strAccount, strFolder)
  Call DebugLog("AssignFolderRights: " & strFolder)

  Select Case True
    Case strAccount = strGroupDBANonSA
      strCmd        = """" & strFolder & """ /T /C /E /G """ & FormatAccount(strAccount) & """:R "
    Case Left(strAccount, Len(strNTService) + 1) = strNTService & "\"
      ' Nothing
    Case Else
      strCmd        = """" & strFolder & """ /T /C /E /G """ & FormatAccount(strAccount) & """:F "
  End Select
  Call RunCacls(strCmd)
  
End Sub


Sub PostPreparation()
  Call SetProcessId("1H", "Post Preparation Tasks")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1HA"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case Else
      Call SystemFolderPermissions()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1HB"
      ' Nothing
    Case Else
      Call GPUpdate()
  End Select

  Call SetProcessId("1HZ", " Post Preparation Tasks" & strStatusComplete)
  Call ProcessEnd("")

End Sub 


Sub SystemFolderPermissions()
  Call SetProcessId("1HA", "Set System Folder Permissions")

  Call SetKB2811566Permissions() 

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetKB2811566Permissions()
  Call SetProcessId("1HAA", "Set KB2811566 Permissions")
  Dim strLogPath

  strLogPath        = strDirSys & "\system32\LogFiles\Sum"

  Select Case True
   Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case strOSVersion < "6.2"
      ' Nothing
    Case Not objFSO.FolderExists(strLogPath)
      ' Nothing
    Case Else
      strCmd        = """" & strLogPath & """ /T /C /E /G """ & FormatAccount(strSQLAccount) & """:R"
      Call RunCacls(strCmd)
      strCmd        = """" & strLogPath & """ /T /C /E /G """ & FormatAccount(strSQLAccount) & """:W"
      Call RunCacls(strCmd)
  End Select

  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case strOSVersion < "6.2"
      ' Nothing
    Case Not objFSO.FolderExists(strLogPath)
      ' Nothing
    Case Else
      strCmd        = """" & strLogPath & """ /T /C /E /G """ & FormatAccount(strAsAccount) & """:R"
      Call RunCacls(strCmd)
      strCmd        = """" & strLogPath & """ /T /C /E /G """ & FormatAccount(strAsAccount) & """:W"
      Call RunCacls(strCmd)
  End Select

End Sub


Sub GPUpdate()
  Call SetProcessId("1HB", "Run GPUpdate to apply permissions")

  strCmd            = "HKLM\SOFTWARE\CLS\ITInfra\FFGPO_Update"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD")

  strCmd            = "GPUPDATE /Target:Computer /Force"
  Call Util_RunExec(strCmd, "", strResponseNo, -1)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub UserPreparation()
  Call SetProcessId("1U", "User Preparation Tasks")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",    strPathFBScripts)
  Call SetXMLParm(objInstParm, "ParmXtra",    GetBuildfileValue("FBParm"))
  Call RunInstall("UserPreparation", GetBuildfileValue("UserPreparationvbs"), objInstParm)

  Call ProcessEnd("")

End Sub 


Function Include(strFile)
  Dim objFSO, objFile
  Dim strFilePath, strFileText

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      err.Raise 8, "", "ERROR: This process must be run by SQLFineBuild.bat"
    Case Else
      Set objFSO        = CreateObject("Scripting.FileSystemObject")
      strFilePath       = strPathFB & "Build Scripts\" & strFile
      Set objFile       = objFSO.OpenTextFile(strFilePath)
      strFileText       = objFile.ReadAll()
      objFile.Close 
      ExecuteGlobal strFileText
  End Select

End Function


Sub RunNTRights(strCmd)
  Call DebugLog("RunNTRights: " & strCmd)

  Call Util_RunExec(strProgNtrights & " -u " & strCmd, "", strResponseYes, -1)
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case intErrSave = 2
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
  End Select

End Sub


Function GetPathLog()

  GetPathLog        = strSetupLog & strInstLog & strProcessIdLabel & " " & strProcessIdDesc & ".txt"""

End Function


End Class