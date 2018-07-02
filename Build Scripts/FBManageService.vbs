'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageService.vbs  
'  Copyright FineBuild Team © 2018.  Distributed under Ms-Pl License
'
'  Purpose:      Start and Stop SQL Server Services 
'
'  Author:       Ed Vassie
'
'  Date:         11 Jul 2017
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     11 Jul 2017  Initial version

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageService: Set FBManageService = New FBManageServiceClass

Class FBManageServiceClass
  Dim objFile, objFolder, objFSO, objShell
  Dim strActionSQLAS, strActionSQLDB, strActionSQLRS, strCmd, strClusterName, strResSuffix, strSetupSQLAS, strSetupSQLDB, strSetupSQLDBCluster, strSetupSQLRS, strSetupSQLRSCluster, strWaitLong, strWaitShort


Private Sub Class_Initialize
  Call DebugLog("FBManageService Class_Initialize:")

  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objShell      = CreateObject ("Wscript.Shell")

  strActionSQLAS    = GetBuildfileValue("ActionSQLAS")
  strActionSQLDB    = GetBuildfileValue("ActionSQLDB")
  strActionSQLRS    = GetBuildfileValue("ActionSQLRS")
  strClusterName    = GetBuildfileValue("ClusterName")
  strResSuffix      = GetBuildfileValue("ResSuffix")
  strSetupSQLAS     = GetBuildfileValue("SetupSQLAS")
  strSetupSQLDB     = GetBuildfileValue("SetupSQLDB")
  strSetupSQLDBCluster = GetBuildfileValue("SetupSQLDBCluster")
  strSetupSQLRS     = GetBuildfileValue("SetupSQLRS")
  strSetupSQLRSCluster = GetBuildfileValue("SetupSQLRSCluster")
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitShort      = GetBuildfileValue("WaitShort")

End Sub


Sub StopSQL()
  Call DebugLog("StopSQL:")

  Call StopSSRS()

  Call StopSSAS()

  Select Case True
    Case GetBuildfileValue("SetupSQLDBAG") <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      Call Util_RunExec("NET STOP " & GetBuildfileValue("InstAgent"), "", "", 2)
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE ""SQL Server Agent" & strResSuffix & """ /OFF"
      Call Util_RunExec(strCmd, "", "", 0)
  End Select

  Call StopSQLServer()

End Sub


Sub StopSQLServer()
  Call DebugLog("StopSQLServer:")

  Select Case True
    Case GetBuildfileValue("SetupAnalytics") <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      Call Util_RunExec("NET STOP " & GetBuildfileValue("InstAnal"), "", "", 2)
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE ""SQL Server Launchpad" & strResSuffix & """ /OFF"
      Call Util_RunExec(strCmd, "", "", 0)
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      Call Util_RunExec("%COMSPEC% /D /C NET STOP " & GetBuildfileValue("InstSQL") & " /Y", "", "", 2)
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE ""SQL Server" & strResSuffix & """ /OFF"
      Call Util_RunExec(strCmd, "", "", -1)
  End Select

End Sub


Sub StartSQL()
  Call DebugLog("StartSQL:")

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      Call Util_RunExec("NET START " & GetBuildfileValue("InstSQL"), "", "", 2)
      Call CheckSQLReady()
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE ""SQL Server" & GetBuildfileValue("ResSuffix") & """ /ON"
      Call Util_RunExec(strCmd, "", "", 0)
      If GetBuildfileValue("Instance") <> "MSSQLSERVER" Then
        strCmd      = "NET START ""SQLBrowser"""
        Call Util_RunExec(strCmd, "", "", 2)  
      End If
      If strActionSQLDB <> "ADDNODE" Then
        Call CheckSQLReady()
      End If
  End Select

End Sub


Private Sub CheckSQLReady()
  Call DebugLog("CheckSQLReady:")
  Dim intFound
  Dim strCmd, strErrorLog, strFileData, strInstRegSQL, strPath, strRestartComplete, strSQLLogReinit

  Wscript.Sleep strWaitShort
  strInstRegSQL     = GetBuildfileValue("InstRegSQL")
  strRestartComplete  = GetBuildfileValue("SQLRecoveryComplete")
  strSQLLogReinit   = GetBuildfileValue("SQLLogReinit")
  strPath           = GetBuildfileValue("HKLMSQL") & strInstRegSQL & "\MSSQLServer\Parameters\SQLArg1"
  strErrorLog       = Mid(objShell.RegRead(strPath), 3)

  Select Case True
    Case objFSO.FileExists(strErrorLog)
      ' Nothing
    Case UCase(Right(strErrorlog, 4)) = ".OUT"
      strErrorlog   = Left(strErrorlog, Len(strErrorlog) - 4)
  End Select

  intFound          = 0
  strDebugMsg1      = "SQL Log File Path: " & strErrorLog
  strCmd            = "%COMSPEC% /D /C FIND /C """ & strSQLLogReinit & """ """ & strErrorLog & """"
  Call Util_RunExec(strCmd, "", "", 1)
  If intErrSave = 0 Then
    intFound        = 1   
  End If
  strCmd            = "%COMSPEC% /D /C FIND /C """ & strRestartComplete & """ """ & strErrorLog & """"
  While intFound = 0
    Wscript.Sleep strWaitLong 
    Call Util_RunExec(strCmd, "", "", 1)
    If intErrSave = 0 Then
      intFound      = 1   
    End If
  WEnd

  Wscript.Sleep strWaitLong 

End Sub


Sub StartSQLAgent()
  Call DebugLog("StartSQLAgent:")

  Select Case True
    Case GetBuildfileValue("SetupSQLDBAG") <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      Call Util_RunExec("NET START " & GetBuildfileValue("InstAgent"), "", "", 2)
      Call CheckSQLAgentReady()
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ GROUP """ & GetBuildfileValue("ClusterGroupSQL") & """ /ON"
      Call Util_RunExec(strCmd, "", "", 0)
      Call CheckSQLAgentReady()
  End Select

End Sub


Private Sub CheckSQLAgentReady()
  Call DebugLog("CheckSQLAgentReady:")
  Dim intFound
  Dim strErrorLog, strFileData, strPath, strRestartComplete

  strRestartComplete = GetBuildfileValue("SQLAgentStart")
  strRestartComplete = Replace(strRestartComplete, "SQLSERVERAGENT", GetBuildfileValue("InstAgent"))
  strPath           = GetBuildfileValue("HKLMSQL") & GetBuildfileValue("InstRegSQL") & "\SQLServerAgent\ErrorLogFile"
  strErrorLog       = objShell.RegRead(strPath)

  Select Case True
    Case strActionSQLDB = "ADDNODE"
      Exit Sub
    Case objFSO.FileExists(strErrorLog)
      ' Nothing
  End Select

  intFound          = 0
  strDebugMsg1      = "SQL Log File Path: " & strErrorLog
  strCmd            = "%COMSPEC% /D /C FIND /C """ & strRestartComplete & """ """ & strErrorLog & """"
  While intFound = 0
    Wscript.Sleep strWaitLong 
    Call Util_RunExec(strCmd, "", "", 1)
    If intErrSave = 0 Then
      intFound      = 1   
    End If
  WEnd

  Wscript.Sleep strWaitLong

End Sub


Sub StopSSAS()
  Call DebugLog("StopSSAS:")

  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case GetBuildfileValue("SetupSQLASCluster") <> "YES"
      Call Util_RunExec("NET STOP " & GetBuildfileValue("InstAS"), "", "", 2)
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE ""Analysis Services" & GetBuildfileValue("ResSuffixAS") & """ /OFF"
      Call Util_RunExec(strCmd, "", "", 0)
  End Select

End Sub


Sub StartSSAS()
  Call DebugLog("StartSSAS:")

  Select Case True
    Case strSetupSQLAS <> "YES" 
      ' Nothing
    Case GetBuildfileValue("SetupSQLASCluster") <> "YES"
      Call Util_RunExec("NET START " & GetBuildfileValue("InstAS"), "", "", 2)
      Call CheckSQLASReady()
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ GROUP """ & GetBuildfileValue("ClusterGroupAS") & """ /ON"
      Call Util_RunExec(strCmd, "", "", 0)
      Call CheckSQLASReady()
  End Select

End Sub


Private Sub CheckSQLASReady()
  Call DebugLog("CheckSQLASReady:")

  Select Case True
    Case strActionSQLAS = "ADDNODE"
      Exit Sub
  End Select

  Wscript.Sleep strWaitLong 

End Sub


Sub StopSSRS
  Call DebugLog("StopSSRS:")

  Select Case True
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case CheckStatus("SQLRSCluster")
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE """ & GetBuildfileValue("ClusterNameRS") & """ /OFF"
      Call Util_RunExec(strCmd, "", "", 0)
    Case Else
      Call Util_RunExec("NET STOP " & GetBuildfileValue("InstRS"), "", "", 2)
  End Select

End Sub


Sub StartSSRS(strOpt)
  Call DebugLog("StartSSRS:")
  Dim strClusterGroupRS

  Select Case True
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case CheckStatus("SQLRSCluster")
      strClusterGroupRS = GetBuildfileValue("ClusterGroupRS")
      strCmd        = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroupRS & """ /MOVETO:""" & GetBuildfileValue("AuditServer") & """" 
      Call Util_RunExec(strCmd, "", GetBuildfileValue("ResponseYes"), 0)
      strCmd        = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroupRS & """ /ON"
      Call Util_RunExec(strCmd, "", "", 0)
      Call CheckRSReady()
    Case strOpt = "FORCE"
      Call Util_RunExec("NET START " & GetBuildfileValue("InstRS"), "", "", 2)
    Case strSetupSQLRSCluster = "YES"
      Call Util_RunExec("NET START " & GetBuildfileValue("InstRS"), "", "", 2)
      Call CheckRSReady()
    Case UCase(Left(GetBuildfileValue("RSInstallMode"), 9)) = UCase("FilesOnly")
      ' Nothing
    Case Else
      Call Util_RunExec("NET START " & GetBuildfileValue("InstRS"), "", "", 2)
      Call CheckRSReady()
  End Select

End Sub


Private Sub CheckRSReady()
  Call DebugLog("CheckRSReady:")
  Dim intFound
  Dim strCmd, strErrorLog, strErrorLogDate, strFileData, strPath, strRestartComplete, strSQLVersion, strWaitLong

  strRestartComplete  = GetBuildfileValue("SQLRSStart")
  strSQLVersion     = GetBuildfileValue("SQLVersion")
  strWaitLong       = GetBuildfileValue("WaitLong")
  strPath           = GetBuildfileValue("HKLMSQL") & GetBuildfileValue("InstRegRS") & GetBuildfileValue("InstRSDir")
  strPath           = objShell.RegRead(strPath)
  strDebugMsg1      = "Path Root: " & strPath
  Select Case True
    Case strActionSQLRS = "ADDNODE"
      Exit Sub
    Case strSQLVersion >= "SQL2017"
      strPath       = strPath & "\" & GetBuildfileValue("InstRSSQL") & "\"  & "LogFiles"
    Case GetBuildfileValue("SetupPowerBI") = "YES"
      strPath       = strPath & "\" & GetBuildfileValue("InstRSSQL") & "\"  & "LogFiles"
    Case Else
      strPath       = strPath & "LogFiles"
  End Select

  Set objFolder     = objFSO.GetFolder(strPath)
  strErrorLogDate   = 0
  For Each objFile In objFolder.Files
    Select Case True
      Case (strSQLVersion = "SQL2005") And (Left(objFile.Name, 25) = "ReportServerService__main")
        ' Nothing
      Case Left(objFile.Name, 20) <> "ReportServerService_"
        ' Nothing
      Case DateDiff("S", strErrorLogDate, objFile.DateCreated) > 0 
        strErrorLogDate = objFile.DateCreated
        strErrorLog     = objFile.Name
    End Select
  Next  

  intFound          = 0
  strDebugMsg1      = "SSRS Log File: " & strErrorLog 
  strCmd            = "%COMSPEC% /D /C FIND /C """ & strRestartComplete & """ """ & strPath & "\" & strErrorLog & """"
  While (intFound = 0) And (strErrorLog > "")
    Call Util_RunExec(strCmd, "", "", 1)
    If intErrSave = 0 Then
      intFound      = 1
    End If
    Wscript.Sleep strWaitLong  
  WEnd

End Sub


End Class


Sub StopSQL()
  Call FBManageService.StopSQL()
End Sub

Sub StopSQLServer()
  Call FBManageService.StopSQLServer()
End Sub

Sub StopSSAS()
  Call FBManageService.StopSSAS()
End Sub

Sub StopSSRS()
  Call FBManageService.StopSSRS()
End Sub

Sub StartSQL()
  Call FBManageService.StartSQL()
End Sub

Sub StartSQLAgent()
  Call FBManageService.StartSQLAgent()
End Sub

Sub StartSSAS()
  Call FBManageService.StartSSAS()
End Sub

Sub StartSSRS(strOpt)
  Call FBManageService.StartSSRS(strOpt)
End Sub