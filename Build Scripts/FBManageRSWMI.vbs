'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageRSWMI.vbs  
'  Copyright FineBuild Team © 2018 - 2019.  Distributed under Ms-Pl License
'
'  Purpose:      Manage RS WMI processes 
'
'  Author:       Ed Vassie
'
'  Date:         26 Oct 2017
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     26 Oct 2017  Initial version

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageRSWMI: Set FBManageRSWMI = New FBManageRSWMIClass

Dim objRSInParam, objRSOutParam
Dim intRSLcid

Class FBManageRSWMIClass
  Dim objRSConfig, objRSWMI, objShell
  Dim strFunction, strHTTP, strInstRS, strInstRSSQL, strOSVersion, strPath, strRSAlias, strRSNamespace, strRSWMIPath, strSetupPowerBI, strSetupSQLRSCluster, strTCPPortRS, strSQLVersion, strWMIPath


Private Sub Class_Initialize
  Call DebugLog("FBManageRSWMI Class_Initialize:")
  Dim strInstRSWMI, strRSVersionNum

  Set objShell      = WScript.CreateObject ("Wscript.Shell")

  strHTTP           = GetBuildfileValue("HTTP")
  strInstRS         = GetBuildfileValue("InstRS")
  strInstRSSQL      = GetBuildfileValue("InstRSSQL")
  strInstRSWMI      = GetBuildfileValue("InstRSWMI")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strRSAlias        = GetBuildfileValue("RSAlias")
  strRSNamespace    = "MSReportServer_ConfigurationSetting"
  strRSVersionNum   = GetBuildfileValue("RSVersionNum")
  strSetupPowerBI   = GetBuildfileValue("SetupPowerBI")
  strSetupSQLRSCluster = GetBuildfileValue("SetupSQLRSCluster")
  strSQLVersion     = GetBuildfileValue("SQLVersion")
  strTCPPortRS      = GetBuildfileValue("TCPPortRS")

  Select Case True
    Case strSQLVersion <= "SQL2005"
      strRSWMIPath  = "winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\SqlServer\ReportServer\v" & strRSVersionNum & "\Admin"
    Case strSQLVersion >= "SQL2017"
      strRSWMIPath  = "winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\SqlServer\ReportServer\" & strInstRSWMI & "\V" & strRSVersionNum & "\Admin"
    Case strSetupPowerBI = "YES"
      strRSWMIPath  = "winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\SqlServer\ReportServer\" & strInstRSWMI & "\V" & strRSVersionNum & "\Admin"
    Case Else
      strRSWMIPath  = "winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\SqlServer\ReportServer\" & strInstRSWMI & "\v" & strRSVersionNum & "\Admin"
  End Select

End Sub


Function RunRSWMI(strFunction, strOK)
  Call DebugLog("RunRSWMI: " & strFunction)
  Dim intErrSave

  strWMIPath        = strRSNamespace & ".InstanceName='" & strInstRSSQL & "'"
  strDebugMsg1      = "WMI Path: " & strWMIPath
  Set objRSOutParam = objRSWMI.ExecMethod(strWMIPath, strFunction, objRSInParam)
  intErrSave        = objRSOutParam.HRESULT
  If intErrSave = -2147023181 Then
    WScript.Sleep GetBuildfileValue("WaitLong")
    WScript.Sleep GetBuildfileValue("WaitLong")
    Set objRSOutParam = objRSWMI.ExecMethod(strWMIPath, strFunction, objRSInParam)
    intErrSave      = objRSOutParam.HRESULT
  End If
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case Instr(" " & strOK & " ", " " & CStr(intErrSave) & " ") > 0
      intErrSave    = 0
    Case intErrSave = -2147023181
      Call SetBuildMessage(strMsgError,   "Unexpected WMI error: " & CStr(intErrSave) & " for " & strFunction & ": " & "RPC Server not available")
    Case strSQLVersion = "SQL2005"
      Call SetBuildMessage(strMsgWarning, "Unexpected WMI error: " & CStr(intErrSave) & " for " & strFunction)
    Case Else
      Call SetBuildMessage(strMsgWarning, "Unexpected WMI error: " & CStr(intErrSave) & " for " & strFunction & ": " & objRSOutParam.Error)
  End Select

  RunRSWMI        =  intErrSave

End Function


Function SetRSInParam(strFunction)
  Call DebugLog("SetRSInParam: " & strFunction)

  If Not IsObject(objRSWMI) Then
    Call SetRSWMI()
  End If

  strWMIPath        = strRSNamespace & ".InstanceName='" & strInstRSSQL & "'"
  strDebugMsg1      = "WMI Path: " & strWMIPath
  Set objRSConfig   = objRSWMI.Get(strWMIPath)
  Set objRSInParam  = objRSConfig.Methods_(strFunction).inParameters.SpawnInstance_()

  SetRSInParam      = strFunction

End Function


Sub SetRSDirectory(strApplication, strDirectory)
  Call DebugLog("SetRSDirectory: " & strApplication & ", " & strDirectory)
  Dim strStoreNamespace

  strStoreNamespace = strRSNamespace

  Select Case True
    Case strSQLVersion <= "SQL2005"
      If strApplication = "ReportManager" Then
        strFunction = SetRSInParam("SetWebServiceIdentity")
        objRSInParam.Properties_.Item("ApplicationPool")  = "DefaultAppPool"
        Call RunRSWMI(strFunction, "")
        strRSNamespace = "MSReportManager_ConfigurationSetting"
      End If
      strFunction   = SetRSInParam("CreateVirtualDirectory")
    Case Else
      strFunction   = SetRSInParam("SetVirtualDirectory")
  End Select

  Select Case True
    Case strSQLVersion <= "SQL2005"
      objRSInParam.Properties_.Item("IISPath")          = "IIS://localhost/w3svc/1/root"
      objRSInParam.Properties_.Item("Name")             = strDirectory
    Case strSetupPowerBI = "YES"
      objRSInParam.Properties_.Item("Application")      = strApplication
      objRSInParam.Properties_.Item("VirtualDirectory") = strDirectory
      objRSInParam.Properties_.Item("Lcid")             = intRSLcid
    Case Else
      objRSInParam.Properties_.Item("Application")      = strApplication
      objRSInParam.Properties_.Item("VirtualDirectory") = strDirectory
      objRSInParam.Properties_.Item("lcid")             = intRSLcid
  End Select
  Call RunRSWMI(strFunction, "-2147220930 -2147220978 -2147220980") ' OK if Directory already exists

  Select Case True
    Case strSQLVersion <= "SQL2005"
      ' Nothing
    Case Else
      Call SetRSURL(strApplication, strDirectory)
  End Select

  strRSNamespace    = strStoreNamespace

End Sub


Private Sub SetRSURL(strApplication, strDirectory)
  Call DebugLog("SetRSURL: " & strApplication)
  Dim strClusterIPV4RS, strClusterIPV6RS, strStoreNamespace, strURL, strURLVar

  strStoreNamespace = strRSNamespace

  Select Case True
    Case strSQLVersion <= "SQL2005"
      strRSNamespace = "MSReportManager_ConfigurationSetting"
      strFunction    = SetRSInParam("SetReportServerURLs")
      objRSInParam.Properties_.Item("ReportServerVirtualDirectory") = strDirectory
      objRSInParam.Properties_.Item("ReportServerExternalURL")      = ""
      strURLVar      = "ReportServerURL"
    Case strSetupPowerBI = "YES"
      strFunction    = SetRSInParam("ReserveURL")
      objRSInParam.Properties_.Item("Application")                  = strApplication
      objRSInParam.Properties_.Item("Lcid")                         = intRSLcid
      strURLVar      = "UrlString"
    Case Else
      strFunction    = SetRSInParam("ReserveURL")
      objRSInParam.Properties_.Item("Application")                  = strApplication
      objRSInParam.Properties_.Item("lcid")                         = intRSLcid
      strURLVar      = "URLString"
  End Select
  strURL             = strHTTP & "://" & GetBuildfileValue("AuditServer") & ":" & strTCPPortRS
  strDebugMsg1       = "URL: " & strURL
  objRSInParam.Properties_.Item(CStr(strURLVar)) = strURL
  Call RunRSWMI(strFunction, "-2147220932") ' OK if URL already exists

  Select Case True
    Case strSetupSQLRSCluster <> "YES"
      ' Nothing
    Case Else
      objRSInParam.Properties_.Item(CStr(strURLVar)) = strHTTP & "://" & GetBuildfileValue("ClusterGroupRS") & ":" & strTCPPortRS
      Call RunRSWMI(strFunction, "-2147220932") ' OK if URL already exists
  End Select

  Select Case True
    Case strRSAlias = ""
      ' Nothing
    Case Else
      objRSInParam.Properties_.Item(CStr(strURLVar)) = strHTTP & "://" & strRSAlias & ":" & strTCPPortRS
      Call RunRSWMI(strFunction, "-2147220932")
  End Select

  strClusterIPV4RS = GetBuildfileValue("ClusterIPV4RS")
  Select Case True
    Case strClusterIPV4RS = ""
      ' Nothing
    Case Else
      objRSInParam.Properties_.Item(CStr(strURLVar)) = strHTTP & "://" & strClusterIPV4RS & ":" & strTCPPortRS
      Call RunRSWMI(strFunction, "-2147220932")
  End Select

  strClusterIPV6RS = GetBuildfileValue("ClusterIPV6RS")
  Select Case True
    Case strClusterIPV6RS = ""
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case Else
      objRSInParam.Properties_.Item(CStr(strURLVar)) = strHTTP & "://" & strClusterIPV6RS & ":" & strTCPPortRS
      Call RunRSWMI(strFunction, "-2147220932")
  End Select

  strRSNamespace    = strStoreNamespace

End Sub


Private Sub SetRSWMI()
  Call DebugLog("SetRSWMI: " & strRSWMIPath)

  Set objRSWMI      = GetObject(strRSWMIPath)

  Select Case True
    Case strSQLVersion >= "SQL2017"
      intRSLcid     = 1033
    Case strSetupPowerBI = "YES"
      intRSLcid     = 1033
    Case Else
      strPath       = GetBuildfileValue("HKLMSQL") & GetBuildfileValue("InstRegRS") & "\Setup\Language"
      intRSLcid     = objShell.RegRead(strPath)
  End Select

End Sub


End Class


Function RunRSWMI(strFunction, strOK)
  RunRSWMI        = FBManageRSWMI.RunRSWMI(strFunction, strOK)
End Function

Sub SetRSDirectory(strApplication, strDirectory)
  Call FBManageRSWMI.SetRSDirectory(strApplication, strDirectory)
End Sub

Function SetRSInParam(strFunction)
  SetRSInParam      = FBManageRSWMI.SetRSInParam(strFunction)
End Function