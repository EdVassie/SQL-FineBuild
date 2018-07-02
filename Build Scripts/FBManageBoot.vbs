'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageBoot.vbs  
'  Copyright FineBuild Team © 2017 - 2018.  Distributed under Ms-Pl License
'
'  Purpose:      Manage the FineBuild Boot processing 
'
'  Author:       Ed Vassie
'
'  Date:         02 Aug 2017
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     02 Aug 2017  Initial version

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageBoot: Set FBManageBoot = New FBManageBootClass

Class FBManageBootClass
  Dim objWMIReg
  Dim strOSVersion, strPathReg, strReboot, strRebootLimit, strRebootLoop, strWaitLong


Private Sub Class_Initialize
  Call DebugLog("FBManageBoot Class_Initialize:")

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

  strOSVersion      = GetBuildfileValue("OSVersion")
  strWaitLong       = GetBuildfileValue("WaitLong")

End Sub


Function CheckReboot()
  Call DebugLog("CheckReboot:")
  Dim arrOperations
  Dim objKey

  strReboot         = GetBuildfileValue("RebootStatus")
  strRebootLoop     = GetBuildfileValue("RebootLoop")
  strRebootLimit    = 2

  If strRebootLoop = "" Then
    strRebootLoop   = "0"
  End If
  strRebootLoop     = strRebootLoop + 1

  If strRebootLoop > 1 Then
    Wscript.Sleep strWaitLong ' Wait to allow pending operations to complete
  End If

  strPathReg        = "SYSTEM\CurrentControlSet\Control\Session Manager"
  strDebugMsg1      = "Source: " & strPathReg
  objWMIReg.GetMultiStringValue strHKLM,strPathReg,"PendingFileRenameOperations",arrOperations
  Select Case True
    Case strRebootLoop > strRebootLimit
      ' Nothing
    Case IsNull(arrOperations)
      ' Nothing
    Case Else
      strReboot     = "Pending"
  End Select

  strPathReg        = "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing"
  strDebugMsg1      = "Source: " & strPathReg
  objWMIReg.EnumKey strHKLM,strPathReg,arrOperations
  If Not IsNull(arrOperations) Then
    For Each objKey In arrOperations
      Select Case True
        Case strRebootLoop > strRebootLimit
          ' Nothing
        Case objKey = "RebootPending"
          strReboot   = "Pending"
      End Select
    Next
  End If

  Select Case True
    Case strRebootLoop > strRebootLimit
      Call DebugLog(strMsgWarning & " Reboot Loop detected in " & strProcessIdLabel & ", reboot request ignored")
    Case strReboot = "Pending"
      Call DebugLog("Reboot required")
  End Select

  Call SetBuildfileValue("RebootStatus", strReboot)
  Call SetBuildfileValue("RebootLoop", strRebootLoop)
  CheckReboot       = strReboot

End Function


Sub SetupReboot(strLabel, strDescription)
  Call DebugLog("SetupReboot:")
  Dim objShell
  Dim intIdx
  Dim strAdminPassword, strCmd, strFBCmd, strFBPath, strFBVol, strPath, strRebootDesc, strRebootLabel, strStopAt

  Call SetProcessId(strLabel, "Preparing server reboot")
  Set objShell      = CreateObject("Wscript.Shell")
  strRebootDesc     = "Reboot in progress: " & GetBuildfileValue("SQLVersion") & " " & strDescription
  strRebootLabel    = strLabel

  strFBCmd          = GetBuildfileValue("FBCmd")
  strFBVol          = Left(strFBCmd, 1)
  strFBPath         = GetBuildfileValue("Vol" & strFBVol & "Path")
  If strFBPath <> "" Then
    strFBPath       = "NET USE " & strFBVol & ": """ & strFBPath & """ /PERSISTENT:NO & "
  End If

  strCmd            = objShell.ExpandEnvironmentStrings("%COMSPEC%")  & " /D /k " 
  strCmd            = strCmd &  strFBPath & """" & strFBCmd & """"
  strCmd            = strCmd & " /Type:" & GetBuildfileValue("Type") & " /SQLVersion:" & GetBuildfileValue("AuditVersion") & " /Instance:" & GetBuildfileValue("Instance") & " /IAcceptLicenseTerms /Restart:Yes " 
  strStopAt         = GetBuildFileValue("StopAt")
  Select Case True
    Case GetBuildfileValue("StopAtFound") <> "YES"
      ' Nothing
    Case strStopAt = "AUTO"
      ' Nothing
    Case strStopAt < strProcessId
      ' Nothing
    Case Else
      strCmd        = strCmd & " /StopAt:" & strProcessId
  End Select

  strPath           = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\FineBuild" & strLabel
  strAdminPassword  = GetBuildfileValue("AdminPassword")
  Call Util_RegWrite(strPath, strCmd, "REG_SZ")
  Call DebugLog("Restart Command: " & strCmd)
  Call SetBuildfileValue("RebootStatus", "Done")
  If strAdminPassword <> "" Then
    strPath         = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon"
    Call Util_RegWrite(strPath, "1", "REG_SZ")
    strPath         = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultDomainName"
    Call Util_RegWrite(strPath, GetBuildfileValue("Domain"), "REG_SZ")
    strPath         = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DisableCAD"
    Call Util_RegWrite(strPath, "0", "REG_DWORD")
    strPath         = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName"
    Call Util_RegWrite(strPath, GetBuildfileValue("AuditUser"), "REG_SZ")
    strCmd          = "REG ADD ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"" /v DefaultPassword /d """ & strAdminPassword & """ /t REG_SZ /f"
    Call Util_RunExec(strCmd, "", strResponseYes, 0)
    strPath         = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoLogonCount"
    Call Util_RegWrite(strPath, GetBuildfileValue("AutoLogonCount"), "REG_DWORD")
  End If

  strCmd            = "SHUTDOWN /r /t 5 /f /d p:4:2 /c """ & strRebootDesc & """"
  Call Util_RunCmdAsync(strCmd, 0)
  Call FBLog("**************************************************")
  Call FBLog("*")
  Call SetProcessId(strRebootLabel, "Server reboot (" & strRebootDesc & ")")
  Call FBLog("*")
  Call FBLog("**************************************************")
  err.Raise 3010, strRebootDesc, "Reboot required"

End Sub


End Class


Function CheckReboot()
  CheckReboot       = FBManageBoot.CheckReboot()
End Function

Sub SetupReboot(strLabel, strDescription)
  Call FBManageBoot.SetupReboot(strLabel, strDescription)
End Sub