'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageInstall.vbs  
'  Copyright FineBuild Team © 2017 - 2019.  Distributed under Ms-Pl License
'
'  Purpose:      Install routines required for the Build 
'
'  Author:       Ed Vassie
'
'  Date:         05 Jul 2017
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     05 Jul 2017  Initial version
'
'  Parameters:
'  Value         Default               Description
'  ParmMain      strPathAddComp        Primary Path to Install Module 
'  ParmXtra                            Additional Parameters for Install Module
'  StatusOption  strStatusComplete     Completion Status
'  InstallError  strStatusWarning      Error given on failed Install
'  LogXtra                             Additional label for Log File
'  MenuOption                          Control Menu setup
'  CleanBoot                           Force Reboot if pending
'  PreConType    String                Data Type of Pre-Con Registry Key
'  PreConValue                         Value to check in PreCon test
'  PathAlt                             Alternative Path to find Install Module
'  PreConKey                           Registry Rey to check in PreCon test
'  InstOption    Install               Install Option
'  SetupOption                         Pre-Install Option
'  InstTarget    strPathTemp           Target Path for Setup processing
'  InstFile      Setup.exe             File to be Installed
'  ParmExtract   /q /x:                Parameter for Setup Extract processing
'  PathLog       GetLogPath()          Path to Log File
'  ParmRetry                           List of Return Codes that allow a Retry
' TBC TBC TBC
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageInstall: Set FBManageInstall = New FBManageInstallClass
Dim strPathInst

Class FBManageInstallClass
Dim objFile, objFolder, objFSO, strLogXtra, objShell, objWMIReg
Dim strPathAddComp, strStatusVar, strWaitShort


Private Sub Class_Initialize
  Call DebugLog("FBManageInstall Class_Initialize:")

  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

  strPathAddComp    = GetBuildfileValue("PathAddComp")
  strPathInst       = ""
  strWaitShort      = GetBuildfileValue("WaitShort")

End Sub


Sub RunInstall(strInstName, strInstFile, objInstParm)
  Call DebugLog("RunInstall: " & strInstName & " using " & strInstFile)
  Dim strInstallError, strMenuOption, strStatusOption

  strInstallError   = UCase(GetXMLParm(objInstParm, "InstallError",  strMsgWarning))
  strLogXtra        = UCase(Replace(GetXMLParm(objInstParm, "LogXtra", "")," ", ""))
  strMenuOption     = UCase(GetXMLParm(objInstParm, "MenuOption", ""))
  strStatusVar      = "Setup" & strInstName & strLogXtra & "Status"
  Select Case True
    Case strLogXtra = ""
      strStatusOption = GetXMLParm(objInstParm, "StatusOption", strStatusComplete)
    Case Else
      strStatusOption = GetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  End Select
  Call SetBuildFileValue(strStatusVar, strStatusProgress)

  Select Case True
    Case RunInstall_PreCon(strInstName, strInstFile, objInstParm) 
      Exit Sub
    Case Not RunInstall_Setup(strInstName, strInstFile, objInstParm) 
      Call DebugLog(" " & strProcessIdDesc & strStatusBypassed & ", no media")
      Call SetBuildFileValue(strStatusVar, strStatusBypassed & ", no media")
      Exit Sub
    Case Not RunInstall_Process(strInstName, objInstParm)
      Select Case True
        Case intErrSave = 0
          ' Nothing
        Case UCase(strInstallError) = "IGNORE"
          ' Nothing
        Case Else
          Call SetBuildMessage(strInstallError, "SETUP" & strInstName & ": " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strPathInst)
          intErrSave = 0
      End Select
      Call DebugLog(" " & strProcessIdDesc & strStatusFail)
      Call SetBuildFileValue(strStatusVar, strStatusFail)
      Exit Sub
    Case strMenuOption <> ""
      Call RunInstall_Menu(strInstName, objInstParm) 
  End Select

  If strStatusOption <> strStatusProgress Then
    Call SetBuildFileValue(strStatusVar, strStatusOption)
    Call DebugLog(" " & strProcessIdDesc & strStatusOption)
  End If
  objInstParm       = ""

End Sub


Private Function RunInstall_PreCon(strInstName, strInstFile, objInstParm)
  Call DebugLog("RunInstall_PreCon:")
  Dim strCleanBoot, strPreConType

  RunInstall_PreCon = False
  strCleanBoot      = UCase(GetXMLParm(objInstParm, "CleanBoot",   ""))
  strPreConType     = UCase(GetXMLParm(objInstParm, "PreConType",  "String"))

  Select Case True
    Case strCleanBoot <> "YES"
      ' Nothing
    Case CheckReboot() = "Pending"
      Call SetupReboot(strProcessIdLabel, "Prepare for " & strInstName)
  End Select

  Select Case True
    Case strPreConType = "FILE"
      RunInstall_PreCon = RunInstall_PreCon_File(strInstName, strInstFile, objInstParm)
      If RunInstall_PreCon Then
        Call DebugLog(" " & strProcessIdDesc & strStatusBypassed)
        Call SetBuildFileValue(strStatusVar, strStatusBypassed)
      End If
    Case Else
      RunInstall_PreCon = RunInstall_PreCon_Registry(objInstParm)
      If RunInstall_PreCon Then
        Call DebugLog(" " & strProcessIdDesc & strStatusPreConfig)
        Call SetBuildFileValue(strStatusVar, strStatusPreConfig)
      End If
  End Select

End Function


Private Function RunInstall_PreCon_File(strInstName, strInstFile, objInstParm)
  Call DebugLog("RunInstall_PreCon_File:")
  Dim arrVerFile, arrVerPreCon
  Dim intIdx
  Dim strPathAlt, strPathMain, strPreConStatus, strPreConValue, strVerFile

  RunInstall_PreCon_File = True
  strPreConValue    = GetXMLParm(objInstParm, "PreConValue", "")
  strPathAlt        = GetXMLParm(objInstParm, "PathAlt", "")
  Select Case True
    Case UCase(strInstFile) = "DISM.EXE"
      strPathMain   = GetXMLParm(objInstParm, "PathMain", GetBuildfileValue("PathSys"))
    Case UCase(strInstFile) = "PKGMGR.EXE"
      strPathMain   = GetXMLParm(objInstParm, "PathMain", GetBuildfileValue("PathSys"))
    Case Else
      strPathMain   = GetXMLParm(objInstParm, "PathMain", FormatFolder(strPathAddComp))
  End Select

  If GetPathInst(strInstFile, strPathMain, strPathAlt) = "" Then
    Exit Function
  End If

  strVerFile        = objFSO.GetFileVersion(strPathInst)
  arrVerFile        = Split(strVerFile, ".")
  arrVerPreCon      = Split(strPreConValue, ".")
  If UBound(arrVerFile) < UBound(arrVerPreCon) Then
    Exit Function
  End If

  For intIdx = 0 To UBound(arrVerPreCon)
    Select Case True
      Case CLng(arrVerFile(intIdx)) < CLng(arrVerPreCon(intIdx))
        Exit Function
      Case intIdx > 0
        ' Nothing
      Case CLng(arrVerFile(intIdx)) > CLng(arrVerPreCon(intIdx))
        Exit Function
    End Select
  Next

  RunInstall_PreCon_File = False

End Function


Private Function RunInstall_PreCon_Registry(objInstParm)
  Call DebugLog("RunInstall_PreCon_Registry:")
  Dim intIdx
  Dim strPreConKey, strPreConStatus, strPreConType, strPreConValue

  RunInstall_PreCon_Registry = False
  strPreConKey      = GetXMLParm(objInstParm,       "PreConKey",   "")
  strPreConType     = UCase(GetXMLParm(objInstParm, "PreConType",  "String"))
  strPreConValue    = GetXMLParm(objInstParm,       "PreConValue", "")

  intIdx            = InstrRev(strPreConKey, "\")
  Select Case True
    Case strPreConKey = ""
      Exit Function
    Case strPreConType = "DWORD"
      objWMIReg.GetDWordValue  strHKLM,Left(strPreConKey, intIdx),Mid(strPreConKey, intIdx + 1),strPreConStatus
    Case strPreConType = "STRING"
      objWMIReg.GetStringValue strHKLM,Left(strPreConKey, intIdx),Mid(strPreConKey, intIdx + 1),strPreConStatus
  End Select

  Call DebugLog("Check Precon: " & strPreConKey)
  Select Case True
    Case IsNull(strPreConStatus)
      Exit Function
    Case strPreConStatus = ""
      Exit Function
    Case CStr(strPreConStatus) < strPreConValue
      Exit Function
  End Select

  RunInstall_PreCon_Registry = True

End Function


Private Function RunInstall_Setup(strInstName, strInstFile, objInstParm)
  Call DebugLog("RunInstall_Setup:")
  Dim objApp, objItem, objTarget
  Dim intItems
  Dim strCmd, strInstOption, strInstType, strNewFile, strNewPath, strParmExtract, strPathMain, strPathAlt, strPathTemp, strZipPath
  Dim strSetupOption, strStatusSetup

  RunInstall_Setup  = False
  Set objApp        = CreateObject ("Shell.Application")
  Call SetBuildfileValue("RebootLoop", "0")
  strNewPath        = ""
  strInstOption     = UCase(GetXMLParm(objInstParm, "InstOption",  "Install"))
  strInstType       = UCase(Right(strInstFile, 4))
  strPathAlt        = GetXMLParm(objInstParm, "PathAlt", "")
  strPathTemp       = GetBuildfileValue("PathTemp")
  strSetupOption    = UCase(GetXMLParm(objInstParm, "SetupOption", ""))

  Select Case True
    Case UCase(strInstFile) = "DISM.EXE"
      strPathMain   = GetXMLParm(objInstParm, "PathMain", GetBuildfileValue("PathSys"))
    Case UCase(strInstFile) = "PKGMGR.EXE"
      strPathMain   = GetXMLParm(objInstParm, "PathMain", GetBuildfileValue("PathSys"))
    Case Else
      strPathMain   = GetXMLParm(objInstParm, "PathMain", FormatFolder(strPathAddComp))
  End Select

  If GetPathInst(strInstFile, strPathMain, strPathAlt) = "" Then
    Exit Function
  End If

  Select Case True
    Case strInstType = ".ZIP"
      Call SetTrustedZone(strPathInst)
      strNewPath    = Replace(GetXMLParm(objInstParm, "InstTarget", strPathTemp) & "\" & strInstName, "\\" & strInstName, "\" & strInstName)
      strZipPath    = ""
      intItems      = 0
      Call CreateSetupFolder(strNewPath, "Y")
      Set objFile   = objFSO.GetFile(strPathInst)
      Set objFolder = objApp.NameSpace(objFile.Path).Items( )
      Set objTarget = objApp.NameSpace(strNewPath)
      objTarget.CopyHere objFolder, 256 + 16
      For Each objItem in objTarget.Items()
        intItems    = intItems + 1
        If objItem.IsFolder = True Then
          strZipPath = "\" & objItem.Name
        End If
      Next 
      Select Case True
        Case strZipPath = ""
          ' Nothing
        Case intItems > 1
          ' Nothing
        Case Else
          strNewPath = strNewPath & strZipPath
      End Select
      strNewFile    = GetXMLParm(objInstParm, "InstFile", "Setup.exe")
      strNewPath    = GetPathInst(strNewFile, strNewPath, "")
      Select Case True
        Case strInstOption = "NONE" 
          ' Nothing
        Case strNewPath = "" 
          Exit Function
      End Select
    Case strInstType = ".CAB"
      strNewPath    = Replace(GetXMLParm(objInstParm, "InstTarget", strPathTemp) & "\" & strInstName, "\\" & strInstName, "\" & strInstName)
      Call CreateSetupFolder(strNewPath, "Y")
      strDebugMsg1  = "Source: " & strPathInst
      strDebugMsg2  = "Target: " & strNewPath
      strCmd        = "EXPAND """ & strPathInst & """ -F:* """ & strNewPath & """"
      Call Util_RunExec(strCmd, "", "", 0)
      strNewFile    = GetXMLParm(objInstParm, "InstFile", "Setup.exe")
      strNewPath    = GetPathInst(strNewFile, strNewPath, "")
      Select Case True
        Case strInstOption = "NONE" 
          ' Nothing
        Case strNewPath = "" 
          Exit Function
      End Select
    Case strInstType = ".PDF"
      strNewPath    = Replace(GetXMLParm(objInstParm, "InstTarget", strPathTemp) & "\" & strInstName, "\\" & strInstName, "\" & strInstName)
      Call CreateSetupFolder(strNewPath, "Y")
      strDebugMsg1  = "Source: " & strPathInst
      strDebugMsg2  = "Target: " & strNewPath
      strNewFile    = Right(strPathInst, Len(strPathInst) - InstrRev(strPathInst, "\"))
      Set objFile   = objFSO.GetFile(strPathInst)
      objFile.Copy strNewPath & "\" & strNewFile, True
      If GetPathInst(strNewFile, strNewPath, "") = "" Then
        Exit Function
      End If
    Case strSetupOption = "EXTRACT" ' Extact contents from Install File before installing
      strNewPath    = Replace(GetXMLParm(objInstParm, "InstTarget", strPathTemp) & "\" & strInstName, "\\" & strInstName, "\" & strInstName)
      Call CreateSetupFolder(strNewPath, "Y")
      strDebugMsg1  = "Source: " & strPathInst
      strDebugMsg2  = "Target: " & strNewPath
      strParmExtract   = GetXMLParm(objInstParm, "ParmExtract", "/q /x:")
      If Right(strParmExtract, 1) <> ":" Then
        strParmExtract = strParmExtract & " "
      End If
      strCmd        = """" & strPathInst & """ " & strParmExtract & """" & strNewPath & """"
      Call Util_RunExec(strCmd, "", "", 0)
      strNewFile    = GetXMLParm(objInstParm, "InstFile", "Setup.exe")
      strNewPath    = GetPathInst(strNewFile, strNewPath, "")
      Select Case True
        Case strInstOption = "NONE" 
          ' Nothing
        Case strNewPath = "" 
          Exit Function
      End Select
    Case strSetupOption = "COPY" ' Copy Install File to local folder before installing
      strNewPath    = Replace(GetXMLParm(objInstParm, "InstTarget", strPathTemp) & "\" & strInstName, "\\" & strInstName, "\" & strInstName)
      Call CreateSetupFolder(strNewPath, "Y") 
      strDebugMsg1  = "Source: " & strPathInst
      strDebugMsg2  = "Target: " & strNewPath
      strNewFile    = Right(strPathInst, Len(strPathInst) - InstrRev(strPathInst, "\"))
      Set objFile   = objFSO.GetFile(strPathInst)
      objFile.Copy strNewPath & "\" & strNewFile, True
      If GetPathInst(strNewFile, strNewPath, "") = "" Then
        Exit Function
      End If
      Call SetTrustedZone(strPathInst)
  End Select

  RunInstall_Setup  = True
  Set objAPP        = Nothing
  Set objTarget     = Nothing

End Function


Private Sub CreateSetupFolder(strPath, strReset)
  Call DebugLog("CreateSetupFolder: " & strPath)
  Dim strPathFull, strPathPrev, strPathTemp

  strPathTemp       = GetBuildfileValue("PathTemp")

  strPathFull       = strPath
  If Right(strPathFull, 1) = "\" Then
    strPathFull     = Left(strPathFull, Len(strPathFull) - 1)
  End If

  Select Case True
    Case strReset <> "Y"
      ' Nothing
    Case Left(strPath, Len(strPath)) <> strPathTemp
      ' Nothing
    Case objFSO.FolderExists(strPathFull) 
      objFSO.DeleteFolder strPathFull, 1 
      Wscript.Sleep strWaitShort ' Wait for NTFS Cache to catch up
  End Select

  strPathPrev       = Left(strPathFull, InstrRev(strPathFull, "\") - 1)
  strDebugMsg1      = "PathPrev: " & strPathPrev
  Select Case True
    Case objFSO.FolderExists(strPathFull)
      ' Nothing
    Case objFSO.FolderExists(strPathPrev)
      objFSO.CreateFolder(strPathFull)
      Wscript.Sleep strWaitShort
    Case Else
      Call CreateSetupFolder(strPathPrev, "N")
      Wscript.Sleep strWaitShort
      objFSO.CreateFolder(strPathFull)
      Wscript.Sleep strWaitShort
  End Select

End Sub


Private Function RunInstall_Process(strInstName, objInstParm)
  Call DebugLog("RunInstall_Process:")
  Dim strCmd, strCompatFlags, strHKCU, strInstFile, strInstOption, strInstPrompt, strInstType
  Dim strMode, strMSILayer, strMSIAutoOS, strOSType, strOSVersion
  Dim strParmLog, strParmRetry, strPath, strPathCmd, strPathLog, strPathTemp

  RunInstall_Process = False
  strInstOption     = UCase(GetXMLParm(objInstParm, "InstOption", "Install"))
  strInstPrompt     = ""
  strInstFile       = Right(strPathInst, Len(strPathInst) - InstrRev(strPathInst, "\"))
  strInstType       = UCase(Right(strPathInst, 4))
  strMSILayer       = ""
  strCompatFlags    = GetBuildfileValue("CompatFlags")
  strHKCU           = &H80000001
  strMode           = GetBuildfileValue("Mode")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPathLog        = GetXMLParm(objInstParm, "PathLog",   GetPathLog(strLogXtra))
  strPathTemp       = GetBuildfileValue("PathTemp")
  strParmRetry      = GetXMLParm(objInstParm, "ParmRetry", "0")

  Select Case True
    Case strInstOption = "NONE"
      RunInstall_Process = True
      Exit Function
    Case strInstOption = "MENU"
      RunInstall_Process = True
      Exit Function
    Case strInstType = ".MSI"
      strMSIAutoOS  = UCase(GetXMLParm(objInstParm, "MSIAutoOS",  ""))
      If (strMSIAutoOS <> "") And (strOSVersion > strMSIAutoOS) Then ' Allow MSI to run on newer versions of Windows
        strMSILayer = strCompatFlags & "Layers"
        Call Util_RegWrite(strMSILayer & "\", "", "REG_SZ")
        objWMIReg.SetStringValue strHKCU, Mid(strMSILayer, 6), strPathInst, "MSIAUTO RUNASADMIN" & GetAppOS(strMSIAutoOS)
      End If
      strPathCmd    = "MSIEXEC /i """ & strPathInst & """ " & GetXMLParm(objInstParm, "ParmReboot", "/norestart")
      If strMode <> "ACTIVE" Then
        strPathCmd  = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmSilent", "/passive")
      End If
      strParmLog    = GetXMLParm(objInstParm, "ParmLog", "/log")
      Select Case True
        Case strParmLog = ""
          ' Nothing
        Case Right(strParmLog, 1) = ":"
          strPathCmd = strPathCmd & " " & strParmLog & strPathLog
        Case Else
          strPathCmd = strPathCmd & " " & strParmLog & " " & strPathLog
      End Select
      strPathCmd    = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmXtra", "")
    Case strInstType = ".MSP"
      strPathCmd    = "MSIEXEC /p """ & strPathInst & """ " & GetXMLParm(objInstParm, "ParmReboot", "/norestart")
      If strMode <> "ACTIVE" Then
        strPathCmd  = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmSilent", "/passive")
      End If
      strParmLog    = GetXMLParm(objInstParm, "ParmLog", "/log")
      Select Case True
        Case strParmLog = ""
          ' Nothing
        Case Right(strParmLog, 1) = ":"
          strPathCmd = strPathCmd & " " & strParmLog & strPathLog
        Case Else
          strPathCmd = strPathCmd & " " & strParmLog & " " & strPathLog
      End Select
      strPathCmd    = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmXtra", "")
    Case strInstType = ".MSU"
      strPathCmd    = "WUSA """ & strPathInst & """ " & GetXMLParm(objInstParm, "ParmReboot", "/norestart")
      If strMode <> "ACTIVE" Then
        strPathCmd  = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmSilent", "/quiet")
      End If
      strPathLog    = Replace(strPathLog, ".txt", ".evt")
      strParmLog    = GetXMLParm(objInstParm, "ParmLog", "/log:")
      Select Case True
        Case strParmLog = ""
          ' Nothing
        Case Right(strParmLog, 1) = ":"
          strPathCmd = strPathCmd & " " & strParmLog & strPathLog
        Case Else
          strPathCmd = strPathCmd & " " & strParmLog & " " & strPathLog
      End Select
      strPathCmd     = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmXtra", "")
    Case strInstType = ".EXE"
      Call SetTrustedZone(strPathInst)
      strPathCmd     = """" & strPathInst & """ " & GetXMLParm(objInstParm, "ParmReboot", "/norestart") 
      Select Case True
        Case UCase(strInstFile) = "DISM.EXE"
          strPathCmd = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmSilent", "")
        Case UCase(strInstFile) = "PKGMGR.EXE"
          strPathCmd = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmSilent", "/QUIET")
        Case strMode = "ACTIVE" 
          ' Nothing
        Case Instr(strOSType, "CORE") > 0
          strPathCmd = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmSilent", "/q")
        Case Else
          strPathCmd = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmSilent", "/passive")
      End Select
      strParmLog     = GetXMLParm(objInstParm, "ParmLog", "/log:")
      Select Case True
        Case UCase(strInstFile) = "DISM.EXE"
          ' Nothing
        Case UCase(strInstFile) = "PKGMGR.EXE"
          strPathCmd = strPathCmd & " " & GetXMLParm(objInstParm, "ParmLog", "/L:") & Replace(strPathLog, ".txt", ".log")
        Case strParmLog = ""
          ' Nothing
        Case Right(strParmLog, 1) = ":"
          strPathCmd = strPathCmd & " " & strParmLog & strPathLog
        Case Else
          strPathCmd = strPathCmd & " " & strParmLog & " " & strPathLog
      End Select
      strPathCmd    = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmXtra", "")
    Case strInstType = ".PS1"
      strPathCmd    = "POWERSHELL """ & strPathInst & """"
      strPathCmd    = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmXtra", "")
      strPathCmd    = strPathCmd & " > " & strPathLog
      strInstPrompt = "EOF"
    Case strInstType = ".SQL"
      strPathCmd    = GetBuildfileValue("CmdSQL") & " -i """ & strPathInst & """"
      strPathCmd    = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmXtra", "")
      strPathCmd    = strPathCmd & " -o " & strPathLog
      strInstPrompt = "EOF"
    Case strInstType = ".VBS"
      strPathCmd    = "%COMSPEC% /D /C CSCRIPT """ & strPathInst & """"
      strPathCmd    = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmXtra", "")
      strPathCmd    = strPathCmd & " > " & strPathLog
      strInstPrompt = "EOF"
    Case Else
      strPathCmd    = """" & strPathInst & """"
      strPathCmd    = strPathCmd & " "  & GetXMLParm(objInstParm, "ParmXtra", "")
      strPathCmd    = strPathCmd & " > " & strPathLog
  End Select

  strDebugMsg1      = "Log file: " & strPathLog
  Call Util_RunExec(strPathCmd, strInstPrompt, strResponseYes, -1)
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case intErrSave = 3010
      Call SetBuildfileValue("RebootStatus", "Pending")
      intErrSave    = 0
    Case (intErrSave = 5) And (strInstType = ".MSU")
      Exit Function
    Case intErrSave = 1642
      Exit Function
    Case intErrSave = -1073741818
      Call DebugLog("Retrying " & strInstName & " Install - Network connectivity workaround")
      Call Util_RunExec(strCmd, "", "", 0)
    Case (strParmRetry <> "0") And (Instr(" " & strParmRetry & " ", " " & Cstr(intErrSave) & " ") > 0)
      Call DebugLog("Retrying " & strInstName & " Install due to code " & Cstr(intErrSave))
      Call Util_RunExec(strCmd, "", "", 0)
    Case intErrSave = -2147205120 ' Install blocked
      Exit Function
    Case Else
      Exit Function
  End Select

  Select Case True
    Case Left(strPathInst, Len(strPathTemp)) <> strPathTemp
      ' Nothing
    Case Left(strPathInst, InstrRev(strPathInst, "\") - 1) = strPathTemp
      ' Nothing
    Case Else
      strPath       = Left(strPathInst, InstrRev(strPathInst, "\") - 1)
      strDebugMsg1  = "Deleting: " & strPath
      Set objFolder = objFSO.GetFolder(strPath)
      objFolder.Delete(1)
      If strMSILayer <> "" Then                
        objWMIReg.DeleteValue strHKCU, Mid(strMSILayer, 6), strPathInst
      End If
  End Select

  RunInstall_Process = True

End Function


Private Function GetAppOS(strAppOS)
  Call DebugLog("GetAppOS: " & strAppOS)

  Select Case True
    Case strAppOS = "5.1"
       GetAppOS     = " WINXPSP3"
    Case strAppOS = "5.2"
       GetAppOS     = " WINSRV03SP1"
    Case strAppOS = "6.0"
       GetAppOS     = " WINSRV08SP1"
    Case Else
       GetAppOS     = ""
  End Select

End Function


Private Sub RunInstall_Menu(strInstName, objInstParm)
  Call DebugLog("RunInstall_Menu:")
  Dim objShortcut
  Dim strMenuError,strMenuName, strMenuOption, strPathOld, strPathNew

  strMenuOption     = UCase(GetXMLParm(objInstParm, "MenuOption", ""))
  strMenuError      = UCase(GetXMLParm(objInstParm, "MenuError", strMsgWarning))

  Select Case True
    Case strMenuOption = "BUILD"
      strMenuName   = GetXMLParm(objInstParm, "MenuName", strInstName)
      strPathOld    = GetXMLParm(objInstParm, "MenuSource", strPathInst)
      strPathNew    = GetXMLParm(objInstParm, "MenuPath",   "")
      strDebugMsg1  = "Source: " & strPathOld
      strDebugMsg2  = "Target: " & strPathNew
      Select Case True
        Case Not objFSO.FileExists(strPathOld)
          If UCase(strMenuError) <> "IGNORE" Then
            Call SetBuildMessage(strMenuError, "SETUP" & strInstName & ": " & strMenuName & " Menu source file not found " & strPathOld)
          End If
          Exit Sub
        Case Not objFSO.FolderExists(strPathNew)
          objFSO.CreateFolder(strPathNew)
      End Select
      If Not objFSO.FileExists(strPathNew & "\" & strMenuName & ".lnk") Then
        Set objShortcut = objShell.CreateShortcut(strPathNew & "\" & strMenuName & ".lnk")
        objShortcut.TargetPath       = strPathOld
        objShortcut.WorkingDirectory = Left(strPathOld, InstrRev(strPathOld, "\") - 1)
        objShortcut.Save()
      End If
    Case strMenuOption = "MOVE"
      strPathOld    = GetXMLParm(objInstParm, "MenuSource", "")
      strPathNew    = GetXMLParm(objInstParm, "MenuPath",   "")
      strDebugMsg1  = "Source: " & strPathOld
      strDebugMsg2  = "Target: " & strPathNew
      If objFSO.FolderExists(strPathOld) Then
        If Not objFSO.FolderExists(strPathNew) Then
          objFSO.CreateFolder(strPathNew)
        End If
        Set objFolder = objFSO.GetFolder(strPathOld)
        objFolder.Copy strPathNew & "\" & objFolder.Name
        objFolder.Delete(1)
      End If
    Case strMenuOption = "REMOVE"
      strPathOld    = GetXMLParm(objInstParm, "MenuPath",   "") & "\" & GetXMLParm(objInstParm, "MenuPath",   "") & ".lnk"
      strDebugMsg1  = "Source: " & strPathOld
      If objFSO.FileExists(strPathOld) Then
        Set objFile     = objFSO.GetFile(strPathOld)
        objFile.Delete(1)
      End If
  End Select

End Sub


Function GetPathInst(strInstFile, strPathMain, strPathAlt)
  Call DebugLog("GetPathInst: " & strInstFile & " in " & strPathMain)
  Dim strInstNLS, strOSLanguage, strPathInstAlt

  strOSLanguage     = GetBuildfileValue("OSLanguage")
  strInstNLS        = Replace(strInstFile, "ENU", strOSLanguage, 1, -1, 1)
  strPathInst       = strPathMain
  strPathInstAlt    = strPathAlt
  If Right(strPathInst, 1) <> "\" Then
    strPathInst     = strPathInst & "\"
  End If
  If Right(strPathInstAlt, 1) <> "\" Then
    strPathInstAlt  = strPathInstAlt & "\"
  End If
  Select Case True
    Case objFSO.FileExists(strPathInst & strInstNLS)
      strPathInst   = strPathInst & strInstNLS
    Case objFSO.FileExists(strPathInst & strInstFile)
      strPathInst   = strPathInst & strInstFile
    Case (strPathInstAlt <> "") And (objFSO.FileExists(strPathInstAlt & strInstNLS))
      strPathInst   = strPathInstAlt & strInstNLS
    Case (strPathInstAlt <> "") And (objFSO.FileExists(strPathInstAlt & strInstFile))
      strPathInst   = strPathInstAlt & strInstFile
    Case Else 
      strPathInst   = ""
  End Select  

  GetPathInst       = strPathInst

End Function


Function GetPathLog(strLogXtra)
  Dim strInstLog, strPathLog

  strInstLog        = GetBuildfileValue("InstLog")
  strPathLog        = strSetupLog & strInstLog & strProcessIdLabel & " " & strProcessIdDesc
  If strLogXtra <> "" Then
    strPathLog      = strPathLog & " " & strLogXtra
  End If
  GetPathLog        = strPathLog & ".txt"""

End Function


Sub SetTrustedZone(strPathExe)
  Call DebugLog("SetTrustedZone: " & strPathExe)
  Dim strAltFile

  Set objFile       = objFSO.GetFile(strPathExe)
  strAltFile        = objFile.Path & ":Zone.Identifier"
  strDebugMsg1      = "Alt Path: " & strAltFile
  If objFSO.FileExists(strAltFile) Then
    Set objFile     = objFSO.CreateTextFile(strAltFile, True)
    objFile.WriteLine "[ZoneTransfer]"
    objFile.WriteLine "ZoneId=1"
    objFile.Close
  End If

End Sub


End Class


Sub RunInstall(strInstName, strInstFile, objInstParm)
  Call FBManageInstall.RunInstall(strInstName, strInstFile, objInstParm)
End Sub

Function GetPathInst(strInstFile, strPathMain, strPathAlt)
  ' This interface will be removed in a future version of SQL FineBuild
  GetPathInst       = FBManageInstall.GetPathInst(strInstFile, strPathMain, strPathAlt)
End Function

Function GetPathLog(strLogXtra)
  GetPathLog        = FBManageInstall.GetPathLog(strLogXtra)
End Function

Sub SetTrustedZone(strPathExe)
  ' This interface will be removed in a future version of SQL FineBuild
  Call FBManageInstall.SetTrustedZone(strPathExe)
End Sub