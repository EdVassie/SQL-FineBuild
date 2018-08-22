'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBUtils.vbs  
'  Copyright FineBuild Team © 2017 - 2018.  Distributed under Ms-Pl License
'
'  Purpose:      Manage the FineBuild Log File 
'
'  Author:       Ed Vassie
'
'  Date:         05 Jul 2017
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     05 Jul 2017  Initial version

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBUtils: Set FBUtils = New FBUtilsClass
Dim intErrSave
Dim strHKCR, strHKLM, strErrSave, strResponseYes, strResponseNo

Class FBUtilsClass

Dim objAutoUpdate, objFSO, objShell, objWMIReg
Dim colPrcEnvVars
Dim strCmd, strCmdPS, strGroupDBA, strGroupDBANonSA, strIsInstallDBA, strPath, strServer, strSIDDistComUsers, strUserAccount
Dim intIdx


Private Sub Class_Initialize
  Call DebugLog("FBUtils Class_Initialize:")

  Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objShell      = CreateObject("Wscript.Shell")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colPrcEnvVars = objShell.Environment("Process")

  intErrSave        = 0
  strErrSave        = ""
  strHKCR           = &H80000000
  strHKLM           = &H80000002
  strCmdPS          = GetBuildfileValue("CmdPS")
  strGroupDBA       = GetBuildfileValue("GroupDBA")
  strGroupDBANonSA  = GetBuildfileValue("GroupDBANonSA")
  strIsInstallDBA   = GetBuildfileValue("IsInstallDBA")
  strServer         = GetBuildfileValue("AuditServer")
  strSIDDistComUsers  = GetBuildfileValue("SIDDistComUsers")
  strResponseNo     = GetBuildfileValue("ResponseNo")
  strResponseYes    = GetBuildfileValue("ResponseYes")
  strUserAccount    = GetBuildfileValue("UserAccount")

End Sub


Function FormatAccount(strAccount)
  Call DebugLog("FormatAccount: " & strAccount)
  Dim intBuiltinDomLen, intNTAuthLen, intServerLen
  Dim strBuiltinDom, strNTAuth

  strBuiltinDom     = GetBuildfileValue("BuiltinDom")
  intBuiltinDomLen  = Len(strBuiltinDom) + 1
  strNTAuth         = GetBuildfileValue("NTAuth")
  intNTAuthLen      = Len(strNTAuth) + 1
  intServerLen      = Len(strServer) + 1
  Select Case True
    Case Left(strAccount, intNTAuthLen) = strNTAuth & "\"
      FormatAccount = Mid(strAccount, intNTAuthLen + 1)
    Case Left(strAccount, intServerLen) = strServer & "\"
      FormatAccount = Mid(strAccount, intServerLen + 1)
    Case Left(strAccount, intBuiltinDomLen) = strBuiltinDom & "\"
      FormatAccount = Mid(strAccount, intBuiltinDomLen + 1)
    Case Else
      FormatAccount = strAccount
  End Select

End Function


Function FormatFolder(strFolder)
  Call DebugLog("FormatFolder: " &strFolder)
  Dim strFBLocal, strFBRemote, strFolderPath

  strFBLocal        = GetBuildfileValue("FBPathLocal")
  strFBRemote       = GetBuildfileValue("FBPathRemote")

  Select Case True
    Case Mid(strFolder, 2, 1) = ":"
      strFolderPath = strFolder
    Case Left(strFolder, 2) = "\\"
      strFolderPath = strFolder
    Case Else
      strFolderPath = GetBuildfileValue(strFolder)
  End Select
  
  Select Case True
    Case strFBLocal = strFBRemote
      ' Nothing
    Case UCase(Left(strFolderPath, Len(strFBRemote))) = strFBRemote
      strFolderPath = strFBLocal & Mid(strFolderPath, Len(strFBRemote) + 1)
  End Select

  FormatFolder      = strFolderPath  
 
End Function


Function FormatFolderURI(strFolder)
  Call DebugLog("FormatFolderURI: " &strFolder)
  Dim strFolderPath

  strFolderPath     = FormatFolder(strFolder)

  strFolderPath     = Replace(Replace(strFolderPath, "\", "/"), "%", "%25")
  strFolderPath     = Replace(Replace(Replace(Replace(Replace(Replace(strFolderPath, " ", "%20"), "#", "%23"), "$", "%24"), "?", "%3F"), "{", "%7B"), "}", "%7D")
  strFolderPath     = "file:///" & strFolderPath

  FormatFolderURI   = strFolderPath

End Function


Function FormatServer(strServer, strProtocol)
  Dim strServerWork

  strServerWork     = strServer
  If strProtocol <> "" Then
    strServerWork   = strProtocol & "://" & strServerWork & "." & GetBuildfileValue("UserDNSDomain")
  End If

  FormatServer      = strServerWork

End Function


Function Max(intA, intB)

  Select Case True
    Case CInt(intA) > CInt(intB)
      Max           = intA
    Case Else
      Max           = intB
  End Select

End Function


Sub ResetDBAFilePerm(strFolder)
  Call DebugLog("ResetDBAFilePerm: " & strFolder)

  Call ResetFilePerm(strFolder, strGroupDBA)

  If strGroupDBANonSA <> "" Then
    Call ResetFilePerm(strFolder, strGroupDBANonSA)
  End If

  If strIsInstallDBA = "1" Then
    Call ResetFilePerm(strFolder, strUserAccount)
  End If

End Sub


Sub ResetFilePerm(strFolder, strAccount)
  Call DebugLog("ResetFilePerm: " & strAccount)

  strPath           = strFolder
  If Right(strPath, 1) = "\" Then
    strPath         = Left(strPath, Len(strPath) - 1)
  End If

  Select Case True
    Case strAccount = strGroupDBA
      strCmd        = """" & strPath & """ /T /C /E /G """ & FormatAccount(strGroupDBA) & """:F"
      Call RunCacls(strCmd)
    Case strAccount = strGroupDBANonSA
      strCmd        = """" & strPath & """ /T /C /E /G """ & FormatAccount(strGroupDBANonSA) & """:R"
      Call RunCacls(strCmd)
    Case Else 
      strCmd        = """" & strPath & """ /T /C /E /G """ & FormatAccount(strAccount) & """:F"
      Call RunCacls(strCmd)
  End Select

End Sub


Sub RunCacls(strCmd)
  Call DebugLog("RunCacls: " & strCmd)
  Dim arrCmd
  Dim intUBound, intIdx, intIdx2
  Dim strNTService

  arrCmd            = Split(strCmd)
  intUBound         = UBound(arrCmd)
  strNTService      = GetBuildfileValue("NTService")
  For intIdx = 0 To intUBound
    Select Case True
      Case Instr(arrCmd(intIdx), """:") = 0 
        ' Nothing
      Case Instr(arrCmd(intIdx), strNTService & "\") > 0 
        arrCmd(intIdx) = ""
      Case Else
        For intIdx2 = intIdx + 1 To intUBound
          Select Case True
            Case Instr(arrCmd(intIdx2), """:") = 0
              ' Nothing
            Case StrComp(Left(arrCmd(intIdx), Instr(arrCmd(intIdx), """:")), Left(arrCmd(intIdx2), Instr(arrCmd(intIdx2), """:")), vbTextCompare) = 0
              arrCmd(intIdx) = ""
          End Select
        Next      
    End Select
  Next  
  strCmd            = Join(arrCmd, " ")

  intIdx2           = 0
  For intIdx = 0 To intUBound
    Select Case True
      Case Instr(arrCmd(intIdx), """:") = 0 
        ' Nothing
      Case Else
        intIdx2     = 1
    End Select
  Next
  If intIdx2 = 0 Then
    Exit Sub
  End If

  Call Util_RunExec(GetBuildfileValue("ProgCacls") & " " & strCmd, "", strResponseYes, -1)
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case intErrSave = 2
      ' Nothing
    Case intErrSave = 13
      ' Nothing
    Case intErrSave = 67     ' Network Name not found
      ' Nothing
    Case intErrSave = 1240   ' Not Authorized - Cannot put permission on remote share root
      ' Nothing
    Case intErrSave = 1332   ' Problem with security descriptor
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
  End Select
  Wscript.Sleep GetBuildfileValue("WaitShort") ' Allow time for CACLS processing to complete

End Sub


Sub SetupFolder(strFolder)
  Call DebugLog("SetupFolder: " & strFolder)
  Dim strPathParent

  strPath           = strFolder
  If Right(strPath, 1) = "\" Then
    strPath         = Left(strPath, Len(strPath) - 1)
  End If
  strPathParent     = Left(strPath, InstrRev(strPath, "\") - 1)

  Select Case True
    Case objFSO.FolderExists(strPath)
      ' Nothing
    Case Not objFSO.FolderExists(strPathParent)
      Call SetupFolder(strPathParent)
      objFSO.CreateFolder(strPath)
    Case Else
      objFSO.CreateFolder(strPath)
      Wscript.Sleep GetBuildfileValue("WaitShort")
  End Select

End Sub


Sub SetDCOMSecurity(strAppId)
  Call DebugLog("SetDCOMSecurity: " & strAppId)
  Dim arrPermDCom
  Dim objHelper, objPermDCom
  Dim strPermDCom, strSDDLDCom

  strSDDLDCom       = "(A;;CCDCLCSWRP;;;" & strSIDDistComUsers & ")"
  strPath           = "winmgmts:{impersonationLevel=impersonate}!\\" & strServer & "\ROOT\cimv2:Win32_securityDescriptorHelper"
  Set objHelper     = GetObject(strPath)
  objWMIReg.GetBinaryValue strHKCR,strAppId,"LaunchPermission",arrPermDCom
  Call objHelper.BinarySDToSDDL(arrPermDCom, strPermDCom)

  intIdx            = Instr(strPermDCom, strSIDDistComUsers)
  If intIdx = 0 Then
    intIdx          = Instr(strPermDCom, "(A;;CCSW;;;BU)")
    If intIdx = 0 Then
      strPermDCom   = strPermDCom & strSDDLDCom 
    Else
      strPermDCom   = Left(strPermDCom, intIdx - 1) & strSDDLDCom & Mid(strPermDCom, intIdx)
    End If
    Call DebugLog("Update DCom security with " & strPermDCom)
    Call objHelper.SDDLToWin32SD(strPermDCom, objPermDCom)
    Call objHelper.Win32SDToBinarySD(objPermDCom, arrPermDCom)
    objWMIReg.SetBinaryValue strHKCR,strAppId,"LaunchPermission",arrPermDCom
  End If

  Set objHelper     = Nothing

End Sub


Sub SetParam(strParamName, strParam, strNewValue, strMessage, ByRef strList)
  Call DebugLog("SetParam: " & strParamName & " to " & strNewValue)
  Dim strBuildValue

  strBuildValue     = GetBuildfileValue(strParamName)
  Select Case True
    Case strParam = strNewValue
      ' Nothing
    Case strParam = "N/A"
      ' Nothing
    Case (strParam = "NO") And (strNewValue = "N/A")
      strParam      = strNewValue
    Case (strParam = "") And (strNewValue = "YES") And (strMessage = "")
      strParam      = strNewValue
      strList       = strList & " " & strParamName
    Case strParam = ""
      strParam      = strNewValue
    Case strBuildValue = strNewValue
      strParam      = strNewValue
    Case strMessage = ""
      strParam      = strNewValue
      strList       = strList & " " & strParamName
    Case Else
      strParam      = strNewValue
      Call SetBuildMessage(strMsgInfo, Left(strParamName & Space(24), Max(Len(strParamName), 24)) & " set to " & strNewValue & ": " & strMessage)
  End Select

End Sub


Sub SetRegPerm(strRegParm, strName, strAccess)
  Call DebugLog("SetRegPerm: " & strRegParm & " for " & strName)
  ' Code based on example posted by ROHAM on www.tek-tips.com/viewthread.cfm?qid=1456390
  Dim objACE, objDACL, objSD, objSDUtil, objWMIReg
  Dim strACEAccessAllow, strACEFullControl, strACEPropogate, strACERead, strPath, strRegKey, strSDFormatIID, strSDPathRegistry, strStatusKB933789, strTrusteeName

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  strPath           = "SOFTWARE\Microsoft\Updates\Windows Server 2003\SP3\KB933789\"
  objWMIReg.GetStringValue strHKLM,strPath,"Description",strStatusKB933789
  Select Case True
    Case GetBuildfileValue("OSVersion") >= "6.0"
      ' Nothing
    Case Instr(Ucase(GetBuildfileValue("OSName")), " XP") > 0
      ' Nothing
    Case strStatusKB933789 > ""
      ' Nothing
    Case Else
      Call DebugLog("SetRegPerm bypassed")
      Exit Sub
  End Select

  strACEAccessAllow = 0
  strACEFullControl = &h10000000
  strACEPropogate   = &h2
  strACERead        = &h80000000
  strSDFormatIID    = 1
  strSDPathRegistry = 3
  strRegKey         = strRegParm
  If Right(strRegKey, 1) = "\" Then
    strRegKey       = Left(strRegKey, Len(strRegKey) - 1)
  End If

  strTrusteeName    = FormatAccount(strName)
  Set objSDUtil     = CreateObject("ADsSecurityUtility")
  Set objSD         = objSDUtil.GetSecurityDescriptor(strRegKey, strSDPathRegistry, strSDFormatIID)
  objSD.Owner       = GetBuildfileValue("LocalAdmin")
  Set objDACL       = objSD.DiscretionaryAcl
  Set objACE        = CreateObject("AccessControlEntry")
  objACE.Trustee    = strTrusteeName
  Select Case True
    Case strAccess = "F"
      objACE.AccessMask = strACEFullControl
    Case Else
      objACE.AccessMask = strACERead
  End Select
  objACE.ACEType    = strACEAccessAllow
  objACE.ACEFlags   = strACEPropogate
  objDACL.AddAce objACE

  objSDUtil.SetSecurityDescriptor strRegKey, strSDPathRegistry, objSD, strSDFormatIID

  Set objACE        = Nothing
  Set objDACL       = Nothing
  Set objSD         = Nothing
  Set objSDUtil     = Nothing

End Sub


Sub SetUpdate(strOnOff)
  Call DebugLog("SetUpdate: messages " & strOnOff)
  On Error Resume Next

  Select Case True
    Case strOnOff = "ON"
      colPrcEnvVars("SEE_MASK_NOZONECHECKS") = 1    ' Prevent Security Warning message hanging quiet install
      err.Number    = objAutoUpdate.Pause()         ' Prevent Windows Update service triggering a reboot prompt
    Case Else
      colPrcEnvVars.Remove("SEE_MASK_NOZONECHECKS") ' Allow Security Warning messages
      err.Number    =  objAutoUpdate.Resume()       ' Resume normal Window Update Service prompts
  End Select

  Select Case True
    Case err.Number = 0
      ' No action
    Case err.Number < 0
      Call SetBuildMessage(strMsgWarning, "Error " & Cstr(err.Number) & " returned by Windows Update Service when setting service to " & strOnOff)
      err.Clear
    Case Else
      Call SetBuildMessage(strMsgError,   "Error " & Cstr(err.Number) & " returned by Windows Update configuration " & strOnOff)
  End Select

End Sub


Function GetXMLParm(objParm, strParm, strDefault)
  Dim strValue

  Select Case True
    Case IsObject(objParm) = 0
      strValue      = strDefault
    Case IsNull(objParm.documentElement.getAttribute(strParm))
      strValue      = strDefault
    Case Else
      strValue      = objParm.documentElement.getAttribute(strParm)
  End Select

  GetXMLParm       = strValue

End Function


Sub SetXMLParm(objParm, strParm, strValue)
  Dim objAttribute

  If IsObject(objParm) = 0 Then
    Set objParm     = CreateObject("MSXML2.DomDocument")
    Set objParm.documentElement = objParm.createElement("ROOT")
  End If

  Select Case True
    Case Not IsNull(objParm.documentElement.getAttribute(strParm))
      objParm.documentElement.setAttribute strParm, strValue
    Case Else
      Set objAttribute  = objParm.createAttribute(strParm)
      objAttribute.Text = strValue
      objParm.documentElement.Attributes.setNamedItem objAttribute
  End Select

End Sub


Sub Util_RegWrite(strRegKey, strRegValue, strRegType)
  Call DebugLog("Util_RegWrite: " & strRegKey)

  err.Number        = objShell.RegWrite(strRegKey, strRegValue, strRegType)
  intErrSave        = err.Number
  strErrSave        = err.Description
  Select Case True
    Case intErrSave = 0 
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by Write " & strRegKey)
  End Select

End Sub


Sub Util_RunCmdAsync(strCmd, strOK)
  Call DebugLog("Util_RunCmdAsync: " & strCmd)
  On Error Resume Next

  err.Number        = objShell.Run(strCmd,7,False)
  intErrSave        = err.Number
  strErrSave        = err.Description
  On Error Goto 0
  Select Case True
    Case strOK      = -1
      Call DebugLog("Command ended with code: " & intErrSave)
    Case intErrSave = 0 
      ' Nothing
    Case Instr(" " & strOK & " ", " " & CStr(intErrSave) & " ") > 0
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
  End Select

  err.Clear

End Sub


Sub Util_ExecSQL(strCmd, strSQL, strOK)
  Call DebugLog("Util_ExecSQL: " & strSQL)
  Dim strSQLCmd

  strSQLCmd         = strCmd & " " & strSQL
  Call Util_RunExec(strSQLCmd, "EOF", strResponseYes, strOK)

End Sub


Sub Util_RunExec(strCmd, strMessage, strResponse, strOK)
  Call DebugLog("Util_RunExec: " & strCmd)
  Dim objCmd
  Dim intCount, intEOS
  Dim strBox1, strBox2, strProgCacls, strStdOut, strWaitShort

  On Error Resume Next
  err.Clear
  strBox1           = "[" & strResponseYes & "/" & strResponseNo & "]"
  strBox2           = "(" & strResponseYes & "/" & strResponseNo & ")?"
  strProgCacls      = GetBuildfileValue("ProgCacls")
  strWaitShort      = GetBuildfileValue("WaitShort")
  Set objCmd        = objShell.Exec(strCmd)
  Select Case True
    Case Not IsObject(objCmd)
      intErrSave    = 8
      strErrSave    = "Command not recognised"
    Case Else
      Select Case True
        Case strMessage = "EOF"
          objCmd.StdIn.Close
        Case Left(strCmd, Len(strCmdPS)) = strCmdPS
          objCmd.StdIn.Close
        Case Left(strCmd, Len(strProgCacls) + 1) = strProgCacls & " "
          objCmd.StdIn.Write strResponse & vbCrLf
          objCmd.StdIn.Close
      End Select
      intEOS        = objCmd.StdOut.AtEndOfStream
      While Not intEOS
        strStdOut   = objCmd.StdOut.ReadLine()
        intEOS      = objCmd.StdOut.AtEndOfStream
        Select Case True
          Case objCmd.Status <> 0
            intEOS  = True
          Case Right(strStdOut, Len(strBox1)) = strBox1
            objCmd.StdIn.Write strResponse & vbCrLf
          Case Right(strStdOut, Len(strBox2)) = strBox2
            objCmd.StdIn.Write strResponse & vbCrLf
          Case Left(strStdOut, Len(strAnyKey)) = strAnyKey
            objCmd.StdIn.Write strResponse & vbCrLf
          Case strMessage = ""
            ' Nothing
          Case Right(strStdOut, Len(strMessage)) = strMessage
            objCmd.StdIn.Write strResponse & vbCrLf
        End Select
      Wend
      objCmd.StdIn.Close
      intCount      = 0
      intEOS        = objCmd.Status
      While intEOS = 0
        Wscript.Sleep strWaitShort
        intCount    = intCount + 1
        intEOS      = objCmd.Status
        If intCount > 10 Then
          intEOS    = intCount
        End If
      WEnd
      intErrsave    = objCmd.ExitCode
      Select Case True
        Case Left(strCmd, Len(strProgCacls) + 1) = strProgCacls & " "
          strErrSave = err.Description
        Case Else
          strErrSave = objCmd.StdErr.ReadAll()
      End Select
  End Select

  On Error Goto 0
  Select Case True
    Case intErrSave = 0 
      ' Nothing
    Case Instr(" " & strOK & " ", " " & CStr(intErrSave) & " ") > 0
      ' Nothing
    Case strOK      = -1
      Call DebugLog("Command ended with code: " & intErrSave)
    Case intErrSave = 3010
      strReboot     = "Pending"
      Call SetBuildfileValue("RebootStatus", strReboot)
    Case intErrSave = -2067723326
      strReboot     = "Pending"
      Call SetBuildfileValue("RebootStatus", strReboot) 
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
  End Select

  err.Clear

End Sub


End Class


Function FormatAccount(strAccount)
  FormatAccount     = FBUtils.FormatAccount(strAccount)
End Function

Function FormatFolder(strFolder)
  FormatFolder      = FBUtils.FormatFolder(strFolder)
End Function

Function FormatFolderURI(strFolder)
  FormatFolderURI   = FBUtils.FormatFolderURI(strFolder)
End Function

Function FormatServer(strServer, strProtocol)
  FormatServer      = FBUtils.FormatServer(strServer, strProtocol)
End Function

Function Max(intA, intB)
  Max               = FBUtils.Max(intA, intB)
End Function

Sub ResetDBAFilePerm(strFolder)
  Call FBUtils.ResetDBAFilePerm(strFolder)
End Sub

Sub ResetFilePerm(strFolder, strAccount)
  Call FBUtils.ResetFilePerm(strFolder, strAccount)
End Sub

Sub RunCacls(strCmd)
  Call FBUtils.RunCacls(strCmd)
End Sub

Sub SetupFolder(strFolder)
  Call FBUtils.SetupFolder(strFolder)
End Sub

Sub SetDCOMSecurity(strAppId)
  Call FBUtils.SetDCOMSecurity(strAppId)
End Sub

Sub SetUpdate(strOnOff)
  Call FBUtils.SetUpdate(strOnOff)
End Sub

Sub SetRegPerm(strRegParm, strName, strAccess)
  Call FBUtils.SetRegPerm(strRegParm, strName, strAccess)
End Sub

Function GetXMLParm(objParm, strParm, strDefault)
  GetXMLParm        = FBUtils.GetXMLParm(objParm, strParm, strDefault)
End Function

Sub SetXMLParm(objParm, strParm, strValue)
  Call FBUtils.SetXMLParm(objParm, strParm, strValue)
End Sub

Sub SetParam(strParamName, strParam, strNewValue, strMessage, ByRef strList)
  Call FBUtils.SetParam(strParamName, strParam, strNewValue, strMessage, strList)
End Sub

Sub Util_RegWrite(strRegKey, strRegValue, strRegType)
  Call FBUtils.Util_RegWrite(strRegKey, strRegValue, strRegType)
End Sub

Sub Util_RunCmdAsync(strCmd, strOK)
  Call FBUtils.Util_RunCmdAsync(strCmd, strOK)
End Sub

Sub Util_ExecSQL(strCmd, strSQL, strOK)
  Call FBUtils.Util_ExecSQL(strCmd, strSQL, strOK)
End Sub

Sub Util_RunExec(strCmd, strMessage, strResponse, strOK)
  Call FBUtils.Util_RunExec(strCmd, strMessage, strResponse, strOK)
End Sub