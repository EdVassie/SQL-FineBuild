'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageBuildFile.vbs  
'  Copyright FineBuild Team © 2017 - 2019.  Distributed under Ms-Pl License
'
'  Purpose:      Manage the FineBuild Buildfile 
'
'  Author:       Ed Vassie
'
'  Date:         05 Jul 2017
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     05 Jul 2017  Initial version
'  1.1      Ed Vassie     12 Nov 2019  Added Statefile processing
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageBuildFile: Set FBManageBuildFile = New FBManageBuildFileClass
Dim objBuildfile
Dim strMsgError, strMsgErrorConfig, strMsgWarning, strMsgIgnore, strMsgInfo

Class FBManageBuildFileClass
Dim colBuildfile, colMessage, colStatefile
Dim objAttribute, objMessages, objShell, objStatefile
Dim intBuildMsg, intFound
Dim strBuildfile, strPathFBStart, strMessageOut, strMessagePrefix, strMessageRead, strProcessId, strStatefile, strValue


Private Sub Class_Initialize
' Perform Initialisation processing

  Set objBuildfile  = CreateObject("Microsoft.XMLDOM") 
  Set objStatefile  = CreateObject("Microsoft.XMLDOM") 
  Set objShell      = CreateObject("Wscript.Shell")

  strBuildfile      = objShell.ExpandEnvironmentStrings("%SQLLOGTXT%")
  If strBuildfile = "%SQLLOGTXT%" Then
    Exit Sub
  End If

  strBuildfile        = Mid(strBuildfile, 2, Len(strBuildfile) - 6) & ".xml"
  objBuildfile.async  = False
  objBuildfile.load(strBuildFile)
  Set colBuildfile    = objBuildfile.documentElement.selectSingleNode("BuildFile")

End Sub


Function GetBuildfileValue(strParam) 
' Get value from Buildfile

  Select Case True
    Case strParam = ""
      strValue      = ""
    Case Not IsObject(colBuildfile)
      strValue      = ""
    Case IsNull(colBuildfile.getAttribute(strParam))
      strValue      = ""
    Case Else
      strValue      = colBuildfile.getAttribute(strParam)
  End Select

  GetBuildfileValue = strValue

End Function


Sub SetBuildfileValue(strName, strValue)
  Call DebugLog("Set Buildfile value " & strName & ": " & strValue)
  ' Code based on http://www.vbforums.com/showthread.php?t=480935

  If IsNull(strValue) Then
    strValue        = ""
  End If

  Select Case True
    Case Not IsObject(colBuildfile)
      ' Nothing
    Case IsNull(colBuildfile.getAttribute(strName))
      colBuildfile.setAttribute strName, strValue
    Case Else
      Set objAttribute  = objBuildFile.createAttribute(strName)
      objAttribute.Text = strValue
      colBuildFile.Attributes.setNamedItem objAttribute
      objBuildFile.documentElement.appendChild colBuildfile
  End Select

  If IsObject(colBuildfile) Then
    objBuildFile.save strBuildFile
  End If

End Sub


Sub LinkBuildfile(strLogFile)
'  "LinkBuildfile:"

  If strLogFile = "" Then
    strLogFile      = objShell.ExpandEnvironmentStrings("%SQLLOGTXT%")
  End If

  strBuildFile      = Mid(strLogFile, 2, Len(strLogFile) - 6) & ".xml"
  objBuildfile.load(strBuildFile)
  Set colBuildFile  = objBuildfile.documentElement.selectSingleNode("BuildFile")

End Sub


Sub SetBuildMessage(strType, strMessage)
  ' Code based on http://www.vbforums.com/showthread.php?t=480935

  strProcessId      = GetBuildfileValue("ProcessId")

  Select Case True
    Case strMessage = ""
      Exit Sub
    Case strType = ""
      strMessagePrefix = ""
    Case strType = strMsgInfo
      strMessagePrefix = ""
    Case strProcessId > "1"
      strMessagePrefix = "(" & strProcessId & ") "
    Case Else
      strMessagePrefix = ""
  End Select
  Select Case True
    Case strType = strMsgErrorConfig
      strMessageOut = strMsgError & ": " & HidePasswords(strMessage)
      Call SetBuildfileValue("ErrorConfig", "YES")
    Case Else
      strMessageOut = strType     & ": " & strMessagePrefix & HidePasswords(strMessage)
  End Select

  Set colMessage    = objBuildfile.documentElement.selectSingleNode("Message")
  Set objMessages   = colMessage.attributes
  intBuildMsg       = 0
  intFound          = 0
  While intBuildMsg  < objMessages.length
    intBuildMsg     = intBuildMsg + 1
    strMessageRead  = colMessage.getAttribute("Msg" & CStr(intBuildMsg))
    If strMessageRead = strMessageOut Then
      intFound      = 1
    End If
  WEnd

  intBuildMsg       = GetBuildfileValue("BuildMsg")
  If intBuildMsg = "" Then
    intBuildMsg     = 0
  End If
  intBuildMsg       = intBuildMsg + 1

  If intFound = 0 Then  
    Set objAttribute  = objBuildFile.createAttribute("Msg" & CStr(intBuildMsg))
    objAttribute.Text = strMessageOut
    colMessage.Attributes.setNamedItem objAttribute
    objBuildFile.documentElement.appendChild colMessage
    objBuildFile.save strBuildFile
    Call SetBuildfileValue("BuildMsg", intBuildMsg)
  End If

  Select Case True
    Case strType = strMsgError 
      Call FBLog(" ")
      Call FBLog(" " & strMessageOut)
      err.Raise 8, "", strMessageOut
    Case Else
      Call FBLog(" " & strMessageOut)
  End Select

End Sub


Function GetStatefileValue(strParam) 
' Get value from Statefile

  If Not IsObject(colStatefile) Then
    strStatefile        = GetBuildfileValue("Statefile")
    objStatefile.async  = False
    objStatefile.load(strStatefile)
    Set colStatefile    = objStatefile.documentElement.selectSingleNode("FineBuildState")    
  End If

  Select Case True
    Case strParam = ""
      strValue      = ""
    Case IsNull(colStatefile.getAttribute(strParam))
      strValue      = ""
    Case Else
      strValue      = colStatefile.getAttribute(strParam)
  End Select

  GetStatefileValue = strValue

End Function


Sub SetStatefileValue(strName, strValue)
  Call DebugLog("Set Statefile value " & strName & ": " & strValue)
  ' Code based on http://www.vbforums.com/showthread.php?t=480935

  If Not IsObject(colStatefile) Then
    strStatefile        = GetBuildfileValue("Statefile")
    objStatefile.async  = False
    objStatefile.load(strStatefile)
    Set colStatefile    = objStatefile.documentElement.selectSingleNode("FineBuildState")  
  End If

  If IsNull(strValue) Then
    strValue        = ""
  End If

  Select Case True
    Case IsNull(colStatefile.getAttribute(strName))
      colStatefile.setAttribute strName, strValue
    Case Else
      Set objAttribute  = objStatefile.createAttribute(strName)
      objAttribute.Text = strValue
      colStatefile.Attributes.setNamedItem objAttribute
      objStatefile.documentElement.appendChild colStatefile
  End Select

  objStatefile.save strStatefile

End Sub


End Class

Function GetBuildfileValue(strParam)
  GetBuildfileValue = FBManageBuildFile.GetBuildfileValue(strParam)
End Function

Sub SetBuildfileValue(strName, strValue)
  Call FBManageBuildFile.SetBuildfileValue(strName, strValue)
End Sub

Sub LinkBuildfile(strLogFile)
  Call FBManageBuildFile.LinkBuildfile(strLogFile)
End Sub

Sub SetBuildMessage(strType, strMessage)
  Call FBManageBuildFile.SetBuildMessage(strType, strMessage)
End Sub

Function GetStatefileValue(strParam)
  GetStatefileValue = FBManageBuildfile.GetStatefileValue(strParam)
End Function

Sub SetStatefileValue(strName, strValue)
  Call FBManageBuildfile.SetStatefileValue(strName, strValue)
End Sub