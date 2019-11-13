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
'  1.1      Ed Vassie     12 Nov 2019  Added Statusfile processing
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageBuildFile: Set FBManageBuildFile = New FBManageBuildFileClass
Dim objBuildFile, objStatusfile

Class FBManageBuildFileClass
Dim colBuildfile, colMessage, colStatusfile
Dim objAttribute, objMessages
Dim intBuildMsg, intFound
Dim strBuildfile, strMessageOut, strMessagePrefix, strMessageRead, strProcessId, strStatusfile, strValue


Private Sub Class_Initialize
' Perform Initialisation processing
  Dim objShell

  Set objBuildFile  = CreateObject("Microsoft.XMLDOM") 
  Set objStatusFile = CreateObject("Microsoft.XMLDOM") 
  Set objShell      = CreateObject("Wscript.Shell")

  strBuildFile      = objShell.ExpandEnvironmentStrings("%SQLLOGTXT%")
  If strBuildFile = "%SQLLOGTXT%" Then
    Exit Sub
  End If

  strBuildFile        = Mid(strBuildFile, 2, Len(strBuildFile) - 6) & ".xml"
  objBuildFile.async  = False
  objBuildfile.load(strBuildFile)
  Set colBuildFile    = objBuildfile.documentElement.selectSingleNode("BuildFile")

'  strStatusfile       = Mid(strStatusfile, 2, Len(strStatusfile) - 6) & ".xml"
'  objStatusfile.async = False
'  objStatusfile.load(strStatusfile)
'  Set colStatusfile   = objStatusfile.documentElement.selectSingleNode("Statusfile")

End Sub


Function GetBuildfileValue(strParam) 
' Get value from Buildfile

  Select Case True
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
    Case IsNull(colBuildfile.getAttribute(strName))
      colBuildfile.setAttribute strName, strValue
    Case Else
      Set objAttribute  = objBuildFile.createAttribute(strName)
      objAttribute.Text = strValue
      colBuildFile.Attributes.setNamedItem objAttribute
      objBuildFile.documentElement.appendChild colBuildfile
  End Select

  objBuildFile.save strBuildFile

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
  strMessageOut     = strMessagePrefix & HidePasswords(strMessage)

  Set colMessage    = objBuildfile.documentElement.selectSingleNode("Message")
  Set objMessages   = colMessage.attributes
  intBuildMsg       = 0
  intFound          = 0
  While intBuildMsg  < objMessages.length
    intBuildMsg     = intBuildMsg + 1
    strMessageRead  = colMessage.getAttribute("Msg" & CStr(intBuildMsg))
    If strMessageRead = strType & ": " & strMessageOut Then
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
    objAttribute.Text = strType & ": " & strMessageOut
    colMessage.Attributes.setNamedItem objAttribute
    objBuildFile.documentElement.appendChild colMessage
    objBuildFile.save strBuildFile
    Call SetBuildfileValue("BuildMsg", intBuildMsg)
  End If

  Select Case True
    Case strType = strMsgError 
      Call FBLog(" ")
      Call FBLog(" " & strType & ": " & strMessageOut)
      err.Raise 8, "", strType & ": " & strMessageOut
    Case strType = strMsgWarning 
      Call FBLog(" ")
      Call FBLog(" " & strType & ": " & strMessageOut)
    Case Else
      Call FBLog(" " & strMessageOut)
  End Select

End Sub


Function GetStatusfileValue(strParam) 
' Get value from Statusfile

  Select Case True
    Case IsNull(colStatusfile.getAttribute(strParam))
      strValue      = ""
    Case Else
      strValue      = colStatusfile.getAttribute(strParam)
  End Select

  GetStatusfileValue = strValue

End Function


Sub SetStatusfileValue(strName, strValue)
  Call DebugLog("Set Statusfile value " & strName & ": " & strValue)
  ' Code based on http://www.vbforums.com/showthread.php?t=480935

  If IsNull(strValue) Then
    strValue        = ""
  End If

  Select Case True
    Case IsNull(colStatusfile.getAttribute(strName))
      colStatusfile.setAttribute strName, strValue
    Case Else
      Set objAttribute  = objStatusfile.createAttribute(strName)
      objAttribute.Text = strValue
      colStatusfile.Attributes.setNamedItem objAttribute
      objStatusfile.documentElement.appendChild colStatusfile
  End Select

  objStatusfile.save strStatusFile

End Sub


End Class

Function GetBuildfileValue(strParam)
  GetBuildfileValue = FBManageBuildFile.GetBuildfileValue(strParam)
End Function

Sub SetBuildfileValue(strName, strValue)
  Call FBManageBuildFile.SetBuildfileValue(strName, strValue)
End Sub

Sub SetBuildMessage(strType, strMessage)
  Call FBManageBuildFile.SetBuildMessage(strType, strMessage)
End Sub

Function GetStatusfileValue(strParam)
  GetStatusfileValue = FBManageBuildfile.GetStatusfileValue(strParam)
End Function

Sub SetStatusfileValue(strName, strValue)
  Call FBManageBuildfile.SetStatusfileValue(strName, strValue)
End Sub