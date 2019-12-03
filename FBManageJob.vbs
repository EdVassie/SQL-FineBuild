'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageJob.vbs  
'  Copyright FineBuild Team © 2017 - 2019.  Distributed under Ms-Pl License
'
'  Purpose:      Manage Creation of SQL Agent and Windows Scheduler Jobs
'
'  Author:       Ed Vassie
'
'  Date:         18 Sep 2017
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     18 Sep 2017  Initial version

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageJob: Set FBManageJob = New FBManageJobClass


Class FBManageJobClass
Dim objFSO, objJobSQL, objJobSQLData, objShell
Dim strCmd, strCmdSQL, strDBA_DB, strDirBackup, strEdition, strHKLMSQL, strInstance, strInstReg, strInstSQL
Dim strJobCategory, strJobCmd, strJobSaAccount, strJobId, strOSVersion, strPath, strPathMaint, strReport
Dim strServInst, strSetupSQLAgent, strSqlAccount, strSqlPassword, strSQLOperator

Private Sub Class_Initialize
  Call DebugLog("FBManageJob Class_Initialize:")

  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objJobSQL     = CreateObject("ADODB.Connection")
  Set objShell      = WScript.CreateObject ("Wscript.Shell")

  strHKLMSQL        = GetBuildfileValue("HKLMSQL")
  strCmdSQL         = GetBuildfileValue("CmdSQL")
  strDBA_DB         = GetBuildfileValue("DBA_DB")
  strEdition        = GetBuildfileValue("AuditEdition")
  strInstance       = GetBuildfileValue("Instance")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strJobCategory    = GetBuildfileValue("JobCategory")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strSetupSQLAgent  = GetBuildfileValue("SetupSQLAgent")
  strServInst       = GetBuildfileValue("ServInst")
  strSqlAccount     = GetBuildfileValue("SqlAccount")
  strSqlPassword    = GetBuildfileValue("SqlPassword")
  strSQLOperator    = GetBuildfileValue("SQLOperator")

  Call DebugLog("Get Backup folder name")
  strPath           = strHKLMSQL & "Instance Names\SQL\" & strInstance
  strInstReg        = objShell.RegRead(strPath)
  strPath           = strHKLMSQL & strInstReg & "\MSSQLServer\BackupDirectory"
  strDirBackup      = objShell.RegRead(strPath)
  Select Case True
    Case Right(strDirBackup, 11) = "AdHocBackup"
      strDirBackup  = Left(strDirBackup, Len(strDirBackup) - 11)
    Case Right(strDirBackup, 1) <> "\"
      strDirBackup  = strDirBackup & "\"
  End Select
  strPathMaint      = strDirBackup & strJobCategory
  If strSetupSQLAgent <> "YES" Then
    Call SetupFolder(strPathMaint)
  End If

  objJobSQL.Provider = "SQLOLEDB"
  objJobSQL.ConnectionString = "Driver={SQL Server};Server=" & GetBuildfileValue("ServInst") & ";Database=master;Trusted_Connection=Yes;"
  Call DebugLog("Connection string: " & objJobSQL.ConnectionString)
  objJobSQL.Open 

  strCmd            = "SELECT name FROM master.dbo.syslogins WHERE sid=1"
  Set objJobSQLData = objJobSQL.Execute(strCmd)
  Do Until objJobSQLData.EOF
    strJobSaAccount = objJobSQLData.Fields("name")
    objJobSQLData.MoveNext
  Loop

End Sub


Sub SetupCommonJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
  Call DebugLog("Setup Job: " & strJobName)

  Select Case True
    Case Instr("EXPRESS", GetBuildfileValue("AuditEdition")) > 0
      Call SetupWindowsJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
    Case Else
      Call SetupSQLJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
  End Select

End Sub


Sub SetupSQLJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
  Call DebugLog("Setup SQL Job: " & strJobName)

  If strJobCategory <> "" Then
    Call SetupJobCategory(strJobCategory)
  End If

  If strSQLOperator <> "" Then
    Call SetupOperator(strSQLOperator)
  End If

  Call DebugLog("Delete existing job")
  strJobId          = ""
  strCmd            = "SELECT j.job_id FROM msdb.dbo.sysjobs j JOIN msdb.dbo.sysjobservers s ON s.job_id = j.job_id AND s.server_id <> 0 WHERE j.name = '" & strJobName & "'"
  Set objJobSQLData    = objJobSQL.Execute(strCmd)
  Do Until objJobSQLData.EOF
    strJobId        = objJobSQLData.Fields("job_id")
    objJobSQLData.MoveNext
  Loop
  If strJobId = "" Then
    strCmd          = strCmdSQL & " -Q"
    strJobCmd       = """EXECUTE msdb.dbo.sp_delete_job @job_name = '" & strJobName & "'"""
    Call Util_ExecSQL(strCmd, strJobCmd, -1)
  End If

  Call DebugLog("Create new job")
  strCmd            = strCmdSQL & " -Q"
  strJobCmd         = """EXECUTE msdb.dbo.sp_add_job @job_name = N'" & strJobName & "', @owner_login_name = N'" & strJobSaAccount & "', @description = N'" & strJobName & "', @category_name = N'" & strJobCategory & "', @enabled = 1, @notify_level_email = 2, @notify_level_page = 0, @notify_level_netsend = 0, @notify_level_eventlog = 3, @delete_level= 0, @notify_email_operator_name = N'" & strSQLOperator & "'"""
  Call Util_ExecSQL(strCmd, strJobCmd, 0)

  Call DebugLog("Select job_id")
  strJobId          = ""
  strCmd            = "SELECT j.job_id FROM msdb.dbo.sysjobs j WHERE j.name = '" & strJobName & "'"
  Set objJobSQLData    = objJobSQL.Execute(strCmd)
  Do Until objJobSQLData.EOF
    strJobId        = Replace(Replace(objJobSQLData.Fields("job_id"), "{", "'"), "}", "'")
    objJobSQLData.MoveNext
  Loop
  Wscript.Sleep 500 ' Wait 1/2 second

  Call DebugLog("Add job step")
  Select Case True
    Case strJobType = "SQL"
      strJobCmd = strCmdSQL & " -d """ & strDBA_DB & """ -Q""" & strStepCmd & """"
    Case Else
      strJobCmd = "" & strStepCmd & ""
  End Select
  strJobCmd     = Replace(Replace(strJobCmd, """", """"""), "'", "''")
  strReport         = GetReportFile(strReportOpt, strJobName, "SQL")
  strCmd            = strCmdSQL & " -x -Q"
  strJobCmd         = """EXECUTE msdb.dbo.sp_add_jobstep @job_id = " & strJobId & ", @step_id = 1, @step_name = N'" & strJobName & "', @command = N'" & strJobCmd & "', @subsystem = N'CMDEXEC', @cmdexec_success_code = 0, @flags = 0, @retry_attempts = 0, @retry_interval = 1, @output_file_name = N'" & strReport & "', @on_success_step_id = 0, @on_success_action = 1, @on_fail_step_id = 0, @on_fail_action = 2"""
  Call Util_ExecSQL(strCmd, strJobCmd, 0)
  strCmd            = strCmdSQL & " -Q"
  strJobCmd         = """EXECUTE msdb.dbo.sp_update_job @job_id = " & strJobId & ", @start_step_id = 1"""
  Call Util_ExecSQL(strCmd, strJobCmd, 0)

  If strFreq <> "" Then
    Call DebugLog("Add job schedule")
    strCmd          = strCmdSQL & " -Q"
    strJobCmd       = """EXECUTE msdb.dbo.sp_add_jobschedule @job_id = " & strJobId & ", @name = N'" & strJobName & "', @enabled = 1, @freq_type = " & GetSQLFreqType(strFreq) & ", @active_start_date = 20000101, @active_start_time = " & GetSQLStartTime(strTime) & ", @freq_interval = " & GetSQLFreqInterval(strFreq) & ", @freq_subday_type = " & GetSQLFreqSubDayType(strFreq) & ", @freq_subday_interval = " & GetSQLFreqSubDayInt(strFreq, strTime) & ", @freq_relative_interval = 0, @freq_recurrence_factor = " & GetSQLFreqRecurrence(strFreq) & ", @active_end_date = 99991231, @active_end_time = 235959"""
    Call Util_ExecSQL(strCmd, strJobCmd, 0)
  End If

  Call DebugLog("Add target servers")
  strCmd            = strCmdSQL & " -x -Q"
  strJobCmd         = """EXECUTE msdb.dbo.sp_add_jobserver @job_id = " & strJobId & ", @server_name = N'(local)' """
  Call Util_ExecSQL(strCmd, strJobCmd, 0)

End Sub


Private Function GetSQLFreqType(strFreq)

  Select Case True
    Case strFreq = "MINUTE"
      GetSQLFreqType = "4"
    Case strFreq = "HOUR"
      GetSQLFreqType = "4"
    Case strFreq = "ALL"
      GetSQLFreqType = "4"
    Case Else
      GetSQLFreqType = "8"
  End Select

End Function


Private Function GetSQLFreqInterval(strFreq)

  Select Case True
    Case strFreq = "MINUTE"
      GetSQLFreqInterval = "1"
    Case strFreq = "HOUR"
      GetSQLFreqInterval = "1"
    Case strFreq = "ALL"
      GetSQLFreqInterval = "1"
    Case Else
      GetSQLFreqInterval = 0
      If Instr(strFreq, "SUN") > 0 Then
        GetSQLFreqInterval = GetSQLFreqInterval Or 1
      End If
      If Instr(strFreq, "MON") > 0 Then
        GetSQLFreqInterval = GetSQLFreqInterval Or 2
      End If
      If Instr(strFreq, "TUE") > 0 Then
        GetSQLFreqInterval = GetSQLFreqInterval Or 4
      End If
      If Instr(strFreq, "WED") > 0 Then
        GetSQLFreqInterval = GetSQLFreqInterval Or 8
      End If
      If Instr(strFreq, "THU") > 0 Then
        GetSQLFreqInterval = GetSQLFreqInterval Or 16
      End If
      If Instr(strFreq, "FRI") > 0 Then
        GetSQLFreqInterval = GetSQLFreqInterval Or 32
      End If
      If Instr(strFreq, "SAT") > 0 Then
        GetSQLFreqInterval = GetSQLFreqInterval Or 64
      End If
  End Select

End Function


Private Function GetSQLFreqRecurrence(strFreq)

  Select Case True
    Case strFreq = "MINUTE"
      GetSQLFreqRecurrence = "0"
    Case strFreq = "HOUR"
      GetSQLFreqRecurrence = "0"
    Case strFreq = "ALL"
      GetSQLFreqRecurrence = "0"
    Case Else
      GetSQLFreqRecurrence = "1"
  End Select

End Function


Private Function GetSQLFreqSubDayType(strFreq)

  Select Case True
    Case strFreq = "MINUTE"
      GetSQLFreqSubDayType = "4"
    Case strFreq = "HOUR"
      GetSQLFreqSubDayType = "8"
    Case Else
      GetSQLFreqSubDayType = "1"
  End Select

End Function


Private Function GetSQLFreqSubDayInt(strFreq, strTime)

  Select Case True
    Case strFreq = "MINUTE"
      GetSQLFreqSubDayInt = strTime
    Case strFreq = "HOUR"
      GetSQLFreqSubDayInt = "1"
    Case Else
      GetSQLFreqSubDayInt = "0"
  End Select

End Function


Private Function GetSQLStartTime(strTime)

  Select Case True
    Case Len(strTime) < 6
      GetSQLStartTime = "000000"
    Case Else
      GetSQLStartTime = Replace(strTime, ":", "")
  End Select

End Function


Sub SetupWindowsJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
  Call DebugLog("Setup Windows Job: " & strJobName)
  Dim objFile
  Dim strPathJob

  Call DebugLog("Delete existing job")
  strJobCmd         = "SCHTASKS /Delete /tn """ & GetWinJobFolder() & GetWinJobName(strJobName) & """ /F"
  Call Util_RunExec(strJobCmd, "", strResponseYes, -1)
  strPathJob        = strPathMaint & "\" & GetWinJobName(strJobname) & ".bat"
  If objFSO.FileExists(strPathJob) Then
    Call objFSO.DeleteFile(strPathJob, True)
  End If

  Call DebugLog("Create new job")
  Select Case True
    Case strJobType = "SQL"
      strJobCmd     = strCmdSQL & " -d """ & GetBuildfileValue("DBA_DB") & """ -Q """ & strStepCmd & """ "
    Case Else
      strJobCmd     = strStepCmd
  End Select
  strJobCmd         = strJobCmd & " -o """ & GetReportFile(strReportOpt, GetWinJobName(strJobName), "WIN") & """"
  Set objFile       = objFSO.CreateTextFile(strPathJob, True)
  objFile.WriteLine "SET RUNDATE=%DATE:/=%"
  objFile.WriteLine "SET RUNTIME=%TIME::=%"
  objFile.WriteLine "SET RUNTIME=%RUNTIME:~0,6%"
  objFile.WriteLine strJobCmd
  objFile.Close
  strJobCmd         = "SCHTASKS /Create /tn """ & GetWinJobFolder() & GetWinJobName(strJobName) & """ /tr """ & GetWinJobPath(strPathJob) & """ /sc """ & GetWinFreqType(strFreq) & """ " & GetWinFreq(strFreq) & " /st """ & GetWinStartTime(strTime) & """ /sd ""01/01/2000"" " & GetWinAccount(strSqlAccount) & " " & GetWinPassword(strSqlPassword)
  Call Util_RunExec(strJobCmd, "", strResponseYes, 0)

  Set objFile       = Nothing

End Sub


Private Function GetWinJobFolder()

  GetWinJobFolder   = strInstSQL & "\"

End Function


Private Function GetWinJobName(strJobName)

  GetWinJobName     = Replace(strJobName, ":", "")

End Function


Private Function GetWinJobPath(strPath)
  Dim strJobPath

  Select Case True
    Case strOSVersion < "6" ' Windows 2003 or below
      strJobPath    = "\""" & strPath & """\"
    Case Else
      strJobPath    = "'" & strPath & "'"
  End Select

  GetWinJobPath     = Replace(strJobPath, ":", "")

End Function


Private Function GetWinFreqType(strFreq)

  Select Case True
    Case strFreq = "MINUTE"
      GetWinFreqType = "HOURLY"
    Case strFreq = "HOUR"
      GetWinFreqType = "HOURLY"
    Case strFreq = "ALL"
      GetWinFreqType = "DAILY"
    Case Else
      GetWinFreqType = "WEEKLY"
  End Select

End Function


Private Function GetWinFreq(strFreq)

  Select Case True
    Case strFreq = "MINUTE"
      GetWinFreq = ""
    Case strFreq = "HOUR"
      GetWinFreq = ""
    Case strFreq = "ALL"
      GetWinFreq = ""
    Case Else
      GetWinFreq = "/d """ & strFreq & """"
  End Select

End Function


Private Function GetWinStartTime(strTime)

  Select Case True
    Case Len(strTime) < 6
      GetWinStartTime   = "00:00:00"
    Case Else
      GetWinStartTime   = strTime
  End Select

End Function


Private Function GetWinAccount(strAccount)

  Select Case True
    Case strAccount = ""
      GetWinAccount = ""
    Case strOSVersion < "6" And strSqlPassword = ""
      GetWinAccount = " /ru ""SYSTEM"" "
    Case Else
      GetWinAccount = " /ru """ & strAccount & """ "
  End Select

End Function


Private Function GetWinPassword(strPassword)

  Select Case True
    Case strPassword = ""
      GetWinPassword = ""
    Case Else
      GetWinPassword = " /rp """ & strPassword & """ "
  End Select

End Function


Private Function GetReportFile(strReportOpt, strJobName, strJobType)
' ReportOpt: (blank)= No report file, N=No report file, S=Single reuseable rport file,  Y=Datestamped report file
  Dim strDirBackup, strDirReport, strJobNameFile

  strJobNameFile    = Replace(strJobName, ":", "-")

  Select Case True
    Case strReportOpt = ""
      strDirReport  = ""
    Case strReportOpt = "N"
      strDirReport  = ""
    Case Else
      strDirBackup      = GetBuildfileValue("DirBackup")
      Select Case True
        Case Instr(strDirBackup, "AdHocBackup") > 0
          strDirBackup  = Left(strDirBackup, Instr(strDirBackup, "AdHocBackup") - 1)
        Case Right(strDirBackup, 1) <> "\"
          strDirBackup  = strDirBackup & "\"
      End Select
      strDirReport      = strDirBackup & "Reports"
  End Select

  Select Case True
    Case strDirReport = ""
      GetReportFile = ""
    Case (strJobType = "WIN") And (strReportOpt = "S")
      GetReportFile = strDirReport & "\" & Replace(strServInst, "\", "-") & "_" & strJobNameFile & ".txt"
    Case strJobType = "WIN"
      GetReportFile = strDirReport & "\" & Replace(strServInst, "\", "-") & "_" & strJobNameFile & "_%RUNDATE%_%RUNTIME%.txt"
    Case strReportOpt = "S"
      GetReportFile = strDirReport & "\" & Replace(strServInst, "\", "-") & "_" & strJobNameFile & ".txt"
    Case Else
      GetReportFile = strDirReport & "\" & Replace(strServInst, "\", "-") & "_" & strJobNameFile & "_$(ESCAPE_SQUOTE(STRTDT))_$(ESCAPE_SQUOTE(STRTTM)).txt"
  End Select

End Function


Sub RunCommonJob(strJobName)
  Call DebugLog("Run Job: " & strJobName)

  Select Case True
    Case Instr("EXPRESS", strEdition) > 0
      Call RunWindowsJob(strJobName)
    Case Else
      Call RunSQLJob(strJobName)
  End Select

End Sub


Sub RunSQLJob(strJobName)

  strCmd            = strCmdSQL & " -Q"
  strJobCmd         = """EXECUTE msdb.dbo.sp_start_job @job_name = N'" & strJobName & "' """
  Call Util_ExecSQL(strCmd, strJobCmd, 0)

End Sub


Sub RunWindowsJob(strJobName)

  strJobCmd         = "SCHTASKS /Run /tn """ & GetWinJobFolder() & GetWinJobName(strJobName) & """ "
  Call Util_RunExec(strJobCmd, "", strResponseYes, 0)

End Sub


Private Sub SetupJobCategory(strJobCategory)
  Call DebugLog("Setup Job Category: " & strJobCategory)

  strCmd            = strCmdSQL & " -Q"
  strJobCmd         = """EXECUTE msdb.dbo.sp_add_category @name = '" & strJobCategory & "'"""
  Call Util_ExecSQL(strCmd, strJobCmd, -1)

End Sub


Private Sub SetupOperator(strOperator)
  Call DebugLog("Setup Operator: " & strOperator)

  strCmd            = strCmdSQL & " -Q"
  strJobCmd         = """EXEC msdb.dbo.sp_add_operator @name = N'" & strOperator & "', @enabled = 1, @email_address = N'" & GetBuildfileValue("SQLEmail") & "', @category_name = N'[Uncategorized]', @weekday_pager_start_time = 80000, @weekday_pager_end_time = 180000, @saturday_pager_start_time = 80000, @saturday_pager_end_time = 180000, @sunday_pager_start_time = 80000, @sunday_pager_end_time = 180000, @pager_days = 62;"""
  Call Util_ExecSQL(strCmd, strJobCmd, -1)

End Sub


End Class


Sub SetupCommonJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
  Call FBManageJob.SetupCommonJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
End Sub

Sub RunCommonJob(strJobName)
  Call FBManageJob.RunCommonJob(strJobName)
End Sub

Sub SetupSQLJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
  Call FBManageJob.SetupSQLJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
End Sub

Sub RunSQLJob(strJobName)
  Call FBManageJob.RunSQLJob(strJobName)
End Sub

Sub SetupWindowsJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
  Call FBManageJob.SetupWindowsJob(strJobName, strJobCategory, strReportOpt, strJobType, strStepCmd, strFreq, strTime)
End Sub

Sub RunWindowsJob(strJobName)
  Call FBManageJob.RunWindowsJob(strJobName)
End Sub