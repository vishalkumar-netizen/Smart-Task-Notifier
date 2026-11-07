' ==============================================================================
' Smart Task Notifier (100% Native Windows VBScript)
' ==============================================================================
' - Zero dependencies: Runs natively on Windows 7, 10, and 11 via wscript.exe.
' - Double-click anywhere to test immediately, or place directly inside:
'   C:\Users<username>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup
' - Survives PC Sleep, Hibernation, and full Restarts by saving timestamps
'   to %LOCALAPPDATA%\Smart_Task_Notifier.
' ==============================================================================

Option Explicit

Dim fso, shell, appDir, configPath, statePath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Create persistent storage folder in %LOCALAPPDATA%
appDir = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Smart_Task_Notifier"
If Not fso.FolderExists(appDir) Then
fso.CreateFolder(appDir)
End If

configPath = appDir & "\config.ini"
statePath = appDir & "\state.ini"

' Configuration Variables
Dim timerHours, timerMinutes, redirectUrl, startTitle, endTitle, endSubtitle

' State Variables
Dim activeTask, startTimeStr, targetTimeStr

' Load settings from disk (or initialize defaults)
LoadConfig

' Check if a task timer is already in progress (handles PC Restarts & Sleep recovery)
LoadState

If activeTask And IsDate(targetTimeStr) And IsDate(startTimeStr) Then
Dim startDt, targetDt, nowDt
startDt = CDate(startTimeStr)
targetDt = CDate(targetTimeStr)
nowDt = Now()

' Safety check: If saved session is older than 18 hours (from yesterday), reset it
If DateDiff("h", startDt, nowDt) > 18 Then
    ClearState
    ShowStartPrompt
ElseIf nowDt >= targetDt Then
    ' Timer finished while PC was asleep or restarting
    ShowEndPrompt
Else
    ' Resume quiet background tracking right where it left off
    TrackBackgroundTask targetDt
End If


Else
ShowStartPrompt
End If

' ==============================================================================
' 1. START PROMPT (3 OPTIONS: Start | Close Notifier | Settings)
' ==============================================================================
Sub ShowStartPrompt()
Dim choice, promptText, estEndTime
estEndTime = FormatDateTime(DateAdd("n", (timerHours * 60) + timerMinutes, Now()), vbShortTime)

promptText = startTitle & vbCrLf & vbCrLf & _
             "Timer Duration: " & timerHours & " hr " & timerMinutes & " min" & vbCrLf & _
             "Estimated Completion Alert: " & estEndTime & vbCrLf & _
             "Redirect URL: " & redirectUrl & vbCrLf & vbCrLf & _
             "Choose an option:" & vbCrLf & _
             "  [ YES ]       ->  Start Task (Open URL & Start Timer)" & vbCrLf & _
             "  [ NO ]        ->  Close Notifier (Exit for now)" & vbCrLf & _
             "  [ CANCEL ] ->  Settings (Customize Timer, URL & Messages)"

' vbYesNoCancel (3) + vbQuestion (32) + vbSystemModal (4096)
choice = MsgBox(promptText, 3 + 32 + 4096, "Smart Task Notifier - Start")

If choice = vbYes Then
    ' Option 1: Start Task
    OpenRedirectUrl
    Dim nowDt, targetDt
    nowDt = Now()
    targetDt = DateAdd("n", (timerHours * 60) + timerMinutes, nowDt)
    SaveState True, CStr(nowDt), CStr(targetDt)
    TrackBackgroundTask targetDt

ElseIf choice = vbNo Then
    ' Option 2: Close Notifier
    ClearState
    WScript.Quit 0

ElseIf choice = vbCancel Then
    ' Option 3: Settings
    OpenSettingsMenu False
End If


End Sub

' ==============================================================================
' 2. QUIET BACKGROUND TRACKING (IMMUNE TO SLEEP & HIBERNATION)
' ==============================================================================
Sub TrackBackgroundTask(targetDt)
' Loops silently in the background with 0% CPU usage, checking system clock every 10 seconds
Do While Now() < targetDt
WScript.Sleep 10000
Loop

' Timer duration reached -> Show Completion Popup
ShowEndPrompt


End Sub

' ==============================================================================
' 3. COMPLETION PROMPT
' ==============================================================================
Sub ShowEndPrompt()
Dim choice, promptText, startedDisplay
If IsDate(startTimeStr) Then
startedDisplay = FormatDateTime(CDate(startTimeStr), vbShortTime)
Else
startedDisplay = "Earlier today"
End If

promptText = endTitle & vbCrLf & _
             ">>> " & endSubtitle & " <<<" & vbCrLf & vbCrLf & _
             "Started At: " & startedDisplay & vbCrLf & _
             "Current Time: " & FormatDateTime(Now(), vbShortTime) & vbCrLf & vbCrLf & _
             "Choose an option:" & vbCrLf & _
             "  [ YES ]       ->  Complete & Open URL (Resets timer)" & vbCrLf & _
             "  [ NO ]        ->  Close Notifier (Exit & Reset)" & vbCrLf & _
             "  [ CANCEL ] ->  Settings (Customize Settings)"

' vbYesNoCancel (3) + vbExclamation (48) + vbSystemModal (4096)
choice = MsgBox(promptText, 3 + 48 + 4096, "Smart Task Notifier - Time Up")

If choice = vbYes Then
    ' Complete & Open URL
    OpenRedirectUrl
    ClearState
    WScript.Quit 0

ElseIf choice = vbNo Then
    ' Close Notifier
    ClearState
    WScript.Quit 0

ElseIf choice = vbCancel Then
    ' Settings
    OpenSettingsMenu True
End If


End Sub

' ==============================================================================
' 4. SETTINGS MENU (CUSTOMIZE ALL OPTIONS & SAVE)
' ==============================================================================
Sub OpenSettingsMenu(isEndMode)
Dim settingChoice, menuText

Do
    menuText = "CURRENT SETTINGS:" & vbCrLf & _
               "--------------------------------------------------" & vbCrLf & _
               "1. Timer Hours:         " & timerHours & " hr" & vbCrLf & _
               "2. Timer Minutes:       " & timerMinutes & " min" & vbCrLf & _
               "3. Redirect URL:        " & redirectUrl & vbCrLf & _
               "4. Start Message:       " & startTitle & vbCrLf & _
               "5. Completion Title:    " & endTitle & vbCrLf & _
               "6. Completion Subtitle: " & endSubtitle & vbCrLf & _
               "7. Reset All to Default (9 hr 0 min)" & vbCrLf & _
               "0. Save & Go Back" & vbCrLf & _
               "--------------------------------------------------" & vbCrLf & _
               "Enter a number (0 to 7) to edit:"

    settingChoice = InputBox(menuText, "Smart Task Notifier - Settings", "0")

    ' If user clicked Cancel or entered 0, save and return
    If settingChoice = "" Or Trim(settingChoice) = "0" Then
        SaveConfig
        Exit Do
    End If

    Select Case Trim(settingChoice)
        Case "1"
            Dim newHrs
            newHrs = InputBox("Enter Timer Hours (0 - 23):" & vbCrLf & "(Default is 9)", "Edit Timer Hours", CStr(timerHours))
            If newHrs <> "" And IsNumeric(newHrs) Then
                If CInt(newHrs) >= 0 And CInt(newHrs) <= 23 Then
                    timerHours = CInt(newHrs)
                    SaveConfig
                Else
                    MsgBox "Please enter a valid number between 0 and 23.", vbExclamation, "Invalid Input"
                End If
            End If

        Case "2"
            Dim newMins
            newMins = InputBox("Enter Timer Minutes (0 - 59):" & vbCrLf & "(Default is 0)", "Edit Timer Minutes", CStr(timerMinutes))
            If newMins <> "" And IsNumeric(newMins) Then
                If CInt(newMins) >= 0 And CInt(newMins) <= 59 Then
                    If timerHours = 0 And CInt(newMins) = 0 Then
                        MsgBox "Total timer duration must be at least 1 minute.", vbExclamation, "Invalid Input"
                    Else
                        timerMinutes = CInt(newMins)
                        SaveConfig
                    End If
                Else
                    MsgBox "Please enter a valid number between 0 and 59.", vbExclamation, "Invalid Input"
                End If
            End If

        Case "3"
            Dim newUrl
            newUrl = InputBox("Enter Redirect URL (e.g., HR Portal, GitHub, Jira):", "Edit Redirect URL", redirectUrl)
            If Trim(newUrl) <> "" Then
                newUrl = Trim(newUrl)
                If Left(LCase(newUrl), 7) <> "http://" And Left(LCase(newUrl), 8) <> "https://" Then
                    newUrl = "https://" & newUrl
                End If
                redirectUrl = newUrl
                SaveConfig
            End If

        Case "4"
            Dim newStartMsg
            newStartMsg = InputBox("Enter Start Prompt Message:", "Edit Start Message", startTitle)
            If Trim(newStartMsg) <> "" Then
                startTitle = Trim(newStartMsg)
                SaveConfig
            End If

        Case "5"
            Dim newEndTitle
            newEndTitle = InputBox("Enter Completion Title (e.g., 'Time is Up!'):", "Edit Completion Title", endTitle)
            If Trim(newEndTitle) <> "" Then
                endTitle = Trim(newEndTitle)
                SaveConfig
            End If

        Case "6"
            Dim newEndSub
            newEndSub = InputBox("Enter Completion Subtitle (e.g., 'Remember to Sign Out'):", "Edit Completion Subtitle", endSubtitle)
            If Trim(newEndSub) <> "" Then
                endSubtitle = Trim(newEndSub)
                SaveConfig
            End If

        Case "7"
            SetDefaultConfig
            SaveConfig
            MsgBox "Settings have been reset to defaults (9 hr 0 min).", vbInformation, "Defaults Restored"

        Case Else
            MsgBox "Please enter a valid option number between 0 and 7.", vbExclamation, "Invalid Option"
    End Select
Loop

' Return to the appropriate prompt
If isEndMode Then
    ShowEndPrompt
Else
    ShowStartPrompt
End If


End Sub

' ==============================================================================
' 5. HELPER FUNCTIONS (BROWSER LAUNCH & FILE PERSISTENCE)
' ==============================================================================
Sub OpenRedirectUrl()
On Error Resume Next
shell.Run Chr(34) & redirectUrl & Chr(34), 1, False
If Err.Number <> 0 Then
MsgBox "Could not open URL: " & redirectUrl & vbCrLf & "Error: " & Err.Description, vbCritical, "Browser Error"
Err.Clear
End If
On Error GoTo 0
End Sub

Sub SetDefaultConfig()
timerHours = 9
timerMinutes = 0
redirectUrl = "https://github.com"
startTitle = "Ready to start your session? Check-in & start timer."
endTitle = "Timer Complete! Time to wrap up."
endSubtitle = "Your scheduled duration has finished - Take action now."
End Sub

Sub LoadConfig()
SetDefaultConfig
On Error Resume Next
If fso.FileExists(configPath) Then
Dim file, line, eqPos, k, v
Set file = fso.OpenTextFile(configPath, 1, False)
Do Until file.AtEndOfStream
line = file.ReadLine
eqPos = InStr(line, "=")
If eqPos > 1 Then
k = Left(line, eqPos - 1)
v = Mid(line, eqPos + 1)
Select Case k
Case "timer_hours": If IsNumeric(v) Then timerHours = CInt(v)
Case "timer_minutes": If IsNumeric(v) Then timerMinutes = CInt(v)
Case "redirect_url": If v <> "" Then redirectUrl = v
Case "start_title": If v <> "" Then startTitle = v
Case "end_title": If v <> "" Then endTitle = v
Case "end_subtitle": If v <> "" Then endSubtitle = v
End Select
End If
Loop
file.Close
Else
SaveConfig
End If
On Error GoTo 0
End Sub

Sub SaveConfig()
On Error Resume Next
Dim file
Set file = fso.CreateTextFile(configPath, True)
file.WriteLine "timer_hours=" & timerHours
file.WriteLine "timer_minutes=" & timerMinutes
file.WriteLine "redirect_url=" & redirectUrl
file.WriteLine "start_title=" & startTitle
file.WriteLine "end_title=" & endTitle
file.WriteLine "end_subtitle=" & endSubtitle
file.Close
On Error GoTo 0
End Sub

Sub LoadState()
activeTask = False
startTimeStr = ""
targetTimeStr = ""
On Error Resume Next
If fso.FileExists(statePath) Then
Dim file, line, eqPos, k, v
Set file = fso.OpenTextFile(statePath, 1, False)
Do Until file.AtEndOfStream
line = file.ReadLine
eqPos = InStr(line, "=")
If eqPos > 1 Then
k = Left(line, eqPos - 1)
v = Mid(line, eqPos + 1)
Select Case k
Case "active_task": activeTask = (LCase(v) = "true")
Case "start_time": startTimeStr = v
Case "target_time": targetTimeStr = v
End Select
End If
Loop
file.Close
End If
On Error GoTo 0
End Sub

Sub SaveState(isActive, startStr, targetStr)
activeTask = isActive
startTimeStr = startStr
targetTimeStr = targetStr
On Error Resume Next
Dim file
Set file = fso.CreateTextFile(statePath, True)
file.WriteLine "active_task=" & LCase(CStr(isActive))
file.WriteLine "start_time=" & startStr
file.WriteLine "target_time=" & targetStr
file.Close
On Error GoTo 0
End Sub

Sub ClearState()
SaveState False, "", ""
End Sub