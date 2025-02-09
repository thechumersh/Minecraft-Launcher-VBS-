Option Explicit

Dim shell, fso, appPath, settingsFile, minecraftPath, javaPath, javaArgs, nickname, ramAmount
Dim logDirectory, sessionLogFile, logFileName, logNumber, currentDate, uiStyle

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' ���������� ����
appPath = fso.GetParentFolderName(WScript.ScriptFullName)
settingsFile = appPath & "\settings.txt"
logDirectory = appPath & "\logs"

' ��������� � ������ ���������� ��� �����, ���� ��� �� ����������
If Not fso.FolderExists(logDirectory) Then
    fso.CreateFolder(logDirectory)
End If

' ����������� �������
Sub LogEvent(eventMessage)
    Dim logFile, timestamp
    timestamp = Now
    Set logFile = fso.OpenTextFile(sessionLogFile, 8, True)
    logFile.WriteLine "[" & FormatDateTime(timestamp, 2) & "] [" & FormatDateTime(timestamp, 4) & "] " & eventMessage
    logFile.Close
End Sub

' ������ ��� ���������� ���-����
Sub InitializeLog
    currentDate = Year(Now) & "-" & Right("00" & Month(Now), 2) & "-" & Right("00" & Day(Now), 2)
    logFileName = logDirectory & "\" & currentDate & "-log.txt"
    logNumber = 1
    While fso.FileExists(logFileName)
        logFileName = logDirectory & "\" & currentDate & "-" & logNumber & "-log.txt"
        logNumber = logNumber + 1
    Wend
    sessionLogFile = logDirectory & "\latest_log.txt"
    If fso.FileExists(sessionLogFile) Then
        If Not fso.FileExists(logFileName) Then
            fso.MoveFile sessionLogFile, logFileName
        End If
    End If
    LogEvent "[INFO] ������� �������"
End Sub

' �������� ��������
Sub LoadSettings
    If fso.FileExists(settingsFile) Then
        Dim settings, line, lines
        Set settings = fso.OpenTextFile(settingsFile, 1)
        lines = Split(settings.ReadAll, vbCrLf)
        settings.Close
        For Each line In lines
            If InStr(line, "nickname=") = 1 Then
                nickname = Mid(line, 10)
            ElseIf InStr(line, "ramAmount=") = 1 Then
                ramAmount = Mid(line, 11)
            ElseIf InStr(line, "minecraftPath=") = 1 Then
                minecraftPath = Mid(line, 15)
            ElseIf InStr(line, "javaPath=") = 1 Then
                javaPath = Mid(line, 10)
            ElseIf InStr(line, "javaArgs=") = 1 Then
                javaArgs = Mid(line, 10)
            ElseIf InStr(line, "uiStyle=") = 1 Then
                uiStyle = Mid(line, 9)
            End If
        Next
        LogEvent "[CLIENT] ��������� ������� ���������"
    Else
        nickname = "Player"
        ramAmount = "1024M"
        minecraftPath = "C:\Minecraft\minecraft.jar"
        javaPath = "java"
        javaArgs = ""
        uiStyle = "Default"
        SaveSettings
        LogEvent "[WARNING] ��������� �� �������. ������� ����� � �������������� ����������"
    End If
End Sub

' ���������� ��������
Sub SaveSettings
    Dim settings
    Set settings = fso.CreateTextFile(settingsFile, True)
    settings.WriteLine "nickname=" & nickname
    settings.WriteLine "ramAmount=" & ramAmount
    settings.WriteLine "minecraftPath=" & minecraftPath
    settings.WriteLine "javaPath=" & javaPath
    settings.WriteLine "javaArgs=" & javaArgs
    settings.WriteLine "uiStyle=" & uiStyle
    settings.Close
End Sub

' ���������� ����� ����������
Function FormatText(title, content)
    Select Case uiStyle
        Case "Default"
            FormatText = title & vbCrLf & content
        Case "Modern"
            FormatText = "==== " & title & " ====" & vbCrLf & content
        Case "Retro"
            FormatText = "+----------------+" & vbCrLf & "| " & title & " |" & vbCrLf & "+----------------+" & vbCrLf & content
        Case "Elegant"
            FormatText = "�� " & title & " ��" & vbCrLf & content
        Case "Cyber"
            FormatText = ">>> " & title & " <<<" & vbCrLf & content
        Case Else
            FormatText = title & vbCrLf & content
    End Select
End Function

' ������������� ������
Function ConfirmExit
    Dim confirmResult
    confirmResult = MsgBox("�� ������������� ������ �����?", vbYesNo + vbQuestion, "������������� ������")
    If confirmResult = vbYes Then
        LogEvent "[INFO] ������� ������" ' ���������� � ��� ����� �������
        ConfirmExit = True
    Else
        ConfirmExit = False
    End If
End Function

' ������� ��������� ����� ����������
Sub ChangeUIStyle
    Dim choice
    choice = InputBox(FormatText("�������� ����� ����������", _
                                 "1. Default" & vbCrLf & _
                                 "2. Modern" & vbCrLf & _
                                 "3. Retro" & vbCrLf & _
                                 "4. Elegant" & vbCrLf & _
                                 "5. Cyber" & vbCrLf & _
                                 "6. �����"))
    Select Case choice
        Case "1"
            uiStyle = "Default"
        Case "2"
            uiStyle = "Modern"
        Case "3"
            uiStyle = "Retro"
        Case "4"
            uiStyle = "Elegant"
        Case "5"
            uiStyle = "Cyber"
        Case "6"
            Exit Sub
        Case Else
            MsgBox "�������� �����!", vbExclamation, "������"
    End Select
    LogEvent "[CLIENT] ����� ���������� ������ ��: " & uiStyle
End Sub

' ���� ��������
Sub SettingsMenu
    Dim choice, input
    Do
        choice = InputBox(FormatText("���������", _ 
                                     "1. �������� ��� (�������: " & nickname & ")" & vbCrLf & _
                                     "2. �������� ��� (�������: " & ramAmount & ")" & vbCrLf & _
                                     "3. �������� ���� ��� ������ ����" & vbCrLf & _
                                     "4. �������� Java ��������� (�������: " & javaArgs & ")" & vbCrLf & _
                                     "5. �������� ����� ���������� (�������: " & uiStyle & ")" & vbCrLf & _
                                     "6. �����"))
        If choice = "" Then Exit Do
        Select Case choice
            Case "1"
                input = InputBox("������� ����� ���:", "��������� ����", nickname)
                If input = "" Then
                    MsgBox "� ���� ������ �� ���� �������, ��������� ���� ��������.", vbExclamation, "������"
                    LogEvent "[WARNING] ��������� ���� ��������, ���� ���� ������."
                Else
                    LogEvent "[CLIENT] ������ ��� � '" & nickname & "' �� '" & input & "'"
                    nickname = input
                End If
            Case "2"
                input = InputBox("������� ����� �������� ��� (��������, 1024M):", "��������� ���", ramAmount)
                If input = "" Then
                    MsgBox "� ���� ������ �� ���� �������, ��������� ��� ��������.", vbExclamation, "������"
                    LogEvent "[WARNING] ��������� ��� ��������, ���� ���� ������."
                Else
                    LogEvent "[CLIENT] �������� ��� � '" & ramAmount & "' �� '" & input & "'"
                    ramAmount = input
                End If
            Case "3"
                ChangePaths
            Case "4"
                input = InputBox("������� ����� Java ���������:", "��������� Java ����������", javaArgs)
                If input = "" Then
                    MsgBox "� ���� ������ �� ���� �������, ��������� Java ���������� ��������.", vbExclamation, "������"
                    LogEvent "[WARNING] ��������� Java ���������� ��������, ���� ���� ������."
                Else
                    LogEvent "[CLIENT] �������� Java ��������� � '" & javaArgs & "' �� '" & input & "'"
                    javaArgs = input
                End If
            Case "5"
                ChangeUIStyle
            Case "6"
                Exit Do
            Case Else
                MsgBox "�������� �����!", vbExclamation, "������"
        End Select
        SaveSettings
    Loop
End Sub

' ��������� �����
Sub ChangePaths
    Dim input
    input = InputBox("������� ����� ���� � Minecraft:", "��������� ���� Minecraft", minecraftPath)
    If input = "" Then
        MsgBox "� ���� ������ �� ���� �������, ��������� ���� Minecraft ��������.", vbExclamation, "������"
        LogEvent "��������� ���� Minecraft ��������, ���� ���� ������."
    ElseIf Not fso.FileExists(input & "\minecraft.jar") Then
        MsgBox "[WARNING] ���� minecraft.jar �� ������ �� ���������� ����.", vbCritical, "������"
        LogEvent "[ERROR] ������: minecraft.jar �� ������"
    Else
        LogEvent "[CLIENT] ������ ���� � Minecraft � '" & minecraftPath & "' �� '" & input & "'"
        minecraftPath = input
    End If
End Sub

' ������ ����
Sub LaunchGame
    If minecraftPath = "" Or javaPath = "" Then
        MsgBox "���� � Minecraft ��� Java �� �������. ��������� ���������.", vbCritical, "������ �������"
		LogEvent "[ERROR] ������: ���� �� ������ Minecraft ��� Java"
        Exit Sub
    End If
    
    Dim commandLine, launcherHwnd, exitCode

    ' ��������� ������� ��� ������� Minecraft
    commandLine = """" & javaPath & """ -Xmx" & ramAmount & " -Djava.library.path=natives -cp """ & _
                  "minecraft.jar;jinput.jar;lwjgl.jar;lwjgl_util.jar""" & _
                  " net.minecraft.client.Minecraft " & nickname & " " & javaArgs
    LogEvent "[CLIENT] ������ ���� � ��������: " & commandLine

    ' �������� ������������� ���� ��������
    On Error Resume Next
    launcherHwnd = shell.AppActivate("������� Minecraft") ' ����� �������� ���� ��������
    On Error GoTo 0

    ' �������� ������� (���� �� �������)
    If launcherHwnd Then
        shell.SendKeys "% n" ' ���������� Alt+Tab ��� ������������
    End If

    ' ��������� Minecraft � ������� ������
    On Error Resume Next
    exitCode = shell.Run(commandLine, 0, False) ' ���� 0 �������� ���������� ����
    If Err.Number <> 0 Then
        MsgBox "������ ��� ������� ����. ��������� � ����� - latest_log.txt" & Err.Description, vbCritical, "������"
        LogEvent "[ERROR] ������ ��� ������� ����. ��������: ����������� Java 8!" & Err.Description
    Else
        LogEvent "[CLIENT] ���� �������� �������"
    End If
    On Error GoTo 0
End Sub

' �������� ����
Sub MainMenu
    Dim choice
    InitializeLog
    LoadSettings
    Do
        choice = InputBox(FormatText("���� ��������", _
                                     "1. ������ ����" & vbCrLf & _
                                     "2. ���������" & vbCrLf & _
                                     "3. �����"))
        If choice = "" Then Exit Do
        Select Case choice
            Case "1"
                LaunchGame
            Case "2"
                SettingsMenu
            Case "3"
                If ConfirmExit Then Exit Do
            Case Else
                MsgBox "�������� �����!", vbExclamation, "������"
        End Select
    Loop
End Sub

' ������ �������� ���������
MainMenu
