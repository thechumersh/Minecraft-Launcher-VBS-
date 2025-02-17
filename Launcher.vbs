Option Explicit

Dim shell, fso, appPath, settingsFile, minecraftPath, javaPath, javaArgs, nickname, ramAmount
Dim logDirectory, sessionLogFile, logFileName, logNumber, currentDate, uiStyle

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Определяем пути
appPath = fso.GetParentFolderName(WScript.ScriptFullName)
settingsFile = appPath & "\settings.txt"
logDirectory = appPath & "\logs"

' Проверяем и создаём директорию для логов, если она не существует
If Not fso.FolderExists(logDirectory) Then
    fso.CreateFolder(logDirectory)
End If

' Логирование событий
Sub LogEvent(eventMessage)
    Dim logFile, timestamp
    timestamp = Now
    Set logFile = fso.OpenTextFile(sessionLogFile, 8, True)
    logFile.WriteLine "[" & FormatDateTime(timestamp, 2) & "] [" & FormatDateTime(timestamp, 4) & "] " & eventMessage
    logFile.Close
End Sub

' Создаём или определяем лог-файл
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
    LogEvent "[MAIN/CLIENT]: Launcher started"
End Sub

' Загрузка настроек
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
        LogEvent "[MAIN/CLIENT]: Settings loaded successfully"
    Else
        nickname = "Player"
        ramAmount = "1024M"
        minecraftPath = "C:\Minecraft\minecraft.jar"
        javaPath = "java"
        javaArgs = ""
        uiStyle = "Default"
        SaveSettings
        LogEvent "[WARNING/CLIENT]: Settings not found. New ones created with default values"
    End If
End Sub

' Сохранение настроек
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

' Применение стиля интерфейса
Function FormatText(title, content)
    Select Case uiStyle
        Case "Default"
            FormatText = title & vbCrLf & content
        Case "Modern"
            FormatText = "==== " & title & " ====" & vbCrLf & content
        Case "Retro"
            FormatText = "+----------------+" & vbCrLf & "| " & title & " |" & vbCrLf & "+----------------+" & vbCrLf & content
        Case "Elegant"
            FormatText = "¦¦ " & title & " ¦¦" & vbCrLf & content
        Case "Cyber"
            FormatText = ">>> " & title & " <<<" & vbCrLf & content
        Case Else
            FormatText = title & vbCrLf & content
    End Select
End Function

' Подтверждение выхода
Function ConfirmExit
    Dim confirmResult
    confirmResult = MsgBox("Are you sure you want to go out?", vbYesNo + vbQuestion, "Confirm exit")
    If confirmResult = vbYes Then
        LogEvent "[MAIN/CLIENT]: Launcher closed" ' Записываем в лог перед выходом
        ConfirmExit = True
    Else
        ConfirmExit = False
    End If
End Function

' Функция изменения стиля интерфейса
Sub ChangeUIStyle
    Dim choice
    choice = InputBox(FormatText("Choose an interface style", _
                                 "1. Default" & vbCrLf & _
                                 "2. Modern" & vbCrLf & _
                                 "3. Retro" & vbCrLf & _
                                 "4. Elegant" & vbCrLf & _
                                 "5. Cyber" & vbCrLf & _
                                 "6. Назад"))
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
            MsgBox "The wrong choice!", vbExclamation, "Error"
    End Select
    LogEvent "[INFO/CLIENT]: Interface style changed to: " & uiStyle
End Sub

' Меню настроек
Sub SettingsMenu
    Dim choice, input
    Do
        choice = InputBox(FormatText("Settings", _ 
                                     "1. Change Nick (current: " & nickname & ")" & vbCrLf & _
                                     "2. Change RAM (current: " & ramAmount & ")" & vbCrLf & _
                                     "3. Change the path for running the game" & vbCrLf & _
                                     "4. Change Java arguments (current: " & javaArgs & ")" & vbCrLf & _
                                     "5. Change UI style (current: " & uiStyle & ")" & vbCrLf & _
                                     "6. Back"))
        If choice = "" Then Exit Do
        Select Case choice
            Case "1"
                input = InputBox("Enter a new nickname:", "Change nickname", nickname)
                If input = "" Then
                    MsgBox "Nothing was entered in the field, nickname change was canceled.", vbExclamation, "Cancel"
                    LogEvent "[WARNING/CLIENT]: Nickname change cancelled, field was empty."
                Else
                    LogEvent "[INFO/CLIENT]: Change Nick from '" & nickname & "' to '" & input & "'"
                    nickname = input
                End If
            Case "2"
                input = InputBox("Enter a new RAM value (example, 1024M):", "Changing RAM", ramAmount)
                If input = "" Then
                    MsgBox "Nothing was entered in the field, RAM change was cancelled.", vbExclamation, "Cancel"
                    LogEvent "[WARNING/CLIENT]: RAM change cancelled, field was empty."
                Else
                    LogEvent "[INFO/CLIENT]: Changed ram from '" & ramAmount & "' to '" & input & "'"
                    ramAmount = input
                End If
            Case "3"
                ChangePaths
            Case "4"
                input = InputBox("Enter new Java arguments:", "Changing Java Arguments", javaArgs)
                If input = "" Then
                    MsgBox "Nothing was entered in the field, Java arguments change was cancelled.", vbExclamation, "Cancel"
                    LogEvent "[WARNING/CLIENT]: Java arguments change cancelled, field was empty."
                Else
                    LogEvent "[INFO/CLIENT]: Changed Java arguments from '" & javaArgs & "' to '" & input & "'"
                    javaArgs = input
                End If
            Case "5"
                ChangeUIStyle
            Case "6"
                Exit Do
            Case Else
                MsgBox "Wrong choice!", vbExclamation, "Error"
        End Select
        SaveSettings
    Loop
End Sub

' Изменение путей
Sub ChangePaths
    Dim input
    input = InputBox("Enter the new path to Minecraft:", "Changing Minecraft Path", minecraftPath)
    If input = "" Then
        MsgBox "Nothing was entered in the field, the change to the Minecraft path was canceled.", vbExclamation, "Cancel"
        LogEvent "[ERROR/CLIENT]: Minecraft path change reverted, field was empty."
    ElseIf Not fso.FileExists(input & "\minecraft.jar") Then
        MsgBox "The file minecraft.jar was not found at the specified path.", vbCritical, "Error"
        LogEvent "[ERROR/CLIENT]: Error: minecraft.jar not found"
    Else
        LogEvent "[INFO/CLIENT]: Changed path to Minecraft from '" & minecraftPath & "' to '" & input & "'"
        minecraftPath = input
    End If
End Sub

' Запуск игры
Sub LaunchGame
    If minecraftPath = "" Or javaPath = "" Then
        MsgBox "Paths to Minecraft or Java are not specified. Check your settings.", vbCritical, "Launch error"
		LogEvent "[ERROR/CLIENT]: Error: Path will not find Minecraft or Java"
        Exit Sub
    End If
    
    Dim commandLine, launcherHwnd, exitCode

    ' Формируем команду для запуска Minecraft
    commandLine = """" & javaPath & """ -Xmx" & ramAmount & " -Djava.library.path=natives -cp """ & _
                  "minecraft.jar;jinput.jar;lwjgl.jar;lwjgl_util.jar""" & _
                  " net.minecraft.client.Minecraft " & nickname & " " & javaArgs
    LogEvent "[INFO/CLIENT]: Launching the game with the team: " & commandLine

    ' Получаем идентификатор окна лаунчера
    On Error Resume Next
    launcherHwnd = shell.AppActivate("Minecraft Launcher") ' Укажи название окна лаунчера
    On Error GoTo 0

    ' Скрываем лаунчер (если он активен)
    If launcherHwnd Then
        shell.SendKeys "% n" ' Отправляем Alt+Tab для сворачивания
    End If

    ' Запускаем Minecraft в скрытом режиме
    On Error Resume Next
    exitCode = shell.Run(commandLine, 0, False) ' Флаг 0 скрывает консольное окно
    If Err.Number <> 0 Then
        MsgBox "Error when starting the game. More details in the logs - latest_log.txt" & Err.Description, vbCritical, "Error"
        LogEvent "[ERROR/CLIENT]: Error starting the game. Details: Java 8 is missing!" & Err.Description
    Else
        LogEvent "[INFO/CLIENT]: The game has been launched successfully."
    End If
    On Error GoTo 0
End Sub

' Основное меню
Sub MainMenu
    Dim choice
    InitializeLog
    LoadSettings
    Do
        choice = InputBox(FormatText("Menu", _
                                     "1. Launch" & vbCrLf & _
                                     "2. Settings" & vbCrLf & _
                                     "3. Exit"))
        If choice = "" Then Exit Do
        Select Case choice
            Case "1"
                LaunchGame
            Case "2"
                SettingsMenu
            Case "3"
                If ConfirmExit Then Exit Do
            Case Else
                MsgBox "Wrong choice!", vbExclamation, "Error"
        End Select
    Loop
End Sub

' Запуск основной программы
MainMenu
