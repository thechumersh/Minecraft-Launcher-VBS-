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
    LogEvent "[INFO] Лаунчер запущен"
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
        LogEvent "[CLIENT] Настройки успешно загружены"
    Else
        nickname = "Player"
        ramAmount = "1024M"
        minecraftPath = "C:\Minecraft\minecraft.jar"
        javaPath = "java"
        javaArgs = ""
        uiStyle = "Default"
        SaveSettings
        LogEvent "[WARNING] Настройки не найдены. Созданы новые с умолчательными значениями"
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
    confirmResult = MsgBox("Вы действительно хотите выйти?", vbYesNo + vbQuestion, "Подтверждение выхода")
    If confirmResult = vbYes Then
        LogEvent "[INFO] Лаунчер закрыт" ' Записываем в лог перед выходом
        ConfirmExit = True
    Else
        ConfirmExit = False
    End If
End Function

' Функция изменения стиля интерфейса
Sub ChangeUIStyle
    Dim choice
    choice = InputBox(FormatText("Выберите стиль интерфейса", _
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
            MsgBox "Неверный выбор!", vbExclamation, "Ошибка"
    End Select
    LogEvent "[CLIENT] Стиль интерфейса изменён на: " & uiStyle
End Sub

' Меню настроек
Sub SettingsMenu
    Dim choice, input
    Do
        choice = InputBox(FormatText("Настройки", _ 
                                     "1. Изменить ник (текущий: " & nickname & ")" & vbCrLf & _
                                     "2. Изменить ОЗУ (текущее: " & ramAmount & ")" & vbCrLf & _
                                     "3. Изменить путь для работы игры" & vbCrLf & _
                                     "4. Изменить Java аргументы (текущие: " & javaArgs & ")" & vbCrLf & _
                                     "5. Изменить стиль интерфейса (текущий: " & uiStyle & ")" & vbCrLf & _
                                     "6. Назад"))
        If choice = "" Then Exit Do
        Select Case choice
            Case "1"
                input = InputBox("Введите новый ник:", "Изменение ника", nickname)
                If input = "" Then
                    MsgBox "В поле нечего не было введено, изменение ника отменено.", vbExclamation, "Отмена"
                    LogEvent "[WARNING] Изменение ника отменено, поле было пустым."
                Else
                    LogEvent "[CLIENT] Изменён ник с '" & nickname & "' на '" & input & "'"
                    nickname = input
                End If
            Case "2"
                input = InputBox("Введите новое значение ОЗУ (например, 1024M):", "Изменение ОЗУ", ramAmount)
                If input = "" Then
                    MsgBox "В поле нечего не было введено, изменение ОЗУ отменено.", vbExclamation, "Отмена"
                    LogEvent "[WARNING] Изменение ОЗУ отменено, поле было пустым."
                Else
                    LogEvent "[CLIENT] Изменена ОЗУ с '" & ramAmount & "' на '" & input & "'"
                    ramAmount = input
                End If
            Case "3"
                ChangePaths
            Case "4"
                input = InputBox("Введите новые Java аргументы:", "Изменение Java аргументов", javaArgs)
                If input = "" Then
                    MsgBox "В поле нечего не было введено, изменение Java аргументов отменено.", vbExclamation, "Отмена"
                    LogEvent "[WARNING] Изменение Java аргументов отменено, поле было пустым."
                Else
                    LogEvent "[CLIENT] Изменены Java аргументы с '" & javaArgs & "' на '" & input & "'"
                    javaArgs = input
                End If
            Case "5"
                ChangeUIStyle
            Case "6"
                Exit Do
            Case Else
                MsgBox "Неверный выбор!", vbExclamation, "Ошибка"
        End Select
        SaveSettings
    Loop
End Sub

' Изменение путей
Sub ChangePaths
    Dim input
    input = InputBox("Введите новый путь к Minecraft:", "Изменение пути Minecraft", minecraftPath)
    If input = "" Then
        MsgBox "В поле нечего не было введено, изменение пути Minecraft отменено.", vbExclamation, "Отмена"
        LogEvent "Изменение пути Minecraft отменено, поле было пустым."
    ElseIf Not fso.FileExists(input & "\minecraft.jar") Then
        MsgBox "[WARNING] Файл minecraft.jar не найден по указанному пути.", vbCritical, "Ошибка"
        LogEvent "[ERROR] Ошибка: minecraft.jar не найден"
    Else
        LogEvent "[CLIENT] Изменён путь к Minecraft с '" & minecraftPath & "' на '" & input & "'"
        minecraftPath = input
    End If
End Sub

' Запуск игры
Sub LaunchGame
    If minecraftPath = "" Or javaPath = "" Then
        MsgBox "Пути к Minecraft или Java не указаны. Проверьте настройки.", vbCritical, "Ошибка запуска"
		LogEvent "[ERROR] Ошибка: Путь не найдет Minecraft или Java"
        Exit Sub
    End If
    
    Dim commandLine, launcherHwnd, exitCode

    ' Формируем команду для запуска Minecraft
    commandLine = """" & javaPath & """ -Xmx" & ramAmount & " -Djava.library.path=natives -cp """ & _
                  "minecraft.jar;jinput.jar;lwjgl.jar;lwjgl_util.jar""" & _
                  " net.minecraft.client.Minecraft " & nickname & " " & javaArgs
    LogEvent "[CLIENT] Запуск игры с командой: " & commandLine

    ' Получаем идентификатор окна лаунчера
    On Error Resume Next
    launcherHwnd = shell.AppActivate("Лаунчер Minecraft") ' Укажи название окна лаунчера
    On Error GoTo 0

    ' Скрываем лаунчер (если он активен)
    If launcherHwnd Then
        shell.SendKeys "% n" ' Отправляем Alt+Tab для сворачивания
    End If

    ' Запускаем Minecraft в скрытом режиме
    On Error Resume Next
    exitCode = shell.Run(commandLine, 0, False) ' Флаг 0 скрывает консольное окно
    If Err.Number <> 0 Then
        MsgBox "Ошибка при запуске игры. Подробнее в логах - latest_log.txt" & Err.Description, vbCritical, "Ошибка"
        LogEvent "[ERROR] Ошибка при запуске игры. Подробно: Отсутствует Java 8!" & Err.Description
    Else
        LogEvent "[CLIENT] Игра запущена успешно"
    End If
    On Error GoTo 0
End Sub

' Основное меню
Sub MainMenu
    Dim choice
    InitializeLog
    LoadSettings
    Do
        choice = InputBox(FormatText("Меню лаунчера", _
                                     "1. Запуск игры" & vbCrLf & _
                                     "2. Настройки" & vbCrLf & _
                                     "3. Выход"))
        If choice = "" Then Exit Do
        Select Case choice
            Case "1"
                LaunchGame
            Case "2"
                SettingsMenu
            Case "3"
                If ConfirmExit Then Exit Do
            Case Else
                MsgBox "Неверный выбор!", vbExclamation, "Ошибка"
        End Select
    Loop
End Sub

' Запуск основной программы
MainMenu
