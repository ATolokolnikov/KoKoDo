Attribute VB_Name = "basTrace"
Option Explicit

'-------------------------------------------------------------------------------
' Класс для обработки багов и оптимизации работы Word.
'-------------------------------------------------------------------------------

' -- Закрытые переменные:
Private mcolIds As Collection
Private mcolTitles As Collection
Private miFile As Integer
Private msErrHistory As String

' Копии параметров Excel:
Private dicAppParams As New Scripting.Dictionary

'-------------------------------------------------------------------------------
' Возвращает стек вызовов в строковом виде.
'-------------------------------------------------------------------------------
Public Function CallStack() As String

    Dim iFrame As Integer
    Dim zResult As String

    zResult = "FPSFED_VBA_" & gsBuild

    If mcolIds Is Nothing Then
    ElseIf mcolIds.Count = 0 Then
    Else
        For iFrame = 1 To mcolIds.Count
            zResult = zResult & "/" & mcolIds.Item(iFrame)
        Next iFrame
    End If

    CallStack = zResult & ":" & Erl

End Function

'-------------------------------------------------------------------------------
' Отключает перерисовку окна Excel и реакцию Excel на действия пользователя,
' чтобы ускорить выполнение процедур по массовому обновлению ячеек.
'-------------------------------------------------------------------------------
Public Sub Dive()

    ' Отключение перерисовки и реакции Excel:
    With Application
        If .DisplayAlerts Then
            .DisplayAlerts = False
        End If
        If .ScreenUpdating Then
            .ScreenUpdating = False
        End If
    End With

End Sub

'-------------------------------------------------------------------------------
' Пополняет стек выполнения записью о текущей процедуре.
'-------------------------------------------------------------------------------
Public Sub EnterFrame(ByVal ModuleId As String _
                    , ByVal RoutineId As String _
                    , ByVal RoutineTitle As String)

    Dim sKey As String

    If mcolIds Is Nothing Then
        Set mcolIds = New Collection
        Set mcolTitles = New Collection
    End If

    ModuleId = Format(ModuleId, ">@@@@")
    RoutineId = Format(RoutineId, ">@@@@@")

    sKey = ModuleId & "-" & RoutineId
    sKey = Replace(sKey, " ", "_")
    mcolIds.Add sKey
    mcolTitles.Add RoutineTitle

    Application.StatusBar = RoutineTitle

End Sub

'-------------------------------------------------------------------------------
' Возвращает ID текущей ошибки ([Модуль]-[Процедура]-[LID строки]-[№ ошибки]).
' Например: "RUIB-BTONA-101-0x1a2b3c4d"
'-------------------------------------------------------------------------------
Public Function ErrId() As String

    Dim zResult As String

    zResult = "[RoutineId]-[LineId]-[Err]"
    zResult = Replace(zResult, "[RoutineId]", SubId)
    zResult = Replace(zResult, "[LineId]", Format(Erl, "000"))
    zResult = Replace(zResult, "[Err]", Err.Number)

    ErrId = zResult

End Function

'-------------------------------------------------------------------------------
' Удаляет последнюю запись из стека выполнения.
'-------------------------------------------------------------------------------
Public Sub LeaveFrame()

    If mcolIds Is Nothing Then
    ElseIf mcolIds.Count = 0 Then
    Else
        If mcolIds.Count > 1 Then
            If dicAppParams.Exists(mcolIds.Count) Then
                dicAppParams.Remove (mcolIds.Count)
            End If
        End If
        mcolIds.Remove mcolIds.Count
        mcolTitles.Remove mcolTitles.Count
    End If

    If mcolIds.Count = 0 Then
        SOS
        Application.StatusBar = False
    End If

End Sub

'-------------------------------------------------------------------------------
' Создаёт в журнале ошибок новую запись о событии [LogText] с типом [LogType].
'-------------------------------------------------------------------------------
Public Sub Log(ByVal LogType As VbMsgBoxStyle _
             , Optional ByVal LogText As String)

    Dim sType As String

    If CBool(LogType And vbCritical) And (LogType And vbQuestion) = 0 Then
        sType = "Ошибка системы"
    ElseIf (LogType And vbExclamation) = vbExclamation Then
        sType = "Проверьте данные"
    ElseIf LogType And vbQuestion Then
        sType = "Вопрос"
    ElseIf LogType And vbInformation Then
        sType = "Информация"
    End If

    If LogType And vbCritical And (LogType And vbQuestion) = 0 Then
        ' TODO: how user feed back?

        Select Case Err.Number
        Case 0, vbObjectError
            If LogText = Empty Then
                LogText = gsMes_ErrUnexpected
            End If
            If msErrHistory <> Empty Then
                LogText = msErrHistory & LogText
            End If
        Case Else
            If LogText = Empty Then
                LogText = gsMes_ErrUnexpected
            End If
            LogText = LogText & vbCrLf & "Код ошибки: " & ErrId
            msErrHistory = msErrHistory & vbCrLf & vbCrLf & LogText
        End Select

    End If

    ' Создать новую запись о событии в журнале ошибок.
    LogEntry UCase(sType)

    ' Описать параметры события.
    If Err.Number Then
        LogAttr "Заголовок JIRA", "VBA. (" & ErrId & ") " & SubTitle & " — " & sType
        LogAttr "Описание ошибки", Err.Description
        Err.Clear
    End If
    LogAttr "Call Stack", CallStack
    LogAttr "Текст", "{noformat}" & LogText & "{noformat}"

    ' Закрыть журнал ошибок.
    Print #miFile, Empty
    Close #miFile
    miFile = 0

End Sub

'-------------------------------------------------------------------------------
' Создаёт строку со значением параметра ошибки в журнале ошибок.
'-------------------------------------------------------------------------------
Public Sub LogAttr(AttrName As String, AttrValue As String)

    Dim zResult As String

    zResult = "* " & AttrName & ": " & AttrValue

    If miFile = 0 Then
        miFile = FreeFile
        Open Path & "\" & gsLog_FileName For Append As #miFile
    End If

    Print #miFile, zResult

End Sub

'-------------------------------------------------------------------------------
' Создаёт строку с текстом события в журнале ошибок.
'-------------------------------------------------------------------------------
Public Sub LogEntry(EntryTitle As String)

    Dim zResult As String

    zResult = Format(Now, "yyyy-mm-dd hh:mm, ss.") _
            & Format((Timer * 1000) Mod 1000, "000") _
            & " - " & EntryTitle

    If miFile = 0 Then
        miFile = FreeFile
        Open Path & "\" & gsLog_FileName For Append As #miFile
    End If

    Print #miFile, zResult

End Sub

'-------------------------------------------------------------------------------
' Отображает окно с типом [LogType] и сообщением [LogText].
'-------------------------------------------------------------------------------
Public Function Msg(ByVal LogType As VbMsgBoxStyle _
                  , Optional ByVal LogText As String) As VbMsgBoxResult

    Dim sType As String

    If CBool(LogType And vbCritical) And (LogType And vbQuestion) = 0 Then
        sType = "Ошибка системы"
    ElseIf (LogType And vbExclamation) = vbExclamation Then
        sType = "Проверьте данные"
    ElseIf LogType And vbQuestion Then
        sType = "Вопрос"
    ElseIf LogType And vbInformation Then
        sType = "Информация"
    End If

    If (LogType And vbCritical) And (LogType And vbQuestion) = 0 Then

        Select Case Err.Number
        Case 0, vbObjectError
            If msErrHistory = Empty Then
            ElseIf LogText = Empty Then
            Else
                LogText = LogText & vbCrLf & vbCrLf & msErrHistory
            End If
            Log LogType, LogText
        Case Else
            If LogText = Empty Then
                LogText = gsMes_ErrUnexpected
                Log LogType, Err.Description
            Else
                Log LogType, LogText
            End If
            LogText = LogText & vbCrLf & "Код ошибки: " & ErrId
            If msErrHistory <> Empty Then
                LogText = LogText & vbCrLf & vbCrLf & msErrHistory
            End If
        End Select

        msErrHistory = Empty
        Err.Clear

        SOS
        frmError.Init LogText, SubTitle & " — " & sType

    Else

        ' Вывести сообщение.
        SOS
        Msg = MsgBox(LogText, LogType, SubTitle & " — " & sType)

    End If

End Function

'-------------------------------------------------------------------------------
' Возобновляет перерисовку окна Excel и реакцию Excel на действия пользователя.
'-------------------------------------------------------------------------------
Public Sub SOS()

    ' Включение перерисовки и реакции Excel:
    With Application
        If Not .DisplayAlerts Then
            .DisplayAlerts = True
        End If
        If Not .ScreenUpdating Then
            .ScreenUpdating = True
        End If
    End With
    DoEvents

End Sub

'-------------------------------------------------------------------------------
' Возвращает запись о текущей процедуре.
'-------------------------------------------------------------------------------
Private Function SubId() As String

    Dim zResult As String

    If mcolIds Is Nothing Then
        zResult = "____-_____"
    ElseIf mcolIds.Count = 0 Then
        zResult = "____-_____"
    Else
        zResult = mcolIds(mcolIds.Count)
    End If

    SubId = zResult

End Function

'-------------------------------------------------------------------------------
' Возвращает запись о текущей процедуре.
'-------------------------------------------------------------------------------
Public Function SubTitle() As String

    If mcolTitles Is Nothing Then
    ElseIf mcolTitles.Count = 0 Then
    Else
        SubTitle = mcolTitles(mcolTitles.Count)
    End If

    If SubTitle = Empty Then
        SubTitle = gsProjectTitle & " " & gsBuild
    End If

End Function
