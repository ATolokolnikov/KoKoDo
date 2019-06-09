Attribute VB_Name = "basMain"
Option Explicit

'-------------------------------------------------------------------------------
' Стандартный модуль с основными процедурами.
'-------------------------------------------------------------------------------

Public Const mCode As String = "MAIB"

'-------------------------------------------------------------------------------
' Инициирует добавление блока.
' #NewLID = 100
'-------------------------------------------------------------------------------
Public Sub Block_Add()

    Dim cc As Word.ContentControl
    Dim colCCs As Word.ContentControls
    Dim i As Long
    Dim pgCurr As Word.Paragraph
    Dim sBlockId As String

100 EnterFrame mCode, "BLKAD", "Добавление блока"

    ' Если выделен текст:
    If Selection.Type <> wdSelectionIP Then
        Msg vbExclamation, "Текст должен быть не выделен. Установите курсор " _
                         & "ввода текста в необходимую позицию."
        ' Завершить процедуру.
        GoTo Result_EXIT
    End If

    ' Получить абзац, в котором расположен курсор.
    Set pgCurr = Selection.Range.Paragraphs.Item(1)

    ' Для формы со списком блоков:
    With frmBlocks

        ' Отобразить форму выбора блока.
        .Show vbModal

        ' Если пользователь нажал "Отмена":
        If .hidResult.Value <> vbOK Then
            ' Завершить процедуру.
            GoTo Result_EXIT
        End If

        ' Для каждого пункта из списка блоков:
        For i = 0 To .lstBlocks.ListCount - 1

            ' Получить ID блока.
            sBlockId = .lstBlocks.List(i, 0)

            ' Если пункт был выбран пользователем:
            If .lstBlocks.Selected(i) Then
                Set colCCs = docThis.SelectContentControlsByTag(sBlockId)

                ' Для каждого КК, который соответствует пункту.
                For Each cc In colCCs
                    
                Next cc

            ' Иначе (если блок не выбран):
            Else
                

            End If
        Next i

    End With

Result_OK:
    GoTo Result_EXIT

Result_BUG:
    Msg vbCritical, gsMes_ErrUnexpected
    GoTo Result_EXIT

Result_EXIT:
    On Error Resume Next
    Set pgCurr = Nothing
999 LeaveFrame

End Sub

'-------------------------------------------------------------------------------
' Удалеяет блок.
' #NewLID = 100
'-------------------------------------------------------------------------------
Public Sub Block_Del()

    EnterFrame mCode, "BLKDE", "Удаление блока"

    

    LeaveFrame

End Sub

'-------------------------------------------------------------------------------
' Проверяет выделенные блоки или весь документ на соответствие действующему
'   законодательству.
' #NewLID = 100
'-------------------------------------------------------------------------------
Public Sub CheckForActive()

    Dim dicMarkers As Scripting.Dictionary
    Dim htmlAnchor As MSHTML.IHTMLAnchorElement
    Dim htmlDoc As MSHTML.IHTMLDocument
    Dim http As WinHttp.WinHttpRequest

100 EnterFrame mCode, "CHECK", "Проверка на соответствие законодательству"

    ' Параметры запроса в поисковую систему.
    Set dicMarkers = New Scripting.Dictionary
    dicMarkers.Add "query", "Заказчик признается исполнившим эту обязанность"

    ' Запрос в поисковую систему.
    HttpResponse "GoogleThis", dicMarkers, http

    ' Парсинг ответа в формате HTML.
    ParseHtml http, htmlDoc

    ' Определение ссылок результатов поиска.
    'htmlDoc.

Result_OK:
    GoTo Result_EXIT

Result_BUG:
    Msg vbCritical, gsMes_ErrUnexpected
    GoTo Result_EXIT

Result_EXIT:
    On Error Resume Next
    Set dicMarkers = Nothing
999 LeaveFrame

End Sub

'-------------------------------------------------------------------------------
' Возвращает путь к текущему файлу.
' #NewLID = 100
'-------------------------------------------------------------------------------
Public Function Path() As String

    If docThis.Path Like "http*" Then
        Path = Environ("OneDrive") _
             & Replace(Mid(docThis.Path, InStr(25, docThis.Path, "/")), "/", "\")
        Path = Replace(Path, "%20", " ")
    Else
        Path = docThis.Path
    End If

End Function
