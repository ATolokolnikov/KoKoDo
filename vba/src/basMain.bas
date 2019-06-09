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
                    cc.Copy
                    docThis.Range(docThis.Range.End, docThis.Range.End).PasteSpecial
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

    If Not Selection.Range.ParentContentControl Is Nothing Then
        Selection.Range.ParentContentControl.Delete False
    Else
        Msg vbExclamation, "В указанном месте блок не найден."
    End If

    LeaveFrame

End Sub

'-------------------------------------------------------------------------------
' Проверяет выделенные блоки или весь документ на соответствие действующему
'   законодательству.
' #NewLID = 100
'-------------------------------------------------------------------------------
Public Sub CheckForActive()

    Dim ccParent As Word.ContentControl
    Dim dicMarkers As Scripting.Dictionary
    Dim htmlAnchor As MSHTML.HTMLAnchorElement
    Dim htmlAct As MSHTML.HTMLDocument
    Dim htmlDoc As MSHTML.HTMLDocument
    Dim http As WinHttp.WinHttpRequest
    Dim rngSent As Word.Range
    Dim rxNd As New clsRegexMatch
    Dim sHtml As String

100 EnterFrame mCode, "CHECK", "Проверка на соответствие законодательству"
    On Error GoTo Result_BUG

    rxNd.SetPattern "&nd=(\d+).+&rdk=(\d+)"

    For Each rngSent In docThis.Sentences

        If Len(rngSent.Text) > 5 Then

            Set rngSent = docThis.Range(rngSent.Start, rngSent.End - 4)

            ' Параметры запроса в поисковую систему.
            Set dicMarkers = New Scripting.Dictionary
            dicMarkers.Add "query", rngSent.Text

            ' Запрос в поисковую систему.
            HttpResponse "GoogleThis", dicMarkers, http

            ' Парсинг ответа в формате HTML.
            ParseHtml http, htmlDoc

            ' Определение ссылок результатов поиска.
            For Each htmlAnchor In htmlDoc.getElementsByTagName("a")
                If Not htmlAnchor.href Like "about:/*" Then
                If Not htmlAnchor.href Like "*google*" Then
                If htmlAnchor.href Like "*gov.ru*" Then
                    rxNd.Execute htmlAnchor.href

                    If rxNd.Count > 0 Then

                        ' Параметры запроса в поисковую систему.
                        Set dicMarkers = New Scripting.Dictionary
                        dicMarkers.Add "nd", rxNd.Subs(0)
                        dicMarkers.Add "rdk", rxNd.Subs(1)

                        ' Запрос в поисковую систему.
                        HttpResponse "OpenThis", dicMarkers, http
                        ParseHtml http, htmlAct

                        sHtml = htmlAct.body.innerText
                        Do While sHtml Like "*  *" Or rngSent.Text Like "*  *"
                            sHtml = Replace(sHtml, "  ", " ")
                            rngSent.Text = Replace(rngSent.Text, "  ", " ")
                        Loop

                        If sHtml Like "*" & rngSent.Text & "*" Then
                            GoTo Next_OK
                        End If

                    End If

                End If
                End If
                End If
            Next htmlAnchor

Next_OK:

            If htmlAnchor Is Nothing Then
    
                If Not rngSent.ParentContentControl Is Nothing Then
                    Set ccParent = rngSent.ParentContentControl
                    If Not ccParent Is Nothing Then
                        ccParent.Delete False
                    End If
                End If
    
                Set ccParent = docThis.ContentControls.Add(wdContentControlRichText, rngSent)
                ccParent.Tag = "Не найдено"
                ccParent.Title = "Не найдено"
                'Options.DefaultHighlightColorIndex = wdRed
                rngSent.HighlightColorIndex = wdRed

            Else
    
                If Not rngSent.ParentContentControl Is Nothing Then
                    Set ccParent = rngSent.ParentContentControl
                    If Not ccParent Is Nothing Then
                        ccParent.Delete False
                    End If
                End If
    
                Set ccParent = docThis.ContentControls.Add(wdContentControlRichText, rngSent)
                ccParent.Tag = htmlAnchor.href
                ccParent.Title = htmlAnchor.href
    
            End If

        End If
        

    Next rngSent

    Set htmlAnchor = Nothing
    Set rxNd = Nothing

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
