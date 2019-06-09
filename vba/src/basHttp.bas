Attribute VB_Name = "basHttp"
Option Explicit

Public rxHtmlDoc As clsRegexMatch

'-------------------------------------------------------------------------------
' ���������� HTTP-������.
'-------------------------------------------------------------------------------
Public Function HttpResponse( _
       ListName As String _
       , Optional DictMarkers As Scripting.Dictionary _
        , Optional http As WinHttp.WinHttpRequest _
         ) As Integer

    Dim dicHeaders As Scripting.Dictionary
    Dim iHeader As Long
    Dim iMarker As Long
    Dim lrRow As Excel.ListRow
    Dim lstHeaders As Excel.ListObject
    Dim rCookie As Excel.Range
    Dim sCookieOld As String
    Dim sCookieSet As String
    Dim sHeaderName As String
    Dim sHeaderValue As String

100 EnterFrame mCode, "HTRES", "������ HTTP � �������"

    On Error GoTo Handle_BUG

    ' ���������� ��������� � �������� ������.
    frmLoading.lblWait.Caption = "����������, ���������..."
    frmLoading.Show vbModeless

    ' -- ������������� ��������.
    Set dicHeaders = New Scripting.Dictionary

    If http Is Nothing Then
        Set http = New WinHttpRequest
    End If

    ' ���� ���� � HTTP �� ������:
    If xlSh_HTTP Is Nothing Then
        GoTo Handle_EXIT
    End If

    ' �������� ���� � HTTP.
    Set lstHeaders = xlSh_HTTP.ListObjects(ListName)

    ' ������������ ������� ����������:
    For Each lrRow In lstHeaders.ListRows
        If Not lrRow.Range.EntireRow.Hidden Then

            ' ��������� ����� � �������� ���������.
            sHeaderName = lrRow.Range(1, 1).Value
            sHeaderValue = lrRow.Range(1, 2).Value

            ' �������������� �������� ���������.
            If Not DictMarkers Is Nothing Then
                For iMarker = 0 To DictMarkers.Count - 1
                    sHeaderValue = Replace(sHeaderValue _
                                         , "[" & DictMarkers.Keys(iMarker) & "]" _
                                         , DictMarkers.Items(iMarker))
                Next iMarker
            End If

            ' ���� �������� ������������� �������:
            If sHeaderValue Like "*[[]*" Then
                Stop ' ������������!
            End If

            ' ��������� ���������� �������� ���������.
            dicHeaders(sHeaderName) = sHeaderValue

        End If
    Next lrRow

    ' ������������ HTTP-�������.
    http.Open dicHeaders("������"), dicHeaders("URL"), True
    For iHeader = 0 To dicHeaders.Count - 1
        If Not dicHeaders.Keys(iHeader) Like "Cookie" Then
        If Not dicHeaders.Keys(iHeader) Like "URL" Then
        If Not dicHeaders.Keys(iHeader) Like "������" Then
        If Not dicHeaders.Keys(iHeader) Like vbNullString Then
            sHeaderName = dicHeaders.Keys(iHeader)
            sHeaderValue = dicHeaders.Items(iHeader)
            http.SetRequestHeader sHeaderName, Replace(Replace(sHeaderValue, vbCr, ""), vbLf, "")
        End If
        End If
        End If
        End If
    Next

    ' ��������� �������� Cookie.
    Set rCookie = xlSh_HTTP.Range("Cookie")
    sCookieOld = CStr(rCookie.Value)

    If sCookieOld <> vbNullString Then
        http.SetRequestHeader "Cookie", sCookieOld
    End If

    ' �������� HTTP-�������.
    http.Send dicHeaders(vbNullString)
    http.WaitForResponse
    HttpResponse = http.Status

    ' ��������� Cookie.
    On Error Resume Next
    If http.getAllResponseHeaders Like "*Set-Cookie*" Then
        sCookieSet = http.GetResponseHeader("Set-Cookie")
        SetCookie sCookieSet
    End If

    If False Then
        Dim iFile As Integer
        iFile = FreeFile
        Open Path & "\debug.html" For Output As #iFile
        Print #iFile, , http.ResponseText;
        Close #iFile
    End If

    GoTo Handle_EXIT

Handle_BUG:     ' ��������� ������ VBA.
    MsgBox "���������� � ������������. " _
         & vbCrLf _
         & vbCrLf & "������ VBA #" & Err.Number _
         & ": " & Err.Description _
           , vbCritical _
            , "�������� �� �����. Http-������"
    GoTo Handle_EXIT

Handle_EXIT:     ' ���������� �������.
    On Error Resume Next
    Set dicHeaders = Nothing
    Set lrRow = Nothing
    Set lstHeaders = Nothing
    frmLoading.Hide
999 LeaveFrame

End Function

'-------------------------------------------------------------------------------
' ���������� ������, � ������� ���� "\u0000" �������� �� ���������������
' ������� �������.
'-------------------------------------------------------------------------------
Public Function JsonUcodeGet(ByVal String_IN As String _
                           , ByVal IsEscaped As Boolean) As String

    Dim iPrevFinish As Long
    Dim rxJsonUcode As clsRegexMatch

    Set rxJsonUcode = New clsRegexMatch

    With rxJsonUcode

        If IsEscaped Then
            .SetPattern "rxJsonUcodeEscaped", rx_gi
            String_IN = Replace(String_IN, "\\\\", "\")
        Else
            .SetPattern "rxJsonUcode", rx_gi
            String_IN = Replace(String_IN, "\\", "\")
        End If

        If .Execute(String_IN) Then
            .MoveFirst
            Do Until .EOF
                If iPrevFinish Then
                    JsonUcodeGet = JsonUcodeGet & Mid(String_IN, iPrevFinish + 2, .Start - iPrevFinish - 2)
                Else
                    JsonUcodeGet = Left(String_IN, .Start - 1)
                End If
                JsonUcodeGet = JsonUcodeGet & ChrW(Val("&h" & .Subs(0)))
                iPrevFinish = .Finish
                .MoveNext
            Loop
            JsonUcodeGet = JsonUcodeGet & Mid(String_IN, iPrevFinish + 2)
            GoTo Handle_EXIT
        Else
            JsonUcodeGet = String_IN
        End If

    End With

Handle_EXIT:
    Set rxJsonUcode = Nothing

End Function

'-------------------------------------------------------------------------------
' ������������� ��������� ������ HTML.
'-------------------------------------------------------------------------------
Public Sub ParseHtml(ByRef http As WinHttp.WinHttpRequest _
                   , ByRef htmlDoc As MSHTML.HTMLDocument)

    Set htmlDoc = New HTMLDocument
    htmlDoc.body.innerHTML = http.ResponseText

End Sub

'-------------------------------------------------------------------------------
' ������������� ��������� ������ HTML.
'-------------------------------------------------------------------------------
Public Sub SetCookie(sCookieSet As String)

    Dim rCookie As Excel.Range
    Dim rxCookieOld As clsRegexMatch
    Dim rxCookieSet As clsRegexMatch
    Dim sCookieOld As String

    Set rCookie = xlSh_HTTP.Range("Cookie")

    Set rxCookieOld = New clsRegexMatch
    rxCookieOld.SetPattern "rxCookie", rx_g

    Set rxCookieSet = New clsRegexMatch
    rxCookieSet.SetPattern "rxCookie", rx_g

    sCookieOld = rCookie.Value

    Call rxCookieOld.Execute(sCookieOld)

    If rxCookieSet.Execute(sCookieSet) Then

        rxCookieOld.MoveFirst
        Do Until rxCookieOld.EOF

            rxCookieSet.MoveFirst
            Do Until rxCookieSet.EOF
                If rxCookieOld.Subs(0) = rxCookieSet.Subs(0) Then
                    Exit Do
                End If
                rxCookieSet.MoveNext
            Loop

            If rxCookieSet.EOF Then
                If rxCookieOld.Subs.Count > 1 Then
                    If rxCookieOld.Subs(1) <> vbNullString Then
                        sCookieSet = sCookieSet & ";" & rxCookieOld.Subs(0) & "=" _
                                                      & rxCookieOld.Subs(1)
                    Else
                        sCookieSet = sCookieSet & ";" & rxCookieOld.Subs(0)
                    End If
                Else
                    sCookieSet = sCookieSet & ";" & rxCookieOld.Subs(0)
                End If
            End If
            rxCookieOld.MoveNext

        Loop

        rCookie.Value = sCookieSet

    End If

End Sub

'-------------------------------------------------------------------------------
' ���������� ������ �� ���� Excel � ����������� HTTP-��������.
'-------------------------------------------------------------------------------
Public Function xlSh_HTTP() As Excel.Worksheet

    Dim appXL As Excel.Application
    Dim bk_HTTP As Excel.Workbook
    Dim sHttpPath As String

100 EnterFrame mCode, "SHTTP", "��������� ���������� HTTP"

    ' ����� �������� ���� � ����������� HTTP.
    On Error Resume Next
    Set appXL = GetObject(, "Excel.Application")
    On Error GoTo Result_EXIT

    sHttpPath = Path & "\HTTP.xlsx"

    ' ���� �� ������ �������� ���� "HTTP.xlsx":
    If appXL Is Nothing Then
        Set appXL = Excel.Application
        Set bk_HTTP = appXL.Workbooks.Open(sHttpPath)
        Set xlSh_HTTP = bk_HTTP.Sheets("HTTP")
        appXL.Visible = True
        GoTo Result_EXIT
    End If

    ' ����� ���������� HTTP ����� �������� ������ Excel.
    For Each bk_HTTP In appXL.Workbooks
        If bk_HTTP.FullName = Path & "\HTTP.xlsx" Then
            Exit For
        End If
    Next bk_HTTP

    ' ���� ���������� HTTP ������ �� ������:
    If bk_HTTP Is Nothing Then
        ' �������.
        Set bk_HTTP = appXL.Workbooks.Open("D:\OneDrive\Desktop\�������\4. ����������" _
                                        & "\HTTP.xlsx")
    End If

    ' ��������� ���� � ����������� HTTP-��������.
    Set xlSh_HTTP = bk_HTTP.Sheets("HTTP")

Result_OK:
    GoTo Result_EXIT

Result_BUG:
    Msg vbCritical, gsMes_ErrUnexpected
    GoTo Result_EXIT

Result_EXIT:
    On Error Resume Next
    Set appXL = Nothing
    Set bk_HTTP = Nothing
999 LeaveFrame

End Function
