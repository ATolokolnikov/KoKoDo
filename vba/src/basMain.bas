Attribute VB_Name = "basMain"
Option Explicit

'-------------------------------------------------------------------------------
' ����������� ������ � ��������� �����������.
'-------------------------------------------------------------------------------

Public Const mCode As String = "MAIB"

'-------------------------------------------------------------------------------
' ���������� ���������� �����.
' #NewLID = 100
'-------------------------------------------------------------------------------
Public Sub Block_Add()

    Dim cc As Word.ContentControl
    Dim colCCs As Word.ContentControls
    Dim i As Long
    Dim pgCurr As Word.Paragraph
    Dim sBlockId As String

100 EnterFrame mCode, "BLKAD", "���������� �����"

    ' ���� ������� �����:
    If Selection.Type <> wdSelectionIP Then
        Msg vbExclamation, "����� ������ ���� �� �������. ���������� ������ " _
                         & "����� ������ � ����������� �������."
        ' ��������� ���������.
        GoTo Result_EXIT
    End If

    ' �������� �����, � ������� ���������� ������.
    Set pgCurr = Selection.Range.Paragraphs.Item(1)

    ' ��� ����� �� ������� ������:
    With frmBlocks

        ' ���������� ����� ������ �����.
        .Show vbModal

        ' ���� ������������ ����� "������":
        If .hidResult.Value <> vbOK Then
            ' ��������� ���������.
            GoTo Result_EXIT
        End If

        ' ��� ������� ������ �� ������ ������:
        For i = 0 To .lstBlocks.ListCount - 1

            ' �������� ID �����.
            sBlockId = .lstBlocks.List(i, 0)

            ' ���� ����� ��� ������ �������������:
            If .lstBlocks.Selected(i) Then
                Set colCCs = docThis.SelectContentControlsByTag(sBlockId)

                ' ��� ������� ��, ������� ������������� ������.
                For Each cc In colCCs
                    
                Next cc

            ' ����� (���� ���� �� ������):
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
' �������� ����.
' #NewLID = 100
'-------------------------------------------------------------------------------
Public Sub Block_Del()

    EnterFrame mCode, "BLKDE", "�������� �����"

    

    LeaveFrame

End Sub

'-------------------------------------------------------------------------------
' ��������� ���������� ����� ��� ���� �������� �� ������������ ������������
'   ����������������.
' #NewLID = 100
'-------------------------------------------------------------------------------
Public Sub CheckForActive()

    Dim dicMarkers As Scripting.Dictionary
    Dim htmlAnchor As MSHTML.IHTMLAnchorElement
    Dim htmlDoc As MSHTML.IHTMLDocument
    Dim http As WinHttp.WinHttpRequest

100 EnterFrame mCode, "CHECK", "�������� �� ������������ ����������������"

    ' ��������� ������� � ��������� �������.
    Set dicMarkers = New Scripting.Dictionary
    dicMarkers.Add "query", "�������� ���������� ����������� ��� �����������"

    ' ������ � ��������� �������.
    HttpResponse "GoogleThis", dicMarkers, http

    ' ������� ������ � ������� HTML.
    ParseHtml http, htmlDoc

    ' ����������� ������ ����������� ������.
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
' ���������� ���� � �������� �����.
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
