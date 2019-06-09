VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBlocks 
   Caption         =   "Список блоков"
   ClientHeight    =   8265.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15195
   OleObjectBlob   =   "frmBlocks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBlocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
' Экранная форма со списком блоков.
'-------------------------------------------------------------------------------

Private Const mCode As String = "BLKF"

'-------------------------------------------------------------------------------
' Обрабатывает нажатие кнопки "Отмена".
' #NewLID = 101
'-------------------------------------------------------------------------------
Private Sub cmdCancel_Click()

    Me.hidResult.Value = vbCancel
    Me.Hide

End Sub

'-------------------------------------------------------------------------------
' Обрабатывает нажатие кнопки "OK".
' #NewLID = 101
'-------------------------------------------------------------------------------
Private Sub cmdOk_Click()

    Me.hidResult.Value = vbOK
    Me.Hide

End Sub

'-------------------------------------------------------------------------------
' Обрабатывает нажатие кнопки в панели инструментов.
' #NewLID = 101
'-------------------------------------------------------------------------------
Private Sub UserForm_Activate()

    Dim cc As Word.ContentControl
    Dim colCC As Word.ContentControls
    Dim docTpl As Word.Document
    Dim sTplFullName As String

100 EnterFrame mCode, "FACTI", "Открытие списка блоков"

    ' Открыть шаблон документа с блоками.
    sTplFullName = Path & "\Шаблон.docx"
    Set docTpl = Application.Documents.Open(sTplFullName _
                                          , ReadOnly:=True _
                                          , Visible:=True)

    ' Обнаружить коллекцию блоков в шаблоне.
    Set colCC = docTpl.ContentControls

    ' Для списка блоков:
    With lstBlocks

        ' Очистить список блоков (далее заполним его заново).
        .Clear

        ' По каждому блоку из шаблона:
        For Each cc In colCC

            ' Добавить пункт в список блоков.
            .AddItem
            .List(.ListCount - 1, 0) = cc.Tag
            .List(.ListCount - 1, 1) = cc.Title

        Next cc

    End With

Result_OK:
    Me.hidResult.Value = Empty
    GoTo Result_EXIT

Result_BUG:
    Msg vbCritical, gsMes_ErrUnexpected
    GoTo Result_EXIT

Result_EXIT:
    On Error Resume Next
    Set docTpl = Nothing
999 LeaveFrame

End Sub

'-------------------------------------------------------------------------------
' Обрабатывает скрытие экранной формы со списком блоков.
' #NewLID = 101
'-------------------------------------------------------------------------------
Private Sub UserForm_Deactivate()

    Dim docTpl As Word.Document
    Dim sTplFullName As String

100 EnterFrame mCode, "FACTI", "Открытие списка блоков"

    ' Открыть шаблон документа с блоками.
    sTplFullName = Path & "\Шаблон.docx"

    For Each docTpl In Application.Documents
        If docTpl.FullName = sTplFullName Then
            docTpl.Close
            GoTo Result_EXIT
        End If
    Next docTpl

Result_OK:
    GoTo Result_EXIT

Result_BUG:
    Msg vbCritical, gsMes_ErrUnexpected
    GoTo Result_EXIT

Result_EXIT:
    On Error Resume Next
    Set docTpl = Nothing
999 LeaveFrame

End Sub
