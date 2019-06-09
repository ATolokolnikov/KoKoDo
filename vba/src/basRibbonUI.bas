Attribute VB_Name = "basRibbonUI"
Option Explicit

Private Const mCode As String = "RUIB"

'-------------------------------------------------------------------------------
' Обрабатывает нажатие кнопки в панели инструментов.
' #NewLID = 119
'-------------------------------------------------------------------------------
Public Sub Button_Click(ByVal ctl As Office.IRibbonControl _
                      , Optional ByVal isPressed As Boolean)

100 EnterFrame mCode, "RBTCL", "Нажатие в панели инструментов"
    On Error GoTo Result_BUG

    Select Case ctl.ID

    ' Нажата кнопка "Добавить блоки".
    Case "btnBlockAdd"
        Block_Add

    ' Нажата кнопка "Удалить блоки".
    Case "btnBlockAdd"
        Block_Del

    ' Нажата кнопка "Проверить на актуальность".
    Case "btnCheck"
        CheckForActive

    ' Нажата неизвестная кнопка.
    Case Else
        Msg vbCritical, "Нажата неизвестная кнопка: " & ctl.ID & ". " & vbCrLf _
                      & gsMes_BugReport

    End Select

Result_OK:
    GoTo Result_EXIT

Result_BUG:
    Msg vbCritical
    GoTo Result_EXIT

Result_EXIT:
    On Error Resume Next
999 LeaveFrame

End Sub
