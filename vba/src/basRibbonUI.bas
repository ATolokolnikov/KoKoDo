Attribute VB_Name = "basRibbonUI"
Option Explicit

Private Const mCode As String = "RUIB"

'-------------------------------------------------------------------------------
' ������������ ������� ������ � ������ ������������.
' #NewLID = 119
'-------------------------------------------------------------------------------
Public Sub Button_Click(ByVal ctl As Office.IRibbonControl _
                      , Optional ByVal isPressed As Boolean)

100 EnterFrame mCode, "RBTCL", "������� � ������ ������������"
    On Error GoTo Result_BUG

    Select Case ctl.ID

    ' ������ ������ "�������� �����".
    Case "btnBlockAdd"
        Block_Add

    ' ������ ������ "������� �����".
    Case "btnBlockAdd"
        Block_Del

    ' ������ ������ "��������� �� ������������".
    Case "btnCheck"
        CheckForActive

    ' ������ ����������� ������.
    Case Else
        Msg vbCritical, "������ ����������� ������: " & ctl.ID & ". " & vbCrLf _
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
