VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "docThis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const mCode As String = "THIS"

'-------------------------------------------------------------------------------
' Обрабатывает открытие документа.
' #NewLID = 100
'-------------------------------------------------------------------------------
Private Sub Document_Open()

    Dim sTplFullName As String
    Dim sXmlFullName As String
    Dim xml As DOMDocument

100 EnterFrame mCode, "DOPEN", "Открытие документа"

    sTplFullName = Path & "\Шаблон.docx"
    sXmlFullName = Path & "\Выгрузка.xml"

    Set xml = New DOMDocument
    xml.Load sXmlFullName

Result_OK:
    GoTo Result_EXIT

Result_BUG:
    Msg vbExclamation, gsMes_ErrUnexpected
    GoTo Result_EXIT

Result_EXIT:
    On Error Resume Next
    Set xml = Nothing
999 LeaveFrame

End Sub
