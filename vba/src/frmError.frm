VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmError 
   Caption         =   "Непредвиденное исключение"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   OleObjectBlob   =   "frmError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
' Форма вывода сообщения о непредвиденном исключении.
'-------------------------------------------------------------------------------

' -- Закрытые константы:
Private Const mCode As String = "ERRF" ' Кодовое имя данного модуля.

'-------------------------------------------------------------------------------
' Обрабатывает щелчок по кнопке "OK".
'-------------------------------------------------------------------------------
Private Sub cmdOk_Click()

    On Error Resume Next
    Me.Hide

End Sub

'-------------------------------------------------------------------------------
' Обрабатывает щелчок по гиперссылке.
'-------------------------------------------------------------------------------
Private Sub lblHyperlink_Click()

    Dim sLogFullName As String
    Dim t As Double

    On Error Resume Next
    sLogFullName = Path & "\" & gsLog_FileName

    If Dir(sLogFullName) = Empty Then
        SOS
        MsgBox gsMes_LogNotFound, vbCritical
    Else
        Dive
        Shell "explorer /select,""" & sLogFullName & """", vbNormalFocus
        t = Timer
        Do
            DoEvents
        Loop Until Timer > t + 1
        SOS
    End If

End Sub

'-------------------------------------------------------------------------------
' Инициализирует форму.
'-------------------------------------------------------------------------------
Public Sub Init(ErrDesc As String, Optional ErrTitle As String)

    On Error Resume Next
    Me.txtDescription.Value = ErrDesc
    If ErrTitle <> Empty Then
        Me.Caption = ErrTitle
    End If
    Me.Show vbModal
    If Err.Number Then
        Err.Clear
        Me.Hide
        Me.Show vbModeless
    End If

End Sub
