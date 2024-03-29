VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegexMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
' ����� ��� ������ � ����������� �����������.
' ��������� ���������� (Tools->References):
' - Microsoft VBScript Regular Expressions 5.5
'-------------------------------------------------------------------------------

' -- �������� ���������:
Private Const mCode As String = "REGX" ' ������� ��� ������� ������.

' -- �������� ����������:
Public CurMatch As Match
Public Index As Long
Public Item As MatchCollection
Public Subs As SubMatches
Public RegEx As RegExp

Public Enum rx_flags
    rx_0 = 0
    rx_g = 1
    rx_i = 2
    rx_gi = 3
    rx_m = 4
    rx_gm = 5
    rx_im = 6
    rx_gim = 7
End Enum

'-------------------------------------------------------------------------------
' ���������� ���������� ����������.
'-------------------------------------------------------------------------------
Public Property Get Count() As Long

    Count = Item.Count

End Property

'-------------------------------------------------------------------------------
' ���������� �����, ��������� �� ��������� ������� ��������� ����������.
'-------------------------------------------------------------------------------
Public Function EOF() As Boolean

    If Index >= Count Then
        EOF = True
    ElseIf Index < 0 Then
        EOF = True
    End If

End Function

'-------------------------------------------------------------------------------
' ��������� ����� ���������� � ������ [SourceText].
' ���������� ���������� ��������� ����������.
'-------------------------------------------------------------------------------
Public Function Execute(ByVal SourceText As String) As Long

    Set Item = RegEx.Execute(SourceText)
    Index = 0
    If Count = 0 Then
        Set CurMatch = Nothing
        Set Subs = Nothing
    Else
        Set CurMatch = Item.Item(0)
        Set Subs = CurMatch.SubMatches
    End If
    Execute = Count

End Function

'-------------------------------------------------------------------------------
' ���������� ������� ����� ����������.
'-------------------------------------------------------------------------------
Public Property Get Finish()

    Finish = Start + Length - 1

End Property

'-------------------------------------------------------------------------------
' ���������� ���������� ����� ����������.
'-------------------------------------------------------------------------------
Public Property Get Length() As Long

    Length = CurMatch.Length

End Property

'-------------------------------------------------------------------------------
' ������� � ������� ����������.
'-------------------------------------------------------------------------------
Public Sub MoveFirst()

    Index = 0
    If Count <= 0 Then
        Set CurMatch = Nothing
        Set Subs = Nothing
    Else
        Set CurMatch = Item.Item(Index)
        Set Subs = CurMatch.SubMatches
    End If

End Sub

'-------------------------------------------------------------------------------
' ������� � ���������� ����������.
'-------------------------------------------------------------------------------
Public Sub MoveLast()

    Index = Count - 1
    If Index < 0 Then
        Set CurMatch = Nothing
        Set Subs = Nothing
    Else
        Set CurMatch = Item.Item(Index)
        Set Subs = CurMatch.SubMatches
    End If

End Sub

'-------------------------------------------------------------------------------
' ������� � ���������� ����������.
'-------------------------------------------------------------------------------
Public Sub MoveNext()

    Index = Index + 1
    If Index >= Count Then
        Set CurMatch = Nothing
        Set Subs = Nothing
    Else
        Set CurMatch = Item.Item(Index)
        Set Subs = CurMatch.SubMatches
    End If

End Sub

'-------------------------------------------------------------------------------
' ������������� ������ ������� � ����� ����������� ���������.
' @Flags - ����� � ������� "[g][i][m]".
'    ��������, "gi" �������� ���������� ����� � ������������ �������.
' @NewValue - ������ �������.
'-------------------------------------------------------------------------------
Public Sub SetPattern(ByVal NewValue As String _
                    , Optional ByVal Flags As rx_flags = rx_0)

    If RegEx Is Nothing Then
        Set RegEx = New RegExp
    End If

    RegEx.Pattern = NewValue
    RegEx.Global = Flags And rx_g
    RegEx.IgnoreCase = Flags And rx_i
    RegEx.MultiLine = Flags And rx_m

End Sub

'-------------------------------------------------------------------------------
' ���������� ������� ������ ����������.
'-------------------------------------------------------------------------------
Public Property Get Start() As Long

    Start = CurMatch.FirstIndex

End Property

'-------------------------------------------------------------------------------
' ������������ ����������� �������.
'-------------------------------------------------------------------------------
Private Sub Class_Terminate()

    Set CurMatch = Nothing
    Set Item = Nothing
    Set Subs = Nothing
    Set RegEx = Nothing

End Sub
