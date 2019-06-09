Attribute VB_Name = "basDev"
Option Explicit

'-------------------------------------------------------------------------------
' ����������� ������ � ����������� ������������.
'-------------------------------------------------------------------------------

' -- �������� ���������:
Private Const mCode As String = "DEVB" ' ������� ��� ������� ������.
Private Const msConfFolder As String = "\conf" ' ��� ����� � �������������.
Private Const msSrcFolder As String = "\src" ' ��� ����� � �������� �����.

Private Enum menumLineStatus
    miLine_EOFOrRCommBeg ' ���������� ������ ��� ������ ����������� � ���������.
    miLine_RCommBody ' ���� ����������� � ���������.
    miLine_RCommBodyInclNewLid ' ���� ����������� � ��������� (� �. �. #NewLID).
    miLine_RCommBodyOrRCommEnd ' ���� ����������� � ��������� ��� ��� �����.
    miLine_RCommEnd ' ����� �����������.
    miLine_RName ' ������ ���������.
    miLine_RNameNewLine ' ������ ��������� (����������� �� ����� ������).
    miLine_EmpAfterRName ' ������ ������ ����� ������ ���������.
    miLine_DimOrEmp ' Dim ��� ������ ������ ����� ����.
    miLine_DimOrEnterFrame
    miLine_EnterFrame ' 100 EnterFrame.
    miLine_PhysOrEmp ' ����������/������ ������.
    miLine_Phys ' ���������� ������.
    miLine_MsgLog ' Msg/Log
    miLine_MsgLogGoTo ' Msg/Log/GoTo Result_EXIT.
    miLine_GoToExit ' GoTo Result_EXIT.
    miLine_Handler ' Result_*
    miLine_HandlerEnd ' ������ ������ ����� GoTo Result_EXIT.
    miLine_IgnErrs ' On Error Resume Next.
    miLine_SetToNoth ' Set = * Nothing.
    miLine_LeaveFrame ' 999 LeaveFrame.
    miLine_EmpBeforeREnd ' ������ ������ ����� End.
    miLine_REnd ' End *.
    miLine_EmpAfterR ' ������ ������ ����� ���������.
End Enum

Private Enum menumRType
    miRType_Sub
    miRType_Func
    miRType_Get
    miRType_Let
    miRType_Set
End Enum

'-------------------------------------------------------------------------------
' ���������, ������� ���������� ��������� ������������ ����� �������� � Git.
'-------------------------------------------------------------------------------
Public Sub BeforeCommit()

    Dim sPwd As String
    Dim xlSheet As Excel.Worksheet

    ' �������� ������.
    If sPwd = Empty Then
        sPwd = InputBox("������� ������:")
    End If

    ' ��������� �������� ��� ������� ������� VBA.
    If Not Linter Then
        Exit Sub
    End If

    ' ������� �������� ��� � ����� [msSrcFolder].
    If Not Export Then
        Exit Sub
    End If

    docThis.Save ' ��������� ������� ����.

End Sub

'-------------------------------------------------------------------------------
' ������������ ��� ���������� VBA � ����� [msSrcFolder].
'-------------------------------------------------------------------------------
Public Function Export() As Boolean

    Dim iFile As Integer
    Dim sDir As String
    Dim sFileExt As String
    Dim sFullName As String
    Dim sName As String
    Dim vbeComp As VBComponent

    On Error GoTo Result_BUG

    sDir = Path & msSrcFolder
    sName = Dir(sDir, vbDirectory)
    If sName = Empty Then
        MkDir sDir
    Else
        sName = Dir(sDir & "\*", vbNormal)
        Do Until sName = Empty
            Kill sDir & "\" & sName
            sName = Dir
        Loop
    End If

    ' ��� ������� ���������� VBA �������� �����:
    For Each vbeComp In docThis.VBProject.VBComponents

        ' ���������� ���������� �� ���� ������.
        If vbeComp.Type = vbext_ct_StdModule Then
            sFileExt = ".bas"
        ElseIf vbeComp.Type = vbext_ct_ClassModule Then
            sFileExt = ".cls"
        ElseIf vbeComp.Type = vbext_ct_MSForm Then
            sFileExt = ".frm"
        ElseIf vbeComp.Type = vbext_ct_ActiveXDesigner Then
            sFileExt = ".frx"
        ElseIf vbeComp.Type = vbext_ct_Document Then
            sFileExt = ".cls"
        Else
            Stop
        End If
        ' ���������� ��� ������������ �����.
        sName = vbeComp.Name

        If sName Like "xlSh_z*" Then
            GoTo Marker_NEXT_COMP
        End If

        ' ���������� ���� ������������ �����.
        sFullName = sDir & "\" & sName & sFileExt

        ' ����������, ������� ������.
        vbeComp.Export sFullName

Marker_NEXT_COMP:
    Next vbeComp

Result_OK:
    Export = True
    MsgBox "������ ������������� � ����� " & msSrcFolder, vbInformation
    GoTo Result_EXIT

Result_BUG:
    MsgBox "������ #" & Err.Number & ": " & Err.Description, vbCritical
    GoTo Result_EXIT

Result_EXIT:
    Set vbeComp = Nothing
    If iFile Then
        Close #iFile
    End If

End Function

'-------------------------------------------------------------------------------
' ������������ ��� ����������� � ��������� ���� � ����� [msConfFolder].
'-------------------------------------------------------------------------------
Private Function ExportComments() As Boolean

    Dim iFile As Integer
    Dim sDir As String
    Dim sFileExt As String
    Dim sFullName As String
    Dim sName As String
    Dim vbeCode As CodeModule
    Dim vbeComp As VBComponent

    On Error GoTo Result_BUG

    sDir = Path & msConfFolder
    sName = Dir(sDir, vbDirectory)
    If sName = Empty Then
        MkDir sDir
    Else
        sName = Dir(sDir & "\*", vbNormal)
        Do Until sName = Empty
            Kill sDir & "\" & sName
            sName = Dir
        Loop
    End If

    For Each vbeComp In ActiveWorkbook.VBProject.VBComponents

        ' ���� ������� ��������� VBA ����� ������ *.frx:
        If vbeComp.Type = vbext_ct_ActiveXDesigner Then
            GoTo Marker_NEXT_COMP ' ������������ ���.
        End If

        ' ���������� ���������� �� ���� ������.
        If vbeComp.Type = vbext_ct_StdModule Then
            sFileExt = ".bas"
        ElseIf vbeComp.Type = vbext_ct_ClassModule Then
            sFileExt = ".cls"
        ElseIf vbeComp.Type = vbext_ct_MSForm Then
            sFileExt = ".frm"
        ElseIf vbeComp.Type = vbext_ct_ActiveXDesigner Then
            sFileExt = ".frx"
        ElseIf vbeComp.Type = vbext_ct_Document Then
            sFileExt = ".cls"
        Else
            Stop
        End If

        ' ���������� ��� ������������ �����.
        sName = vbeComp.Name

        ' ���� ��������� �������� ������� ����� ��� �����:
        If sName Like "xl*" Then
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.
        End If

        ' ���������� ���� ������������ �����.
        sFullName = sDir & "\" & sName & sFileExt

        ' ����������, ������� ������.
        vbeComp.Export sFullName

Marker_NEXT_COMP:
    Next vbeComp

Result_OK:
    ExportComments = True
    MsgBox "������ ������������� � ����� " & msSrcFolder, vbInformation
    GoTo Result_EXIT

Result_BUG:
    MsgBox "������ #" & Err.Number & ": " & Err.Description, vbCritical
    GoTo Result_EXIT

Result_EXIT:
    Set vbeComp = Nothing
    If iFile Then
        Close #iFile
    End If

End Function

'-------------------------------------------------------------------------------
' ���������� ������� ��������� ���� �� ��� ������������ ID (SID).
'-------------------------------------------------------------------------------
Public Sub FindSid(Optional ByVal sFindMid As String _
                  , Optional ByVal sFindRid As String _
                  , Optional ByVal sFindLid As String _
                  , Optional ByVal sSelectLn As Long)

    Dim iLine As Long
    Dim iCompLineCount As Long
    Dim iLineLast As Long
    Dim sErrDesc As String
    Dim sFind As String
    Dim sFindMask As String
    Dim sLine As String
    Dim vbeCode As CodeModule
    Dim vbeComp As VBComponent

    If sFindMid <> vbNullString Then
        GoTo Marker_PARAMS_NOT_EMPTY
    End If

    If sFindRid <> Empty Then
        GoTo Marker_PARAMS_NOT_EMPTY
    End If

    If sFindLid <> Empty Then
        GoTo Marker_PARAMS_NOT_EMPTY
    End If

    If sSelectLn <> 0 Then
        GoTo Marker_PARAMS_NOT_EMPTY
    End If

Marker_INPUT:
    sFind = InputBox(sErrDesc, "������� SID", "PARB-PANEV-100")
    If sFind = Empty Then
        GoTo Result_EXIT
    End If

    sFindMid = Mid(sFind, 1, 4)
    sFindRid = Mid(sFind, 6, 5)
    sFindLid = Mid(sFind, 12, 3)

Marker_PARAMS_NOT_EMPTY:

    If Not sFindMid Like "[A-z_0-9][A-z_0-9][A-z_0-9][A-z_0-9]" Then
        sErrDesc = "MID ������ �������� �� 4 �������� A-z, _, 0-9."
        GoTo Marker_INPUT
    End If

    If Not sFindRid Like Empty Then
        If Not sFindRid Like "[A-z_0-9][A-z_0-9][A-z_0-9][A-z_0-9][A-z_0-9]" Then
            sErrDesc = "RID ������ �������� �� 5 �������� A-z, _, 0-9."
            GoTo Marker_INPUT
        End If
        sFindRid = "100 EnterFrame mCode, """ & sFindRid & """*"
    End If

    If Not sFindLid Like Empty Then
        If Not sFindLid Like "###" Then
            sErrDesc = "LID ������ �������� �� 3 ����."
            GoTo Marker_INPUT
        End If
        sFindLid = Val(sFindLid) & " *"
    End If

    sFindMid = "Private Const mCode As String = """ & sFindMid & """"

    For Each vbeComp In ActiveWorkbook.VBProject.VBComponents

        Set vbeCode = vbeComp.CodeModule
        iCompLineCount = vbeCode.CountOfLines

        If Not vbeCode.Find(sFindMid, 1, 1, 20, -1) Then
            GoTo Marker_NEXT_COMP
        Else
            ' ������� ��������� � ���� VBE.
            vbeComp.Activate
            vbeCode.CodePane.SetSelection 1, 1, 1, 1

            If sFindRid Like Empty Then
                GoTo Result_EXIT
            End If

            If Not vbeCode.Find(sFindRid, 1, 1, -1, -1, , , True) Then
                sErrDesc = "RID �� ������."
                GoTo Marker_INPUT
            Else
                ' �������� ������������.
                For iLine = 1 To iCompLineCount
                    sLine = vbeCode.Lines(iLine, 1)
                    If sLine Like sFindRid Then
                        vbeCode.CodePane.SetSelection iLine, Len(sLine) + 1, iLine, Len(sLine) + 1
                        Exit For
                    End If
                Next iLine

                ' ���������� ��������� ������ ������������.
                For iLineLast = iLine + 1 To iCompLineCount
                    sLine = vbeCode.Lines(iLineLast, 1)
                    If Trim(sLine) Like "End Function" Then
                        Exit For
                    ElseIf Trim(sLine) Like "End Property" Then
                        Exit For
                    ElseIf Trim(sLine) Like "End Sub" Then
                        Exit For
                    End If
                Next iLineLast

                If sFindLid Like Empty Then
                    GoTo Result_EXIT
                End If

                If Not vbeCode.Find(sFindLid, iLine, 1, iLineLast, -1, , , True) Then
                    sErrDesc = "LID �� ������."
                    GoTo Marker_INPUT
                Else
                    ' �������� ������.
                    For iLine = iLine To iLineLast
                        sLine = vbeCode.Lines(iLine, 1)
                        If sLine Like sFindLid Then
                            vbeCode.CodePane.SetSelection iLine, 1, iLine, 4
                            Exit For
                        End If
                    Next iLine
                    GoTo Result_EXIT
                End If
            End If
        End If

Marker_NEXT_COMP:
    Next vbeComp

    sErrDesc = "MID �� ������."
    GoTo Marker_INPUT

Result_EXIT:

End Sub

'-------------------------------------------------------------------------------
' ��������� �������� ��� �� ������������ �������� �����������.
' ��� ����������� ������� �������� ���������� �, ����� ������� ������,
'   ��������� ��������������� ������ ��������� ����.
'-------------------------------------------------------------------------------
Private Function Linter() As Boolean

    Dim bNextCaseIsFirst As Boolean
    Dim bNextLineIsCont As Boolean
    Dim dicMids As New Scripting.Dictionary
    Dim dicObjects As New Scripting.Dictionary
    Dim dicRids As New Scripting.Dictionary
    Dim dicRNames As New Scripting.Dictionary
    Dim iCurRType As menumRType
    Dim iHandlerStatus As Long
    Dim iLine As Long
    Dim iCompLineCount As Long
    Dim iTotalLineCount As Long
    Dim iLineNewLid As Long
    Dim iLineStatus As menumLineStatus
    Dim iNewLid As Long
    Dim iObject As Byte
    Dim iRParamStartPos As Byte
    Dim rxDim As New clsRegexMatch
    Dim rxNewLid As New clsRegexMatch
    Dim rxMid As New clsRegexMatch
    Dim rxRid As New clsRegexMatch
    Dim sCommDelim As String
    Dim sCurMid As String
    Dim sCurRType As String
    Dim sLastDim As String
    Dim sLine As String
    Dim sProcName As String
    Dim vbeCode As CodeModule
    Dim vbeComp As VBComponent

    On Error GoTo Result_BUG

    sCommDelim = "'" & String(79, "-")
    rxDim.SetPattern "    Dim ([^ ]+) As (New )?([^\s]+)"
    rxNewLid.SetPattern "' #NewLID = (\d\d\d)"
    rxMid.SetPattern "(Private )(Const )mCode( As String = ""(\S{4})"")"
    rxRid.SetPattern "100 EnterFrame mCode, ""(.....)"", ""(.+)"""

    ' ��� ������� ���������� VBA �������� ����� (����� - ���������):
    For Each vbeComp In ActiveWorkbook.VBProject.VBComponents

        ' ���� ��������� ����� ������ *.frx:
        If vbeComp.Type = vbext_ct_ActiveXDesigner Then
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.
        End If

        ' ���� ����� ��� ����������?
        Select Case vbeComp.Name

        ' ���� "basConstants":
        Case "basConstants"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "basDev":
        Case "basDev"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "basParams":
        Case "basParams"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "basSugar":
        Case "basSugar"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "basTrace":
        Case "basTrace"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "clsApp":
        Case "clsApp"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "clsList":
        Case "clsList"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "clsRegexMatch":
        Case "clsRegexMatch"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "frmError":
        Case "frmError"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "xlAddin_Maps":
        Case "xlAddin_Maps"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "xlSh_HTTP":
        Case "xlSh_HTTP"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "xlSh_RegExp":
        Case "xlSh_RegExp"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        ' ���� "xlSh_zParams":
        Case "xlSh_zParams"
            GoTo Marker_NEXT_COMP ' ������� � ���������� ����������.

        End Select

        ' ������ ��������� ���� ���������� (����� - ���).
        Set vbeCode = vbeComp.CodeModule
        iCompLineCount = vbeCode.CountOfLines
        iTotalLineCount = iTotalLineCount + iCompLineCount
        dicRids.RemoveAll
        dicRNames.RemoveAll

        ' ���� ���������� ����� ���� ������ 6:
        If iCompLineCount < 6 Then
            SelectCode vbeComp.Name, 1 ' �������� ������ ������ ����������.
            MsgBox "��������� ������ 6 ����� ����." _
                  , vbExclamation, vbeComp.Name
            GoTo Result_EXIT
        End If

        sCurMid = Empty

        iLine = 1
        sLine = vbeCode.Lines(iLine, 1)
 
        ' ���������� ������ ����� � ������ ����������.
        Do While sLine = Empty
            Call vbeCode.DeleteLines(iLine, 1)
            sLine = vbeCode.Lines(iLine, 1)
        Loop

        ' �������� �� ������� ��������� Option Explicit.
        If sLine <> "Option Explicit" Then
            vbeCode.InsertLines iLine, "Option Explicit"
'            SelectCode vbeComp.Name, iLine
'            MsgBox "�������� Option Explicit." _
'                  , vbExclamation, vbeComp.Name & ":" & iLine
'            GoTo Result_EXIT
        End If

        ' �������� �� ������� ������ ������ ����� Option Explicit.
        iLine = iLine + 1
        sLine = vbeCode.Lines(iLine, 1)
        If sLine <> Empty Then
            vbeCode.InsertLines iLine, Empty
'            SelectCode vbeComp.Name, iLine
'            MsgBox "��������� ������ ������ ����� Option Explicit." _
'                  , vbExclamation, vbeComp.Name & ":" & iLine
'            GoTo Result_EXIT
        End If

        ' �������� �� ������� ���������� ����������� ����������� ����������.
        iLine = iLine + 1
        sLine = vbeCode.Lines(iLine, 1)
        If sLine <> sCommDelim Then
            vbeCode.InsertLines iLine, sCommDelim
'            SelectCode vbeComp.Name, iLine
'            MsgBox "�������� ��������� ����������� ����������� ����������." _
'                  , vbExclamation, vbeComp.Name & ":" & iLine
'            GoTo Result_EXIT
        End If

        ' �������� �� ������� ����������� ����������.
        iLine = iLine + 1
        sLine = vbeCode.Lines(iLine, 1)
        If Not sLine Like "' *" Then
            vbeCode.InsertLines iLine, "'"
            vbeCode.InsertLines iLine + 1, sCommDelim
            SelectCode vbeComp.Name, iLine
            MsgBox "�������� ����������� ����������." _
                  , vbExclamation, vbeComp.Name & ":" & iLine
            GoTo Result_EXIT
        End If

        ' ������� ����������� � ����������.
        For iLine = iLine + 1 To iCompLineCount
            sLine = vbeCode.Lines(iLine, 1)
            If Not sLine Like "' *" And sLine <> "'" Then
                Exit For
            End If
        Next iLine

        ' �������� �� ������� ��������� ����������� ����������� ����������.
        If sLine <> sCommDelim Then
            vbeCode.InsertLines iLine, sCommDelim
'            SelectCode vbeComp.Name, iLine
'            MsgBox "�������� �������� ����������� ����������� ����������." _
'                  , vbExclamation, vbeComp.Name & ":" & iLine
'            GoTo Result_EXIT
        End If

        ' �������� �� ������� ������ ������ ����� ����������� ����������.
        iLine = iLine + 1
        sLine = vbeCode.Lines(iLine, 1)
        If sLine <> Empty Then
            vbeCode.InsertLines iLine, Empty
'            SelectCode vbeComp.Name, iLine
'            MsgBox "��������� ������ ������ ����� ����������� ����������." _
'                  , vbExclamation, vbeComp.Name & ":" & iLine
'            GoTo Result_EXIT
        End If

        ' �������� �� ������� �������� ����� ����������.
        For iLine = iLine + 1 To iCompLineCount
            sLine = vbeCode.Lines(iLine, 1)
            If Len(sLine) > 80 Then
                SelectCode vbeComp.Name, iLine
                MsgBox "��������� 80 ��� ����� �������� � ������." _
                      , vbExclamation, vbeComp.Name & ":" & iLine
                GoTo Result_EXIT
            End If
            If sLine Like "* Declare *" Then
            ElseIf sLine Like "* Sub *" Then
                GoTo Marker_NO_MCODE_ERROR
            ElseIf sLine Like "* Function *" Then
                GoTo Marker_NO_MCODE_ERROR
            ElseIf sLine Like "* Property Get *" Then
                GoTo Marker_NO_MCODE_ERROR
            ElseIf sLine Like "* Property Let *" Then
                GoTo Marker_NO_MCODE_ERROR
            ElseIf sLine Like "* Property Set *" Then
                GoTo Marker_NO_MCODE_ERROR
            ElseIf sLine = sCommDelim Then
                Exit For
            End If
            If rxMid.Execute(sLine) Then
                If rxMid.Subs(0) = Empty Then
                    GoTo Marker_NO_MCODE_ERROR
                ElseIf rxMid.Subs(1) = Empty Then
                    GoTo Marker_NO_MCODE_ERROR
                ElseIf Not rxMid.Subs(2) Like " As String *" Then
                    GoTo Marker_NO_MCODE_ERROR
                End If
                sCurMid = rxMid.Subs(3)
            End If
        Next iLine

        ' �������� �� ������� �������� ����� ����������.
        If sCurMid = Empty Then
Marker_NO_MCODE_ERROR:
            SelectCode vbeComp.Name, iLine
            MsgBox "��������� Private Const mCode As String = ""????""." _
                  , vbExclamation, vbeComp.Name & ":" & iLine
            GoTo Result_EXIT
        End If

        ' �������� �� ������������ �������� ����� ����������.
        If dicMids.Exists(sCurMid) Then
            SelectCode vbeComp.Name, iLine
            MsgBox "������� ��� ���������� (mCode) �� ���������." _
                  , vbExclamation, vbeComp.Name & ":" & iLine
            GoTo Result_EXIT
        End If
        dicMids.Add sCurMid, Empty ' �������� ������� ��� ����������.

        iLine = iLine - 1
        iLineStatus = miLine_EOFOrRCommBeg
        For iLine = iLine + 1 To iCompLineCount

            sLine = vbeCode.Lines(iLine, 1)
            If Len(sLine) > 80 Then
                SelectCode vbeComp.Name, iLine
                MsgBox "��������� 80 ��� ����� �������� � ������." _
                      , vbExclamation, vbeComp.Name & ":" & iLine
                GoTo Result_EXIT
            End If

            Select Case iLineStatus

            ' ������� ������ ����������� � ���������.
            Case miLine_EOFOrRCommBeg
                If Not sLine Like sCommDelim Then
                    vbeCode.InsertLines iLine, sCommDelim
'                    SelectCode vbeComp.Name, iLine
'                    MsgBox "��������� ������ ����������� � ���������." _
'                          , vbExclamation, vbeComp.Name & ":" & iLine
'                    GoTo Result_EXIT
                End If
                iLineStatus = miLine_RCommBodyInclNewLid

            ' ������� ���� ����������� � ���������, � �. �. � #NewLID.
            Case miLine_RCommBodyInclNewLid
                If Not sLine Like "'*" Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������� ����������� � #NewLID." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If
                If sLine Like "' [#]NewLID = ###" Then
                    iLineStatus = miLine_RCommEnd
                    If rxNewLid.Execute(sLine) Then
                        iNewLid = rxNewLid.Subs(0)
                        iLineNewLid = iLine
                    Else
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� ���������� #NewLID." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                End If

            ' ������� ����� ����������� � ���������.
            Case miLine_RCommEnd
                If Not sLine Like sCommDelim Then
                    vbeCode.InsertLines iLine, sCommDelim
'                    SelectCode vbeComp.Name, iLine
'                    MsgBox "�������� ����� ����������� � ���������." _
'                          , vbExclamation, vbeComp.Name & ":" & iLine
'                    GoTo Result_EXIT
                End If
                iLineStatus = miLine_RName

            ' ������� ��������� ���������.
            Case miLine_RName
                If sLine Like "* Sub *" Then
                    iCurRType = miRType_Sub
                    sCurRType = "Sub"
                ElseIf sLine Like "* Function *" Then
                    iCurRType = miRType_Func
                    sCurRType = "Function"
                ElseIf sLine Like "* Property Get *" Then
                    iCurRType = miRType_Get
                    sCurRType = "Property"
                ElseIf sLine Like "* Property Let *" Then
                    iCurRType = miRType_Let
                    sCurRType = "Property"
                ElseIf sLine Like "* Property Set *" Then
                    iCurRType = miRType_Set
                    sCurRType = "Property"
                Else
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������� ��������� ���������." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If
                iRParamStartPos = InStr(sLine, "(") - 2

Marker_CHECK_RNAME:
                If sLine Like "*'*" Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������� ��������� ���������." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If
                If sLine Like "*)" Then
                    iLineStatus = miLine_EmpAfterRName
                ElseIf sLine Like "*) As *" Then
                    iLineStatus = miLine_EmpAfterRName
                ElseIf Trim(sLine) Like "* _" Then
                    iLineStatus = miLine_RNameNewLine
                Else
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������������� ������ ����������� ���������." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If

            ' ������� ���������� ����������� ���������.
            Case miLine_RNameNewLine
                If sLine Like Space(iRParamStartPos) & ", *" Then
                ElseIf sLine Like Space(iRParamStartPos) & ")*" Then
                Else
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������������� ������ ����������� ���������. " _
                         & "��������� ����������� ������ � ����� ������ ���� " _
                         & "��� ���������� ��������� ������ ��� ����������, " _
                         & "� ������� � ������ ����������� ������." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If
                GoTo Marker_CHECK_RNAME

            ' ������� ������ ������ ����� ������ ���������.
            Case miLine_EmpAfterRName
                If Not sLine Like Empty Then
                    vbeCode.InsertLines iLine, Empty
'                    SelectCode vbeComp.Name, iLine
'                    MsgBox "��������� ������ ������ ����� ������ ���������." _
'                          , vbExclamation, vbeComp.Name & ":" & iLine
'                    GoTo Result_EXIT
                End If
                iLineStatus = miLine_DimOrEnterFrame
                dicObjects.RemoveAll
                sLastDim = Empty

            ' ������� Dim ��� EnterFrame.
            Case miLine_DimOrEnterFrame
                If sLine Like "    Dim *" Then

Marker_GET_DIM:
                    If rxDim.Execute(sLine) = 0 Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� ���������� Dim." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If

                    Select Case rxDim.Subs(2)
                    Case "Boolean"
                    Case "Byte"
                    Case "Currency"
                    Case "Date"
                    Case "Decimal"
                    Case "Double"
                    Case "Integer"
                    Case "Long"
                    Case "LongLong"
                    Case "LongPtr"
                    Case "Single"
                    Case "String"
                    Case "Variant"
                    Case "VbMsgBoxResult"
                    Case Else
                        If rxDim.Subs(2) Like "menum*" Then
                        ElseIf rxDim.Subs(2) Like "penum*" Then
                        Else
                            dicObjects.Add rxDim.Subs(0), rxDim.Subs(2)
                        End If
                    End Select

                    If sLastDim > rxDim.Subs(0) Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� ���������� ������� Dim." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                    sLastDim = rxDim.Subs(0)
                    iLineStatus = miLine_DimOrEmp

                ElseIf sLine Like "100 EnterFrame*" Then

                    ' �������� ������ EnterFrame �� ������������.
                    If rxRid.Execute(sLine) = 0 Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� ���������� ����� EnterFrame." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If

                    ' �������� �������� ����� ��������� �� ������������.
                    If dicRids.Exists(rxRid.Subs(0)) Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "��� ��������� ����������." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                    dicRids.Add rxRid.Subs(0), Empty

                    ' �������� �������� ��������� �� ������������.
                    If dicRNames.Exists(rxRid.Subs(1)) Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� ��������� �����������." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                    dicRNames.Add rxRid.Subs(1), Empty

                    ' �������� �� ������� On Error ����� ������ EnterFrame.
                    iLine = iLine + 1
                    sLine = vbeCode.Lines(iLine, 1)
                    If Not sLine Like "    On Error *" Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� On Error." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                    iLineStatus = miLine_PhysOrEmp

                ElseIf Trim(sLine) Like "[#]*" Then
                Else
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������� Dim ��� 100 EnterFrame mCode." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If

            ' ������� ������ ������ ��� Dim.
            Case miLine_DimOrEmp
                If sLine Like Empty Then
                    iLineStatus = miLine_DimOrEnterFrame
                ElseIf sLine Like "    Dim *" Then
                    GoTo Marker_GET_DIM
                ElseIf Trim(sLine) Like "[#]*" Then
                Else
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������� Dim ��� ������ ������." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If

            ' ������� ���������� ��� ������ ������.
            Case miLine_PhysOrEmp, miLine_Phys
                If sLine Like Empty Then
                    If iLineStatus <> miLine_Phys Then
                        iLineStatus = miLine_Phys
                    Else
                        SelectCode vbeComp.Name, iLine
                        MsgBox "��������� ���������� ������." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                ElseIf Trim(sLine) Like "Dim *" Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "Dim ��������� � ������ ���������." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                ElseIf Trim(sLine) Like "[#]*" Then
                ElseIf Trim(sLine) Like "End Sub" Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������� ������ Result_OK." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                ElseIf Trim(sLine) Like "End Function" Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������� ������ Result_OK." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                ElseIf Trim(sLine) Like "End Property" Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "�������� ������ Result_OK." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                ElseIf sLine Like "Result_*:" Then
                    GoTo Marker_CHECK_HANDLER
                ElseIf Not sLine Like "### *" Then
                    If Trim(sLine) Like "'*" Then
                    ElseIf Trim(sLine) Like "On Error *" Then
                    ElseIf Trim(sLine) Like "GoTo *" Then
                    ElseIf Trim(sLine) Like "End *" Then
                    ElseIf Trim(sLine) Like "Msg *" Then
                    ElseIf Trim(sLine) Like "Log *" Then
                    ElseIf bNextLineIsCont Then
                    ElseIf Trim(sLine) Like "Select Case *" Then
                        bNextCaseIsFirst = True
                        GoTo Marker_SET_LID
                    ElseIf Trim(sLine) Like "Case *" Then
                        bNextCaseIsFirst = False
                    ElseIf Trim(sLine) Like "*:" Then
                    ElseIf Trim(sLine) Like "*: '*" Then
                    ElseIf Trim(sLine) Like "#" Then
                        GoTo Marker_CLEAR_LINE
                    ElseIf Trim(sLine) Like "##" Then
                        GoTo Marker_CLEAR_LINE
                    ElseIf Trim(sLine) Like "###" Then
                        GoTo Marker_CLEAR_LINE
                    ElseIf Trim(sLine) Like "####" Then
                        GoTo Marker_CLEAR_LINE
                    Else
Marker_SET_LID:
                        Mid(sLine, 1, 3) = iNewLid
                        iNewLid = iNewLid + 1
Marker_SET_NEWLID:
                        If iNewLid > 998 Then
                            SelectCode vbeComp.Name, iLine
                            MsgBox "LID �������� 999." _
                                  , vbExclamation, vbeComp.Name & ":" & iLine
                            GoTo Result_EXIT
                        End If
                        vbeCode.ReplaceLine iLine, sLine
                        vbeCode.ReplaceLine iLineNewLid, "' #NewLID = " & iNewLid
                    End If
                ElseIf sLine Like "*Select Case *" Then
                    bNextCaseIsFirst = True
                ElseIf sLine Like "### *" Then
                    If iNewLid <= Val(Mid(sLine, 1, 3)) Then
                        iNewLid = Val(Mid(sLine, 1, 3)) + 1
                        GoTo Marker_SET_NEWLID
                    End If
                ElseIf sLine Like "#" Then
Marker_CLEAR_LINE:
                    If iLineStatus = miLine_PhysOrEmp Then
                        vbeCode.ReplaceLine iLine, ""
                        iLineStatus = miLine_Phys
                    Else
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�� ��������� ������ ������ � LID." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                ElseIf sLine Like "##" Then
                    GoTo Marker_CLEAR_LINE
                ElseIf sLine Like "###" Then
                    GoTo Marker_CLEAR_LINE
                ElseIf sLine Like "####" Then
                    GoTo Marker_CLEAR_LINE
                End If

                iLineStatus = miLine_PhysOrEmp
                If sLine Like "* _" And Not sLine Like "*'*" Then
                    bNextLineIsCont = True
                Else
                    bNextLineIsCont = False
                End If

            ' ������� ���������� ������ ��� GoTo Result_EXIT.
            Case miLine_GoToExit
                If Trim(sLine) Like "GoTo Result_EXIT" Then
                    iLineStatus = miLine_HandlerEnd
                ElseIf Trim(sLine) Like "GoTo Result_OK" Then
                    iLineStatus = miLine_HandlerEnd
                ElseIf sLine Like Empty Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "��������� ���������� ������ ��� GoTo Result_EXIT." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If

            ' ������� Result_*.
            Case miLine_Handler, miLine_HandlerEnd
Marker_CHECK_HANDLER:
                Select Case sLine
                Case "Result_OK:"
                    iHandlerStatus = 1
                    iLineStatus = miLine_GoToExit
                Case "Result_BUG:"
                    If iHandlerStatus <> 1 Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� ������ Result_OK." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                    'iHandlerStatus = 2
                    iLineStatus = miLine_GoToExit
                Case "Result_EXIT:"
                    If iHandlerStatus <> 1 Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� ������ Result_BUG." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                    iHandlerStatus = 3
                    iLineStatus = miLine_IgnErrs
                Case Empty
                    If iLineStatus = miLine_Handler Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� ������ Result_EXIT." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    Else
                        iLineStatus = miLine_Handler
                    End If
                Case Else
                    If iHandlerStatus <> 1 Then
                        SelectCode vbeComp.Name, iLine
                        MsgBox "�������� ������ Result_BUG." _
                              , vbExclamation, vbeComp.Name & ":" & iLine
                        GoTo Result_EXIT
                    End If
                    iLineStatus = miLine_GoToExit
                End Select

            ' ������� On Error Resume Next.
            Case miLine_IgnErrs
                If Trim(sLine) <> "On Error Resume Next" Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "��������� On Error Resume Next." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If
                iObject = 0
                If dicObjects.Count > 0 Then
                    iLineStatus = miLine_SetToNoth
                Else
                    iLineStatus = miLine_LeaveFrame
                End If

            ' ������� Set * = Nothing.
            Case miLine_SetToNoth
                If sLine <> "    Set " & dicObjects.Keys()(iObject) & " = Nothing" Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "��������� Set " & dicObjects.Keys()(iObject) & " = Nothing" _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                ElseIf sLine Like "End *" Then
                    iLineStatus = miLine_EmpAfterR
                Else
                    iObject = iObject + 1
                    iLineStatus = miLine_SetToNoth
                End If
                If iObject >= dicObjects.Count Then
                    iLineStatus = miLine_LeaveFrame
                End If

            ' ������� 999 LeaveFrame.
            Case miLine_LeaveFrame
                If sLine <> "999 LeaveFrame" Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "��������� 999 LeaveFrame." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If
                iLineStatus = miLine_EmpBeforeREnd

            ' ������� ������ ������ ����� End.
            Case miLine_EmpBeforeREnd
                If sLine <> Empty Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "��������� ������ ������ ����� End." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If
                iLineStatus = miLine_REnd

            ' ������� End ���������.
            Case miLine_REnd
                If sLine <> "End " & sCurRType Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "��������� " & "End " & sCurRType & "." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If
                iLineStatus = miLine_EmpAfterR

            ' ������� ������ ������ ����� ���������.
            Case miLine_EmpAfterR
                If sLine <> Empty Then
                    SelectCode vbeComp.Name, iLine
                    MsgBox "��������� ������ ������." _
                          , vbExclamation, vbeComp.Name & ":" & iLine
                    GoTo Result_EXIT
                End If
                iLineStatus = miLine_EOFOrRCommBeg

            Case Else
                SelectCode vbeComp.Name, iLine
                MsgBox "����������� ������ ������." _
                      , vbExclamation, vbeComp.Name & ":" & iLine
                GoTo Result_EXIT

            End Select

        Next iLine

        If iLineStatus <> miLine_EmpAfterR Then
            SelectCode vbeComp.Name, iLine - 1
            MsgBox "�������� ����� ����������."
            GoTo Result_EXIT
        End If

Marker_NEXT_COMP:
    Next vbeComp

    Linter = True
    MsgBox "�������� ��� ������� ��������. ����� ���� ����������: " _
         & iTotalLineCount, vbInformation
    GoTo Result_EXIT

Result_BUG:
    Msg vbCritical, vbeComp.Name & ":" & iLine
    GoTo Result_EXIT

Result_EXIT:
    Debug.Print "����� ���� ����������: " & iTotalLineCount
    Set vbeCode = Nothing
    Set vbeComp = Nothing

End Function

'-------------------------------------------------------------------------------
' ���������� �������� ������ ��������� ����.
'-------------------------------------------------------------------------------
Private Sub SelectCode(Optional ByVal sModuleName As String _
                     , Optional ByVal iSelectLn As Long = 0 _
                     , Optional ByVal iSelectA As Long = 1 _
                     , Optional ByVal iSelectB As Long = 1)

    Dim iLine As Long
    Dim iCompLineCount As Long
    Dim iLineLast As Long
    Dim sErrDesc As String
    Dim sFind As String
    Dim sFindMask As String
    Dim sLine As String
    Dim vbeCode As CodeModule
    Dim vbeComp As VBComponent

    For Each vbeComp In ActiveWorkbook.VBProject.VBComponents

        Set vbeCode = vbeComp.CodeModule
        iCompLineCount = vbeCode.CountOfLines

        If vbeComp.Name = sModuleName Then

            vbeComp.Activate
            vbeCode.CodePane.SetSelection 1, 1, 1, 1

            If iSelectLn = 0 Then
                GoTo Result_EXIT
            End If

            If iSelectLn > iCompLineCount Then
                Msg vbExclamation, "������ �� �������."
                GoTo Result_EXIT
            End If

            iLine = iSelectLn
            vbeCode.CodePane.SetSelection iLine, iSelectA, iLine, iSelectB
            GoTo Result_EXIT

        End If

    Next vbeComp

    Msg vbExclamation, "������ �� ������."
    GoTo Result_EXIT

Result_EXIT:

End Sub

'-------------------------------------------------------------------------------
' ���������� ������ ��� ������� ������������ ����.
'-------------------------------------------------------------------------------
Sub ShowFaceIds()

  Dim xBar As CommandBar
  Dim xBarPop As CommandBarPopup
  Dim bCreatedNew As Boolean
  Dim n As Integer, m As Integer
  Dim k As Integer
  Const APP_NAME = "FaceIDs (Browser)"
  Const ICON_SET = 30 ' The number of icons to be displayed in a set.

  On Error Resume Next
  ' Try to get a reference to the 'FaceID Browser' toolbar if it exists and delete it:
  Set xBar = Application.CommandBars(APP_NAME)
  On Error GoTo 0
  If Not xBar Is Nothing Then
    xBar.Delete
    Set xBar = Nothing
  End If

  Set xBar = CommandBars.Add(Name:=APP_NAME, Temporary:=True) ', Position:=msoBarLeft
  With xBar
    .Visible = True
    '.Width = 80
    For k = 0 To 4 ' 5 dropdowns, each for about 1000 FaceIDs
      Set xBarPop = .Controls.Add(Type:=msoControlPopup) ', Before:=1
      With xBarPop
        .BeginGroup = True
        If k = 0 Then
          .Caption = "Face IDs " & 1 + 1000 * k & " ... "
        Else
          .Caption = 1 + 1000 * k & " ... "
        End If
        n = 1
        Do
          With .Controls.Add(Type:=msoControlPopup) '34 items * 30 items = 1020 faceIDs
            .Caption = 1000 * k + n & " ... " & 1000 * k + n + ICON_SET - 1
            For m = 0 To ICON_SET - 1
              With .Controls.Add(Type:=msoControlButton) '
                .Caption = "ID=" & 1000 * k + n + m
                .FaceId = 1000 * k + n + m
              End With
            Next m
          End With
          n = n + ICON_SET
        Loop While n < 1000 ' or 1020, some overlapp
      End With
    Next k
  End With 'xBar

End Sub
