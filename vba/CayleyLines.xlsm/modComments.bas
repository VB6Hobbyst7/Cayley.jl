Attribute VB_Name = "modComments"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : ApplyComments
' Author    : Philip Swannell
' Date      : 31-Oct-2016
' Purpose   : Attached to button on sheet CommentEditor
'---------------------------------------------------------------------------------------
Sub ApplyComments()
          Dim anyChangesMade As Boolean
          Dim i As Long
          Dim NewComment As Variant
          Dim OldComment As String
          Dim Prompt As String
          Dim SourceData As Range
          Dim SPH As clsSheetProtectionHandler
          Dim TargetCell As Range
          Dim TargetSheet As Worksheet

1         Application.ScreenUpdating = False
2         shComments.Calculate

3         On Error GoTo ErrHandler
4         Set SourceData = RangeFromSheet(shComments, "TheDataWithHeaders")

5         For i = 2 To SourceData.Rows.Count
6             Set TargetSheet = ThisWorkbook.Worksheets(SourceData.Cells(i, 1).Value)
7             Set TargetCell = RangeFromSheet(TargetSheet, SourceData.Cells(i, 3).Value)
8             Set TargetCell = TargetCell.Cells(SourceData.Cells(i, 4).Value, SourceData.Cells(i, 5).Value)

9             If TargetCell.Value <> SourceData.Cells(i, 2).Value Then Throw "Assertion failed at row " + CStr(i) + vbLf + " Expecting " + _
                 TargetCell.Parent.Name + "!" + TargetCell.Address + " to contain '" + SourceData.Cells(i, 2).Value + "' but it contains '" + TargetCell.Value + "'"

10            If SourceData.Cells(i, 1).Value <> SourceData.Cells(i - 1, 1).Value Then
11                Set SPH = CreateSheetProtectionHandler(TargetSheet)
12            End If

13            OldComment = ""
14            On Error Resume Next
15            OldComment = TargetCell.Comment.Text
16            On Error GoTo ErrHandler

17            NewComment = SourceData.Cells(i, 6).Value
18            If sArrayIsNonTrivialText(NewComment)(1, 1) Then
19                OldComment = ""
20                On Error Resume Next
21                OldComment = TargetCell.Comment.Text
22                On Error GoTo ErrHandler
                  Dim MsgBoxResult As VbMsgBoxResult
                  Dim ReallyDifferent As Boolean
23                ReallyDifferent = Replace(Replace(Replace(OldComment, vbLf, ""), " ", ""), vbCr, "") <> Replace(Replace(Replace(NewComment, vbLf, ""), " ", ""), vbCr, "")
24                If OldComment <> NewComment Then
25                    If Not ReallyDifferent Then
26                        MsgBoxResult = vbYes
27                    Else
28                        Prompt = "Change comment at " + TargetCell.Parent.Name + "!" + Replace(TargetCell.Address, "$", "") + ": '" + CStr(TargetCell.Value) + "'?" + vbLf + vbLf + _
                                   "OldComment:" + vbLf + OldComment + vbLf + vbLf + "NewComment: " + vbLf + NewComment
                          Dim CheckBoxValue As Boolean
                          Dim DoTheSame As Variant
29                        If IsEmpty(DoTheSame) Then
30                            MsgBoxResult = MsgBoxPlus(Prompt, vbQuestion + vbYesNoCancel + vbDefaultButton2, "Apply Comments", "Yes, change comment", "No, don't change", "Exit this method now", , 500, "Do the same for all other comments", CheckBoxValue)
31                            If CheckBoxValue Then
32                                DoTheSame = MsgBoxResult
33                            End If
34                        Else
35                            MsgBoxResult = DoTheSame
36                        End If

37                    End If
38                    Select Case MsgBoxResult

                          Case vbYes
39                            SetCellComment TargetCell, CStr(NewComment), False
40                            anyChangesMade = True
41                        Case vbNo
                              'Nothing to do
42                        Case Else
43                            Exit Sub

44                    End Select
45                End If
46            ElseIf OldComment <> "" Then
47                Prompt = "Delete comment at " + TargetCell.Parent.Name + "!" + Replace(TargetCell.Address, "$", "") + ": '" + CStr(TargetCell.Value) + "'?" + vbLf + vbLf + _
                           "OldComment:" + vbLf + OldComment
48                TargetCell.Comment.Delete
49            End If
50        Next i

51        If anyChangesMade Then
52            FixCellComments shSummary
57        End If

58        Exit Sub
ErrHandler:
59        Throw "#ApplyComments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FixCellComments
' Author    : Philip Swannell
' Date      : 20-Jul-2016
' Purpose   : Work in progress, make the comments on sheets have consistent layout and formatting
'---------------------------------------------------------------------------------------
Sub FixCellComments(ws As Worksheet)
          Dim C As Range
          Dim cmt As Comment
          Dim R As Range
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         Set SPH = CreateSheetProtectionHandler(ws)
4         Set R = Nothing
5         On Error Resume Next
6         Set R = ws.UsedRange.SpecialCells(xlCellTypeComments)
7         On Error GoTo ErrHandler
8         If Not R Is Nothing Then

9             For Each C In R.Cells
10                Set cmt = Nothing
11                On Error Resume Next
12                Set cmt = C.Comment
13                On Error GoTo ErrHandler
14                If Not cmt Is Nothing Then
15                    With cmt.Shape
16                        .TextFrame.Characters.Font.Name = "Calibri"
17                        .TextFrame.Characters.Font.Size = 11
18                        .Width = 200
                          'Unfortunately I can't get WordWrap and AutoSize to play ball...
                          '   .TextFrame2.WordWrap = msoTrue
                          '  .TextFrame2.AutoSize = True
19                        .Height = sCommentHeight(cmt.Text, .Width)
20                        If .Height > 350 Then
21                            .Width = 400
22                            .Height = sCommentHeight(cmt.Text, .Width)
23                        End If
24                        .Left = C.Left + C.Width
25                        .Top = C.Top
26                    End With
27                End If
28            Next C
29        End If

30        Exit Sub
ErrHandler:
31        SomethingWentWrong "#FixCellComments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

