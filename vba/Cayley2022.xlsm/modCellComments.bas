Attribute VB_Name = "modCellComments"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetCellComment
' Author    : Philip Swannell
' Date      : 31-Oct-2016
' Purpose   : Utility function used to build sheet CommentEditor
' -----------------------------------------------------------------------------------------------------------------------
Function GetCellComment(R As Range)
1         On Error GoTo ErrHandler
2         GetCellComment = R.Comment.Text
3         Exit Function
ErrHandler:
4         GetCellComment = ""
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ApplyCellComments
' Author    : Philip Swannell
' Date      : 31-Oct-2016
' Purpose   : Attached to button on CommentEditor sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub ApplyCellComments()
1         On Error GoTo ErrHandler
2         ApplyCommmentsFromCommentEditor
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#ApplyCellComments (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CountWordsInRange
' Author    : Philip Swannell
' Date      : 31-Oct-2016
' Purpose   : Utility used on the CommentsEditor sheet - move to SolumAddin?
' -----------------------------------------------------------------------------------------------------------------------
Function CountWordsInRange(R As Range)
          Dim c As Range
          Dim N As Long

1         On Error GoTo ErrHandler
2         For Each c In R.Cells
3             If VarType(c.Value) = vbString Then
4                 N = N + (Len(sRegExReplace(c.Value, "\b", "x")) - Len(c.Value)) / 2
5             End If
6         Next c

7         CountWordsInRange = N

8         Exit Function
ErrHandler:
9         Throw "#CountWordsInRange (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ApplyCommmentsFromCommentEditor
' Author    : Philip Swannell
' Date      : 31-Oct-2016
' Purpose   : Attached to button on sheet CommentEditor
' -----------------------------------------------------------------------------------------------------------------------
Sub ApplyCommmentsFromCommentEditor()
          Dim i As Long
          Dim N As Long
          Dim NewComment As Variant
          Dim OldComment As String
          Dim Prompt As String
          Dim SourceData As Range
          Dim SPH As clsSheetProtectionHandler
          Dim TargetCell As Range
          Dim TargetSheet As Worksheet

1         Application.ScreenUpdating = False
2         shCommentEditor.Calculate

3         On Error GoTo ErrHandler
4         Set SourceData = RangeFromSheet(shCommentEditor, "TheDataWithHeaders")

5         N = SourceData.Rows.Count

6         For i = 2 To N
7             Set TargetSheet = ThisWorkbook.Worksheets(SourceData.Cells(i, 1).Value)
8             Set TargetCell = RangeFromSheet(TargetSheet, SourceData.Cells(i, 3).Value)
9             Set TargetCell = TargetCell.Cells(SourceData.Cells(i, 4).Value, SourceData.Cells(i, 5).Value)

10            If TargetCell.Value <> SourceData.Cells(i, 2).Value Then
11                Throw "Assertion failed at row " & CStr(i) & vbLf & " Expecting " & _
                      TargetCell.Parent.Name & "!" & TargetCell.Address & " to contain '" & _
                      SourceData.Cells(i, 2).Value & "' but it contains '" & TargetCell.Value & "'"
12            End If

13            If SourceData.Cells(i, 1).Value <> SourceData.Cells(i - 1, 1).Value Then
14                Set SPH = CreateSheetProtectionHandler(TargetSheet)
15            End If

16            OldComment = ""
17            On Error Resume Next
18            OldComment = TargetCell.Comment.Text
19            On Error GoTo ErrHandler

20            NewComment = SourceData.Cells(i, 6).Value
21            If sArrayIsNonTrivialText(NewComment)(1, 1) Then
22                OldComment = ""
23                On Error Resume Next
24                OldComment = TargetCell.Comment.Text
25                On Error GoTo ErrHandler
                  Dim MsgBoxResult As VbMsgBoxResult
                  Dim ReallyDifferent As Boolean
26                ReallyDifferent = Replace(Replace(Replace(OldComment, vbLf, ""), " ", ""), vbCr, "") <> _
                      Replace(Replace(Replace(NewComment, vbLf, ""), " ", ""), vbCr, "")
27                If OldComment <> NewComment Then
28                    If Not ReallyDifferent Then
29                        MsgBoxResult = vbYes
30                    Else
31                        Prompt = "Change comment at " & TargetCell.Parent.Name & "!" & _
                              Replace(TargetCell.Address, "$", "") & ": '" & CStr(TargetCell.Value) & "'?" & _
                              vbLf & vbLf & "OldComment:" & vbLf & OldComment & vbLf & vbLf & _
                              "NewComment: " & vbLf & NewComment
                          Dim CheckBoxValue As Boolean
                          Dim DoTheSame As Variant
32                        If IsEmpty(DoTheSame) Then
33                            MsgBoxResult = MsgBoxPlus(Prompt, vbQuestion + vbYesNoCancel + vbDefaultButton2, _
                                  "Apply Comments", _
                                  "Yes, change comment", _
                                  "No, don't change", _
                                  "Exit this method now" _
                                  , , 500, "Do the same for all other comments", CheckBoxValue)
34                            If CheckBoxValue Then
35                                DoTheSame = MsgBoxResult
36                            End If
37                        Else
38                            MsgBoxResult = DoTheSame
39                        End If

40                    End If
41                    Select Case MsgBoxResult

                          Case vbYes
42                            SetCellComment TargetCell, CStr(NewComment), False
43                            ResizeComment TargetCell
44                        Case vbNo
                              'Nothing to do
45                        Case Else
46                            Exit Sub

47                    End Select
48                End If
49            ElseIf OldComment <> "" Then
50                Prompt = "Delete comment at " & TargetCell.Parent.Name & "!" & _
                      Replace(TargetCell.Address, "$", "") & ": '" & CStr(TargetCell.Value) & "'?" & vbLf & vbLf & _
                      "OldComment:" & vbLf & OldComment
51                TargetCell.Comment.Delete
52            End If
53        Next i

54        Exit Sub
ErrHandler:
55        Throw "#ApplyCommmentsFromCommentEditor (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Sub ResizeComment(TheCell As Range)

          Dim cmt As Comment
          Dim HeightShouldBe As Double

1         On Error GoTo ErrHandler
2         Set cmt = Nothing
3         On Error Resume Next
4         Set cmt = TheCell.Comment

5         If Not cmt Is Nothing Then
6             With cmt.Shape
7                 With .TextFrame.Characters.Font
8                     If .Name <> "Calibri" Then
9                         .Name = "Calibri"
10                    End If
11                    If .Size <> 11 Then
12                        .Size = 11
13                    End If
14                End With
15                If Abs(.Width - 200) > 1 Then
16                    .Width = 200
17                End If
                  'Unfortunately I can't get WordWrap and AutoSize to play ball...
                  '   .TextFrame2.WordWrap = msoTrue
                  '  .TextFrame2.AutoSize = True
18                HeightShouldBe = sCommentHeight(cmt.Text, .Width)
19                If .Height <> HeightShouldBe Then
20                    .Height = HeightShouldBe
21                End If
22                If .Height > 350 Then
23                    If Abs(.Width - 400) > 1 Then
24                        .Width = 400
25                    End If
26                    HeightShouldBe = sCommentHeight(cmt.Text, .Width)
27                    If .Height <> HeightShouldBe Then
28                        .Height = HeightShouldBe
29                    End If
30                End If
31                If .Left <> TheCell.Left + TheCell.Width Then
32                    .Left = TheCell.Left + TheCell.Width
33                End If
34                If .Top <> TheCell.Top Then
35                    .Top = TheCell.Top
36                End If
37            End With
38        End If

39        Exit Sub
ErrHandler:
40        Throw "#ResizeComment (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResizeCommentsOnSheet
' Author    : Philip Swannell
' Date      : 20-Jul-2016
' Purpose   : Work in progress, make the comments on sheets have consistent layout and formatting
' -----------------------------------------------------------------------------------------------------------------------
Sub ResizeCommentsOnSheet(ws As Worksheet)
          Dim c As Range
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
9             For Each c In R.Cells
10                ResizeComment c
11            Next c
12        End If
13        Exit Sub
ErrHandler:
14        SomethingWentWrong "#ResizeCommentsOnSheet (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

