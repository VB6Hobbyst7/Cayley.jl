Attribute VB_Name = "modImport"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : Menu
' Author    : Philip Swannell
' Date      : 17-Nov-2016
' Purpose   : Menu on Summary sheet.
'---------------------------------------------------------------------------------------
Sub Menu()
          Const chImport = "Import data from another copy of the lines workbook..."
          Dim chAccessMode
          Const chHelp = "Help on Notional Weights"
          Const FidImport = 0
          Dim FidAccessMode As Long
          Const FidHelp = 49
          Dim Res

1         On Error GoTo ErrHandler

2         If ThisWorkbook.ReadOnly Then
3             chAccessMode = "The workbook is ReadOnly. Make it ReadWrite"
4             FidAccessMode = 16368
5         Else
6             chAccessMode = "The workbook is ReadWrite. Make it ReadOnly"
7             FidAccessMode = 16371
8         End If

9         Res = ShowCommandBarPopup(sArrayStack(chImport, chAccessMode, "--" & chHelp), sArrayStack(FidImport, FidAccessMode, FidHelp))

10        Select Case Res
              Case chImport
11                ImportData
12            Case chAccessMode
13                If ThisWorkbook.ReadOnly Then
14                    ThisWorkbook.ChangeFileAccess xlReadWrite
15                Else
16                    ThisWorkbook.ChangeFileAccess xlReadOnly
17                End If
18            Case chHelp
19                HelpButton
20        End Select
21        Exit Sub
ErrHandler:
22        SomethingWentWrong "#Menu (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsLinesBook
' Author    : Philip Swannell
' Date      : 17-Nov-2016
' Purpose   : Identifies if a workbook appears to be a version of the Cayley Lines Workbook
'---------------------------------------------------------------------------------------
Function IsLinesBook(wb As Workbook) As Boolean
          Dim lo As ListObject
          Dim ws As Worksheet
          Dim HeaderRange As Range

1         On Error GoTo ErrHandler
2         If IsInCollection(wb.Worksheets, "Summary") Then
3             Set ws = wb.Worksheets("Summary")
4             If ws.ListObjects.Count = 1 Then
5                 Set lo = ws.ListObjects(1)
6                 Set HeaderRange = lo.HeaderRowRange
7                 IsLinesBook = IsNumber(sMatch("CPTY_PARENT", sArrayTranspose(HeaderRange.Value)))
8             End If
9         End If
10        Exit Function
ErrHandler:
11        Throw "#IsLinesBook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ImportData
' Author    : Philip Swannell
' Date      : 17-Nov-2016
' Purpose   : Makes it easy to import data from another copy of the workbook
'             User can do a "Dummy run" to see what would be changed (in a log file)
'             If list of banks don't match throws an error, user needs to fix this problem
'             Cell comments are also copied across.
'             If columns in source and header don't match
'---------------------------------------------------------------------------------------
Sub ImportData()
          Dim wb As Workbook
          Dim N As Long
          Dim sourceBook As Workbook
          Dim SourceBookName As String
          Dim SourceBooknames As Variant
          Dim SourceSheet As Worksheet, TargetSheet As Worksheet
          Dim SourceComment As String, TargetComment As String
          Dim SourceHeaders As Range, TargetHeaders As Range, SourceBanks As Range, TargetBanks As Range
          Dim SourceDataRange As Range, TargetDataRange As Range
          Dim RowMatches, ColumnMatches
          Dim TargetShortNames As Range
          Dim i As Long, j As Long, SourceCell As Range, TargetCell As Range
          Dim k As Long
          Dim SPH As clsSheetProtectionHandler
          Dim logSheet As Worksheet
          Dim ThisBank As String, ThisCol As String, Prompt As String, ThisShortName As String
          Dim ForReal As Boolean, Res

1         On Error GoTo ErrHandler
2         SourceBooknames = CreateMissing()
3         For Each wb In Application.Workbooks
4             If Not wb Is ThisWorkbook Then
5                 If IsLinesBook(wb) Then
6                     N = N + 1
7                     Set sourceBook = wb
8                     SourceBookName = wb.Name
9                 End If
10            End If
11        Next wb

12        If N = 0 Then
13            Throw "Please open the copy of the lines workbook from which you want to import data", True
14        ElseIf N = 1 Then
              'nothing to do
15        ElseIf N > 1 Then
16            SourceBookName = ShowMultipleChoiceDialog(SourceBooknames, , , "Select lines workbook from which to import")
17            If SourceBookName = "#User Cancel!" Then Exit Sub
18            Set sourceBook = Application.Workbooks(SourceBookName)
19        End If

20        Prompt = "Import data from:" + vbLf + _
                   sourceBook.FullName + vbLf + _
                   "To:" + vbLf + _
                   ThisWorkbook.FullName + vbLf + vbLf + _
                   "Do you want to:" + vbLf + _
                   "a) Do a ""Dummy run"" to see a log file showing what data would be updated; or" + vbLf + _
                   "b) Import the data"

21        Res = MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion, "Cayley Lines Book", "Dummy", "Import", "Exit, do nothing", , 600)
22        Select Case Res
          Case vbYes
23            ForReal = False
24        Case vbNo
25            ForReal = True
26        Case Else
27            Exit Sub
28        End Select

29        Set SourceSheet = sourceBook.Worksheets("Summary")
30        Set SourceDataRange = SourceSheet.ListObjects(1).DataBodyRange
31        Set SourceHeaders = SourceSheet.ListObjects(1).HeaderRowRange
32        Set SourceBanks = SourceDataRange.Columns(sMatch("CPTY_PARENT", sArrayTranspose(SourceHeaders.Value)))

33        Set TargetSheet = ThisWorkbook.Worksheets("Summary")
34        Set TargetDataRange = TargetSheet.ListObjects(1).DataBodyRange
35        Set TargetHeaders = TargetSheet.ListObjects(1).HeaderRowRange
36        Set TargetBanks = TargetDataRange.Columns(sMatch("CPTY_PARENT", sArrayTranspose(TargetHeaders.Value)))
37        Set TargetShortNames = TargetDataRange.Columns(sMatch("Very short name", sArrayTranspose(TargetHeaders.Value)))

38        RowMatches = sMatch(TargetBanks.Value, SourceBanks.Value)
39        ColumnMatches = sMatch(sArrayTranspose(TargetHeaders.Value), sArrayTranspose(SourceHeaders.Value))

40        If Not sArraysIdentical(sSortedArray(SourceBanks.Value), sSortedArray(TargetBanks.Value)) Then
41            g sCompareTwoArrays(SourceBanks.Value, TargetBanks.Value)
42            Throw "List of banks in column ""CPTY_PARENT"" do not match between the source and target workbooks. Please fix before proceeding"
43        End If

44        If Not sArraysIdentical(sSortedArray(sArrayTranspose(SourceHeaders.Value)), sSortedArray(sArrayTranspose(TargetHeaders.Value))) Then
45            Prompt = "Note that headers in the Source and Target books don't match. Only columns that appear in both workbooks will be updated."
46            Res = sCompareTwoArrays(sArrayTranspose(SourceHeaders.Value), sArrayTranspose(TargetHeaders.Value), "In1AndNotIn2")
47            If snrows(Res) > 1 Then
48                Prompt = Prompt + vbLf + vbLf + "Columns in the Source but not the Target:" + vbLf + _
                           sConcatenateStrings(sSubArray(Res, 2))
49            End If
50            Res = sCompareTwoArrays(sArrayTranspose(SourceHeaders.Value), sArrayTranspose(TargetHeaders.Value), "In2AndNotIn1")
51            If snrows(Res) > 1 Then
52                Prompt = Prompt + vbLf + vbLf + "Columns in the Target but not the Source:" + vbLf + _
                           sConcatenateStrings(sSubArray(Res, 2))
53            End If
54            Prompt = Prompt + vbLf + vbLf + "Do you want to proceed" + IIf(ForReal, " and import the data?", " with this dummy run?")

55            If MsgBoxPlus(Prompt, vbQuestion + vbYesNo, , "Yes,Proceed", "No, Exit") <> vbYes Then Exit Sub
56        End If

57        Set logSheet = Application.Workbooks.Add.Worksheets(1)
58        With logSheet
59            .Cells(1, 1).Value = "Log for update of Cayley Lines Workbook"
60            .Cells(2, 1).Value = "Source"
61            .Cells(2, 2).Value = "'" + sourceBook.FullName
62            .Cells(3, 1).Value = "Target"
63            .Cells(3, 2).Value = "'" + ThisWorkbook.FullName
64            .Cells(5, 1) = "Bank"
65            .Cells(5, 2) = "Column"
66            .Cells(5, 3) = "Cell in Source"
67            .Cells(5, 4) = "Cell in Target"
68            .Cells(5, 5) = "Value from Source"
69            .Cells(5, 6) = "Overwrote value in target"
70            .Cells(5, 7) = "Comment from Source"
71            .Cells(5, 8) = "Overwrote comment in target"

72        End With
73        k = 6
          
74        If ForReal Then
75            Set SPH = CreateSheetProtectionHandler(shSummary)
76        End If

77        For j = 1 To TargetDataRange.Columns.Count
78            For i = 1 To TargetDataRange.Rows.Count
79                ThisBank = TargetBanks.Cells(i, 1).Value
80                ThisCol = TargetHeaders.Cells(1, j).Value
81                ThisBank = TargetBanks.Cells(i, 1).Value
82                ThisShortName = TargetShortNames.Cells(i, 1)
83                Set TargetCell = TargetDataRange.Cells(i, j)
84                If IsNumber(RowMatches(i, 1)) Then
85                    If IsNumber(ColumnMatches(j, 1)) Then
86                        Set SourceCell = SourceDataRange(RowMatches(i, 1), ColumnMatches(j, 1))
87                        SourceComment = ""
88                        TargetComment = ""
89                        On Error Resume Next
90                        SourceComment = SourceCell.Comment.Text
91                        TargetComment = TargetCell.Comment.Text
92                        On Error GoTo ErrHandler

93                        If SourceCell.Value <> TargetCell.Value Or SourceComment <> TargetComment Then
94                            With logSheet
95                                .Cells(k, 1) = ThisShortName
96                                .Cells(k, 2) = ThisCol
97                                .Cells(k, 3) = Replace(SourceCell.Address, "$", "")
98                                .Cells(k, 4) = Replace(TargetCell.Address, "$", "")
99                                .Cells(k, 5) = SourceCell.Value
100                               .Cells(k, 6) = TargetCell.Value
101                               .Cells(k, 7) = SourceComment
102                               .Cells(k, 8) = TargetComment
103                               k = k + 1
104                           End With
105                           If ForReal Then
106                               If TargetCell.Value <> SourceCell.Value Then
107                                   TargetCell.Value = SourceCell.Value
108                               End If
109                               If TargetComment <> SourceComment Then
110                                   If SourceComment = "" Then
111                                       TargetCell.Comment.Delete
112                                   Else
113                                       SetCellComment TargetCell, SourceComment, False
114                                   End If
115                               End If
116                           End If
117                       End If
118                   End If
119               End If
120           Next i
121       Next j

122       With logSheet.Cells(5, 1).CurrentRegion
123           .VerticalAlignment = xlVAlignCenter
124           .HorizontalAlignment = xlHAlignCenter
125           .Rows(1).Font.Bold = True
126           .Columns.AutoFit
127       End With

128       Exit Sub
ErrHandler:
129       SomethingWentWrong "#ImportData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetCellComment
' Author    : Philip Swannell
' Date      : 20-May-2016
' Purpose   : Adds a comment to a cell and makes it appear in Calibri 11. Comment must be
'             passed including line feed characters
'---------------------------------------------------------------------------------------
Function SetCellComment(C As Range, ByVal Comment As String, InsertBreaks As Boolean)
1         On Error GoTo ErrHandler

2         If InsertBreaks Then
3             Comment = sConcatenateStrings(sJustifyText(Comment, "Calibri", 11, 300), vbLf)
4         End If

5         C.ClearComments
6         C.AddComment
7         C.Comment.Visible = False
8         C.Comment.Text Text:=Comment
9         With C.Comment.Shape.TextFrame
10            .Characters.Font.Name = "Calibri"
11            .Characters.Font.Size = 11
12            .AutoSize = True
13        End With
14        Exit Function
ErrHandler:
15        Throw "#SetCellComment (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

