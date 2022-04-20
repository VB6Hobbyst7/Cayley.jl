Attribute VB_Name = "modSearchFormulas"
Option Explicit

Private Function IsRegExValid(RegEx As String)
          Dim Res
1         On Error GoTo ErrHandler
2         Res = sIsRegMatch(RegEx, "Foo", False)
3         IsRegExValid = (VarType(Res) = vbBoolean)
4         Exit Function
ErrHandler:
5         IsRegExValid = False
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SearchWorkbookFormulas
' Author    : Philip Swannell
' Date      : 2 Dec 2015
' Purpose   : Search all formulas of a workbook. Sometimes more useful than Excel's native
'             search capability (HOME > Find and Select > Find). For example:
'             Excel's FindAll does not allow copy and paste out of the dialog listing all hits.
'             Excel's FindAll lists many hits for a single array formula.
'             Excel's FindAll has no way of "seach for formulas containing string A and not string B", which I have often found to be useful.
' -----------------------------------------------------------------------------------------------------------------------
Sub SearchWorkbookFormulas(Optional InputSearchRegEx As String, Optional InputOnlyOldArrayFormulas As Variant, Optional WorkbookToSearch As Workbook, Optional CalledFromReleaseScript As Boolean)
          Dim c As Range
          Dim processThis As Boolean
          Dim Prompt As String
          Dim RangeToProcess As Range
          Dim SPH As clsSheetProtectionHandler
          Dim STK As clsStacker
          Dim ws As Worksheet
          Static SearchRegEx As String
          Static OnlyOldArrayFormulas As Boolean
          Dim DefaultRegEx As String
          Dim NFS As Long
          
          Dim rx As VBScript_RegExp_55.RegExp

          Dim DataToPaste As Variant
          Dim R As Range
          Dim wb As Excel.Workbook
          Dim wbName As String
          Const Title = "Search Workbook Formulas (" + gAddinName + ")"
          Dim ThisLine() As String

1         On Error GoTo ErrHandler

2         If CalledFromReleaseScript Then
3             InputSearchRegEx = "="
4             InputOnlyOldArrayFormulas = True
5         End If

6         If InputSearchRegEx <> "" Then
7             SearchRegEx = InputSearchRegEx
8         Else

TryAgain1:
9             If SearchRegEx = vbNullString Then
10                DefaultRegEx = "."
11            Else
12                DefaultRegEx = SearchRegEx
13            End If

14            Prompt = "This method searches the active workbook for formulas" + vbLf + _
                  "that match a regular expression. Results are shown in a" + vbLf + _
                  "new workbook." + vbLf + vbLf + _
                  "Example 1:" + vbLf + _
                  "to whole-word search for:" + vbLf + "MyFunctionName" + vbLf + _
                  "use the regular expresson:" + vbLf + "\bMyFunctionName\b" + vbLf + vbLf + _
                  "Example 2:" + vbLf + _
                  "to search for formulas that contain MyFunction or" + vbLf + _
                  "contain YourFunction use the regular expresson:" + vbLf + "MyFunction|YourFunction" + vbLf + vbLf + _
                  "Enter a regular expression or double-click for help" + vbLf + _
                  "with building one."

15            SearchRegEx = InputBoxPlus(Prompt, Title, DefaultRegEx, "Search Now", , 280, , , , True)
16            If SearchRegEx = "False" Then
17                SearchRegEx = vbNullString
18                GoTo EarlyExit
19            End If

20            If Not IsRegExValid(SearchRegEx) Then
21                MsgBoxPlus "That's not a valid Regular Expression", vbOKOnly + vbCritical, Title
22                GoTo TryAgain1
23            ElseIf SearchRegEx = vbNullString Then
24                GoTo TryAgain1
25            End If
26        End If

27        If VarType(InputOnlyOldArrayFormulas) = vbBoolean Then
28            OnlyOldArrayFormulas = InputOnlyOldArrayFormulas
29        Else
              Dim Prompt2 As String, Res2
30            Prompt2 = "What types of fomulas to search?"
31            Res2 = ShowOptionButtonDialog(sArrayStack("All formulas", "Only old-style array formulas"), Title, Prompt2, , , True)
32            If Res2 = 0 Then Exit Sub
33            OnlyOldArrayFormulas = Res2 = 2
34        End If

35        AddFilterToMRU "SearchFormulas", SearchRegEx

36        If WorkbookToSearch Is Nothing Then
37            If ActiveWorkbook Is Nothing Then Throw "There is no active workbook to search.", True
38            Set WorkbookToSearch = ActiveWorkbook
39        End If

40        For Each ws In WorkbookToSearch.Worksheets
41            If SheetIsProtectedWithPassword(ws) Then Throw "Cannot proceed since sheet '" + ws.Name + "' is protected with a password. Please remove all password protection from sheets before proceeding", True
42        Next ws

43        Set STK = CreateStacker
44        Set rx = New RegExp

45        With rx
46            .IgnoreCase = True
47            .Pattern = SearchRegEx
48            .Global = False        'Find first match only
49        End With

50        STK.Stack2D sArrayRange("Book", "Sheet", "Cell", "Formula")
51        ReDim ThisLine(1 To 1, 1 To 4)
52        wbName = "'" + WorkbookToSearch.Name
53        ThisLine(1, 1) = WorkbookToSearch.Name

54        For Each ws In WorkbookToSearch.Worksheets
55            ThisLine(1, 2) = "'" + ws.Name
56            Set SPH = CreateSheetProtectionHandler(ws)
57            Application.StatusBar = "Searching " + ws.Name

58            Set RangeToProcess = Nothing
59            On Error Resume Next
60            If ws.UsedRange.Cells.CountLarge = 1 Then
61                Set RangeToProcess = ws.UsedRange
62            Else
63                Set RangeToProcess = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
64            End If
65            On Error GoTo ErrHandler
66            If Not RangeToProcess Is Nothing Then
67                NFS = NFS + RangeToProcess.Cells.CountLarge
68                For Each c In RangeToProcess.Cells
69                    processThis = True
70                    If c.HasArray Then
71                        If c.address <> c.CurrentArray.Cells(1, 1).address Then
72                            processThis = False
73                        End If
74                    ElseIf OnlyOldArrayFormulas Then
75                        processThis = False
76                    End If
77                    If processThis Then
78                        If rx.Test(c.Formula) Then
79                            processThis = True
80                        Else
81                            processThis = False
82                        End If

83                        If processThis Then
84                            ThisLine(1, 4) = "'" + c.Formula
85                            If c.HasArray Then
86                                ThisLine(1, 3) = "'" + AddressND(c.CurrentArray)
87                            Else
88                                ThisLine(1, 3) = "'" + AddressND(c)
89                            End If
90                            STK.Stack2D ThisLine
91                        End If
92                    End If
93                Next c
94            End If
95        Next ws

96        Set SPH = Nothing
97        Application.StatusBar = False

98        If CalledFromReleaseScript Then
99            If STK.NumRows <= 1 Then
100               MsgBoxPlus "No old-style array formulas found in workbook '" + WorkbookToSearch.Name + "'", vbInformation
101               Exit Sub
102           End If
103       End If

104       DataToPaste = STK.Report

105       Set wb = Application.Workbooks.Add
106       Set R = wb.Worksheets(1).Range("A4").Resize(UBound(DataToPaste, 1) - LBound(DataToPaste, 1) + 1, UBound(DataToPaste, 2) - LBound(DataToPaste, 2) + 1)
107       R.Value = DataToPaste
108       R.Rows(1).Font.Bold = True
109       R.Columns.AutoFit
110       AddSortButtons R.Rows(0), 1
111       R.Parent.Range("A1").Value = "Search " + IIf(OnlyOldArrayFormulas, "old-stle array-", "all ") + "formulas of workbook " + WorkbookToSearch.Name + " for formulas matching regular expression: " + _
              SearchRegEx + "     Number of formulas searched: " + Format$(NFS, "###,###")
112       R.Parent.Range("A1").Select
EarlyExit:
113       Application.OnRepeat "Repeat Search Workbook Formulas", "SearchWorkbookFormulas"
114       Exit Sub
ErrHandler:
115       SomethingWentWrong "#SearchWorkbookFormulas (line " & CStr(Erl) + "): " & Err.Description & "!", , Title
116       Application.StatusBar = False
117       Application.OnRepeat "Repeat Search Workbook Formulas", "SearchWorkbookFormulas"
End Sub

