Attribute VB_Name = "modISDASIMMAmendSheets"
' -----------------------------------------------------------------------------------------------------------------------
' Name: modISDASIMMAmendSheets
' Kind: Module
' Purpose: Near throw-away code for doing edits to multiple excel workbooks used for the ISDA SIMM project
' Author: Philip Swannell
' Date: 10-Apr-2020
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit

Sub InsertDataFolderCallsToBook(wb As Workbook, ReleaseAsWell As Boolean)
          Dim ws As Worksheet
          Dim NumCellsAmended As Long
1         For Each ws In wb.Worksheets
2             InsertDataFolderCallsToSheet ws, NumCellsAmended
3         Next
4         If NumCellsAmended > 0 Then
5             Application.Goto wb.Worksheets("Audit").Cells(1, 1)
6             AddLineToAuditSheet wb.Worksheets("Audit"), False, "Added calls to ISDASIMMDataFolder"
7             If ReleaseAsWell Then
8                 RunReleaseCleanup wb
9                 ReleaseWorkbook wb, , , True
10                wb.Close False
11            End If
12            MsgBoxPlus "Amended " + CStr(NumCellsAmended) + " cells. Workbook released.", vbExclamation
13        Else
14            MsgBoxPlus "Amended no cells", vbInformation
15        End If
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InsertDataFolderCallsToSheet
' Author     : Philip Swannell
' Date       : 10-Apr-2020
' Purpose    : Find cells containing names of files under the data folder and replace them with calls to
'               ISDASIMMDataFolder
' -----------------------------------------------------------------------------------------------------------------------
Sub InsertDataFolderCallsToSheet(ws As Worksheet, ByRef NumCellsAmended As Long)
          Dim RangeToProcess As Range
1         On Error GoTo ErrHandler
          Dim SPH As clsSheetProtectionHandler
2         Set SPH = CreateSheetProtectionHandler(ws)
3         On Error Resume Next
4         If ws.UsedRange.CountLarge = 1 Then
5             Set RangeToProcess = ws.UsedRange
6         Else
7             Set RangeToProcess = ws.UsedRange.SpecialCells(xlCellTypeConstants)
8         End If
9         On Error GoTo ErrHandler
          
10        If Not RangeToProcess Is Nothing Then
11            ISDASIMMInsertDataFolderCalls RangeToProcess, NumCellsAmended
12        End If

13        Exit Sub
ErrHandler:
14        Throw "#InsertDataFolderCallsToSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMInsertDataFolderCalls
' Author     : Philip Swannell
' Date       : 09-Apr-2020
' Purpose    : Convenience function. Morphs a range of cells such that each cell that's text and points to a file or folder
'              under the current ISDASIMMDataFolder is replced by a call to ISDASIMMDataFolder. Parent sheet must have a range "TheYear" who's value is the current year
' Parameters :
'  RangeToProcess :
'  NumCellsAmended:
' -----------------------------------------------------------------------------------------------------------------------
Sub ISDASIMMInsertDataFolderCalls(RangeToProcess As Range, Optional NumCellsAmended As Long)
          Dim BaseFolder As String
          Dim c As Range
          Dim SubPath1 As String
          Dim Formula As String
          Dim OldValue As String
          Dim NewValue As String
          Dim TheYear As Long
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Set ws = RangeToProcess.Parent
3         If IsInCollection(ws.Names, "TheYear") Then
4             If IsNumber(ws.Range("TheYear")) Then
5                 TheYear = ws.Range("TheYear")
6                 BaseFolder = ISDASIMMDataFolder(TheYear)
7                 If Not sIsErrorString(TheYear) Then

8                     For Each c In RangeToProcess.Cells
9                         If VarType(c.Value) = vbString Then
10                            If LCase(Left(c.Value, Len(BaseFolder))) = LCase(BaseFolder) Then
11                                SubPath1 = Mid(c.Value, Len(BaseFolder) + 1)
12                                If LCase(ISDASIMMDataFolder(TheYear, SubPath1)) = LCase(c.Value) Then
13                                    Formula = "=ISDASIMMDataFolder(TheYear,""" + SubPath1 + """)"
14                                    OldValue = c.Value
15                                    c.Formula2 = Formula
16                                    NewValue = c.Value
17                                    FormatRangeAsLinkToTextFile c
18                                    NumCellsAmended = NumCellsAmended + 1
19                                    Debug.Print "============================================================================================="
20                                    Debug.Print "BookName = ", ws.Parent.Name
21                                    Debug.Print "SheetName = ", ws.Name
22                                    Debug.Print "CellAddress = ", c.address
23                                    Debug.Print "OldValue = ", OldValue
24                                    Debug.Print "NewValue = ", NewValue
25                                    If CStr(OldValue) <> CStr(NewValue) Then
26                                        Throw "Assertion failed..."
27                                    End If
28                                End If
29                            End If
30                        End If
31                    Next c
32                End If
33            End If
34        End If

35        Exit Sub
ErrHandler:
36        Throw "#ISDASIMMInsertDataFolderCalls (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub FormatManyBookLinksToTextFile()

          Const Folder = "C:\SolumWorkbooks\ISDA SIMM 2020\"
          Dim FileName As Variant
          Dim FileNames
          Dim BookNames
          Dim wb As Workbook

1         On Error GoTo ErrHandler
2         BookNames = Selection.Value
3         FileNames = sArrayConcatenate(Folder, BookNames)
4         If Not sAll(sFileExists(FileNames)) Then Throw "One or more files do not exist"

5         For Each FileName In FileNames
6             Set wb = Application.Workbooks.Open(CStr(FileName))
7             FormatBookLinksToTextFile wb, True, True
8         Next

9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#FormatManyBookLinksToTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub TestFormatBookLinksToTextFile()
1         On Error GoTo ErrHandler
2         FormatBookLinksToTextFile ActiveWorkbook, True, False

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestFormatBookLinksToTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub FormatBookLinksToTextFile(wb As Workbook, ReleaseAsWell As Boolean, CloseWhenDone As Boolean)
          Dim ws As Worksheet
          Dim NumCellsAmended1 As Long
          Dim NumCellsAmended2 As Long
          Dim BookName As String
1         On Error GoTo ErrHandler
2         BookName = wb.Name
3         For Each ws In wb.Worksheets
4             InsertDataFolderCallsToSheet ws, NumCellsAmended1
5         Next
6         If NumCellsAmended1 > 0 Then
7             AddLineToAuditSheet wb.Worksheets("Audit"), False, "Added calls to ISDASIMMDataFolder (via script in SolumAddin)"
8         End If
9         For Each ws In wb.Worksheets
10            FormatSheetLinksToTextFile ws, NumCellsAmended2
11        Next
12        If NumCellsAmended2 > 0 Then
13            AddLineToAuditSheet wb.Worksheets("Audit"), False, "Formatted calls to ISDASIMMDataFolder (via script in SolumAddin)"
14        End If

15        If NumCellsAmended1 + NumCellsAmended2 > 0 Then
16            If ReleaseAsWell Then
17                Application.Goto wb.Worksheets("Audit").Cells(1, 1)
18                RunReleaseCleanup wb
19                ReleaseWorkbook wb, , , True
20                If CloseWhenDone Then wb.Close False
21            End If
22            MsgBoxPlus "Amended " + CStr(NumCellsAmended1 + NumCellsAmended2) + " cells in workbook '" + BookName + "'. Workbook released.", vbExclamation
23        Else
24            MsgBoxPlus "Amended no cells in workbook '" + BookName + "'", vbInformation
25            If CloseWhenDone Then wb.Close False
26        End If

27        Exit Sub
ErrHandler:
28        Throw "#FormatBookLinksToTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub TestFormatSheetLinksToTextFile()
1         On Error GoTo ErrHandler
2         FormatSheetLinksToTextFile ActiveSheet, 0

3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestFormatSheetLinksToTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub FormatRangeAsLinkToTextFile(R As Range)

1         On Error GoTo ErrHandler
2         R.Locked = True
3         With R.Font
4             .ThemeColor = xlThemeColorAccent2
5             .TintAndShade = -0.249977111117893
6         End With

7         Exit Sub
ErrHandler:
8         Throw "#FormatRangeAsLinkToTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub FormatSheetLinksToTextFile(ws As Worksheet, ByRef NumCellsAmended As Long)
          Dim RangeToProcess As Range, c As Range
          
1         On Error GoTo ErrHandler
          Dim SPH As clsSheetProtectionHandler
2         Set SPH = CreateSheetProtectionHandler(ws)
3         On Error Resume Next
4         If ws.UsedRange.CountLarge = 1 Then
5             Set RangeToProcess = ws.UsedRange
6         Else
7             Set RangeToProcess = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
8         End If
9         On Error GoTo ErrHandler
10        If Not RangeToProcess Is Nothing Then
11            For Each c In RangeToProcess.Cells
12                If Left(LCase(c.Formula), 20) = LCase("=ISDASIMMDataFolder(") Then
13                    FormatRangeAsLinkToTextFile c
14                    NumCellsAmended = NumCellsAmended + 1
15                End If
16            Next
17        End If
18        Exit Sub
ErrHandler:
19        Throw "#FormatSheetLinksToTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

