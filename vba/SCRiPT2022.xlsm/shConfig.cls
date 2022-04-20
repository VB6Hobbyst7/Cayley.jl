VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function ConfigRange() As Range
1         On Error GoTo ErrHandler
2         Set ConfigRange = sExpandDown(RangeFromSheet(Me, "Headers").Cells(2, 1)).Resize(, 2)
3         Exit Function
ErrHandler:
4         Throw "#ConfigRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : SafeAddress
' Author    : Philip Swannell
' Date      : 11-Jan-2017
' Purpose   : Returns the address of the top left cell of a named range, or an error string
'             if the name does not exist or does not refer to a range.
'---------------------------------------------------------------------------------------
Private Function SafeAddress(RangeName As String)
1         On Error GoTo ErrHandler
2         SafeAddress = RangeFromSheet(Me, RangeName).Cells(1, 1).Address
3         Exit Function
ErrHandler:
4         SafeAddress = "#Not found!"
End Function

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
1         On Error GoTo ErrHandler
          Dim Alternatives
          Dim FullPath As String
          Dim Res
          Dim TargetTopLeft As Range

2         Cancel = True
3         Set TargetTopLeft = Target.Cells(1, 1)        'Need the .Cells(1,1) to cope with merged cells

4         Select Case TargetTopLeft.Address

              Case SafeAddress("MarketDataWorkbook")
5                 FullPath = FileFromConfig("MarketDataWorkbook")
6                 If sFileExists(FullPath) Then AddFileToMRU "SCRiPT_MarketDataWorkbooks", FullPath
7                 Res = GetOpenFilenameWrap("SCRiPT_MarketDataWorkbooks", "MarketDataWorkbooks (*MarketData*.xlsm),*MarketData*.xlsm", , "Choose Market Data Workbook", , , True, Target)
8                 If Res <> False Then
9                     Res = sRelativePath(Res, ThisWorkbook.Path)
10                    SafeSetCellValue TargetTopLeft, Res
11                End If
12                Exit Sub
13            Case SafeAddress("LinesWorkbook")
14                FullPath = FileFromConfig("LinesWorkbook")
                  'Note how we share the Registry key with the Cayley workbook....
15                If sFileExists(FullPath) Then AddFileToMRU "CayleyLinesFiles", FullPath
16                Res = GetOpenFilenameWrap("CayleyLinesFiles", , , "Choose Lines Workbook", , False, True, Target)
17                If Res <> False Then
18                    Res = sRelativePath(Res, ThisWorkbook.Path)
19                    SafeSetCellValue TargetTopLeft, Res
20                End If
21            Case SafeAddress("NumSimsCVA"), SafeAddress("NumSims")
22                Alternatives = sArraySubtract(sArrayPower(2, sGrid(5, 14, 10)), 1)
23            Case SafeAddress("PartitionByTrade"), SafeAddress("SavePaths"), SafeAddress("RestoreTradesAtStartup"), SafeAddress("DeveloperMode"), _
                  SafeAddress("UseCachedModel"), SafeAddress("BuildModelFromDFsAndSurvProbs"), SafeAddress("UseThreads"), SafeAddress("UseLinux")
24                Alternatives = sArrayStack(True, False)
25            Case SafeAddress("OnValuationErrors")
26                Alternatives = sArrayStack("Continue", "Stop")
27            Case SafeAddress("TimeGap")
28                Alternatives = sArrayStack(1 / 12, 0.25, 0.5)
29            Case Else
30                Cancel = False
31        End Select

32        If Not IsEmpty(Alternatives) Then
33            Cancel = True
34            Select Case sNRows(Alternatives)
                  Case 1
35                    SafeSetCellValue TargetTopLeft, Alternatives
36                Case 2
37                    If sEquals(TargetTopLeft.Value, Alternatives(1, 1)) Then
38                        Res = Alternatives(2, 1)
39                    Else
40                        Res = Alternatives(1, 1)
41                    End If
42                    SafeSetCellValue TargetTopLeft, Res
43                Case 3 To 10
44                    Res = ShowCommandBarPopup(Alternatives, FaceIDsFromAlternatives(Alternatives, CStr(Target.Cells(1, 1).Value)))
45                    If Res <> "#Cancel!" Then
46                        SafeSetCellValue TargetTopLeft, Res
47                    End If
48                Case Else
49                    Res = ShowSingleChoiceDialog(sArrayMakeText(Alternatives), , , , , MsgBoxTitle(), , TargetTopLeft.Offset(0, 1))
50                    If Not IsEmpty(Res) Then
51                        SafeSetCellValue TargetTopLeft, Res
52                    End If
53            End Select
54            Me.Calculate
55        End If
56        Exit Sub
ErrHandler:
57        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!", , MsgBoxTitle()
End Sub

Private Function FaceIDsFromAlternatives(Alternatives, CurrentChoice)

          Dim i As Long
          Dim NC As Long
          Dim NR As Long
          Dim Res As Variant
1         On Error GoTo ErrHandler
2         Force2DArrayR Alternatives, NR, NC
3         Res = sReshape(0, NR, 1)
4         For i = 1 To NR
5             If Unembellish(CStr(Alternatives(i, 1))) = CStr(CurrentChoice) Then
6                 Res(i, 1) = 1087
7             End If
8         Next i

9         FaceIDsFromAlternatives = Res

10        Exit Function
ErrHandler:
11        Throw "#FaceIDsFromAlternatives (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Worksheet_Change
' Author    : Philip Swannell
' Date      : 06-Nov-2016
' Purpose   : When user changes an item on the Config sheet, write it immediately to the Registry
'---------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
          Static InEvent As Boolean
1         On Error GoTo ErrHandler
2         If Not gBlockChangeEvent Then
3             If Not InEvent Then
4                 InEvent = True
5                 If Not Application.Intersect(ConfigRange, Target) Is Nothing Then
6                     SaveToRegistry
7                     getTradesRange(1).Rows(-1).Resize(2).Calculate    'update headers on portfolio sheet
8                 End If
9                 Me.Calculate
10                InEvent = False
11            End If
12        End If
13        Exit Sub
ErrHandler:
14        SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub GetFromRegistry()
          Dim CopyOfErr As String
          Dim oldBlockChange
1         On Error GoTo ErrHandler
2         oldBlockChange = gBlockChangeEvent
3         gBlockChangeEvent = True
4         GetConfigFromRegistry ConfigRange(), gProjectName
5         shConfig.Calculate
6         gBlockChangeEvent = oldBlockChange
7         Exit Sub
ErrHandler:
8         CopyOfErr = "#GetFromRegistry (line " & CStr(Erl) + "): " & Err.Description & "!"
9         gBlockChangeEvent = oldBlockChange
10        Throw CopyOfErr
End Sub

Sub SaveToRegistry()
1         On Error GoTo ErrHandler
2         SaveConfigToRegistry ConfigRange(), gProjectName
3         Exit Sub
ErrHandler:
4         Throw "#SaveToRegistry (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

