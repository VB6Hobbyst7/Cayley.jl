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
Private Const BookNameForRegistry = "CayleyMarketData"
Private gBlockChangeEvent As Boolean

Private Function ConfigRange() As Range
1         On Error GoTo ErrHandler
2         Set ConfigRange = sExpandDown(RangeFromSheet(Me, "SCRiPTWorkbook").Cells(1, 0)).Resize(, 2)
3         Exit Function
ErrHandler:
4         Throw "#ConfigRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim FullPath As String
          Dim Res

1         On Error GoTo ErrHandler
2         Select Case Target.Cells(1, 1).Address    'Need the .Cells(1,1) to cope with merged cells
              Case Me.Range("SCRiPTWorkbook").Address
3                 Cancel = True

4                 FullPath = FileFromConfig("SCRiPTWorkbook")
5                 If sFileExists(FullPath) Then AddFileToMRU "SCRiPTWorkbooks", FullPath
6                 Res = GetOpenFilenameWrap("SCRiPTWorkbooks", "SCRiPTWorkbooks (*SCRiPT*.xlsm),*SCRiPT*.xlsm", , "Choose SCRiPT Workbook", , , True, Target)
7                 If Res <> False Then
8                     Res = sRelativePath(Res, ThisWorkbook.Path)
9                     SafeSetCellValue Target, Res
10                End If
11                Exit Sub
12            Case Me.Range("MarketDataFile").Address
13                Cancel = True
14                FullPath = FileFromConfig("MarketDataFile")
15                If sFileExists(FullPath) Then AddFileToMRU "AirbusMarketDataFiles", FullPath
16                Res = GetOpenFilenameWrap("AirbusMarketDataFiles", ".out files,*.out", , , , False, True, Target)
17                If Res <> False Then
18                    Res = sRelativePath(Res, ThisWorkbook.Path)
19                    SafeSetCellValue Target, Res
20                End If
21                Exit Sub
22            Case Me.Range("Numeraire").Address
23                Cancel = True
24                Res = ShowOptionButtonDialog(AllowedNumeraires(), "Select Numeraire Currency", , Target.Value, Target.Offset(, 1))
25                If Not IsEmpty(Res) Then
26                    Target.Value = Res
27                End If
28        End Select
29        Exit Sub
ErrHandler:
30        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function AllowedNumeraires()
          Dim Res As Variant
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         Res = CreateMissing()
3         For Each ws In ThisWorkbook.Worksheets
4             If Left(ws.Name, 14) = "HistoricalCorr" Then
5                 Res = sArrayStack(Res, UCase(Right(ws.Name, 3)))
6             End If
7         Next
8         AllowedNumeraires = sSortedArray(Res)
9         Exit Function
ErrHandler:
10        Throw "#AllowedNumeraires (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub Worksheet_Deactivate()
          'We may have changed the numeraire currency in which case the FxVol sheet should be reformatted since we grey out currency pairs that don't involve the numeraire
1         On Error GoTo ErrHandler
2         FormatFxVolSheet False
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#Worksheet_Deactivate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Worksheet_Change
' Author    : Philip Swannell
' Date      : 06-Nov-2016
' Purpose   : When user changes an item on the Config sheet, write it immediately to the Registry
' -----------------------------------------------------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
          Static InEvent As Boolean
1         On Error GoTo ErrHandler
2         If Not gBlockChangeEvent Then
3             If Not InEvent Then
4                 InEvent = True
5                 If Not Application.Intersect(ConfigRange, Target) Is Nothing Then
6                     SaveToRegistry
7                 End If
8                 InEvent = False
9             End If
10        End If
11        Exit Sub
ErrHandler:
12        SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub GetFromRegistry()
          Dim CopyOfErr As String
          Dim oldBlockChange
1         On Error GoTo ErrHandler
2         oldBlockChange = gBlockChangeEvent
3         gBlockChangeEvent = True
4         GetConfigFromRegistry ConfigRange(), BookNameForRegistry
5         gBlockChangeEvent = oldBlockChange
6         Exit Sub
ErrHandler:
7         CopyOfErr = "#GetFromRegistry (line " & CStr(Erl) + "): " & Err.Description & "!"
8         gBlockChangeEvent = oldBlockChange
9         Throw CopyOfErr
End Sub

Sub SaveToRegistry()
1         On Error GoTo ErrHandler
2         SaveConfigToRegistry ConfigRange(), BookNameForRegistry
3         Exit Sub
ErrHandler:
4         Throw "#SaveToRegistry (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
