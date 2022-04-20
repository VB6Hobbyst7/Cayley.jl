Attribute VB_Name = "modAudit"
'---------------------------------------------------------------------------------------
' Module    : modAudit
' Author    : Philip Swannell
' Date      : 16-May-2016
' Purpose   : Code called from the menu on the Audit sheet
'---------------------------------------------------------------------------------------
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: ClearPortfolioSheet
' Purpose: Encapsulate putting the Portfolio sheet into good order with no trades on it
' Author: Philip Swannell
' Date: 01-Dec-2017
' ----------------------------------------------------------------
Sub ClearPortfolioSheet()
          Dim SPH As SolumAddin.clsSheetProtectionHandler
1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shPortfolio)
3         getTradesRange(1).EntireRow.Delete
4         SetTradesRangeColumnWidths
5         FormatTradesRange    ' Ensures that <Doubleclick to add trade> message is present
6         ResetSortButtons RangeFromSheet(shPortfolio, "PortfolioHeader").Offset(-2, 0).Resize(, RangeFromSheet(shHiddenSheet, "InterestRateSwapTemplate").Columns.Count), False, False
7         RangeFromSheet(shPortfolio, "TradesFileName").ClearContents
8         RangeFromSheet(shPortfolio, "TheFilters").ClearContents
9         Exit Sub
ErrHandler:
10        Throw "#ClearPortfolioSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function RestoreConfigCell(RangeName As String, NewValue As Variant, ByRef ChangesMade)
          Dim c As Range
1         Set c = ConfigRange(RangeName)
2         If Not sNearlyEquals(c.Value, NewValue) Then
3             ChangesMade = sArrayStack(ChangesMade, sArrayRange(RangeName, c.Value, NewValue))
4             SafeSetCellValue c, NewValue
5         End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReleaseCleanup
' Author    : Philip Swannell
' Date      : 02-Nov-2015
' Purpose   : Put workbook in a fit state to be released.
'---------------------------------------------------------------------------------------
Function ReleaseCleanup()
          Dim ChangesMade As Variant
          Dim CopyOfErr As String
          Dim CountOfRows As Long
          Dim i As Long
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim VisibleShouldBe As Variant
          Dim ws As Worksheet

1         On Error GoTo ErrHandler

2         gBlockChangeEvent = True
3         Set SUH = CreateScreenUpdateHandler()
4         If Not gDoValidation Then Throw "Please amend VBA code to set constant ModMain.gDoValidation to TRUE"

5         shConfig.Unprotect
6         ChangesMade = sArrayRange("Parameter", "Changed From", "Changed To")
7         RestoreConfigCell "OurName", "Airbus", ChangesMade
8         RestoreConfigCell "NumSimsCVA", 2047, ChangesMade
9         RestoreConfigCell "NumSims", 255, ChangesMade
10        RestoreConfigCell "TimeGap", 1 / 12, ChangesMade
11        RestoreConfigCell "PFEPercentile", 0.95, ChangesMade
12        RestoreConfigCell "SavePaths", False, ChangesMade
13        RestoreConfigCell "PartitionByTrade", False, ChangesMade
14        RestoreConfigCell "RestoreTradesAtStartup", True, ChangesMade
15        RestoreConfigCell "OnValuationErrors", "Continue", ChangesMade
16        RestoreConfigCell "DeveloperMode", False, ChangesMade
17        RestoreConfigCell "UseCachedModel", True, ChangesMade
18        RestoreConfigCell "BuildModelFromDFsAndSurvProbs", False, ChangesMade
19        RestoreConfigCell "UseLinux", False, ChangesMade
20        RestoreConfigCell "MarketDataWorkbook", "CayleyMarketData.xlsm", ChangesMade
21        RestoreConfigCell "LinesWorkbook", "CayleyLines.xlsm", ChangesMade
22        shConfig.SaveToRegistry 'Necessary since gBlockChangeEvent is True

23        If sNRows(ChangesMade) > 1 Then
              Dim j As Long
              Dim Prompt As String
24            For i = 2 To sNRows(ChangesMade)
25                For j = 2 To 3
26                    Select Case VarType(ChangesMade(i, j))
                          Case vbBoolean
27                            ChangesMade(i, j) = UCase(ChangesMade(i, j))
28                        Case vbString
29                            ChangesMade(i, j) = "'" + ChangesMade(i, j) + "'"
30                        Case Else
31                            ChangesMade(i, j) = CStr(ChangesMade(i, j))
32                    End Select
33                Next
34            Next
35            Prompt = "Release cleanup made the following changes on the Config sheet:" + vbLf + vbLf + _
                  sConcatenateStrings(sJustifyArrayOfStrings(ChangesMade, "SegoeUI", 9, vbTab), vbLf)
36            MsgBoxPlus Prompt, vbInformation, MsgBoxTitle(), , , , , 1000, , , 60, vbOK
37        End If

38        BackUpTrades
39        ClearPortfolioSheet
40        AlignCharts

          'Order the sheets
          Dim SheetNames
          Dim SheetNames2
41        SheetNames = sReshape("", ThisWorkbook.Worksheets.Count, 1)
42        i = 0
43        For Each ws In ThisWorkbook.Worksheets
44            i = i + 1
45            SheetNames(i, 1) = ws.Name
46        Next
47        SheetNames2 = RangeFromSheet(shHiddenSheet, "SheetSettings").Columns(1).Value
48        If Not sArraysIdentical(sSortedArray(SheetNames), sSortedArray(SheetNames2)) Then
49            g sCompareTwoArrays(SheetNames, SheetNames2), ExMthdDebugWindow
50            Throw "Mismatch between the names of the sheets in the workbook and the names held on the HiddenSheet!SheetSettings. Please correct this before release. See VBA debug window for details"
51        End If

          'Activate top-left of each visible sheet
52        For Each ws In ThisWorkbook.Worksheets
53            ws.Protect , False, False
54            CountOfRows = ws.UsedRange.Rows.Count
              Dim ShapeName As String
55            For i = 1 To 2
56                ShapeName = Choose(i, "SolumLogo", "ButtonMenu")
57                If IsInCollection(ws.Shapes, ShapeName) Then
58                    With ws.Shapes(ShapeName)
59                        .Placement = xlFreeFloating
60                        .Top = shPortfolio.Shapes(ShapeName).Top
61                        .Left = shPortfolio.Shapes(ShapeName).Left
62                        .Width = shPortfolio.Shapes(ShapeName).Width
63                        .Height = shPortfolio.Shapes(ShapeName).Height
64                    End With
65                End If
66            Next i
67            ws.Protect , True, True
68            GroupingButtonDoAllOnSheet ws, True
69            gBlockCalculateEvent = True
70            ws.Calculate
71            gBlockCalculateEvent = False
72            VisibleShouldBe = RangeFromSheet(shHiddenSheet, "SheetSettings").Cells(sMatch(ws.Name, RangeFromSheet(shHiddenSheet, "SheetSettings").Columns(1)), 2)
73            Select Case LCase(CStr(VisibleShouldBe))
                  Case "true", "visible"
74                    ws.Visible = xlSheetVisible
75                Case "false", "hidden"
76                    ws.Visible = xlSheetHidden
77                Case "veryhidden", "very hidden"
78                    ws.Visible = xlSheetVeryHidden
79                Case Else
80                    Throw "Unrecognised value in Visible? column of range SheetSettings on worksheet " + shHiddenSheet.Name
81            End Select

82            If ws.Visible = xlSheetVisible Then
83                Application.GoTo ws.Cells(1, 1)
84                ActiveWindow.DisplayGridlines = False
85                ActiveWindow.DisplayHeadings = False
86                ActiveWindow.Zoom = 100
87            End If
88        Next ws

89        CleanOutDashboard
90        ClearCounterpartyViewerSheet
91        ClearTradeViewerSheet
92        ClearCashflowDrilldownSheet

93        Application.GoTo shxVADashboard.Cells(1, 1)
94        Application.GoTo shPortfolio.Cells(1, 1)        ' ensures that Portfolio tab is visible

95        Exit Function
ErrHandler:
96        CopyOfErr = "#ReleaseCleanup (line " & CStr(Erl) + "): " & Err.Description & "!"
97        gBlockChangeEvent = False
98        ReleaseCleanup = CopyOfErr    'This method called via Application.Run so have to return error, not throw
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetSheetCodeNames
' Author    : Philip Swannell
' Date      : 02-Nov-2015
' Purpose   : Make worksheet CodeNames correspond to the names that appear in the sheet
'             tabs, that way they are sorted helpfully in the project browser...
'---------------------------------------------------------------------------------------
Sub SetSheetCodeNames()
          Dim CodeName As String
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         For Each ws In ThisWorkbook.Worksheets
3             CodeName = "sh" + Replace(ws.Name, " ", "_")    'There may be characters other than the space character that also need to be aliased...
4             If ws.CodeName <> CodeName Then
5                 ThisWorkbook.VBProject.VBComponents(ws.CodeName).Name = CodeName
6             End If
7         Next ws

8         Exit Sub
ErrHandler:
9         SomethingWentWrong "#SetSheetCodeNames (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AlignCharts
' Author    : Philip Swannell
' Date      : 19-Apr-2016
' Purpose   : Align the two similar charts on the CounterpartyViewer and TradeViewer sheets
'---------------------------------------------------------------------------------------
Sub AlignCharts()
          Dim chOb1 As ChartObject
          Dim chOb2 As ChartObject
          Dim SPH1 As SolumAddin.clsSheetProtectionHandler
          Dim SPH2 As SolumAddin.clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Set SPH1 = CreateSheetProtectionHandler(shCounterpartyViewer)
3         Set SPH2 = CreateSheetProtectionHandler(shTradeViewer)
4         Set chOb1 = shTradeViewer.ChartObjects(1)
5         Set chOb2 = shCounterpartyViewer.ChartObjects(1)

6         chOb1.Placement = xlFreeFloating
7         chOb2.Placement = xlFreeFloating

8         chOb2.Top = chOb1.Top
9         chOb2.Left = chOb1.Left
10        chOb2.Height = chOb1.Height
11        chOb2.Width = chOb1.Width

12        Exit Sub
ErrHandler:
13        Throw "#AlignCharts (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UnprotectAllSheets
' Author    : Philip Swannell
' Date      : 05-Nov-2015
' Purpose   : Unprotects all sheets, attached to button on Audit sheet
'---------------------------------------------------------------------------------------
Sub UnprotectAllSheets()
          Dim oldVisState
          Dim origSheet As Worksheet
          Dim SUH As SolumAddin.clsScreenUpdateHandler
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler()
3         Set origSheet = ActiveSheet
4         For Each ws In ThisWorkbook.Worksheets
5             ws.Unprotect
6             oldVisState = ws.Visible
7             ws.Visible = xlSheetVisible
8             ws.Activate
9             ActiveWindow.DisplayHeadings = True
10            ws.Visible = oldVisState
11        Next
12        origSheet.Activate
13        Exit Sub
ErrHandler:
14        Throw "#UnprotectAllSheets (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
