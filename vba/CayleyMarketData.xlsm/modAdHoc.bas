Attribute VB_Name = "modAdHoc"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : Module1
' Author    : Philip
' Date      : 07-Sep-2017
' Purpose   : Ad-hoc code for amending the currency sheets
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit






Sub MoreAmendments()
          Dim SPH As clsSheetProtectionHandler
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         For Each ws In ThisWorkbook.Worksheets
3             If IsCurrencySheet(ws) Then
4                 Set SPH = CreateSheetProtectionHandler(ws)

13                FormatCurrencySheet ws, False, Empty

14            End If
15        Next
16        Exit Sub
ErrHandler:
17        SomethingWentWrong "#MoreAmendments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub ShuffleSwapColsOnAllSheets()
          Dim ws As Worksheet
1         For Each ws In ThisWorkbook.Worksheets
2             If IsCurrencySheet(ws) Then
3                 ShuffleSwapCols ws
4             End If
5         Next
End Sub

Sub TestShuffleSwapCols()
1         ShuffleSwapCols ActiveSheet
End Sub

Sub ShuffleSwapCols(ws As Worksheet)
          Dim NewData
          Dim SPH As clsSheetProtectionHandler
          Dim Target As Range

1         On Error GoTo ErrHandler
2         Set Target = sExpandDown(ws.Range("SwapRatesInit"))
3         With Target
4             Set Target = .Offset(-1).Resize(.Rows.Count + 1, .Columns.Count + 1)
5         End With
6         NewData = ShuffleHeaders(Target.Value, sArrayRange("Tenor", "Rate", "FixFreq", "FixDCT", "FloatFreq", "FloatDCT", "BloombergCode"))

7         Set SPH = CreateSheetProtectionHandler(ws)
8         Target.Value = sArrayexcelString(NewData)
9         Target.HorizontalAlignment = xlHAlignCenter
10        FormatCurrencySheet ws, False, Empty

11        Exit Sub
ErrHandler:
12        Throw "#ShuffleSwapCols (line " & CStr(Erl) + "): " & Err.Description & "!"

End Sub

Function ShuffleHeaders(ArrayWithHeaders, NewHeaders)
          Dim i As Long
          Dim j As Long
          Dim MatchIDs
          Dim NC As Long
          Dim Result

1         On Error GoTo ErrHandler
2         MatchIDs = sMatch(sArrayTranspose(NewHeaders), sArrayTranspose(sSubArray(ArrayWithHeaders, 1, 1, 1)))

3         NC = sNCols(NewHeaders)

4         Result = sArrayStack(NewHeaders, sReshape(0, sNRows(ArrayWithHeaders) - 1, sNCols(NewHeaders)))

5         For i = 2 To sNRows(ArrayWithHeaders)
6             For j = 1 To NC
7                 Result(i, j) = ArrayWithHeaders(i, MatchIDs(j, 1))
8             Next j
9         Next i
10        ShuffleHeaders = Result
11        Exit Function
ErrHandler:
12        Throw "#ShuffleHeaders (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

Sub RemoveAllDFs()
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         For Each ws In ThisWorkbook.Worksheets
3             If IsInCollection(ws.Names, "LiborInit") Then

4                 RemoveDFs ws
5             End If
6         Next

7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#RemoveAllDFs (line " & CStr(Erl) + "): " & Err.Description & "!"

End Sub

Sub TestRemoveDFs()
1         On Error GoTo ErrHandler
2         RemoveDFs ActiveSheet
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestRemoveDFs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub RemoveDFs(ws As Worksheet)
          Dim b As Button
          Dim NumCols As Long
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         If IsCurrencySheet(ws) Then
3             Set SPH = CreateSheetProtectionHandler(ws)

4             ws.Range("Title").Cut ws.Range("SwapRatesInit").Cells(-2, 1)

5             For Each b In ws.Buttons
6                 If b.Caption = "Menu..." Then Exit For
7             Next
8             b.Placement = xlFreeFloating

9             NumCols = ws.Range("SwapRatesInit").Column - 2
10            If NumCols > 0 Then
11                ws.Cells(1, 1).Resize(, NumCols).EntireColumn.Delete
12            End If

13            If IsInCollection(ws.Names, "FundingInit") Then ws.Names("FundingInit").Delete
14            If IsInCollection(ws.Names, "LiborInit") Then ws.Names("LiborInit").Delete

15            If IsInCollection(ws.Names, "DiscountFactorParameters") Then
16                ws.Range("DiscountFactorParameters").Clear
17                ws.Range("DiscountFactorParameters").Cells(0, 1).Clear
18                ws.Names("DiscountFactorParameters").Delete
19            End If

20            FormatCurrencySheet ws, False, Empty

21        End If
22        Exit Sub
ErrHandler:
23        Throw "#RemoveDFs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
