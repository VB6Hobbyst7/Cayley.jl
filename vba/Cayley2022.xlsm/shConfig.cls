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
Private Const BookNameForRegistry = "Cayley3"

Private Function ConfigRange() As Range
1         On Error GoTo ErrHandler
2         Set ConfigRange = sExpandDown(RangeFromSheet(Me, "Headers").Cells(2, 1)).Resize(, 2)
3         Exit Function
ErrHandler:
4         Throw "#ConfigRange (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Sub FormatConfig()
          Dim R As Range
          Dim SPH As clsSheetProtectionHandler

1         On Error GoTo ErrHandler
2         Set R = ConfigRange()
3         Set SPH = CreateSheetProtectionHandler(Me)

4         CayleyFormatAsInput R.Columns(2)
5         R.Columns(1).HorizontalAlignment = xlHAlignRight
6         R.Columns(2).HorizontalAlignment = xlHAlignLeft

7         Exit Sub
ErrHandler:
8         Throw "#FormatConfig (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub FileNameDoubleClick(Target As Range, RegKey As String, Title As String, Optional FileFilter As String)

          Dim FullFileName As String
          Dim RelativePath As String
          Dim Res As Variant
          
1         On Error GoTo ErrHandler
2         FullFileName = sJoinPath(ThisWorkbook.Path, CStr(Target.Value))

3         If sFileExists(FullFileName) Then AddFileToMRU RegKey, FullFileName
4         Res = GetOpenFilenameWrap(RegKey, FileFilter, , Title, , , True)
5         If Res <> False Then
6             RelativePath = sRelativePath(Res, ThisWorkbook.Path)
7             Target.Value = "'" & RelativePath
8         End If

9         Exit Sub
ErrHandler:
10        Throw "#FileNameDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim Res As Variant
          Dim SPH As clsSheetProtectionHandler
          
1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

3         Set SPH = CreateSheetProtectionHandler(Me)

4         Cancel = True
5         Select Case Target.Address
              Case RangeFromSheet(Me, "FxTradesCSVFile").Address
6                 FileNameDoubleClick Target, "CayleyFxTradesFiles", "Choose Fx Trades File", "Fx Trades Files (*.csv),*.csv"
7             Case RangeFromSheet(Me, "RatesTradesCSVFile").Address
8                 FileNameDoubleClick Target, "CayleyRatesTradesFiles", "Choose Rates Trades File", "Rates Trades Files (*.csv),*.csv"
9             Case RangeFromSheet(Me, "AmortisationCSVFile").Address
10                FileNameDoubleClick Target, "CayleyAmortisationFiles", "Choose Amortisation File", "Amortisation Files (*.csv),*.csv"
11            Case RangeFromSheet(Me, "LinesWorkbook").Address
12                FileNameDoubleClick Target, "CayleyLinesWorkbooks", "Choose Lines Workbook", "LinesWorkbooks (*.xlsm),*.xlsm"
13            Case RangeFromSheet(Me, "MarketDataWorkbook").Address
14                FileNameDoubleClick Target, "CayleyMarketDataWorkbooks", "Choose Market Data Workbook", "MarketDataWorkbooks (*MarketData*.xlsm),*MarketData*.xlsm"
15            Case RangeFromSheet(Me, "ScenarioResultsDirectory").Address
16                Res = FolderPicker(Target.Value, , "Select ScenarioResultsDirectory")
17                If VarType(Res) = vbString Then
18                    Target.Value = "'" & Res
19                End If
20            Case RangeFromSheet(Me, "CurrenciesToInclude").Address
21                CurrenciesDoubleclick Target
22            Case Else
23                Cancel = False
24        End Select

25        Exit Sub
ErrHandler:
26        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) & "): " & _
              Err.Description & "!", vbExclamation, "Worksheet_BeforeDoubleClick"
End Sub

Sub CurrenciesDoubleclick(Target As Range)
          Dim FinalChoices
          Dim InitialChoices
          Dim LookupTable
          Dim TheChoices
          Const MainCurrenciesOnly As Boolean = False
          Dim Res

1         On Error GoTo ErrHandler
2         TheChoices = sCurrencies(True, MainCurrenciesOnly)
3         LookupTable = sArrayRange(sCurrencies(False, MainCurrenciesOnly), TheChoices)

4         InitialChoices = sTokeniseString(Target.Value)
5         InitialChoices = sVLookup(InitialChoices, LookupTable)

6         Res = ShowMultipleChoiceDialog(TheChoices, _
              InitialChoices, _
              "Currencies to Include", _
              "Trades in the currencies you select are valued. Trades " & vbLf & "in all other " & _
              "currencies are excluded from valuation.", _
              , Target.offset(0, 1))

7         If Not sArraysIdentical(Res, "#User Cancel!") Then
8             FinalChoices = sVLookup(Res, LookupTable, 1, 2)
9             Target.Value = sConcatenateStrings(FinalChoices)
10        End If

11        Exit Sub
ErrHandler:
12        Throw "#CurrenciesDoubleclick (line " & CStr(Erl) + "): " & Err.Description & "!"
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
7                     FormatConfig
8                 End If
9                 InEvent = False
10            End If
11        End If
12        Exit Sub
ErrHandler:
13        SomethingWentWrong "#Worksheet_Change (line " & CStr(Erl) & "): " & Err.Description & "!"
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
7         CopyOfErr = "#GetFromRegistry (line " & CStr(Erl) & "): " & Err.Description & "!"
8         gBlockChangeEvent = oldBlockChange
9         Throw CopyOfErr
End Sub

Sub SaveToRegistry()
1         On Error GoTo ErrHandler
2         SaveConfigToRegistry ConfigRange(), BookNameForRegistry
3         Exit Sub
ErrHandler:
4         Throw "#SaveToRegistry (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

