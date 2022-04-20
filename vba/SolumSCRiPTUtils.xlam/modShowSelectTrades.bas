Attribute VB_Name = "modShowSelectTrades"
Option Explicit

Sub TestShowSelectTrades()
1         On Error GoTo ErrHandler
          Dim twb As Workbook

2         If IsInCollection(Application.Workbooks, "SQL_FX IRD portfolio 30 11 2016.xlsx") Then
3             Set twb = Application.Workbooks("SQL_FX IRD portfolio 30 11 2016.xlsx")
4         Else
5             Set twb = Application.Workbooks.Open("C:\SolumWorkbooks\SQL_FX IRD portfolio 30 11 2016.xlsx")
6         End If

          '      ShowSelectTrades twb, "", "", "", "", "", "", Empty
7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#TestShowSelectTrades (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowSelectTrades
' Author    : Philip Swannell
' Date      : 21-Dec-2016
' Purpose   : Displays a dialog for filtering trades from an excel workbook in Airbus
'             Calypso Extract format.
'             twb must be the trades workbook. Arguments passed in used to initialise the dialog.
'             ByRef arguments are set for a subsequent call to GetTradesFromAirbusFile (method in SCRiPT workbook)
' -----------------------------------------------------------------------------------------------------------------------
Function ShowSelectTrades(twb As Workbook, ByRef FilterBy1 As String, ByRef Filter1Value As Variant, ByRef FilterBy2 As String, ByRef Filter2Value As Variant, _
                          ByRef IncludeAssetClasses As String, ByRef CurrenciesToInclude, ByRef CompressTrades As Variant, LinesBook As Workbook)

          Dim TheFrm As FrmSelectTrades
          Static PrevFilterBy1 As String
          Static PrevFilter1Value As Variant
          Static PrevFilterBy2 As Variant
          Static PrevFilter2Value As Variant
          Static PrevIncludeAssetClasses As String
          Static PrevCurrenciesToInclude As String
          Static PrevCompressTrades As Variant
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler

2         If FilterBy1 = "" Then
3             If PrevFilterBy1 = "" Then
4                 FilterBy1 = "CPTY_PARENT"
5             Else
6                 FilterBy1 = PrevFilterBy1
7             End If
8         End If

9         If IsEmpty(Filter1Value) Or sArraysIdentical(Filter1Value, "") Then Filter1Value = PrevFilter1Value

10        If FilterBy2 = "" Then
11            If PrevFilterBy2 = "" Then
12                FilterBy2 = "None"
13            Else
14                FilterBy2 = PrevFilterBy2
15            End If
16        End If

17        If IsEmpty(Filter2Value) Or sArraysIdentical(Filter2Value, "") Then Filter2Value = PrevFilter2Value
18        If IncludeAssetClasses = "" Then
19            If PrevIncludeAssetClasses <> "" Then
20                IncludeAssetClasses = PrevIncludeAssetClasses
21            Else
22                IncludeAssetClasses = "Rates and Fx"
23            End If
24        End If
25        If CurrenciesToInclude = "" Then
26            If PrevCurrenciesToInclude <> "" Then
27                CurrenciesToInclude = PrevCurrenciesToInclude
28            Else
29                CurrenciesToInclude = FirstElementOf(sParseArrayString(GetSetting("SolumConfig", "Cayley", "CurrenciesToInclude", "{""EUR,USD,GBP,CHF""}")))
30            End If
31        End If
32        If VarType(CompressTrades) <> vbBoolean Then
33            If VarType(PrevCompressTrades) = vbBoolean Then
34                CompressTrades = PrevCompressTrades
35            Else
36                CompressTrades = True
37            End If
38        End If

39        Set TheFrm = New FrmSelectTrades
40        TheFrm.Initialise twb, FilterBy1, Filter1Value, FilterBy2, Filter2Value, IncludeAssetClasses, CurrenciesToInclude, CBool(CompressTrades), LinesBook
41        Set SUH = CreateScreenUpdateHandler(True)
42        TheFrm.Show
43        Set SUH = Nothing
44        If TheFrm.ButtonClicked = "OK" Then
45            FilterBy1 = TheFrm.ComboFilterBy1.Value
46            Filter1Value = CastNumberStringToNumber(TheFrm.TextBoxFilter1Value.Value)
47            FilterBy2 = TheFrm.ComboFilterBy2.Value
48            Filter2Value = CastNumberStringToNumber(TheFrm.TextBoxFilter2Value.Value)
49            IncludeAssetClasses = TheFrm.ComboAssetClasses.Value
50            CurrenciesToInclude = TheFrm.TextBoxCurrencies.Value
51            CompressTrades = TheFrm.CheckBoxCompress.Value
52            PrevFilterBy1 = FilterBy1
53            PrevFilter1Value = Filter1Value
54            PrevFilterBy2 = FilterBy2
55            PrevFilter2Value = Filter2Value
56            PrevCurrenciesToInclude = CurrenciesToInclude
57            PrevCompressTrades = CompressTrades
58            PrevIncludeAssetClasses = IncludeAssetClasses
59            ShowSelectTrades = "OK"
60        Else
61            ShowSelectTrades = "#User Cancel!"
62        End If
63        Set TheFrm = Nothing
64        Exit Function
ErrHandler:
65        Throw "#ShowSelectTrades (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Private Function CastNumberStringToNumber(TheInput As String)
1         On Error GoTo ErrHandler
2         CastNumberStringToNumber = CDbl(TheInput)
3         Exit Function
ErrHandler:
4         CastNumberStringToNumber = TheInput
End Function

