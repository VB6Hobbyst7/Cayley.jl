Attribute VB_Name = "modHistoricData"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsBloombergInstalled
' Author    : Philip
' Date      : 05-Oct-2017
' Purpose   : Test if Bloomberg Addin is installed. Would be better to also test if data
'             is available (e.g. what if user is not logged in to Bloomberg?)
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsBloombergInstalled() As Boolean
          Dim Res
1         On Error GoTo ErrHandler
2         Res = Application.Evaluate("=BToday(TRUE)")
3         IsBloombergInstalled = Not (IsError(Res))

4         Exit Function
ErrHandler:
5         Throw "#IsBloombergInstalled (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PrevWeekDay
' Author     : Philip Swannell
' Date       : 09-Apr-2020
' Purpose    : Returns InputDate if that's a weekday or preceding Friday otherwise.
' -----------------------------------------------------------------------------------------------------------------------
Private Function PrevWeekDay(InputDate)
1         On Error GoTo ErrHandler
2         Select Case InputDate Mod 7
              Case 0 'Sat
3                 PrevWeekDay = InputDate - 1
4             Case 1 'Sun
5                 PrevWeekDay = InputDate - 2
6             Case Else
7                 PrevWeekDay = InputDate
8         End Select
9         Exit Function
ErrHandler:
10        Throw "#PrevWeekDay (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FollWeekDay
' Author     : Philip Swannell
' Date       : 09-Apr-2020
' Purpose    : Returns InputDate if thats a weekday or following Monday if weekend
' -----------------------------------------------------------------------------------------------------------------------
Private Function FollWeekDay(InputDate)
1         On Error GoTo ErrHandler
2         Select Case InputDate Mod 7
              Case 0 'Sat
3                 FollWeekDay = InputDate + 2
4             Case 1 'Sun
5                 FollWeekDay = InputDate + 1
6             Case Else
7                 FollWeekDay = InputDate
8         End Select
9         Exit Function
ErrHandler:
10        Throw "#FollWeekDay (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MenuHistoricData
' Author     : Philip Swannell
' Date       : 07-Jan-2021
' Purpose    : Attached to "Menu..." button on HistoricData sheet
' -----------------------------------------------------------------------------------------------------------------------
Sub MenuHistoricData()
          Dim Choices
          Dim Chosen
          Dim Enabled
          Dim FaceIDs

1         On Error GoTo ErrHandler

2         RunThisAtTopOfCallStack

3         If IsBloombergInstalled Then
4             Choices = "Update Historic Data from Bloomberg..."
5             FaceIDs = 349
6             Enabled = True
7         Else
8             Choices = "Update Historic Data from Bloomberg. Disabled since Bloomberg Addin is not available."
9             FaceIDs = 349
10            Enabled = False
11        End If

12        Chosen = ShowCommandBarPopup(Choices, FaceIDs, Enabled, , ChooseAnchorObject(), True)
13        If Chosen = 1 Then
14            UpdateHistoricData
15        End If

16        Exit Sub
ErrHandler:
17        SomethingWentWrong "#MenuHistoricData (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : UpdateHistoricData
' Author     : Philip Swannell
' Date       : 07-Jan-2021
' Purpose    : Updates the data on sheet HistoricalDatafor EURUSD spot and 3 year Fx vol. Requires Bloomberg addin to be
'              installed. Appends data to data already on sheet to minimisa BBG data usage.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub UpdateHistoricData()
          Dim ChooseVector
          Dim DatesToGet
          Dim FirstDateToGet As Long
          Dim Formula1 As String
          Dim Formula2 As String
          Dim i As Long
          Dim LastDateToGet As Long
          Dim Prompt As String
          Dim SPH As clsSheetProtectionHandler
          Dim SUH As clsScreenUpdateHandler
          Dim TargetRange As Range
          Const Title = "Update HistoricData"

1         On Error GoTo ErrHandler

2         If IsBloombergInstalled() Then
3             Prompt = "Do you want to bring the data on this sheet up to date?"
4             If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, Title) <> vbOK Then Exit Sub
5         Else
6             Prompt = "This functionality to bring the data on this sheet up to date requires the Bloomberg Addin for Excel, but that does not seem to be installed."
7             MsgBoxPlus Prompt, vbOKOnly + vbExclamation, Title
8             Exit Sub
9         End If

10        FirstDateToGet = FollWeekDay(RangeFromSheet(shHistoricalData, "TheDates").Cells(1, 1).End(xlDown).Value + 1)
11        LastDateToGet = PrevWeekDay(Date - 1)

12        If LastDateToGet <= FirstDateToGet Then Throw "Data is already up to date", True

13        Set SUH = CreateScreenUpdateHandler()
14        Set SPH = CreateSheetProtectionHandler(shHistoricalData)

15        DatesToGet = sArrayAdd(FirstDateToGet - 1, sIntegers(LastDateToGet - FirstDateToGet + 1))

16        ChooseVector = sReshape(True, sNRows(DatesToGet), 1)
17        For i = 1 To sNRows(ChooseVector)
18            If DatesToGet(i, 1) Mod 7 <= 1 Then ChooseVector(i, 1) = False
19        Next
20        DatesToGet = sMChoose(DatesToGet, ChooseVector)

21        Set TargetRange = RangeFromSheet(shHistoricalData, "TheDates").Cells(1, 1).End(xlDown).offset(1).Resize(sNRows(DatesToGet), 3)

22        Formula1 = "=BDH(""EURUSD Curncy"",""PX_LAST"",StartDate,EndDate,""ARRAY=TRUE"",""Days=W"")"
23        Formula1 = Replace(Formula1, "StartDate", CStr(FirstDateToGet))
24        Formula1 = Replace(Formula1, "EndDate", CStr(LastDateToGet))
25        TargetRange.Resize(, 2).FormulaArray = Formula1

26        Formula2 = "=BDH(""EURUSDV3Y Curncy"",""PX_LAST"",StartDate,EndDate,""ARRAY=TRUE"",""Days=W"",""Dates=FALSE"",""Factor=0.01"")"
27        Formula2 = Replace(Formula2, "StartDate", CStr(FirstDateToGet))
28        Formula2 = Replace(Formula2, "EndDate", CStr(LastDateToGet))

29        TargetRange.Resize.offset(, 2).Resize(, 1).FormulaArray = Formula2

30        While Not isCalculated(TargetRange)
31            TargetRange.Calculate
32            For i = 1 To 100
33                DoEvents
34            Next i
35            TargetRange.Calculate
36        Wend

37        If Not sAll(sArrayIsNumber(TargetRange.Value2)) Then
38            TargetRange.Clear
39            Throw "Calls to Bloomberg function BDH returned non-numeric values. Is Bloomberg Addin installed and are you logged in to Bloomberg?"
40        End If

41        If Not sArraysIdentical(DatesToGet, TargetRange.Columns(1).Value2) Then Throw "Calls to Bloomberg function BDH did not return the expected set of weekdays in the first column"

42        TargetRange.Value = TargetRange.Value

43        FormatDatesAndFixGraph

44        Prompt = "All done. Imported data from " & Format(TargetRange.Cells(1, 1).Value, "d-mmm-yyyy") & " to " & _
              Format(TargetRange.Cells(TargetRange.Rows.Count, 1).Value, "d-mmm-yyyy") & " to range " & Replace(TargetRange.Address, "$", "") & vbLf & vbLf & _
              "Please save this workbook so that the new data persists."

45        MsgBoxPlus Prompt, vbOKOnly + vbInformation, Title

46        Exit Sub
ErrHandler:
47        SomethingWentWrong "#UpdateHistoricData (line " & CStr(Erl) & "): " & Err.Description & "!", , Title
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FormatDatesAndFixGraph
' Author     : Philip Swannell
' Date       : 07-Jan-2021
' Purpose    : Sub of UpdateHistoricData, updates chart, fixes up number formatting and one range name
' -----------------------------------------------------------------------------------------------------------------------
Private Sub FormatDatesAndFixGraph()

          Dim SourceRange As Range
          Dim SPH As clsSheetProtectionHandler
1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(shHistoricalData)

3         Set SourceRange = RangeFromSheet(shHistoricalData, "TheDates").Cells(1, 1)
4         Set SourceRange = Range(SourceRange, SourceRange.End(xlDown).offset(0, 2))
5         shHistoricalData.Names.Add "TheDates", SourceRange.Columns(1)

6         SourceRange.ClearFormats
7         SourceRange.Columns(1).NumberFormat = "dd-mmm-yyyy"
8         AddGreyBorders SourceRange, True

9         With shHistoricalData.ChartObjects("Chart 1").Chart
10            .FullSeriesCollection(1).xValues = "=" & shHistoricalData.Name & "!" & SourceRange.Columns(1).Address
11            .FullSeriesCollection(1).Values = "=" & shHistoricalData.Name & "!" & SourceRange.Columns(2).Address
12            .FullSeriesCollection(2).xValues = "=" & shHistoricalData.Name & "!" & SourceRange.Columns(1).Address
13            .FullSeriesCollection(2).Values = "=" & shHistoricalData.Name & "!" & SourceRange.Columns(3).Address
14            .Axes(xlCategory).MinimumScale = SourceRange.Cells(1, 1).Value2
15            .Axes(xlCategory).MaximumScale = SourceRange.Cells(SourceRange.Rows.Count, 1).Value2
16        End With

17        Exit Sub
ErrHandler:
18        Throw "#FormatDatesAndFixGraph (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : isCalculated
' Author     : Philip Swannell
' Date       : 07-Jan-2021
' Purpose    : Test for whether a range of cells containing calls to BDH function have "resolved"
' -----------------------------------------------------------------------------------------------------------------------
Private Function isCalculated(R As Range) As Boolean
          Dim c As Range

1         For Each c In R.Cells
2             If InStr(CStr(c.Value), "Requesting") > 0 Then Exit Function
3         Next
4         isCalculated = True
End Function

