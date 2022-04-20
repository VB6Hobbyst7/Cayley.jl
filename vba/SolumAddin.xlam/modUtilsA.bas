Attribute VB_Name = "modUtilsA"
Option Explicit
Private m_StatusBarText As Variant
Private Const m_SuppressCallStackIndicator = "<SDoCS>"
Public g_AppObject As New clsApp
Private m_tictime As Double

'PGS 16 Nov 2017 added 64 bit versions of the declarations, using "Windows API Viewer for MS Excel"
'http://www.rondebruin.nl/win/dennis/windowsapiviewer.htm

Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function SetCursorPos Lib "USER32" (ByVal x As Long, ByVal y As Long) As Long
Declare PtrSafe Sub mouse_event Lib "USER32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As LongPtr)

Public LastShiftF10Time As Double
Public LastAltBacktickTime As Double
Public LastAltBacktickButton As Button
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP As Long = &H10

Private Type SYSTEMTIME
    wYear          As Integer
    wMonth         As Integer
    wDayOfWeek     As Integer
    wDay           As Integer
    wHour          As Integer
    wMinute        As Integer
    wSecond        As Integer
    wMilliseconds  As Integer
End Type

Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (ByRef lpLocalTime As SYSTEMTIME)

'From https://stackoverflow.com/questions/1470632/vba-show-clock-time-with-accuracy-of-less-than-a-second
Private Function NowMilli() As String
          Dim sOut As String
          Dim sThree As String
          Dim sTwo As String
          Dim tTime As SYSTEMTIME
1         On Error GoTo ErrHandler
2         sOut = "yyyy-mm-dd hh:mm:ss.mmm"
3         sTwo = "00": sThree = "000"
4         GetLocalTime tTime
5         Mid(sOut, 1, 4) = tTime.wYear
6         Mid(sOut, 6, 2) = Format$(tTime.wMonth, sTwo)
7         Mid(sOut, 9, 2) = Format$(tTime.wDay, sTwo)
8         Mid(sOut, 12, 2) = Format$(tTime.wHour, sTwo)
9         Mid(sOut, 15, 2) = Format$(tTime.wMinute, sTwo)
10        Mid(sOut, 18, 2) = Format$(tTime.wSecond, sTwo)
11        Mid(sOut, 21, 3) = Format$(tTime.wMilliseconds, sThree)
12        NowMilli = sOut

13        Exit Function
ErrHandler:
14        Throw "#NowMilli (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sExcelWorkingSetSize
' Author    : Philip Swannell
' Date      : 06-Feb-2017
' Purpose   : Returns the amount of memory that Excel is currently using (the so-called "Working Set
'             Size").
' Notes     : copied from http://stackoverflow.com/questions/21189865/out-of-memory-error-trying-to-use-win-api-to-check-memory-usage-in-errorhandle
' -----------------------------------------------------------------------------------------------------------------------
Function sExcelWorkingSetSize()
Attribute sExcelWorkingSetSize.VB_Description = "Returns the amount of memory that Excel is currently using (the so-called ""Working Set Size"")."
Attribute sExcelWorkingSetSize.VB_ProcData.VB_Invoke_Func = " \n28"
          Dim objSWbemServices As Object
1         On Error GoTo ErrHandler
2         Set objSWbemServices = GetObject("winmgmts:")

3         sExcelWorkingSetSize = CDbl(objSWbemServices.Get( _
              "Win32_Process.Handle='" & _
              GetCurrentProcessId & "'").WorkingSetSize)

4         Set objSWbemServices = Nothing
5         Exit Function
ErrHandler:
6         sExcelWorkingSetSize = "#sExcelWorkingSetSize (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : Throw
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : Helps implement consistent error handling...
' -----------------------------------------------------------------------------------------------------------------------
Sub Throw(ByVal ErrorString As String, Optional SuppressCallStackDisplay As Boolean)
1         If SuppressCallStackDisplay Then ErrorString = m_SuppressCallStackIndicator + ErrorString
          '"Out of stack space" errors can lead to enormous error strings, _
           but we cannot handle strings longer than 32767, so just take the right part...
2         If Len(ErrorString) > 32000 Then
3             Err.Raise vbObjectError + 1, , Left$(ErrorString, 1) & Right$(ErrorString, 31999)
4         Else
5             Err.Raise vbObjectError + 1, , Right$(ErrorString, 32000)
6         End If
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ThrowIfError
' Author    : Philip Swannell
' Date      : 24-Jun-2013
' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
'             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
'             methods to be used from VBA code while keeping error handling robust
'             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
' -----------------------------------------------------------------------------------------------------------------------
Function ThrowIfError(Data As Variant)
1         ThrowIfError = Data
2         If VarType(Data) = vbString Then
3             If Left$(Data, 1) = "#" Then
4                 If Right$(Data, 1) = "!" Then
5                     Throw CStr(Data)
6                 End If
7             End If
8         End If
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SomethingWentWrong
' Author    : Philip Swannell
' Date      : 20-May-2015
' Purpose   : Unravels an error string produced by errors burbling up the error handling
'             stack where every method in the stack has the standard error handler:
'             Throw "#NameOfMethod (line " & CStr(Erl) + "): " & Err.Description & "!"
'             then posts a (reasonably) friendly MsgBoxPlus
'        Tip: To avoid this method posting information about the call stack, then use Throw
'             with its second argument - SuppressCallstackDisplay - set to True
' -----------------------------------------------------------------------------------------------------------------------
Sub SomethingWentWrong(ByVal ErrorString As String, Optional Buttons As VbMsgBoxStyle = vbExclamation, Optional Title As String)
          Dim CallStack As String
          Dim MethodName As String
          Dim origErrorString As String
          Dim Prompt As String
          Dim SuppressDisplayOfCallStack As Boolean
          Dim TextWidth As Long
          Dim Res As VbMsgBoxResult

1         On Error GoTo ErrHandler

2         If InStr(ErrorString, m_SuppressCallStackIndicator) > 0 Then
3             ErrorString = Replace(ErrorString, m_SuppressCallStackIndicator, vbNullString)
4             SuppressDisplayOfCallStack = True
5         End If

6         origErrorString = ErrorString

7         Do While InStr(InStr(ErrorString, "#") + 1, ErrorString, "):") > 0
8             MethodName = sStringBetweenStrings(ErrorString, "#", "):") + ")"
9             CallStack = CallStack + vbLf + MethodName
10            ErrorString = Trim$(sStringBetweenStrings(ErrorString, "):"))
11            If Right$(ErrorString, 1) = "!" Then ErrorString = SafeLeft(ErrorString, -1)
12        Loop
13        If Left$(ErrorString, 1) <> "#" Then
14            ErrorString = "#" + ErrorString
15        End If
16        If Right$(ErrorString, 1) <> "!" Then
17            ErrorString = ErrorString + "!"
18        End If

19        If SuppressDisplayOfCallStack Then
20            If Buttons = vbExclamation Then Buttons = vbInformation
21            If Title = vbNullString Then Title = gCompanyName
22            Prompt = ErrorString
23            If Left$(Prompt, 1) = "#" Then Prompt = Mid$(Prompt, 2)
24            If Right$(Prompt, 1) = "!" Then Prompt = Left$(Prompt, Len(Prompt) - 1)
25        Else
26            If Title = vbNullString Then Title = "Uh-oh"
27            Prompt = "Something went wrong:" + vbLf + ErrorString + vbLf + vbLf + _
                  "This was the call stack:" + CallStack
28        End If

29        If InStr(Prompt, vbLf) > 0 Then
30            If LongestSubString(Prompt) < 90 Then

31                TextWidth = 800        'Assume the error message has been formatted to insert line breaks at sensible points, _
                                          the lines can be quite wide so don't squish them down.
32            Else
33                TextWidth = 285
34            End If
35        Else
36            TextWidth = 285
37        End If

38        Res = MsgBoxPlus(Prompt, Buttons + vbOKCancel + vbDefaultButton2, Title, "Copy to clipboard", , , , TextWidth)
If Res = vbOK Then CopyStringToClipboard Prompt

39        Exit Sub
ErrHandler:
40        MsgBoxPlus origErrorString, Buttons, Title
End Sub

Private Function LongestSubString(x As String) As Long
          Dim i As Long
          Dim parts As Variant
          Dim Res As Long

1         parts = VBA.Split(x, vbLf)

2         For i = LBound(parts) To UBound(parts)
3             If Len(parts(i)) > Res Then Res = Len(parts(i))
4         Next
5         LongestSubString = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sElapsedTime
' Author    : Philip Swannell
' Date      : 16-Jun-2013
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
'
'             See http://msdn.microsoft.com/en-us/library/windows/desktop/ms644904(v=vs.85).aspx
' -----------------------------------------------------------------------------------------------------------------------
Function sElapsedTime() As Double
Attribute sElapsedTime.VB_Description = "Retrieves the current value of the performance counter, which is a high resolution (<1us) time stamp that can be used for time-interval measurements."
Attribute sElapsedTime.VB_ProcData.VB_Invoke_Func = " \n28"
          Dim a As Currency
          Static b As Currency
1         On Error GoTo ErrHandler

2         QueryPerformanceCounter a
3         If b = 0 Then QueryPerformanceFrequency b
6         sElapsedTime = a / b

7         Exit Function
ErrHandler:
8         Throw "#sElapsedTime (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedures : tic and toc
' Author     : Philip Swannell
' Date       : 22-Jun-2018
' Purpose    : Timer functions inspired by MATLAB functions of the same names.
' -----------------------------------------------------------------------------------------------------------------------
Sub tic()
1         m_tictime = sElapsedTime()
End Sub
Sub toc(Optional WhatWasTimed As String)
1         Debug.Print (IIf(m_tictime = 0, "Call tic() before calling toc()", IIf(WhatWasTimed = "", "Elapsed time: ", "Elapsed time for " + WhatWasTimed + ": ") & CStr(sElapsedTime() - m_tictime) & " seconds"))
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreateMissing
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Returns a variant of type Missing
' -----------------------------------------------------------------------------------------------------------------------
Function CreateMissing()
1         CreateMissing = CM2()
End Function
Private Function CM2(Optional OptionalArg As Variant)
1         CM2 = OptionalArg
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddLineToAuditSheet
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : Assigned to button on Audit sheet. Can be used from other workbooks too.
' -----------------------------------------------------------------------------------------------------------------------
Sub AddLineToAuditSheet(Optional ws As Worksheet, Optional withDialog As Boolean, Optional Comment As String)
          Dim ColNoAuthor
          Dim ColNoComment
          Dim ColNoDate
          Dim ColNoTime
          Dim ColNoVersion
          Dim HeaderRange As Range
          Dim Headers
          Dim UserName As String

1         On Error GoTo ErrHandler

2         If ws Is Nothing Then Set ws = ActiveSheet

3         ws.Unprotect
4         Application.CutCopyMode = False

5         Application.ScreenUpdating = False
6         With ws
7             If Not IsInCollection(ws.Names, "Headers") Then Throw "Active sheet must have range called Headers with values Version, Date, Time, Author, Comment"
8             Set HeaderRange = ws.Range("Headers")
9             Headers = sArrayTranspose(HeaderRange.Value2)
10            ColNoVersion = sMatch("Version", Headers): If Not IsNumber(ColNoVersion) Then Throw "Header Range must have header titled Version"
11            ColNoDate = sMatch("Date", Headers): If Not IsNumber(ColNoDate) Then Throw "Header Range must have header titled Date"
12            ColNoTime = sMatch("Time", Headers): If Not IsNumber(ColNoTime) Then Throw "Header Range must have header titled Time"
13            ColNoAuthor = sMatch("Author", Headers): If Not IsNumber(ColNoAuthor) Then Throw "Header Range must have header titled Author"
14            ColNoComment = sMatch("Comment", Headers): If Not IsNumber(ColNoComment) Then Throw "Header Range must have header titled Comment"

15            UserName = Application.UserName
16            If UserName = "BBUser" Then
17                UserName = InputBoxPlus("Please enter your name", "Add line to Audit sheet", .Range("Headers").Cells(2, ColNoAuthor))
18                If UserName = "False" Then Exit Sub
19            End If

20            .Range("Headers").Offset(1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
21            .Range("Headers").Cells(2, ColNoVersion).Value = .Range("Headers").Cells(3, ColNoVersion) + 1
22            .Range("Headers").Cells(2, ColNoVersion).NumberFormat = "#,##0;[Red]-#,##0"
23            .Range("Headers").Cells(2, ColNoDate).Value = Date
24            .Range("Headers").Cells(2, ColNoTime).Value = Now() - Date
25            .Range("Headers").Cells(2, ColNoAuthor).Value = UserName
26            .Range("Headers").Rows(2).Font.Bold = False
27            .Range("Headers").Rows(2).Font.ColorIndex = xlColorIndexAutomatic        'in case header row has been formatted to use colors...
28            .Range("Headers").Rows(2).Interior.ColorIndex = xlColorIndexAutomatic

29            With sExpandDown(.Range("Headers"))
30                .HorizontalAlignment = xlHAlignLeft
31                .Columns(1).HorizontalAlignment = xlHAlignCenter
32                .BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
33                With .Borders(xlInsideVertical)
34                    .LineStyle = xlContinuous
35                    .ThemeColor = 1
36                    .TintAndShade = -0.14996795556505
37                    .Weight = xlThin
38                End With
39                With .Borders(xlInsideHorizontal)
40                    .LineStyle = xlContinuous
41                    .ThemeColor = 1
42                    .TintAndShade = -0.14996795556505
43                    .Weight = xlThin
44                End With
45                .VerticalAlignment = xlCenter
46                .Columns(2).NumberFormat = "dd-mmm-yyyy"
47                .Columns(3).NumberFormat = "hh:mm"
48                .Columns(4).WrapText = False 'because prior to 29/5/19 was incorrectly setting to True
49                .Columns(5).WrapText = True
50                If Comment <> vbNullString Then
51                    .Cells(2, 5).Value = "'" + Comment
52                ElseIf withDialog Then
                      Dim AllComments As Variant
                      Dim ChosenComment
53                    AllComments = GetReleaseCommentsFromMRU()
54                    If Not sArraysIdentical(AllComments, "Not found") Then
55                        ChosenComment = ShowSingleChoiceDialog(AllComments, , , , , "Recent Release Comments", "Search")
56                        If Not IsEmpty(ChosenComment) Then
57                            .Cells(2, 5).Value = "'" + ChosenComment
58                        End If
59                    End If
60                End If
61            End With
62        End With

63        Exit Sub
ErrHandler:
64        SomethingWentWrong "#AddLineToAuditSheet (line " & CStr(Erl) + "): " & Err.Description & "!", vbExclamation, "AddLineToAuditSheet"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: GetReleaseCommentsFromMRU
' Purpose:
' Author: Philip Swannell
' Date: 04-Dec-2017
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetReleaseCommentsFromMRU()
          Dim Comments As Variant
1         On Error GoTo ErrHandler
2         Comments = GetSetting(gAddinName, "ReleaseComments", "MRU", "Not found")
3         If Comments <> "Not found" Then
4             Comments = sParseArrayString(CStr(Comments))
5         End If
6         GetReleaseCommentsFromMRU = Comments
7         Exit Function
ErrHandler:
8         Throw "#GetReleaseCommentsFromMRU (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AddAuditSheetToBook
' Author     : Philip Swannell
' Date       : 29-Mar-2018
' Purpose    : Adds an Audit sheet to a workbook, formatted in a consistent way...
' -----------------------------------------------------------------------------------------------------------------------
Sub AddAuditSheetToBook(wb As Excel.Workbook, Optional FirstMessage As String)
          Dim TopLeftCell As Range
          Dim ws As Worksheet
1         On Error GoTo ErrHandler
2         If IsInCollection(wb.Worksheets, "Audit") Then Exit Sub
3         If FirstMessage = vbNullString Then
4             FirstMessage = InputBoxPlus("Please enter a comment for the Audit sheet of '" + wb.Name + "'.", "Audit sheet comment", , , , 400, 60)
5             If FirstMessage = "False" Then Throw "Release aborted", True
6         End If

7         wb.Unprotect
8         Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
9         ActiveWindow.DisplayGridlines = False
10        ActiveWindow.DisplayHeadings = False
11        With ws
12            .Name = "Audit"
13            With .Cells(1, 2)
14                .Value = "Audit"
15                .Font.Size = 22
16            End With
17            With ws.Buttons.Add(0, 0, 10, 10)
18                .Top = ws.Cells(2, 2).Top
19                .Left = ws.Cells(2, 2).Left
20                .Width = 127.8
21                .Height = 28.2
22                .OnAction = "AuditMenu"
23                .caption = "Menu..."
24                .Placement = xlMove
25            End With
26            .Columns(1).ColumnWidth = 1.89
27            .Columns(2).ColumnWidth = 7.11
28            .Columns(3).ColumnWidth = 11.22
29            .Columns(4).ColumnWidth = 4.78
30            .Columns(5).ColumnWidth = 14.11
31            .Columns(6).ColumnWidth = 154.89
32            Set TopLeftCell = .Range("B5")
33            With TopLeftCell.Resize(1, 5)
34                .Value = sArrayRange("Version", "Date", "Time", "Author", "Comment")
35                ws.Names.Add "Headers", .Offset(0)
36            End With
37            If FirstMessage <> vbNullString Then
38                AddLineToAuditSheet ws, False, FirstMessage
39            End If
40        End With
41        Exit Sub
ErrHandler:
42        Throw "#AddAuditSheetToBook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ToggleAutoTrace
' Author    : Philip Swannell
' Date      : 07-Nov-2013
' Purpose   : Switch AutoTracing of arrows on or off - called from the Ribbon
' -----------------------------------------------------------------------------------------------------------------------
Sub ToggleAutoTrace()
1         EnsureAppObjectExists
2         If g_AppObject.AutoTraceIsOn Then
3             g_AppObject.SwitchOffAutoTrace
4         Else
5             g_AppObject.SwitchOnAutoTrace
6         End If
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : EnsureAppObjectExists
' Author    : Philip Swannell
' Date      : 08-Nov-2013
' Purpose   : Unfortunately if ever there is an unhandled error in any VBA code then the
'             g_AppObject ceases to exist :-(. So we need a method to bring it back to life.
' -----------------------------------------------------------------------------------------------------------------------
Sub EnsureAppObjectExists()
1         Application.EnableEvents = True
2         If Not g_AppObject.AppEvents Is Application Then
3             Set g_AppObject.AppEvents = Application
4         End If
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AmendRightClickCommandBar
' Author    : Philip Swannell
' Date      : 08-Nov-2013
' Purpose   : Adds an option to do Paste Duplicate Range to the Cell right-click,
'             but only if Excel is in Copy mode.
' -----------------------------------------------------------------------------------------------------------------------
Sub AmendRightClickCommandBar()
          Dim Bar As CommandBar
          Dim i As Long
          Dim InsertionPoint As Variant
          Dim NewControl As CommandBarButton

          Const caption = "Paste Duplicate Range"

1         On Error GoTo ErrHandler

2         Set Bar = Application.CommandBars("Cell")

3         On Error Resume Next
4         Bar.Controls(caption).Delete
5         On Error GoTo ErrHandler

6         If Application.CutCopyMode <> xlCopy Then Exit Sub

7         For i = 1 To Bar.Controls.Count
8             If InStr(LCase$(Replace(Bar.Controls(i).caption, "&", vbNullString)), "paste") Then
9                 InsertionPoint = i + 1
10            End If
11        Next

12        Set NewControl = Bar.Controls.Add(msoControlButton, , , InsertionPoint, True)
13        With NewControl
14            .caption = caption
15            .OnAction = ThisWorkbook.Name & "!PasteDuplicateRange"
16            .Picture = Application.CommandBars.GetImageMso("Paste", 16, 16)
17            .Style = msoButtonIconAndCaption
18        End With
19        Exit Sub
ErrHandler:
20        Throw "#AmendRightClickCommandBar (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CtrlForwardslashResponse
' Author     : Philip Swannell
' Date       : 06-Aug-2019
' Purpose    : Excel's native Ctrl + forwardslash only works for old-style CSE array formulas. This simulates the same behaviour
'              also supporting new-style dynamic array formulas.
'              This method is only assigned to Ctrl + / in dynamic-array-aware Excel - see method AssignKeys.
' -----------------------------------------------------------------------------------------------------------------------
Sub CtrlForwardslashResponse()
1         On Error GoTo ErrHandler
2         If ActiveSheet Is Nothing Then Exit Sub
3         If ActiveCell Is Nothing Then Exit Sub
4         ExpandRangeToIncludeEntireArrayFormulas(ActiveCell).Select
5         Exit Sub
ErrHandler:
6         SomethingWentWrong "#CtrlForwardslashResponse (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AltBacktickResponse
' Author    : Philip Swannell
' Date      : 14-May-2015
' Purpose   : Run the double-click event of the active sheet, and run macro associated with button overlaying the active cell
' -----------------------------------------------------------------------------------------------------------------------
Sub AltBacktickResponse()
          Dim retVal As Boolean

1         On Error GoTo ErrHandler

2         If ActiveSheet Is Nothing Then Exit Sub
3         If ActiveCell Is Nothing Then Exit Sub
4         LastAltBacktickTime = sElapsedTime()
5         RunButtonAtActiveCell retVal
6         If retVal Then Exit Sub
7         retVal = RunActiveSheetDoubleClickHandler()
8         If retVal Then Exit Sub
9         retVal = RunActiveBookDoubleClickHandler()
10        If retVal Then Exit Sub
11        DoubleClickActiveCell
12        Exit Sub
ErrHandler:
13        SomethingWentWrong "#AltBacktickResponse (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DoubleClickActiveCell
' Author    : Philip Swannell
' Date      : 20-Jun-2017
' Purpose   : Double-click the active cell via windows API calls, partly from https://excelhelphq.com/how-to-move-and-click-the-mouse-in-vba/
' -----------------------------------------------------------------------------------------------------------------------
Private Sub DoubleClickActiveCell()
          'Double click as a quick series of two clicks
          Dim RC As RECT
          Dim x As Double
          Dim y As Double

1         On Error GoTo ErrHandler

2         If Not ActiveCell Is Nothing Then
3             If Not Application.Intersect(ActiveCell, ActiveWindow.VisibleRange) Is Nothing Then
4                 RC = GetRangeRect(ActiveCell)
5                 x = (RC.Left + RC.Right) / 2
6                 y = (RC.Top + RC.Bottom) / 2
7                 SetCursorPos x, y
8                 mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
9                 mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
10                mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
11                mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
12            End If
13        End If
14        Exit Sub
ErrHandler:
15        Throw "#DoubleClickActiveCell (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function RunActiveSheetDoubleClickHandler() As Boolean
          Dim Cancel As Boolean
          Dim MacroName
1         On Error GoTo ErrHandler
2         MacroName = "'" & Replace(ActiveWorkbook.Name, "'", "''") & "'!" & ActiveSheet.CodeName + ".Worksheet_BeforeDoubleClick"
3         Application.Run MacroName, ActiveCell, Cancel
4         RunActiveSheetDoubleClickHandler = True
5         Exit Function
ErrHandler:
6         RunActiveSheetDoubleClickHandler = False
End Function

Private Function RunActiveBookDoubleClickHandler() As Boolean
          Dim Cancel As Boolean
          Dim MacroName
1         On Error GoTo ErrHandler
2         MacroName = "'" & Replace(ActiveWorkbook.Name, "'", "''") & "'!" & ActiveWorkbook.CodeName + ".Workbook_SheetBeforeDoubleClick"
3         Application.Run MacroName, ActiveSheet, ActiveCell, Cancel
4         RunActiveBookDoubleClickHandler = True
5         Exit Function
ErrHandler:
6         RunActiveBookDoubleClickHandler = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RunButtonAtActiveCell
' Author    : Philip Swannell
' Date      : 26-Feb-2016
' Purpose   : Runs the macro associated with the button overlapping the active cell. If
'             more than one button overlaps the active cell then runs macro associated
'             with the button whose centre is closest to the centre of the active cell.
'             retVal is set to True if a method is run
' -----------------------------------------------------------------------------------------------------------------------
Public Sub RunButtonAtActiveCell(Optional retVal As Boolean)   'Must be Public, since called from SCRiPT.xlsm
          Dim acb As Double
          Dim acl As Double
          Dim acr As Double
          Dim act As Double
          Dim b As Button
          Dim bb As Double
          Dim bl As Double
          Dim br As Double
          Dim bt As Double
          Dim bToRun As Button
          Dim MinDist As Double
          Dim ThisDist As Double
          Const Epsilon = 0.0001

1         On Error GoTo ErrHandler
2         If Not ActiveCell Is Nothing Then
3             act = ActiveCell.Top: acb = act + ActiveCell.Height: acl = ActiveCell.Left: acr = acl + ActiveCell.Width
4             act = act + Epsilon: acb = acb - Epsilon: acl = acl + Epsilon: acr = acr - Epsilon
5             For Each b In ActiveSheet.Buttons
6                 If Not (b.OnAction = vbNullString) Then
7                     bt = b.Top: bb = bt + b.Height: bl = b.Left: br = bl + b.Width
8                     bt = bt + Epsilon: bb = bb - Epsilon: bl = bl + Epsilon: br = br - Epsilon
9                     If Intersect(acl, acr, bl, br) Then
10                        If Intersect(act, acb, bt, bb) Then
                              'b overlaps the active cell
11                            If bToRun Is Nothing Then
12                                Set bToRun = b
13                                MinDist = ((bt + bb) / 2 - (act + acb) / 2) ^ 2 + ((bl + br) / 2 - (acl + acr) / 2) ^ 2
14                            Else
15                                ThisDist = ((bt + bb) / 2 - (act + acb) / 2) ^ 2 + ((bl + br) / 2 - (acl + acr) / 2) ^ 2
16                                If ThisDist < MinDist Then
17                                    Set bToRun = b
18                                    MinDist = ThisDist
19                                End If
20                            End If
21                        End If
22                    End If
23                End If
24            Next b

25            If Not bToRun Is Nothing Then
26                retVal = True
27                Set LastAltBacktickButton = bToRun
28                LastAltBacktickTime = sElapsedTime()
29                Application.Run bToRun.OnAction
30            End If
31        End If
32        Exit Sub
ErrHandler:
33        SomethingWentWrong "#RunButtonAtActiveCell (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : KeyboardAccessToButtons
' Author    : Philip Swannell
' Date      : 15-May-2017
' Purpose   : Run code associated with a button on the active sheet, assigned to
'             Ctrl Alt B, for those who really don't like using the mouse.
' -----------------------------------------------------------------------------------------------------------------------
Sub KeyboardAccessToButtons()
          Dim b As Button
          Dim ButtonData As Variant
          Dim FaceIDs As Variant
          Dim Methods As Variant
          Dim Res As Long
          Dim STK As clsStacker
          Dim TheChoices As Variant
          Dim ThisRow As Variant
1         On Error GoTo ErrHandler
2         Set STK = CreateStacker()
3         If Not ActiveCell Is Nothing Then
4             ThisRow = sReshape(vbNullString, 1, 4)
5             For Each b In ActiveSheet.Buttons
6                 If b.OnAction <> vbNullString Then
7                     If Not IsSortButton(b) Then
8                         ThisRow(1, 1) = b.TopLeftCell.row
9                         ThisRow(1, 2) = b.TopLeftCell.Column
10                        ThisRow(1, 3) = b.caption
11                        ThisRow(1, 4) = b.OnAction
12                        STK.Stack2D ThisRow
13                    End If
14                End If
15            Next b
16            ButtonData = STK.Report
17            If Not sIsErrorString(ButtonData) Then
18                ButtonData = sSortedArray(ButtonData, 1, 2)
19                TheChoices = sSubArray(ButtonData, 1, 3, , 1)
20                Methods = sSubArray(ButtonData, 1, 4, , 1)
21                FaceIDs = sReshape(14696, sNRows(Methods), 1)        'FaceID that resembles a button
22                LastAltBacktickTime = sElapsedTime()        ' to arrange that menu appears near active cell, not near mouse pointer
23                Res = ShowCommandBarPopup(TheChoices, FaceIDs, , , ActiveCell, True)
24                If Res > 0 Then
25                    LastAltBacktickTime = sElapsedTime()        'to arrange that menu that might be displayed by code executed appears near active cell, not near mouse pointer
26                    Application.Run (Methods(Res, 1))
27                End If
28            End If
29        End If
30        Exit Sub
ErrHandler:
31        SomethingWentWrong "#KeyboardAccessToButtons (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : intersect
' Author    : Philip Swannell
' Date      : 29-Mar-2017
' Purpose   : Assume a<=b and c<=d returns TRUE iif (a,b) intersects (c,d)
' -----------------------------------------------------------------------------------------------------------------------
Private Function Intersect(a As Double, b As Double, c As Double, D As Double) As Boolean
1         Intersect = Not (b <= c Or D <= a)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShiftF10Response
' Author    : Philip Swannell
' Date      : 20-Jan-2016
' Purpose   : If a worksheet has its own right-click event, we want that also to be triggered
'             by Shift F10 (equivalently the context menu key to the left of the (right of the two) Ctrl keys.)
' -----------------------------------------------------------------------------------------------------------------------
Sub ShiftF10Response()
          Dim Cancel As Boolean
          Dim MacroName
1         On Error GoTo ErrHandler
2         MacroName = "'" & Replace(ActiveWorkbook.Name, "'", "''") & "'!" & ActiveSheet.CodeName + ".Worksheet_BeforeRightClick"
3         LastShiftF10Time = sElapsedTime()
4         Application.Run MacroName, ActiveCell, Cancel

5         Exit Sub
ErrHandler:
6         If Err.Number = 1004 Then        ' there is no custom right-click in the sheet - what about workbook-level or even application level....
7             Application.OnKey "+{F10}"
8             Application.SendKeys "+{F10}"
9             Application.OnTime Now() + TimeValue("00:00:00") / 4, "SwitchF10ResponseOn"
10            Exit Sub
11        End If
12        SomethingWentWrong "#ShiftF10Response (line " & CStr(Erl) + "): " & Err.Description & "!", "ShiftF10Response"
End Sub

Sub SwitchF10ResponseOn()
1         Application.OnKey "+{F10}", "!ShiftF10Response"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GrowScreen
' Author    : Philip Swannell
' Date      : 07-Jun-2016
' Purpose   : On some PCs trying to set the window width to stretch accross both screens
'             fails silently (and sets the window width to the width of one screen) this method
'             somewhat slowly sets the window to stretch across both screens.
' -----------------------------------------------------------------------------------------------------------------------
Sub GrowScreen()
          Dim i As Long
1         Application.SendKeys "% s"
2         For i = 1 To 200
3             Application.SendKeys "{RIGHT}"
4         Next
5         Application.SendKeys "{RETURN}"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FindScreenDimensions
' Author     : Philip Swannell
' Date       : 26-Jun-2018
' Purpose    : Assumes the user has two screens arranged side-by-side and figures out in a rather hacky way the coordinates
'              of the left and right screen, via ByRef arguments. I have a vague memory that I did it this way because GetSystemMetrics was unreliable
' -----------------------------------------------------------------------------------------------------------------------
Sub FindScreenDimensions(ByRef LeftWidth As Long, ByRef LeftHeight As Long, ByRef LeftTop As Long, ByRef LeftLeft As Long, _
        ByRef RightWidth As Long, ByRef RightHeight As Long, ByRef RightTop As Long, ByRef RightLeft As Long, ByRef HaveTwoScreens As Boolean)

          Dim DimensionsArray As Variant
          Dim i As Long
          Dim origWindow As Window
          Dim wb As Excel.Workbook
          Dim wn As Window
          Static LeftScreenHeight As Long
          Static LeftScreenLeft As Long
          Static LeftScreenTop As Long
          Static LeftScreenWidth As Long
          Static ReallyHaveTwoScreens As Boolean
          Static RightScreenHeight As Long
          Static RightScreenLeft As Long
          Static RightScreenTop As Long
          Static RightScreenWidth As Long '

1         On Error GoTo ErrHandler
2         If ActiveWindow Is Nothing Then Exit Sub

3         If LeftScreenWidth = 0 Then
4             Application.EnableEvents = False
5             DimensionsArray = CreateMissing()
6             Application.ScreenUpdating = False
7             Set origWindow = ActiveWindow
8             Set wb = Application.Workbooks.Add
9             Set wn = wb.Windows(1)
10            With wn
                  'We hunt in 3 places for windows, i.e. place a window and maximise it to see where it snaps to

11                For i = 1 To 3
12                    .WindowState = xlNormal
13                    .Width = 100
14                    .Height = 100
15                    .Top = 0
16                    If i = 1 Then
17                        .Left = 0        'hunt at 0,0
18                    ElseIf i = 2 Then        'hunt at 0,-200
19                        .Left = -200
20                    Else
                          'hunt to right of the first window found
21                        .Left = DimensionsArray(1, 2) + DimensionsArray(1, 3) + 1000 '<- this offset needs to be large to cope with the possibility that the _
                                                                                        Windows task bar is on the right of the screen rather than the more _
                                                                                        usual position at the bottom.
22                    End If
23                    .WindowState = xlMaximized
24                    DimensionsArray = sArrayStack(DimensionsArray, sArrayRange(wn.Top, wn.Left, wn.Width, wn.Height))
25                Next i
26                DimensionsArray = sRemoveDuplicateRows(DimensionsArray)
27                DimensionsArray = sSortedArray(DimensionsArray, 2)
28                ReallyHaveTwoScreens = sNRows(DimensionsArray) > 1

29                If ReallyHaveTwoScreens Then
30                    LeftScreenTop = DimensionsArray(1, 1)
31                    LeftScreenLeft = DimensionsArray(1, 2)
32                    LeftScreenWidth = DimensionsArray(1, 3)
33                    LeftScreenHeight = DimensionsArray(1, 4)
34                    RightScreenTop = DimensionsArray(2, 1)
35                    RightScreenLeft = DimensionsArray(2, 2)
36                    RightScreenWidth = DimensionsArray(2, 3)
37                    RightScreenHeight = DimensionsArray(2, 4)
38                Else
39                    LeftScreenTop = DimensionsArray(1, 1)
40                    LeftScreenLeft = DimensionsArray(1, 2)
41                    LeftScreenWidth = DimensionsArray(1, 3) / 2
42                    LeftScreenHeight = DimensionsArray(1, 4)
43                    RightScreenTop = DimensionsArray(1, 1)
44                    RightScreenLeft = DimensionsArray(1, 2) + DimensionsArray(1, 3) / 2
45                    RightScreenWidth = DimensionsArray(1, 3) / 2
46                    RightScreenHeight = DimensionsArray(1, 4)
47                End If
48            End With
49            origWindow.Activate
50            wb.Close False
51            Application.EnableEvents = True
52        End If

53        LeftTop = LeftScreenTop
54        LeftLeft = LeftScreenLeft
55        LeftWidth = LeftScreenWidth
56        LeftHeight = LeftScreenHeight
57        RightTop = RightScreenTop
58        RightLeft = RightScreenLeft
59        RightWidth = RightScreenWidth
60        RightHeight = RightScreenHeight
61        HaveTwoScreens = ReallyHaveTwoScreens

62        Exit Sub
ErrHandler:
63        Throw "#FindScreenDimensions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TemporaryMessage
' Author    : Philip Swannell
' Date      : 05-Oct-2015
' Purpose   : Display text in the Status bar at the bottom of the Excel screen for a set number of seconds.
'             After that time the status bar reverts to RevertTo, or whatever it previously read if RevertTo is omitted.
' -----------------------------------------------------------------------------------------------------------------------
Sub TemporaryMessage(StatusBarText, Optional NumSeconds = 4, Optional RevertTo As Variant)
1         On Error GoTo ErrHandler
2         If IsMissing(RevertTo) Then
3             m_StatusBarText = Application.StatusBar
4         Else
5             m_StatusBarText = RevertTo
6         End If
7         Application.StatusBar = StatusBarText
8         Application.OnTime Now + NumSeconds / 24 / 60 / 60, "RevertStatusBar"

9         MessageLogWrite "StatusBar message:" + vbLf + StatusBarText

10        Exit Sub
ErrHandler:
11        Throw "#TemporaryMessage (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MessageLogWrite
' Author     : Philip Swannell
' Date       : 12-Jan-2017
' Purpose    : Writes to the log file, text written is pre-pended with a date and time stamp,
'             to write multiple lines pass in a string with vbLf characters at the line breaks
' Parameters :
'  TextToWrite    :
'  CloseFileHandle: If TRUE then function is slower, but Excel does not maintain a lock on the file.
' -----------------------------------------------------------------------------------------------------------------------
Sub MessageLogWrite(ByVal TextToWrite As String, Optional CloseFileHandle As Boolean)
          Dim FSO As Scripting.FileSystemObject
          Static LogFileName As String
          Static TS As Scripting.TextStream
          Dim Reopen As Boolean

          Const NumSp = 27        'Changing this? Then also change constant of the same name in method Examine
1         On Error GoTo ErrHandler

2         If TS Is Nothing Or LogFileName = "" Then
3             Reopen = True
4         ElseIf Right$(LogFileName, 14) <> (Format$(Date, "yyyy-mm-dd") & ".txt") Then
5             If Not TS Is Nothing Then TS.Close
6             Reopen = True
7         End If

8         If Reopen Then
9             LogFileName = MessagesLogFileName()        'creates a file if it doesn't already exist
10            Set FSO = New FileSystemObject
11            Set TS = FSO.OpenTextFile(LogFileName, ForAppending, , TristateTrue)
12        End If

13        TextToWrite = NowMilli() + "   " + Replace(TextToWrite, vbLf, vbLf + String(NumSp, " "))

14        TS.WriteLine TextToWrite

15        If CloseFileHandle Then
16            TS.Close
17            Set TS = Nothing
18        End If

19        Exit Sub
ErrHandler:
          'Throwing errors is too disruptive if (say) some other process has a lock on the file.
          'Example: R has a lock on the file when running the ISDA SIMM Correlation Generator
20        Debug.Print "#MessageLogWrite (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : StatusBarWrap
' Author    : Philip Swannell
' Date      : 12-Jan-2017
' Purpose   : The Excel Application StatusBar seems somewhat flakey these days...
'             This method also writes to the log file
' -----------------------------------------------------------------------------------------------------------------------
Sub StatusBarWrap(text As Variant)
          Static LastDoEventsTime As Double
          Static TimeNow As Double
          Const WaitBetweenDoEvents = 5

1         On Error GoTo ErrHandler
2         If Not Application.DisplayStatusBar Then Application.DisplayStatusBar = True

3         If VarType(text) = vbBoolean Then
4             Application.StatusBar = text
5         Else
6             Application.StatusBar = CStr(text)
7             MessageLogWrite "StatusBar message:" + vbLf + CStr(text)
8         End If

9         TimeNow = sElapsedTime()
10        If TimeNow - LastDoEventsTime > WaitBetweenDoEvents Then
11            DoEvents
12        End If

13        Exit Sub
ErrHandler:
14        Throw "#StatusBarWrap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowMessages
' Author    : Philip Swannell
' Date      : 04-Nov-2015
' Purpose   : Let the user see those fleeting temporary messages and messages posted by SomethingWentWrong
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowMessages()
          Dim LogFile As String
1         LogFile = MessagesLogFileName()
2         If sFileExists(LogFile) Then
3             ShowFileInSnakeTail LogFile
4         Else
5             MsgBoxPlus "Cannot find file " + LogFile, vbOKOnly
6         End If
7         Application.OnRepeat "Repeat Show Messages", "ShowMessages"
End Sub

Sub RevertStatusBar()
1         If VarType(m_StatusBarText) <> vbString Then
2             Application.StatusBar = False
3         ElseIf m_StatusBarText = "FALSE" Or m_StatusBarText = vbNullString Then
4             Application.StatusBar = False
5         Else
6             Application.StatusBar = m_StatusBarText
7         End If
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRunningProcesses
' Author    : Philip Swannell
' Date      : 22-Jan-2016
' Purpose   : Returns a list of the running processes on the PC, the same list as appears in the
'             Processes tab of the Windows Task Manager.
' Arguments
' Note      : written with the help of http://stackoverflow.com/questions/26277214/vba-getting-program-names-and-task-id-of-running-processes
' -----------------------------------------------------------------------------------------------------------------------
Public Function sRunningProcesses() As Variant
Attribute sRunningProcesses.VB_Description = "Returns a list of the running processes on the PC, the same list as appears in the Processes tab of the Windows Task Manager."
Attribute sRunningProcesses.VB_ProcData.VB_Invoke_Func = " \n28"
1         Application.Volatile
          Dim objProcessSet As Object
          Dim objServices As Object
          Dim Process As Object
          Dim STK As clsStacker
          Dim strComputer As String
2         On Error GoTo ErrHandler
3         Set STK = CreateStacker()
4         strComputer = "."
5         Set objServices = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
6         Set objProcessSet = objServices.ExecQuery("SELECT Name FROM Win32_Process", , 48)
7         For Each Process In objProcessSet
8             STK.Stack0D Process.Name
9         Next
10        sRunningProcesses = sSortedArray(STK.Report)
11        Set STK = Nothing
12        Set objProcessSet = Nothing
13        Set objServices = Nothing

14        Exit Function
ErrHandler:
15        sRunningProcesses = "#sRunningProcesses (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
