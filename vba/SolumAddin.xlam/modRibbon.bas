Attribute VB_Name = "modRibbon"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modRibbon
' Author    : Philip Swannell
' Date      : 28-Nov-2015
' Purpose   : Moved Ribbon code to this module. Implemented enabling and disabling of the ribbon
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Public g_rbxIRibbonUI As IRibbonUI
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef Source As Any, ByVal length As Long)

Private Function ObjectFromPointer(ByVal lRibbonPointer As LongPtr) As Object
          Dim TheObject As Object
1         CopyMemory TheObject, lRibbonPointer, LenB(lRibbonPointer)
2         Set ObjectFromPointer = TheObject
3         Set TheObject = Nothing
End Function

Public Sub TheRibbon_OnLoad(ribbon As IRibbonUI)
1         On Error GoTo ErrHandler
          ' Code for onLoad callback. Ribbon control customUI
2         Set g_rbxIRibbonUI = ribbon
          'Keep a copy of the pointer to the Ribbon so that if global variables get lost (e.g. an unhandled error is encountered) _
           then we can use ObjectFromPointer to get the ribbon back. _
           Trick copied from "losing the state of the global IRibbonUI ribbon object" _
           Ron de Bruin, http://www.rondebruin.nl/win/s2/win015.htm
3         shAudit.Range("PointerToRibbon") = ObjPtr(ribbon)

4         Exit Sub
ErrHandler:
5         SomethingWentWrong "#TheRibbon_OnLoad (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RefreshRibbon
' Author    : Philip Swannell
' Date      : 04-Dec-2015
' Purpose   : Call to "refresh" the ribbon, in fact the ribbon is marked as invalid so all
'             the callbacks for getEnabled etc happen the next time they "need" to happen
' -----------------------------------------------------------------------------------------------------------------------
Sub RefreshRibbon()
1         On Error GoTo ErrHandler
2         If g_rbxIRibbonUI Is Nothing Then
3             If Not IsEmpty(shAudit.Range("PointerToRibbon")) Then
4                 Set g_rbxIRibbonUI = ObjectFromPointer(shAudit.Range("PointerToRibbon"))
5             End If
6         End If
7         If Not g_rbxIRibbonUI Is Nothing Then
8             On Error Resume Next
9             g_rbxIRibbonUI.Invalidate
10            On Error GoTo ErrHandler
11        End If
12        Exit Sub
ErrHandler:
13        SomethingWentWrong "#RefreshRibbon (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UnhandledError
' Author    : Philip Swannell
' Date      : 29-Feb-2016
' Purpose   : Run this method to demonstrate that all global variables get wiped when an
'             unhandled error is encountered and check that the ribbon refreshing still works afterwards
' -----------------------------------------------------------------------------------------------------------------------
Private Sub UnhandledError()
          Dim i As Long
1         i = "Foo"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ActivateMainTab
' Author     : Philip Swannell
' Date       : 04-Nov-2018
' Purpose    : Activates the Solum tab of the ribbon
' -----------------------------------------------------------------------------------------------------------------------
Sub ActivateMainTab()
1         On Error GoTo ErrHandler
2         If g_rbxIRibbonUI Is Nothing Then
3             If Not IsEmpty(shAudit.Range("PointerToRibbon")) Then
4                 Set g_rbxIRibbonUI = ObjectFromPointer(shAudit.Range("PointerToRibbon"))
5             End If
6         End If
7         If Not g_rbxIRibbonUI Is Nothing Then
8             On Error Resume Next
9             g_rbxIRibbonUI.ActivateTab "MainTab"
10            On Error GoTo ErrHandler
11        End If
12        Exit Sub
ErrHandler:
13        Throw "#ActivateMainTab (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Public Sub TheRibbon_getLabel(control As IRibbonControl, ByRef returnedVal)
1         On Error GoTo ErrHandler
2         Select Case control.Tag
              Case "MainTab"
3                 Select Case Val(Application.Version)
                      Case 15        'Office 2013
4                         returnedVal = UCase$(gCompanyName)        ' I think office 2013 is the only version that used ALLCAPS in the ribbon tabs, _
                                                                      Excel 2016 has reverted to Proper
5                     Case Else
6                         returnedVal = gCompanyName
7                 End Select
8             Case "ToggleAutoTrace"
9                 EnsureAppObjectExists
10                If g_AppObject.AutoTraceIsOn Then
11                    returnedVal = "AutoTrace" + vbLf + "is On" + vbLf
12                Else
13                    returnedVal = "AutoTrace" + vbLf + "is Off" + vbLf
14                End If
15            Case "TogglePageBreaks"
16                If ActiveSheet Is Nothing Then
17                    returnedVal = "Show Page Brea&ks (Ctrl Shift K)"
18                ElseIf ActiveSheet.DisplayPageBreaks Then
19                    returnedVal = "Hide Page Brea&ks (Ctrl Shift K)"
20                Else
21                    returnedVal = "Show Page Brea&ks (Ctrl Shift K)"
22                End If
23            Case "LeftGroup"
                  Dim FormatString As String
                  Dim ReleaseDate
24                ReleaseDate = ThrowIfError(sAddinReleaseDate())
25                If Application.WorksheetFunction.RoundDown(ReleaseDate, 0) = Date Then
26                    FormatString = "dd-mmm-yyyy hh:mm"
27                Else
28                    FormatString = "dd-mmm-yyyy"
29                End If
30                If Year(ReleaseDate) = Year(Date) Then
31                    FormatString = Replace(FormatString, "-yyyy", vbNullString)
32                End If
33                returnedVal = Format$(ThrowIfError(sAddinReleaseDate()), FormatString)
34            Case "FixLinksForActiveBook"
35                returnedVal = "Fi&x Links to " & gAddinName
36            Case Else
37                MsgBoxPlus "Unrecognised IRRibbonControl.Tag in call to TheRibbon_getLabel: " + control.Tag, vbExclamation, gAddinName
38        End Select

39        Exit Sub
ErrHandler:
40        SomethingWentWrong "#TheRibbon_getLabel (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TheRibbon_getImage
' Author    : Philip Swannell
' Date      : 04-Dec-2015
' Purpose   : Grabs image from the "Custom Icons" sheet to use on the ribbon!
' -----------------------------------------------------------------------------------------------------------------------
Public Sub TheRibbon_getImage(control As IRibbonControl, ByRef returnedVal)
1         On Error GoTo ErrHandler
2         Select Case control.Tag
              Case "WhiteBoldOnBlue", "AddSortButtons0", "AddSortButtons1", "AddSortButtonsMore", "Numeric Array"
3                 shCustomIcons.Shapes(control.Tag).CopyPicture xlScreen, xlBitmap
4                 Set returnedVal = PastePicture(xlBitmap)
5             Case "ToggleAutoTrace"
6                 EnsureAppObjectExists
7                 If g_AppObject.AutoTraceIsOn Then
8                     returnedVal = "TracePrecedents"
9                 Else
10                    returnedVal = "TracePrecedentsRemoveArrows"
11                End If
12            Case Else
13                Throw "Unrecognised IRRibbonControl.Tag: " + control.Tag
14        End Select

15        Exit Sub
ErrHandler:
16        SomethingWentWrong "#TheRibbon_getImage (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TheRibbon_onAction
' Author    : Philip Swannell
' Date      : 15-Oct-2013
' Purpose   : Callback for various Ribbon elements
' -----------------------------------------------------------------------------------------------------------------------
Sub TheRibbon_onAction(control As IRibbonControl)
          Dim AllXLs As Collection
          Dim MatchRes As Variant
1         On Error GoTo ErrHandler
2         EnsureAppObjectExists
3         AssignKeys True
4         SetApplicationCaptions
          'Application Cursor and EnableEvents can be messed up during debugging sessions (by stopping code). Doesn't hurt to fix it here...
5         Application.Cursor = xlDefault
6         Application.EnableEvents = True
7         MessageLogWrite gCompanyName & " Ribbon: " + control.Tag
8         If Left$(control.Tag, 13) = "ActivateBook_" Then
              Dim AppNumber As Long
              Dim BookName As String
9             AppNumber = sStringBetweenStrings(control.Tag, "_", "_")
10            BookName = sStringBetweenStrings(control.Tag, CStr(AppNumber) & "_")
11            If AppNumber = 1 Then
12                ActivateBook BookName
13                If Application.CommandBars("Ribbon").Height > 100 Then
14                    ActivateMainTab
15                    Application.OnTime Now, ThisWorkbook.Name & "!ActivateMainTab" 'Try again!
16                End If
17            Else
18                GetExcelInstances AllXLs
19                ActivateBook AllXLs(AppNumber).Workbooks(BookName)
20            End If
21            Exit Sub
22        ElseIf Left$(control.Tag, 14) = "ActivateSheet_" Then
23            ActivateSheet Right$(control.Tag, Len(control.Tag) - 14)
24            Exit Sub
25        End If

26        Select Case control.Tag
              Case "AddSortButtons0"
27                AddSortButtons , 0
28            Case "AddSortButtons1"
29                AddSortButtons , 1
30            Case "AddSortButtonsMore"
31                AddSortButtons , -1
32            Case "RemoveSortButtons"
33                RemoveSortButtons
34            Case "InsertFileNames"
35                InsertFileNames
36            Case "InsertFolderName"
37                InsertFolderName
38            Case "ResizeArrayFormula"
39                ResizeArrayFormula
40            Case "FitArrayFormula"
41                FitArrayFormula
42            Case "ShowHelpBrowser"
43                ShowHelpBrowser
44            Case "ToggleAutoTrace"
45                ToggleAutoTrace
46                RefreshRibbon
47            Case "PasteDuplicateRange"
48                PasteDuplicateRange
49            Case "PasteValues"
50                PasteValues
51            Case "FormatAsInput"
52                FormatAsInput
53            Case "FlipNumberFormats"
54                FlipNumberFormats
55            Case "LocaliseGlobalNames"
56                LocaliseGlobalNames
57            Case "CalcSelection"
58                CalcSelection
59            Case "AboutMe"
60                AboutMe
61            Case "AddGroupingButton"
62                AddGroupingButtons
63            Case "ShowTemporaryMessages"
64                ShowMessages
65            Case "ToggleWindow"
66                ToggleWindow
67            Case "SwitchBook"
68                SwitchBook
69            Case "SwitchSheet"
70                SwitchSheet
71            Case "CalcActiveSheet"
72                CalcActiveSheet
73            Case "SearchWorkbookFormulas"
74                SearchWorkbookFormulas
75            Case "GreyBorders"
76                AddGreyBorders , IsShiftKeyDown()
77            Case "WhiteBoldOnBlue"
78                WhiteBoldOnBlue
79            Case "QuitExcel"
80                QuitExcel
81            Case "FixLinksForActiveBook"
82                FixLinksForActiveBook
83            Case "Preferences"
84                ThisAddinPreferences
85            Case "ArrangeWindows"
86                ArrangeWindows
87            Case "TogglePageBreaks"
88                TogglePageBreaks
89            Case "JustifyText"
90                JustifyText
91            Case "OpenTextFileNoRecalc"
92                OpenTextFileNoRecalc
93            Case Else
94                MatchRes = sMatch(control.Tag, shHelp.Range("TheData").Columns(1).Value, True)
95                If IsNumber(MatchRes) Then
96                    ShowHelpForFunction control.Tag, IsExcelReadyForFormula(True)
97                    RefreshRibbon 'Not sure why this is necessary, but the ribbon element mysteriously greys when the function wizard appears out and needs to be kicked back to life
98                Else
99                    MsgBoxPlus "Unrecognised IRRibbonControl.Tag in call to TheRibbon_OnAction: " + control.Tag, vbExclamation, gAddinName
100               End If
101       End Select
102       Exit Sub
ErrHandler:
103       SomethingWentWrong "#TheRibbon_onAction (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TheRibbon_getVisible
' Author    : Philip Swannell
' Date      : 09-Nov-2016
' Purpose   : Allow us to hide or show ribbon elements. Currently used only for the Quit Excel element
' -----------------------------------------------------------------------------------------------------------------------
Sub TheRibbon_getVisible(control As IRibbonControl, ByRef returnedVal)
1         On Error GoTo ErrHandler
2         Select Case control.Tag
              Case "QuitExcel"
3                 returnedVal = Val(Application.Version) > 14        'Excel 2010 (=14) has its own "Quit" option
4             Case Else
5                 returnedVal = True
6         End Select

7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#TheRibbon_getVisible (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TheRibbon_getEnabled
' Author    : Philip Swannell
' Date      : 04-Dec-2015
' Purpose   : All the controls of the ribbon use this method to get their enabled status.
' -----------------------------------------------------------------------------------------------------------------------
Sub TheRibbon_getEnabled(control As IRibbonControl, ByRef returnedVal)
1         On Error GoTo ErrHandler

2         Select Case control.Tag
              Case "StartAutoTrace"
3                 EnsureAppObjectExists
4                 returnedVal = Not (g_AppObject.AutoTraceIsOn)
5             Case "StopAutoTrace"
6                 EnsureAppObjectExists
7                 If g_AppObject.AutoTraceIsOn Then
8                     returnedVal = True
9                 Else
10                    returnedVal = False
11                End If
12            Case "SwitchBook", "SwitchDialogLaunch"
13                If InDeveloperMode() Then
14                    returnedVal = True
15                Else
16                    returnedVal = Not (ActiveSheet Is Nothing)
17                End If
18            Case "AboutMe", "ShowTemporaryMessages", "OpenTextFileNoRecalc"
19                returnedVal = True        'Always show these two
20            Case Else
21                returnedVal = AnyVisibleWindows()
22        End Select
23        Exit Sub
ErrHandler:
24        SomethingWentWrong "#TheRibbon_getEnabled (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Private Function AnyVisibleWindows() As Boolean
          Dim w As Window
1         On Error Resume Next
2         For Each w In Application.Windows
3             If w.Visible Then
4                 AnyVisibleWindows = True
5                 Exit For
6             End If
7         Next
8         On Error GoTo 0
End Function

Public Sub TheRibbon_getSupertip(control As IRibbonControl, ByRef returnedVal)
          Dim MatchRes As Variant
1         On Error GoTo ErrHandler
2         Select Case control.Tag
              Case "SwitchBook"
3                 If InDeveloperMode() Then
4                     returnedVal = "Quickly switch to another open workbook, or to an addin." + vbLf + vbLf + "Lets you switch to workbooks in all running Excel instances."
5                 Else
6                     returnedVal = "Quickly switch to another open workbook, even if it's open in a different Excel instance." + vbLf + vbLf + "When " + gAddinName + " is in ""Developer Mode"" (see Preferences), you can also switch to open addins."
7                 End If
8             Case "SwitchSheet"
9                 If InDeveloperMode() Then
10                    returnedVal = "Quickly switch to another sheet of the active workbook, even hidden or very hidden ones."
11                Else
12                    returnedVal = "Quickly switch to another sheet of the active workbook, even hidden ones." + vbLf + vbLf + "When " + gAddinName + " is in ""Developer Mode"" (see Preferences), you can also switch to very hidden sheets."
13                End If
14            Case "AboutMe"
15                returnedVal = "You have version " + Format$(sAddinVersionNumber, "###,###") + _
                      " of " & gAddinName & ", which was released " + Format$(sAddinReleaseDate, "d-mmm-yyyy hh:mm")
16            Case "HyperlinksVerify"
17                returnedVal = "Ensure that Excel links in the active workbook to " & gAddinName & " and " & gAddinName2 & " point to the correct place. Also handles VBA References to these two addins."
18            Case "ShowTemporaryMessages"
19                returnedVal = "See a list of messages posted today by " & gAddinName & ". Two styles of messages appear: a) text that was shown in the status bar at the bottom of the Excel screen; and b) text that was shown in dialog boxes."
20            Case "ShowHelpBrowser"
21                returnedVal = "Browse the help for all " & gAddinName & " functions and utilities. Choose a function and insert it into your sheet."
22            Case "FixLinksForActiveBook"
23                returnedVal = "Ensure that Excel links in the active workbook to " & gAddinName & " and " & gAddinName2 & " point to the correct place. Also handles VBA References to these two addins."
24            Case "Array", "Comparison and Selection", "File", "Maths", "Numeric Array", "Options", "Range", "String", "Utilities"
25                returnedVal = "Add a " + gAddinName + " " + control.Tag + " function to your worksheet."
26            Case "TogglePageBreaks"
27                returnedVal = "Page Breaks on all sheets. Alternative to Excel's File > Options > Advanced > Display options for this worksheet > Show page breaks."
28                If ActiveSheet Is Nothing Then
29                    returnedVal = "Show " & returnedVal
30                ElseIf ActiveSheet.DisplayPageBreaks Then
31                    returnedVal = "Hide " & returnedVal
32                Else
33                    returnedVal = "Show " & returnedVal
34                End If
35            Case "ResizeArrayFormula"
36                If ExcelSupportsSpill() Then
37                    returnedVal = "Change the formula at the active cell to be to be an old-style Ctrl-Shift-Enter (CSE) array formula the same size as the current selection. Undo with Ctrl+Z." + vbLf + vbLf + _
                          "CSE array formulas have been superceded by Dynamic Array Formulas, so this functionality is mostly redundant."
38                Else
39                    returnedVal = "Change the formula at the active cell to be to be an array formula the same size as the current selection. Undo with Ctrl+Z."
40                End If
41            Case "FitArrayFormula"
42                If ExcelSupportsSpill() Then
43                    returnedVal = "Replace old-style Ctrl-Shift-Enter array formula at the active cell with an equivalent new-style Dynamic Array Formula. Undo with Ctrl+Z."
44                Else
45                    returnedVal = "Change the formula at the active cell to be an array formula the same size as the array that the formula returns. Undo with Ctrl+Z."
46                End If
47            Case Else
48                MatchRes = sMatch(control.Tag, shHelp.Range("TheData").Columns(1).Value, True)
49                If IsNumber(MatchRes) Then
50                    returnedVal = shHelp.Range("TheData").Cells(MatchRes, 7).Value ' + vbLf + String(55, "_") + vbLf + _
                                                                                       "Select to insert formula." + vbLf + "Shift select for more help."
51                Else
52                    MsgBoxPlus "Unrecognised IRRibbonControl.Tag in call to TheRibbon_getSupertip: " + control.Tag, vbExclamation, gAddinName
53                End If
54        End Select

55        Exit Sub
ErrHandler:
56        SomethingWentWrong "#TheRibbon_getSupertip (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Public Sub TheRibbon_getScreentip(control As IRibbonControl, ByRef returnedVal)
1         On Error GoTo ErrHandler
          Dim MatchRes As Variant
2         Select Case control.Tag
              Case "AboutMe"
3                 returnedVal = "About " & gAddinName
4             Case "Preferences"
5                 returnedVal = gAddinName & " Preferences"
6             Case Else
7                 MatchRes = sMatch(control.Tag, shHelp.Range("TheData").Columns(1).Value, True)
8                 If IsNumber(MatchRes) Then
9                     returnedVal = shHelp.Range("TheData").Cells(MatchRes, 2).Value
10                Else
11                    MsgBoxPlus "Unrecognised IRRibbonControl.Tag in call to TheRibbon_getScreentip: " + control.Tag, vbExclamation, gAddinName
12                End If
13        End Select

14        Exit Sub
ErrHandler:
15        SomethingWentWrong "#TheRibbon_getScreentip (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TheRibbon_getContent
' Author    : Philip Swannell
' Date      : 25-Oct-2018
' Purpose   : Call back for dynamic menus.
' -----------------------------------------------------------------------------------------------------------------------
Sub TheRibbon_getContent(control As IRibbonControl, ByRef returnedVal)
          Dim Category As String
          Dim i As Long
          Dim R As Range
          Dim SA As clsStringAppend

1         On Error GoTo ErrHandler
2         Set SA = New clsStringAppend
3         SetApplicationCaptions
4         If Application.Cursor <> xlDefault Then Application.Cursor = xlDefault

5         Select Case control.Tag
              Case "Array", "Maths", "Numeric Array", "Comparison and Selection", "String", "Options", "Utilities", "File", "Range"

6                 Category = control.Tag

                  Const ButtonTemplate = "<button  getScreentip = ""TheRibbon_getScreentip"" getSupertip=""TheRibbon_getSupertip"" onAction=""TheRibbon_onAction"" id=""fnTHE_COUNTER"" tag=""THE_FUNCTION"" label=""THE_FUNCTION""/>"
7                 Set R = shHelp.Range("TheData")

8                 SA.Append "<menu xmlns=""" & _
                      "http://schemas.microsoft.com/office/2009/07/customui"">"

9                 For i = 1 To R.Rows.Count
10                    If R.Cells(i, 4).Value = Category Then
11                        SA.Append Replace(Replace(ButtonTemplate, "THE_FUNCTION", R.Cells(i, 1).Value), "THE_COUNTER", CStr(i))
12                    End If
13                Next
14                SA.Append "</menu>"
15                returnedVal = SA.Report
16            Case "SwitchBook"
17                returnedVal = XMLForSwitchBook()
18            Case "SwitchSheet"
19                returnedVal = XMLForSwitchSheet()
20            Case Else
21                Throw "Unrecognised Control.Tag: " & control.Tag
22        End Select

23        RefreshRibbon

24        Exit Sub
ErrHandler:
25        SomethingWentWrong "#TheRibbon_getContent (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : EscapeChars
' Author    : Philip Swannell
' Date      : 29-Oct-2018
' Purpose   : creates XML that is renders as the literal str passed in.
' -----------------------------------------------------------------------------------------------------------------------
Private Function EscapeChars(Str As String, DoubleAmpersand As Boolean)
          Dim Res As String
1         If DoubleAmpersand Then
2             Res = Replace(Str, "&", "&amp;&amp;")
3         Else
4             Res = Replace(Str, "&", "&amp;")
5         End If
6         Res = Replace(Res, "<", "&lt;")
7         Res = Replace(Res, ">", "&gt;")
8         Res = Replace(Res, """", "&quot;")
9         EscapeChars = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IndexToAccelerator
' Author     : Philip Swannell
' Date       : 14-Nov-2018
' Purpose    : The labels in the menu are prefixed with accelerator key characters. This method returns the character
'              in appropriate syntax, with attention paid top alignment since the accelerator characters are not all of equal width.
'              NumSpaces calculated assuming the font in Segoe UI 9 point. Tried using tab characters &#9; to no avail.
'See PGS OneDrive/Excel Sheets/Working sheet for method IndexToAccelerator.xlsx
' -----------------------------------------------------------------------------------------------------------------------
Private Function IndexToAccelerator(ByVal index As Long)
1         On Error GoTo ErrHandler
          Dim NumSpaces  As Long

2         index = (index - 1) Mod 57 + 1
3         Select Case index
              Case 23, 33, 37, 41 'Widest characters
                  'M   W   @   %
4                 NumSpaces = 1
5             Case 11, 12, 13, 14, 17, 18, 21, 24, 25, 26, 27, 28, 30, 31, 32, 34, 35, 36, 38, 42, 45, 46, 47, 56, 57
                  'A   B   C   D   G   H   K   N   O   P   Q   R   T   U   V   X   Y   Z   #   ^   +   ~   =   <   >
6                 NumSpaces = 2
7             Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 16, 20, 22, 29, 39, 40, 43, 44, 48, 49, 50, 51, 52, 53, 54, 55
                  '1  2  3  4  5  6  7  8  9   0   E   F   J   L   S   £   $   *   -   /   \   {   }   [   ]   (   )
8                 NumSpaces = 3
9             Case 19   'Narrowest character
                  'I
10                NumSpaces = 4
11        End Select
12        Select Case index
              Case 1 To 9
13                IndexToAccelerator = "&amp;" & CStr(index) & String(NumSpaces, " ")
14            Case 10
15                IndexToAccelerator = "&amp;0" & String(NumSpaces, " ")
16            Case 11 To 36
17                IndexToAccelerator = "&amp;" & Chr$(index + 54) & String(NumSpaces, " ")
18            Case 37 To 57
19                IndexToAccelerator = "&amp;" & EscapeChars(Mid$("@#£$%^*-+~=/\{}[]()<>", index - 36, 1), False) & String(NumSpaces, " ")
20        End Select

21        Exit Function
ErrHandler:
22        Throw "#IndexToAccelerator (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function isBookHidden(wb As Excel.Workbook) As Boolean
1         On Error GoTo ErrHandler
2         isBookHidden = True
          Dim wn As Window
3         For Each wn In wb.Windows
4             If wn.Visible = True Then
5                 isBookHidden = False
6                 Exit Function
7             End If
8         Next
9         Exit Function
ErrHandler:
10        Throw "#isBookHidden (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ActivateBook
' Author     : Philip Swannell
' Date       : 13-Nov-2018
' Purpose    : Activates a workbook, even if that workbook is in a different Excel instance
' -----------------------------------------------------------------------------------------------------------------------
Sub ActivateBook(BookOrBookName As Variant)
1         On Error GoTo ErrHandler
          Dim Activated As Boolean
          Dim WasAddin As Boolean
          Dim wb As Excel.Workbook
          Dim wn As Window

2         If ActiveWindow Is Nothing Then
3             If Application.ProtectedViewWindows.Count > 0 Then
4                 Throw "Method ActivateBook fails when a ProtectedViewWindow is active.", True
5             End If
6         End If

7         Select Case TypeName(BookOrBookName)
              Case "Workbook"
8                 Set wb = BookOrBookName
9             Case "String"
10                If IsInCollection(Application.Workbooks, CStr(BookOrBookName)) Then
11                    Set wb = Application.Workbooks(CStr(BookOrBookName))
12                Else
13                    Exit Sub
14                End If
15            Case Else
16                Exit Sub
17        End Select

          'This is tricky and intended to cope with case when a workbook is open in both instances, so that the application's caption might be the same...
18        If Not wb.Parent Is Application Then
              Dim OldCaption As String
19            OldCaption = wb.Parent.caption
20            wb.Parent.caption = "SwitchToMe"
21            AppActivate wb.Parent.caption
22            wb.Parent.caption = vbNullString
23            If wb.Parent.caption <> OldCaption Then
                  'TODO if the other application has a customised caption then the code above will have reset it and I should write code to restore, but that's not straightforward
24            End If
25        End If
          
26        WasAddin = wb.isAddin
27        wb.isAddin = False
28        For Each wn In wb.Windows
29            If wn.Visible = True Then
30                If wn.WindowState = xlMinimized Then
31                    wn.WindowState = xlNormal
32                End If
33                wn.Activate
34                Activated = True
35                Exit For
36            End If
37        Next
38        If Not Activated Then
39            Set wn = wb.Windows(1)
40            wn.Visible = True
41            If wn.WindowState = xlMinimized Then
42                wn.WindowState = xlNormal
43            End If
44            wn.Activate
45        End If
46        If wb Is ThisWorkbook Then
47            If WasAddin Then
48                shAudit.Activate
49            End If
50        End If

51        Exit Sub
ErrHandler:
52        Throw "#ActivateBook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub ActivateSheet(SheetName As String)
          Dim Res As VbMsgBoxResult
          Dim ws As Object

1         On Error GoTo ErrHandler
2         If ActiveWorkbook Is Nothing Then Exit Sub
3         If Not IsInCollection(ActiveWorkbook.Sheets, SheetName) Then Exit Sub
4         Set ws = ActiveWorkbook.Sheets(SheetName)
5         If ws.Visible <> xlSheetVisible Then
6             If ws.Parent.ProtectStructure Then
7                 If BookIsProtectedWithPassword(ws.Parent) Then
8                     Throw "Worksheet '" + SheetName + "' is hidden in a password-protected workbook. Please remove protection using 'Review' > 'Protect Workbook'. You will be prompted for the password.", True
9                 Else
10                    Res = MsgBoxPlus("Worksheet '" + SheetName + "' is hidden in a protected workbook. Would you like to unprotect it now?", vbOKCancel + vbQuestion, gAddinName, "Yes, unprotect", "No, do nothing")
11                    If Res <> vbOK Then
12                        Exit Sub
13                    Else
14                        ws.Parent.Unprotect
15                    End If
16                End If
17            End If
18        End If

19        ws.Visible = xlSheetVisible
20        ws.Activate
21        Exit Sub
ErrHandler:
22        Throw "#ActivateSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function BookIsProtectedWithPassword(wb As Excel.Workbook) As Boolean
1         On Error GoTo ErrHandler
2         If wb.ProtectStructure = False Then
3             BookIsProtectedWithPassword = False
4         Else
5             On Error Resume Next
6             wb.Unprotect "UnlikelyPassword" + CStr(Rnd())
7             On Error GoTo ErrHandler
8             If wb.ProtectStructure Then
9                 BookIsProtectedWithPassword = True
10            Else
11                wb.Protect , True
12                BookIsProtectedWithPassword = False
13            End If
14        End If
15        Exit Function
ErrHandler:
16        Throw "#BookIsProtectedWithPassword (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : XMLForSwitchBook
' Author     : Philip Swannell
' Date       : 13-Nov-2018
' Purpose    : Returns the Ribbon XML for the Switch Book menu
' -----------------------------------------------------------------------------------------------------------------------
Private Function XMLForSwitchBook() As String
          Dim Addins As Variant
          Dim i As Long
          Dim ItemNum As Long
          Dim k As Long
          Dim NumToShow As Long
          Dim SA As clsStringAppend
          Dim ShowAddins As Boolean
          Dim ShowSeparators As Boolean
          Dim STKH As clsStacker
          Dim STKV As clsStacker
          Dim ThisButton As String
          Dim wb As Excel.Workbook
          Const NumToShowLimit = 15
          Dim AllXLs As Collection
          Dim supertipInstances As String
          Dim TheseObjs As Variant
          Dim xl As Application
          Dim LWN As String
                  
1         On Error GoTo ErrHandler
2         Set SA = New clsStringAppend
3         Set STKH = CreateStacker()
4         Set STKV = CreateStacker()

5         ShowAddins = InDeveloperMode()
6         GetExcelInstances AllXLs

7         Select Case AllXLs.Count
              Case 0, 1
8             Case 2
9                 supertipInstances = "There are two Excel instances running. This workbook is open in the other instance."
10            Case Is > 2
11                supertipInstances = "There are " + CStr(AllXLs.Count) + " Excel instances running. This workbook is open in one of the other instances."
12        End Select

13        SA.Append "<menu xmlns=""" & _
              "http://schemas.microsoft.com/office/2009/07/customui"">" & vbLf

14        NumToShow = 0

15        i = 0
16        For Each xl In AllXLs
17            i = i + 1
18            For Each wb In xl.Workbooks
19                LWN = wb.FullName
20                On Error Resume Next
                  'Not sure how robust LocalWorkbookName is, so ignore errors. PGS 25/4/2022
21                LWN = LocalWorkbookName(wb)
22                On Error GoTo ErrHandler

23                NumToShow = NumToShow + 1
24                If isBookHidden(wb) Then
25                    ShowSeparators = True
26                    STKH.Stack2D sArrayRange(wb.Name, i, LWN)
27                Else
28                    STKV.Stack2D sArrayRange(wb.Name, i, LWN)
29                End If
30            Next wb
31        Next xl
32        If ShowAddins Then
33            Addins = WorkbookAndAddInList(2)
34            If Not (IsEmpty(Addins)) Then
35                ShowSeparators = True
36                NumToShow = NumToShow + sNRows(Addins)
37            End If
38        End If

39        ItemNum = 0
40        If NumToShow > NumToShowLimit Then
              Dim supertip As String
41            ItemNum = ItemNum + 1
42            supertip = """ A searchable list of the " + CStr(NumToShow) + " open workbooks. (Ctrl Shift B)"""
43            SA.Append "<menuSeparator id=""MenuseparatorSearchBookNames"" title=""Searchable""/>" & vbLf
44            ThisButton = "<button label=""" & IndexToAccelerator(ItemNum) & "Search..." & """ onAction=""TheRibbon_onAction"" id=""ActivateBook_" & CStr(ItemNum) & _
                  """ tag=""SwitchBook"" screentip=""Search"" supertip = " & supertip & " imageMso=""FindText""/>"
45            SA.Append ThisButton & vbLf
46            ShowSeparators = True
47        End If

48        For k = 1 To IIf(ShowAddins, 3, 2)
49            If k = 1 Then
50                TheseObjs = sSortedArray(STKV.Report)
51                If ShowSeparators Then
52                    If Not sIsErrorString(TheseObjs) Then
53                        SA.Append "<menuSeparator id=""MenuseparatorVisibleBooks"" title=""Visible""/>" & vbLf
54                    End If
55                End If
56            ElseIf k = 2 Then
57                TheseObjs = sSortedArray(STKH.Report)
58                If Not sIsErrorString(TheseObjs) Then
59                    SA.Append "<menuSeparator id=""MenuseparatorHiddenBooks"" title=""Hidden""/>" & vbLf
60                End If
61            ElseIf k = 3 Then
62                TheseObjs = WorkbookAndAddInList(2)
63                TheseObjs = sArrayRange(TheseObjs, sReshape(1, sNRows(TheseObjs), 1))
64                If Not IsEmpty(TheseObjs) Then
65                    SA.Append "<menuSeparator id=""MenuseparatorAddins"" title=""Addins""/>" & vbLf
66                End If
67            End If
68            If Not sIsErrorString(TheseObjs) And Not IsEmpty(TheseObjs) Then
69                For i = 1 To sNRows(TheseObjs)
70                    ItemNum = ItemNum + 1
                      Dim AppNum As Long
                      Dim BookFullName As String
                      Dim BookName As String
71                    BookName = TheseObjs(i, 1)
72                    AppNum = TheseObjs(i, 2)
73                    If AppNum = 1 Then
74                        BookFullName = Application.Workbooks(BookName).FullName
75                        On Error Resume Next
76                        BookFullName = LocalWorkbookName(Application.Workbooks(BookName))
77                        On Error GoTo ErrHandler
78                    Else
79                        BookFullName = TheseObjs(i, 3)
80                    End If
81                    ThisButton = "<button label=""" & IndexToAccelerator(ItemNum) & EscapeChars(BookName, True) & _
                          IIf(AppNum = 1, vbNullString, "  " & String(AppNum - 1, "*")) & """ onAction=""TheRibbon_onAction"" id=""ActivateBook_" & CStr(ItemNum) & _
                          """ tag=""ActivateBook_" & CStr(AppNum) & "_" & EscapeChars(BookName, False) & """"
82                    If AppNum = 1 Then
83                        If Not ActiveWorkbook Is Nothing Then
84                            If BookName = ActiveWorkbook.Name Then
85                                ThisButton = ThisButton & " imageMso=""AcceptProposal"""
86                            End If
87                        End If
88                    End If
89                    supertip = vbNullString
90                    If BookName <> BookFullName Then
91                        If AppNum <> 1 Then
92                            supertip = BookFullName + "&#13;&#10;" + String(50, "_") + "&#13;&#10;" + supertipInstances
93                        Else
94                            supertip = BookFullName
95                        End If
96                    ElseIf AppNum <> 1 Then
97                        supertip = supertipInstances
98                    End If
99                    If supertip <> vbNullString Then
100                       ThisButton = ThisButton & " screentip=""" & BookName & """ supertip=""" & supertip & """"
101                   End If
102                   ThisButton = ThisButton & "/>"
103                   SA.Append ThisButton & vbLf
104               Next i
105           End If
106       Next k
107       SA.Append "</menu>"
108       XMLForSwitchBook = SA.Report
109       Exit Function
ErrHandler:
110       Throw "#XMLForSwitchBook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : XMLForSwitchSheet
' Author     : Philip Swannell
' Date       : 13-Nov-2018
' Purpose    : Returns the Ribbon XML for the Switch Sheet menu
' -----------------------------------------------------------------------------------------------------------------------
Private Function XMLForSwitchSheet() As String
          Dim AdvancedMode As Boolean
          Dim i As Long
          Dim k As Long
          Dim SA As clsStringAppend
          Dim ShowSeparators As Boolean
          Dim STKC As clsStacker
          Dim STKD As clsStacker
          Dim STKH As clsStacker
          Dim STKM As clsStacker
          Dim STKV As clsStacker
          Dim STKVH As clsStacker
          Dim supertip As String
          Dim ws As Object
          Const NumToShowLimit = 15
          Dim ObjName As Variant
          Dim TheseObjs As Variant
          Dim ThisButton As String
                  
1         On Error GoTo ErrHandler

2         Set SA = New clsStringAppend
3         Set STKV = CreateStacker()
4         Set STKH = CreateStacker()
5         Set STKVH = CreateStacker()
6         Set STKC = CreateStacker()
7         Set STKM = CreateStacker()
8         Set STKD = CreateStacker()

9         If ActiveWorkbook Is Nothing Then
10            XMLForSwitchSheet = vbNullString
11            Exit Function
12        End If

13        AdvancedMode = InDeveloperMode()

14        For Each ws In ActiveWorkbook.Worksheets
15            If ws.Visible = xlSheetVisible Then
16                STKV.Stack0D ws.Name
17            ElseIf ws.Visible = xlSheetHidden Then
18                ShowSeparators = True
19                STKH.Stack0D ws.Name
20            ElseIf ws.Visible = xlSheetVeryHidden Then
21                If AdvancedMode Then
22                    ShowSeparators = True
23                    STKVH.Stack0D ws.Name
24                End If
25            End If
26        Next ws
27        For Each ws In ActiveWorkbook.Charts
28            ShowSeparators = True
29            STKC.Stack0D ws.Name
30        Next

31        If AdvancedMode Then
32            For Each ws In ActiveWorkbook.Excel4MacroSheets
33                ShowSeparators = True
34                STKM.Stack0D ws.Name
35            Next
36            For Each ws In ActiveWorkbook.Excel4IntlMacroSheets
37                ShowSeparators = True
38                STKM.Stack0D ws.Name
39            Next
40            For Each ws In ActiveWorkbook.DialogSheets
41                ShowSeparators = True
42                STKD.Stack0D ws.Name
43            Next
44        End If
45        SA.Append "<menu xmlns=""" & _
              "http://schemas.microsoft.com/office/2009/07/customui"">" & vbLf
46        i = 0
47        If ActiveWorkbook.Sheets.Count > NumToShowLimit Then
48            supertip = """A searchable list of the " + CStr(ActiveWorkbook.Sheets.Count) + " sheets in the active workbook. (Ctrl Shift T)"""
49            SA.Append "<menuSeparator id=""MenuseparatorSearchSheetNames"" title=""Searchable""/>" & vbLf
50            ThisButton = "<button label=""" & IndexToAccelerator(i) & "Search..." & """ onAction=""TheRibbon_onAction"" id=""ActivateSheet_" & CStr(i) _
                  & """ tag=""SwitchSheet"" screentip=""Search"" supertip = " & supertip & " imageMso=""FindText""/>"
51            SA.Append ThisButton & vbLf
52            i = i + 1
53            ShowSeparators = True
54        End If

55        For k = 1 To 6
56            If k = 1 Then
57                TheseObjs = sSortedArray(STKV.Report)
58                If Not sIsErrorString(TheseObjs) Then
59                    If ShowSeparators Then
60                        SA.Append "<menuSeparator id=""MenuseparatorVisibleSheets"" title=""Visible""/>" & vbLf
61                    End If
62                End If
63            ElseIf k = 2 Then
64                TheseObjs = sSortedArray(STKH.Report)
65                If Not sIsErrorString(TheseObjs) Then
66                    SA.Append "<menuSeparator id=""MenuseparatorHiddenSheets"" title=""Hidden""/>" & vbLf
67                End If
68            ElseIf k = 3 Then
69                TheseObjs = sSortedArray(STKVH.Report)
70                If Not sIsErrorString(TheseObjs) Then
71                    SA.Append "<menuSeparator id=""MenuseparatorVeryHiddenSheets"" title=""Very Hidden""/>" & vbLf
72                End If
73            ElseIf k = 4 Then
74                TheseObjs = sSortedArray(STKC.Report)
75                If Not sIsErrorString(TheseObjs) Then
76                    SA.Append "<menuSeparator id=""MenuseparatorCharts"" title=""Charts""/>" & vbLf
77                End If
78            ElseIf k = 5 Then
79                TheseObjs = sSortedArray(STKM.Report)
80                If Not sIsErrorString(TheseObjs) Then
81                    SA.Append "<menuSeparator id=""MenuseparatorMacroSheets"" title=""Macro sheets""/>" & vbLf
82                End If
83            Else
84                TheseObjs = sSortedArray(STKD.Report)
85                If Not sIsErrorString(TheseObjs) Then
86                    SA.Append "<menuSeparator id=""MenuseparatorDialogSheets"" title=""Dialog sheets""/>" & vbLf
87                End If
88            End If
89            If Not sIsErrorString(TheseObjs) Then
90                For Each ObjName In TheseObjs
91                    Set ws = ActiveWorkbook.Sheets(ObjName)
92                    i = i + 1
93                    ThisButton = "<button label=""" & IndexToAccelerator(i) & EscapeChars(ws.Name, True) & """ onAction=""TheRibbon_onAction"" id=""ActivateSheet_" & _
                          CStr(i) & """ tag=""ActivateSheet_" & EscapeChars(ws.Name, False) & """"
94                    If ws Is ActiveSheet Then
95                        ThisButton = ThisButton & " imageMso=""AcceptProposal"""
96                    End If
97                    ThisButton = ThisButton & "/>"
98                    SA.Append ThisButton & vbLf
99                Next ObjName
100           End If
101       Next k
102       SA.Append "</menu>"
103       XMLForSwitchSheet = SA.Report
104       Exit Function
ErrHandler:
105       Throw "#XMLForSwitchSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

