Attribute VB_Name = "modFormsA"
Option Explicit
Private Const SM_CXVIRTUALSCREEN = 78        'The width of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors. The SM_XVIRTUALSCREEN metric is the coordinates for the left side of the virtual screen.
Private Const SM_CYVIRTUALSCREEN = 79        'The height of the virtual screen, in pixels. The virtual screen is the bounding rectangle of all display monitors. The SM_YVIRTUALSCREEN metric is the coordinates for the top of the virtual screen.
Private Const SM_XVIRTUALSCREEN = 76        'The coordinates for the left side of the virtual screen. The virtual screen is the bounding rectangle of all display monitors
Private Const SM_YVIRTUALSCREEN = 77        'The coordinates for the top of the virtual screen. The virtual screen is the bounding rectangle of all display monitors.

Declare PtrSafe Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowSingleChoiceDialog
' Author    : Philip Swannell
' Date      : 17-Oct-2013
' Purpose   : Present a dialog allowing the user to select a value from a list. The list
'             is shown in a list box whose contents changes dynamically as the user
'             types into a text box.
'TheChoices:  Column array of text
'TheCategories: Optional column array of text, same number of rows as TheChoices. Each
'             element of TheCategories gives the category name for the corresponding
'             element of TheChoices.
'TheExplanations: Column array of text, same number of rows as TheChoices. As the user
'             selects items in the list box the corresponding Explanation text appears
'             below the list box.
'TheExplanationTitles: Optional. Column array of text. As the user selects items in the
'             list box the corresponding ExplanationTitle text appears in bold below
'             the list box and above the explanation. If not passed (and TheExplanations
'             is passed) then TheChoices themselves appear in bold i.e. TheExplanationTitles
'             defaults to TheChoices.
'InitialFilter: The "Filter" text box at the top of the dialog is initially populated with
'             this string.
'Title:       Text to appear in the blue Caption bar at the top of the form.
'TopText:     Text to appear above the File box in the form.
'CategoryLabel Text to appear to the left of the drop-down for switching categories
'AnchorObject: Optional. Dialog appears with its top left at the top left of this cell.
'             If not passed then the dialog appears in the centre of the Excel application.
'RegistryString: Specifies a location in the Registry for persisting the user's choices for the
'             SearchOptions for the dialog
'SettingsMenu: A three column array to specify additional items that can appear in the "Settings" menu
'              that appears when the user clicks the button with a "gear" icon. First column should be FaceID (zero for no FaceID)
'              second column should be text to appear on the menu and third should be fully specified macro to
'              be run via Application.Run
'Return is either the selected string or, if the user cancels out of the dialog, Empty is returned.
' -----------------------------------------------------------------------------------------------------------------------
Function ShowSingleChoiceDialog(ByVal TheChoices As Variant, _
        Optional ByVal TheExplanations As Variant, _
        Optional TheCategories As Variant, _
        Optional ByVal TheExplanationTitles As Variant, _
        Optional InitialFilter As String, _
        Optional Title As String = "Select an item", _
        Optional TopText As String = "Search for an item:", _
        Optional AnchorObject As Variant, _
        Optional CategoryLabel As String = "or select a category:", _
        Optional RegistryString As String, _
        Optional SettingsMenu As Variant, _
        Optional InitialCategory As String)

          Dim frm As frmSingleChoice
          Dim i As Long
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler

2         Set SUH = CreateScreenUpdateHandler(True)

3         Force2DArrayR TheChoices
4         If sNCols(TheChoices) > 1 Then Throw "TheChoices must have only one column."
5         For i = 1 To sNRows(TheChoices)
6             If VarType(TheChoices(i, 1)) <> vbString Then
7                 Throw "TheChoices must be a one-column array of strings"
8             End If
9         Next

10        If Not IsMissing(TheExplanations) Then
11            If sNCols(TheExplanations) > 1 Or sNRows(TheExplanations) <> sNRows(TheChoices) Then
12                Throw ("TheExplanations must be a column array of text with the same number of rows as TheChoices")
13            End If
14            If IsMissing(TheExplanationTitles) Then
15                TheExplanationTitles = TheChoices
16            End If
17            If sNCols(TheExplanationTitles) > 1 Or sNRows(TheExplanationTitles) <> sNRows(TheChoices) Then
18                Throw ("TheExplanationTitles must be a column array of text with the same number of rows as TheChoices")
19            End If
20        End If
21        If Not IsMissing(TheCategories) Then
22            If sNCols(TheCategories) > 1 Or sNRows(TheCategories) <> sNRows(TheChoices) Then
23                Throw ("TheCategories must be a column array of text with the same number of rows as TheChoices")
24            End If
25        End If

26        Set frm = New frmSingleChoice

27        frm.InitialiseAsChooserDialog TheChoices, TheExplanations, TheExplanationTitles, InitialFilter, Title, TopText, TheCategories, CategoryLabel, RegistryString, SettingsMenu, InitialCategory
28        SetFormPosition frm, AnchorObject
29        frm.Show

30        ShowSingleChoiceDialog = frm.ReturnValue

31        Set frm = Nothing

          'Don't understand why this line is necessary, but without it Excel does not properly _
           get the focus back if the user hits the Enter Key in the dialog rather than using the mouse.
32        On Error Resume Next
33        AppActivate Application.caption
34        On Error GoTo ErrHandler

35        Exit Function
ErrHandler:
36        Throw "#ShowSingleChoiceDialog (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SetFormPosition
' Author    : Philip Swannell
' Date      : 07-Nov-2013
' Purpose   : Position a form. This is designed to work on multi-monitor PCs but I have
'             not been able to test that.
'             If NearThisObject is a Range then the form is positioned with top left of
'             the form at top left of the range, otherwise if NearThisObject is an object that
'             has .top, .left, .width and .height properties (in the same coordinate system as forms use)
'             then the form is positioned with its top left at the center of NearThisObject
'             More control can be provided by passing NearThisObject as a clsPositionInstructions
' -----------------------------------------------------------------------------------------------------------------------
Sub SetFormPosition(frm As Object, Optional ByVal NearThisObject As Variant)
          Dim Left As Double
          Dim PI As clsPositionInstructions
          Dim R As Range
          Dim RC As RECT
          Dim Top As Double
          Dim VirtualScreenBottom As Double
          Dim VirtualScreenLeft As Double
          Dim VirtualScreenRight As Double
          Dim VirtualScreenTop As Double
          Dim X_Nudge As Double
          Dim Y_Nudge As Double

1         On Error GoTo ErrHandler

2         VirtualScreenLeft = GetSystemMetrics(SM_XVIRTUALSCREEN) / fx
3         VirtualScreenRight = VirtualScreenLeft + GetSystemMetrics(SM_CXVIRTUALSCREEN) / fx
4         VirtualScreenTop = GetSystemMetrics(SM_YVIRTUALSCREEN) / fY
5         VirtualScreenBottom = VirtualScreenTop + GetSystemMetrics(SM_CYVIRTUALSCREEN) / fY

6         If TypeName(NearThisObject) = "clsPositionInstructions" Then
7             Set PI = NearThisObject
8             X_Nudge = PI.X_Nudge
9             Y_Nudge = PI.Y_Nudge
10            Set NearThisObject = PI.AnchorObject
11            Set PI = Nothing
12        End If

13        With frm
14            If TypeName(NearThisObject) = "Range" Then
15                Set R = NearThisObject
                  'Position at the Cell
16                RC = GetRangeRect(R)
17                Left = RC.Left / fx + 2        'small nudge of 2 gets alignment better
18                Top = RC.Top / fY
19            ElseIf HasTopAndLeft(NearThisObject) Then
                  'Position centrally over the form
20                Top = NearThisObject.Top + NearThisObject.Height / 2 - .Height / 2
21                Left = NearThisObject.Left + NearThisObject.Width / 2 - .Width / 2
22            ElseIf Not ActiveWindow Is Nothing Then
                  'Position over active window
23                Top = ActiveWindow.Top + ActiveWindow.Height / 2 - .Height / 2
24                Left = ActiveWindow.Left + ActiveWindow.Width / 2 - .Width / 2
25            Else
                  'Position in the middle of the Excel application
26                Top = Application.Top + Application.Height / 2 - .Height / 2
27                Left = Application.Left + Application.Width / 2 - .Width / 2
28            End If

              'Avoid going off to the right of the screen
29            If Left > VirtualScreenRight - .Width Then
30                Left = VirtualScreenRight - .Width
31            End If
              'Avoid going off to the left of the screen
32            If Left < VirtualScreenLeft Then
33                Left = VirtualScreenLeft
34            End If
              'Avoid going off the bottom of the screen
35            If Top > VirtualScreenBottom - .Height Then
36                Top = VirtualScreenBottom - .Height
37            End If
              'Avoid going off the top of the screen
38            If Top < VirtualScreenTop Then
39                Top = VirtualScreenTop
40            End If

41            .StartUpPosition = 0
42            .Left = Left + X_Nudge
43            .Top = Top + Y_Nudge

44        End With

45        Exit Sub
ErrHandler:
46        Throw "#SetFormPosition (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HideForm
' Author    : Philip Swannell
' Date      : 04-May-2016
' Purpose   : This method is an attempt to work-around a problem seen on Windows 10, but
'             not previous versions of Windows. When several Excel windows are open, one can
'             hover the mouse over the Excel icon in the task bar and see a pop-up
'             showing small images (approx 1/8 size) of all open windows placed
'             side-by-side. Hover the mouse over one of those images and a full-size
'             preview of the window is displayed. BUT in Windows 10 that full-size preview
'             contains "Ghosts" of many VBA forms that were previously displayed on top of
'             that window. The "Ghosts" vanish when you activate the previewed window by
'             clicking it. This despite the fact that the VBA code (e.g. in method MsgBoxPlus)
'             has set the form to Nothing. My efforts to Google this problem yielded nothing.
' -----------------------------------------------------------------------------------------------------------------------
Sub HideForm(frm As Object)
1         On Error Resume Next
2         frm.Top = Application.Top - 2000 - frm.Height
3         On Error GoTo ErrHandler
4         frm.Hide
5         Exit Sub
ErrHandler:
6         Throw "#HideForm (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function HasTopAndLeft(o As Variant) As Boolean
          Dim L As Double
          Dim t As Double
1         On Error GoTo ErrHandler
2         t = o.Top
3         L = o.Left
4         HasTopAndLeft = True
5         Exit Function
ErrHandler:
6         HasTopAndLeft = False
End Function

Private Sub TestShowMultipleChoiceDialog()
          Dim ButtonClicked As String
          Dim Res As Variant
1         On Error GoTo ErrHandler
2         Res = ShowMultipleChoiceDialog(sArrayConcatenate("Yabba", sIntegers(80)), , , , , , "< Back", "Yabba", False, "Next >", ButtonClicked, "Hello world this is a very long check box caption that will probably run off the end of the form")
3         Exit Sub
ErrHandler:
4         SomethingWentWrong "#TestShowMultipleChoiceDialog (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowMultipleChoiceDialog
' Author    : Philip Swannell
' Date      : 22-Oct-2013
' Purpose   : Dialog to allow user to multi-select from a list of strings.
' Return is:  a) String "#User Cancel! if the user cancels out of the dialog or hits the right button
'             b) Empty if user hits OK (Left button) or middle button having selected no strings
'             c) Two-dimensional 1-column array of the selected strings. Even if only one
'                string is selected the return is still two dimensional.
'             The ByRef argument ButtonClicked is populated with the caption of the button that the
'             user clicks or the caption of the right button if the user dismisses the dialog by hitting
'             the escape key or clicking the red x in the top right of the dialog.
'             AllowNoneChosen can be a Boolean or a two element 1-d array of Booleans, first element to
'             control the left button, second element to control the middle button
' -----------------------------------------------------------------------------------------------------------------------
Function ShowMultipleChoiceDialog(ByVal TheChoices As Variant, _
        Optional InitialChoices As Variant, _
        Optional Title As String = "Select", _
        Optional TopText As String, _
        Optional ShowCheckBoxes As Boolean = False, _
        Optional AnchorObject As Variant, _
        Optional LeftCaption As String = "OK", _
        Optional RightCaption As String = "Cancel", _
        Optional AllowNoneChosen As Variant = True, _
        Optional MiddleCaption As String = vbNullString, _
        Optional ByRef ButtonClicked As String, _
        Optional CheckBoxCaption As String, _
        Optional ByRef CheckBoxValue As Boolean)

          Dim frm As frmMultipleChoice
          Dim SUH As clsScreenUpdateHandler

1         On Error GoTo ErrHandler
2         Set SUH = CreateScreenUpdateHandler(True)
          Dim AllowNoneLeftButton As Boolean
          Dim AllowNoneMiddleButton As Boolean

3         If sNCols(TheChoices) > 1 Then Throw "TheChoices must have only one column."

4         If IsEmpty(AllowNoneChosen) Or IsMissing(AllowNoneChosen) Then
5             AllowNoneLeftButton = True
6             AllowNoneMiddleButton = True
7         ElseIf VarType(AllowNoneChosen) = vbBoolean Then
8             AllowNoneLeftButton = AllowNoneChosen
9             AllowNoneMiddleButton = AllowNoneChosen
10        Else
11            AllowNoneLeftButton = AllowNoneChosen(LBound(AllowNoneChosen))
12            AllowNoneMiddleButton = AllowNoneChosen(UBound(AllowNoneChosen))
13        End If

14        Set frm = New frmMultipleChoice

15        frm.Initialise TheChoices, InitialChoices, Title, TopText, ShowCheckBoxes, LeftCaption, RightCaption, MiddleCaption, AllowNoneLeftButton, AllowNoneMiddleButton, True, CheckBoxCaption, CheckBoxValue
16        SetFormPosition frm, AnchorObject
17        frm.Show
18        ShowMultipleChoiceDialog = frm.ReturnValue
19        ButtonClicked = frm.m_ButtonClicked
20        CheckBoxValue = frm.CheckBox1.Value
21        Set frm = Nothing

22        Exit Function
ErrHandler:
23        Throw "#ShowMultipleChoiceDialog (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowHelpBrowser
' Author    : Philip Swannell
' Date      : 12-Nov-2013
' Purpose   : Allow the user to browse all functions in TigerLib (as originally written, now SolumAddin) and to enter a function
'             onto the sheet via a hook into the Excel fuction wizard.
'             Implementation of Undo is tricky in this method. When the user chooses to enter an array
'             formula we make two changes to the sheet - first to enter the formula, and second to resize it.
'             But at the time we make the first change, it's not possible to know how
'             large an area will need to be backed up in order to undo the second change. The solution is
'             via RestoreRangeTwice - we make two back ups and in the undo we restore twice.
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowHelpBrowser()

          Dim ChosenFunction As String
          Dim frm As frmSingleChoice
          Const MSGBOXTITLE = "Browse " & gAddinName & " Functions"
          Dim HelpC1
          Dim HelpC2
          Dim HelpC3
          Dim HelpC4
          Dim HelpData As Variant
          Dim TheCategories
          Dim TheChoices
          Dim TheExplanations
          Dim TheExplanationTitles
          Dim UseCtrlShiftEnter As Boolean

1         On Error GoTo ErrHandler

2         HelpData = GetHelpData()
3         HelpC1 = sSubArray(HelpData, 1, 1, , 1)
4         HelpC2 = sSubArray(HelpData, 1, 2, , 1)
5         HelpC3 = sSubArray(HelpData, 1, 3, , 1)
6         HelpC4 = sSubArray(HelpData, 1, 4, , 1)
7         TheChoices = HelpC1
8         TheExplanationTitles = sArrayConcatenate("Help for ", HelpC1)
9         TheExplanations = sArrayConcatenate("Category: ", HelpC4, vbLf, sArrayIf(sArrayEquals(HelpC4, "Keyboard Shortcuts"), vbNullString, "Syntax:" + vbLf), HelpC2, vbLf + vbLf, HelpC3)
10        TheCategories = HelpC4

11        Set frm = New frmSingleChoice

12        frm.InitialiseAsHelpBrowser TheChoices, TheExplanations, TheExplanationTitles, MSGBOXTITLE, _
              "Search for:", TheCategories, "Category:", "FunctionBrowser"
13        SetFormPosition frm
14        frm.Show

          'Note that code in method frm.ClickedOK will check that cells are selected, there is only _
           one Area in the selection, and that the sheet is unprotected.

15        ChosenFunction = frm.ReturnValue
16        UseCtrlShiftEnter = frm.UseCtrlShiftEnter
17        Set frm = Nothing
18        If ChosenFunction <> vbNullString Then
19            InsertFunctionAtActiveCell ChosenFunction, UseCtrlShiftEnter
20        End If
21        Application.OnRepeat "Repeat Browse Functions", "ShowHelpBrowser"

22        Exit Sub
ErrHandler:
23        SomethingWentWrong "#ShowHelpBrowser (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical, MSGBOXTITLE
24        Application.OnRepeat "Repeat Browse Functions", "ShowHelpBrowser"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShowHelpForFunction
' Author     : Philip Swannell
' Date       : 09-Nov-2018
' Purpose    : Pops up a dialog with help for a particular function
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowHelpForFunction(FunctionName As String, Optional WithInsertFormula As Boolean)
          Dim HelpText As String
          Dim MatchRes As Variant
          Dim Res
1         On Error GoTo ErrHandler
2         MatchRes = sMatch(FunctionName, shHelp.Range("TheData").Columns(1).Value)
3         If Not IsNumber(MatchRes) Then Throw "No help available for '" + FunctionName + "'"
4         With shHelp.Range("TheData")
5             HelpText = "Help for " + FunctionName + vbLf + vbLf + _
                  "Category:" + vbLf + .Cells(MatchRes, 4) + vbLf + vbLf + _
                  "syntax:" + vbLf + shHelp.Range("TheData").Cells(MatchRes, 2) + vbLf + vbLf + _
                  HelpFromFunctionAndArgDescriptions(.Cells(MatchRes, 2), .Cells(MatchRes, 7), .Cells(MatchRes, 8).Resize(1, .Cells(MatchRes, 5)), .Cells(MatchRes, 6))
6         End With
              
7         If WithInsertFormula Then
8             Res = MsgBoxPlus(HelpText, vbYesNoCancel + vbDefaultButton2 + vbInformation, gAddinName & " Function " + FunctionName, "Browse functions...", "Insert formula...", "Cancel")
9             If Res = vbCancel Then
10                Exit Sub
11            ElseIf Res = vbYes Then
12                ShowHelpBrowser
13            Else
14                InsertFunctionAtActiveCell FunctionName, IsShiftKeyDown
15            End If
16        Else
17            Res = MsgBoxPlus(HelpText, vbOKCancel + vbDefaultButton2 + vbInformation, gAddinName & " Function " + FunctionName, "Browse functions...", "OK")
18            If Res = vbOK Then
19                ShowHelpBrowser
20            End If
21        End If

22        Exit Sub
ErrHandler:
23        SomethingWentWrong "#ShowHelpForFunction (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowOptionButtonDialog
' Author    : Philip Swannell
' Date      : 15-Nov-2015
' Purpose   : Displays a dialog with a number (up to 60) of option buttons. The user can select one of them.
'      TheChoices  - a column array of strings. Can optionally use & character to indicate that "next character is accelerator key". To acually display a &, use &&.
'      Title appears in blue bar at top
'      TopText appears above all the option buttons
'      CurrentChoice - pass in the text of a button to be pre-selected.
'      AnchorObject - pass in a range object and the dialog appears near that range, otherwise dialog appears at the centre of the Excel application
'      If True then the return (if user hits OK or button2) is a number 1 = first option button etc. Otherwise the return is the element of TheChoices selected by the user.
'      If user hits the Cancel button then function return is zero (if ReturnIndex = True) or Empty (if ReturnIndex = False)
'      To distinguish between user hitting OK button or button 2, use ByRef argument ButtonClicked, which takes the value of the caption of the button the user clicks
' NEW FUNCTIONALITY JULY 2018
'----------------------------
' Arguments TheChoices and TopText can now have multiple columns in which case the displayed form has multiple frames, each frame containing option buttons
' Elements of TheChoices which are errors are ignored and hence the number of option buttons in each frame can differ, though if a column of TheChoices contains
' error values then all those error values should be at the bottom (as per sArrayRange)- since otherwise (if ReturnIndex = True) the indices returned are indices
' into the "columns with errors removed",
' The return from the function in this case is a one-row array of either indices (if ReturnIndex = True) or else of chosen elements from TheChoices
' -----------------------------------------------------------------------------------------------------------------------
Function ShowOptionButtonDialog(ByVal TheChoices, Optional Title As String, Optional ByVal TopText As Variant, _
        Optional ByVal CurrentChoice As Variant, Optional AnchorObject, _
        Optional ReturnIndex As Boolean = False, Optional CheckBoxText As String, _
        Optional ByRef CheckBoxValue As Boolean, Optional HelpMethodName As String, _
        Optional Caption1 As String = "&OK", Optional Caption2 As String = vbNullString, Optional Caption3 As String = "&Cancel", _
        Optional ByRef ButtonClicked As String)

1         On Error GoTo ErrHandler

          Dim theFrm As frmOptionButton
2         Set theFrm = New frmOptionButton
          Dim Result
          Dim SUH As clsScreenUpdateHandler
3         Force2DArrayRMulti TheChoices, CurrentChoice

4         Set SUH = CreateScreenUpdateHandler(True)

5         If IsNumber(CurrentChoice) Then If CurrentChoice > 0 Then If CurrentChoice <= sNRows(TheChoices) Then CurrentChoice = TheChoices(CLng(CurrentChoice), 1)
          'Pad out CurrentChoice if necessary
6         If sNCols(CurrentChoice) < sNCols(TheChoices) Then
7             CurrentChoice = sArrayRange(CurrentChoice, sReshape(vbNullString, 1, sNCols(TheChoices) - sNCols(CurrentChoice)))
8         End If

9         If IsMissing(TopText) Or IsEmpty(TopText) Then TopText = vbNullString
10        Force2DArray TopText
          'Pad out TopText if necessary
11        If sNCols(TopText) < sNCols(TheChoices) Then
12            TopText = sArrayRange(TopText, sReshape(vbNullString, 1, sNCols(TheChoices) - sNCols(TopText)))
13        End If

14        If sNCols(TheChoices) > 1 And sNRows(TheChoices) = 1 Then
15            TheChoices = sArrayTranspose(TheChoices)
16        End If

17        theFrm.Initialise TheChoices, Title, CurrentChoice, TopText, CheckBoxText, CheckBoxValue, HelpMethodName, Caption1, Caption2, Caption3
18        SetFormPosition theFrm, AnchorObject
19        theFrm.Show

          'User hits cancel
20        If IsEmpty(theFrm.ChosenIndices) Then
21            If ReturnIndex Then
22                Result = sReshape(0, 1, sNCols(TheChoices))
23            Else
24                Result = Empty
25            End If
26        Else        'User hits OK or button2
27            If ReturnIndex Then
28                Result = theFrm.ChosenIndices
29            Else
30                Result = theFrm.ChosenValues
31            End If
32        End If
33        If IsArray(Result) Then
34            If sNCols(TheChoices) = 1 Then
35                Result = Result(1, 1)    'For backward compatibility we don't return an array in this case
36            End If
37        End If

38        ButtonClicked = theFrm.ButtonClicked
39        CheckBoxValue = theFrm.CheckBox1.Value
40        ShowOptionButtonDialog = Result
41        Set theFrm = Nothing
42        Exit Function
ErrHandler:
43        Throw "#ShowOptionButtonDialog(line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ShowMsgBoxIcons
' Author     : Philip Swannell
' Date       : 08-Nov-2018
' Purpose    : Want to "harvest" the icons at different screen resolutions, by running this on various PCs
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowMsgBoxIcons()
          Static i As Long
1         i = (i + 1) Mod 4
2         MsgBox "      DPI = " & CStr(ScreenDPI(True)), Choose(i + 1, vbInformation, vbQuestion, vbExclamation, vbCritical)
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MsgBoxPlus
' Author    : Philip Swannell
' Date      : 22-Nov-2015
' Ownership : This code was copied from TigerLib.xlam, copyright Philip Swannell. Used in SolumAddin with permission.
' Purpose   : A dialog designed to be a replacement for VBA-native MsgBox. First three arguments
'             are identical to those of MsgBox
'             Remaining arguments are:
'         Caption1: Text to appear on left button, defaults from value of Buttons. To specify an accelerator
'             key for the button place a "&" character before the accelerator key character. To make "&" appear on the button use "&&".
'         Caption2: Text to appear on the second button (if there is one), defaults from value of Buttons.
'             Accelerator for Caption2 key as per Caption1.
'         Caption3: Text to appear on the third button (if there is one), defaults from value of Buttons.
'             Accelerator key for Caption3 as per Caption1.
'         ButtonWidth: Width in points of the buttons. All buttons appear at the same size and
'             their height adjusts to accommodate their text.
'         TextWidth: Width of textbox in points. The dialog auto-sizes for long messages and scroll bars
'             appear if necessary.
'         CheckBoxCaption If passed, then a checkbox appears under the prompt and above the buttons.  Accelerator
'             key as per Caption1.
'         CheckBoxValue Passed by reference and used to set the initial value of the checkbox if it's visible.
'             Is set to the value of the checkbox at the time the dialog is dismissed (even if the dialog
'             is cancelled).
'         SecondsToSelfDestruct If passed the dialog vanishes after this number of seconds, with a countdown
'             indicator on the button defined by SelfDestructButton.
'         SelfDestructButton - the button which is automatically clicked if the user does not click one (and
'             SecondsToSelfDestruct is passed >0)
'         AnchorObject - passed as a range next to which the dialog appears. Otherwise the dialog
'             appears at the centre of the Excel application, or else pass as a Form object (useful when this method is called from a the code of a form.
'
'             Advantages over MsgBox:
'          1) Can set text on the buttons so no more "For option A click Yes, for option B click No" or such-like - just
'             make the buttons read "Option A" and "Option B"
'          2) Checkbox allows for "Do not show this dialog again" type behaviour.
'          3) Self destruct can be useful to prevent dialogs halting processes unnecessarily.
'          4) Larger text - Segoe UI 9 point rather than Windows system default, probably Microsoft Sans Serif 8 point.
'          5) Text can be copied out of the dialog - right click or Ctrl+Insert or Ctrl+C
'          6) Accelerator keys. It's not necessary to hold down the Alt key with the accelerator key.
'          7) Buttons are highlighted in the MouseOver events.
'          8) Roll-our-own hyperlinks: If Prompt includes web-addresses (starts with www. or http:// or https://) then left
'             mouse-click on that address pops up a menu item - "Go to <web-address>" Nice but not very discoverable since we
'             cannot use underlined font to indicate such "hyperlink" is available.
'
'            Disadvantage versus MsgBox:
'          1) Esoteric options for Buttons are not supported: vbApplicationModal, vbSystemModal, vbMsgBoxHelpButton,
'             VbMsgBoxSetForeground, vbMsgBoxRight, vbMsgBoxRtlReading
' -----------------------------------------------------------------------------------------------------------------------
Function MsgBoxPlus(Optional ByVal Prompt As String, _
          Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
          Optional ByVal Title As String = "Microsoft Excel", _
          Optional Caption1 As String = vbNullString, _
          Optional Caption2 As String = vbNullString, _
          Optional Caption3 As String = vbNullString, _
          Optional ButtonWidth As Double = 70, _
          Optional TextWidth As Long = 285, _
          Optional ByVal CheckBoxCaption As String = vbNullString, _
          Optional ByRef CheckBoxValue As Boolean, _
          Optional ByVal SecondsToSelfDestruct As Long, _
          Optional SelfDestructButton As VbMsgBoxResult, _
          Optional AnchorObject As Variant) As VbMsgBoxResult

          Dim ButtonTexts As String
          Dim SUH As clsScreenUpdateHandler
          Dim TextForLog As String
          Dim theForm As frmMsgBoxPlus

1         On Error GoTo ErrHandler

2         Set SUH = CreateScreenUpdateHandler(True)
3         Set theForm = New frmMsgBoxPlus

4         theForm.Initialise Prompt, Buttons, Title, Caption1, Caption2, Caption3, _
              ButtonWidth, TextWidth, CheckBoxCaption, CheckBoxValue, _
              SecondsToSelfDestruct, SelfDestructButton

          'Turns out to be useful for monitoring certain processes to print to the log file once when the dialog is posted, then again when the user clicks on a button.
5         ButtonTexts = "'" + Replace(Replace(theForm.but1.caption, vbLf, vbNullString), vbCr, vbNullString) + "'"
6         If theForm.but2.Visible Then
7             ButtonTexts = ButtonTexts + String(2, " ") + "'" + Replace(Replace(theForm.but2.caption, vbLf, vbNullString), vbCr, vbNullString) + "'"
8         End If
9         If theForm.but3.Visible Then
10            ButtonTexts = ButtonTexts + String(2, " ") + "'" + Replace(Replace(theForm.but3.caption, vbLf, vbNullString), vbCr, vbNullString) + "'"
11        End If
12        TextForLog = Prompt + vbLf + "Buttons: " + ButtonTexts
13        TextForLog = "Message box:" + vbLf + TextForLog '+ vbLf + "User clicked: '" + Replace(Replace(theForm.m_ReturnCaption, vbLf, vbNullString), vbCr, vbNullString) + "'"
14        MessageLogWrite TextForLog

15        SetFormPosition theForm, AnchorObject
16        theForm.Show

17        MsgBoxPlus = theForm.m_ReturnValue
18        If CheckBoxCaption <> vbNullString Then
19            CheckBoxValue = theForm.CheckBox1.Value
20        End If
21        TextForLog = "User clicked: '" + Replace(Replace(theForm.m_ReturnCaption, vbLf, vbNullString), vbCr, vbNullString) + "'"
22        MessageLogWrite TextForLog

23        Set theForm = Nothing

24        Exit Function
ErrHandler:
25        Throw "#MsgBoxPlus (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TextBoxMouseEvent
' Author    : Philip
' Date      : 28-Mar-2016
' Purpose   : Common code for handling MouseDown and MouseUp events in a textbox of a form.
'             Implements copying the selected text via right-click and going to
'             any web addresses mentioned in the text box via left-click.
' -----------------------------------------------------------------------------------------------------------------------
Sub TextBoxMouseEvent(tb As Object, Button As Integer, Shift As Integer, x As Single, y As Single, IsDownClick As Boolean)
          Dim Choices
          Dim CurrentWord As String
          Dim FaceIDs
          Dim Res
          Const chCopy = "&Copy         (Ctrl C)"        'Ctrl C will only work if we have also implemented a KeyDown event for the textbox
          Dim address As String
          Dim chGoTo As String
          Dim NewSelStart As Long
1         On Error GoTo ErrHandler

2         Choices = CreateMissing()
3         FaceIDs = CreateMissing()

4         If Button = 2 And IsDownClick = True Then        'Right down-click
5             If tb.SelLength > 0 Then
6                 Choices = sArrayStack(Choices, chCopy)
7                 FaceIDs = sArrayStack(FaceIDs, 22)
8             End If
9         End If
10        If Button = 1 And IsDownClick = False Then        'Left up-click
11            GetCurrentWordOfTextBox CurrentWord, tb, NewSelStart
12            If LCase$(Left$(CurrentWord, 4)) = "www." Or _
                  LCase$(Left$(CurrentWord, 7)) = "http://" Or _
                  LCase$(Left$(CurrentWord, 8)) = "https://" Then
13                chGoTo = "&Go to " + CurrentWord
14                Choices = sArrayStack(Choices, chGoTo)
15                FaceIDs = sArrayStack(FaceIDs, 9026)
16                If LCase$(Left$(CurrentWord, 4)) = "www." Then
17                    address = "http://" + CurrentWord
18                Else
19                    address = CurrentWord
20                End If
21                tb.SelStart = NewSelStart
22                tb.SelLength = Len(CurrentWord)
23            End If
24        End If
25        If IsMissing(Choices) Then Exit Sub

          Dim PI As clsPositionInstructions
26        CreatePositionInstructions PI, tb, _
              (-tb.Width / 2 + x + 60) * fx, _
              (-tb.Height / 2 + y + 20) * fY

27        Res = ShowCommandBarPopup(Choices, FaceIDs, , , PI)

28        Select Case Res
              Case Unembellish(chCopy)
29                CopyStringToClipboard tb.SelText
30            Case Unembellish(chGoTo)
31                ThisWorkbook.FollowHyperlink address:=address, NewWindow:=True
32        End Select

33        Exit Sub
ErrHandler:
34        Throw "#TextBoxMouseEvent (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetCurrentWordOfTextBox
' Author    : Philip
' Date      : 28-Mar-2016
' Purpose   : Returns the currently selected "word" in a text box. Words are delimited by
'             space characters and vbCr characters (no support for other whitespace characters)
'             If the word starts and ends with double quotes then those are stripped off.
'             ByRef argument NewSelStart is set so that calling code can amend the selection in the text box
' -----------------------------------------------------------------------------------------------------------------------
Sub GetCurrentWordOfTextBox(ByRef CurrentWord As String, tb As Object, ByRef NewSelStart As Long)
          Dim EndOfWord As Long
          Dim isHyperlink As Boolean
          Dim LongText As String
          Dim Res As String
          Dim StartOfWord As Long
          Dim startPoint As Long

1         On Error GoTo ErrHandler
2         startPoint = tb.SelStart
3         LongText = tb.text
4         LongText = Replace(LongText, vbCrLf, vbCr)        'Tricky to explain why this line is necessary, but the SelStart property of a text box appears to count vbCrLf as one character...
5         LongText = " " + vbCr + LongText + " " + vbCr        'Ensures that searches below always find the searched character
6         startPoint = startPoint + 2

7         StartOfWord = InStrRev(LongText, " ", startPoint)
8         StartOfWord = SafeMax(StartOfWord, InStrRev(LongText, vbCr, startPoint))

9         EndOfWord = InStr(startPoint + 1, LongText, " ")
10        EndOfWord = SafeMin(EndOfWord, InStr(startPoint + 1, LongText, vbCr))

11        Res = Mid$(LongText, StartOfWord + 1, EndOfWord - StartOfWord - 1)
12        NewSelStart = StartOfWord - 2

          'Strip surrounding double-quotes or single quotes
13        If sIsRegMatch("^'.+'$|^"".+""$", Res) Then
14            Res = Mid$(Res, 2, Len(Res) - 2)
15            NewSelStart = NewSelStart + 1
16        End If

17        isHyperlink = LCase$(Left$(Res, 4)) = "www." Or _
              LCase$(Left$(Res, 7)) = "http://" Or _
              LCase$(Left$(Res, 8)) = "https://"

18        If isHyperlink Then
19            CurrentWord = Res
20        Else
              'if it's not a hyperlink, we want to be stricter about what characters form word-boundaries
              Dim goBack As Long
              Dim goForward As Long
21            goBack = 0: goForward = 0
22            Do While IsAlphaNumeric(Mid$(LongText, startPoint + goBack - 1, 1))
23                goBack = goBack - 1
24            Loop
25            Do While IsAlphaNumeric(Mid$(LongText, startPoint + goForward + 1, 1))
26                goForward = goForward + 1
27            Loop
28            CurrentWord = Mid$(LongText, startPoint + goBack, goForward - goBack + 1)
29            NewSelStart = startPoint + goBack - 3
30        End If

31        Exit Sub
ErrHandler:
32        Throw "#GetCurrentWordOfTextBox (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function IsAlphaNumeric(ch As String) As Boolean
1         Select Case Asc(ch)
              Case 97 To 122, 65 To 90, 48 To 57
2                 IsAlphaNumeric = True
3         End Select
End Function

Private Sub TestInputBoxPlus()
          Dim Res
1         Res = InputBoxPlus("Grab a range", "Your text here", , , , , , , , , , , True)
2         If Res <> False Then
3             Application.GoTo ActiveSheet.Range(Res)
4         End If
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : InputBoxPlus
' Author    : Philip Swannell
' Date      : 23-Jun-2016
' Purpose   : Roll-my-own version of VBA.InputBox or Application.InputBox
'             Advantages:
'             Text bigger on high-res screens.
'             Form is resizeable
'             Can distinguish between clicking Cancel and clicking OK when no text has been entered
'                (since in the former case the method returns Boolean False)
'             Can use for password entry (with PasswordMode = True)
'             Can have third button, placed between the OK and Cancel buttons, tell which button the user clicked by
'             checking the value of ByRef argument ButtonClicked that's set to the caption of the button that the user clicks.
' -----------------------------------------------------------------------------------------------------------------------
Function InputBoxPlus(Prompt As String, Optional Title As String = "Input", Optional Default As String, _
        Optional OKText As String = "&OK", Optional CancelText As String = "&Cancel", _
        Optional TextBoxWidth As Double = 200, Optional TextBoxHeight As Double = 20, _
        Optional AnchorObject As Object, Optional PasswordMode As Boolean, _
        Optional RegExMode As Boolean, Optional MiddleText As String = vbNullString, _
        Optional ByRef ButtonClicked As String, Optional RefEditMode As Boolean = False)

1         On Error GoTo ErrHandler

2         If TypeName(Application.Caller) = "Range" Then
3             InputBoxPlus = "#Function InputBoxPlus cannot be called from a spreadsheet!"
4             Exit Function
5         End If

          Dim SUH As clsScreenUpdateHandler
          Dim theForm As frmInputBoxPlus
6         Set SUH = CreateScreenUpdateHandler(True)

7         Set theForm = New frmInputBoxPlus
8         theForm.Initialise Prompt, Title, Default, OKText, CancelText, TextBoxWidth, TextBoxHeight, PasswordMode, RegExMode, MiddleText, RefEditMode
9         SetFormPosition theForm, AnchorObject
10        theForm.Show
11        ButtonClicked = theForm.m_ButtonClicked
12        InputBoxPlus = theForm.m_ReturnValue
13        Set theForm = Nothing
14        Exit Function
ErrHandler:
15        Throw "#InputBoxPlus (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HighlightFormControl
' Author    : Philip Swannell
' Date      : 17-Nov-2013
' Purpose   : Form dialogs are a bit old-school, for example buttons don't get highlighted
'             as the mouse is moved over them. This method and Unhighlight make that happen.
' -----------------------------------------------------------------------------------------------------------------------
Sub HighlightFormControl(o As Object)
          Dim TheCol As Double
1         On Error GoTo ErrHandler

2         TheCol = &H80000014
3         If o.BackColor <> TheCol Then
4             o.BackColor = TheCol
5         End If
6         Exit Sub
ErrHandler:
7         Throw "#HighlightFormControl (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
Sub UnHighlightFormControl(o As Object)
1         On Error GoTo ErrHandler

2         If o.BackColor <> &H8000000F Then
3             o.BackColor = &H8000000F
4         End If

5         Exit Sub
ErrHandler:
6         Throw "#UnHighlightFormControl (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
