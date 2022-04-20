Attribute VB_Name = "modDatePicker"
Option Explicit
Private m_Chosen_Date As Long
Private Const m_ChunkSize = 25    'sets how many years appear when the user clicks "Earlier" or "Later"
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DatePicker
' Author    : Philip
' Date      : 20-Jun-2017
' Purpose   : I've seen many VBA date-pickers built using forms, but all rather clunky...
'             This one is built using command bars. Maybe less clunky?
'             FirstYear - minimum "allowed" year, but user can browse for "Earlier" years (if WithEarlier is TRUE)
'             LastYear -  maximum "allowed" year, but user can browse for "Later" years (if WithLater is TRUE)
'             AnchorObject - object e.g. Range or button over which the command bars appear
'
'     Note that we add controls to the command-bar "just-in-time" since  otherwise
'             the command bar would take too long to construct...
' -----------------------------------------------------------------------------------------------------------------------
Sub DatePicker(ByRef ChosenDate As Variant, FirstYear As Long, LastYear As Long, Optional AnchorObject As Object, _
        Optional WithEarlier As Boolean = False, Optional WithLater As Boolean = True)
          Dim caption As String
          Dim i As Long
          Dim isFromKeyboard As Boolean
          Dim TempCommandBar As Office.CommandBar
          Dim UseOld As Boolean
          Dim x As Double
          Dim y As Double
          Dim YCtrl As CommandBarControl

1         On Error GoTo ErrHandler
2         isFromKeyboard = (sElapsedTime() - SafeMax(LastShiftF10Time, LastAltBacktickTime)) < 0.5
3         Application.EnableCancelKey = xlDisabled    'otherwise it's possible to interrupt macro execution when hitting the escape key to "pop levels" in the command bar

4         On Error Resume Next
5         Set TempCommandBar = Application.CommandBars("DatePicker")
6         On Error GoTo ErrHandler

7         If Not TempCommandBar Is Nothing Then
8             UseOld = CommandBarGood(TempCommandBar, FirstYear, LastYear, WithEarlier, WithLater)
9         End If

10        If Not UseOld Then
11            On Error Resume Next
12            Application.CommandBars("DatePicker").Delete
13            On Error GoTo ErrHandler
14            Set TempCommandBar = CommandBars.Add(Name:="DatePicker", Position:=msoBarPopup, Temporary:=True)

15            If FirstYear > LastYear Then
16                Throw "FirstYear must be less than or equal to LastYear"
17            ElseIf FirstYear = LastYear And Not (WithEarlier Or WithLater) Then
18                DatePickerAddMonthsToYear FirstYear
19            Else
20                If WithEarlier Then
21                    Set YCtrl = TempCommandBar.Controls.Add(msoControlPopup)
22                    YCtrl.caption = "&Earlier"
23                    YCtrl.Tag = CStr(FirstYear - 1) + "AndEarlier"
24                    YCtrl.OnAction = "'DatePickerAddEarlierYears " + CStr(FirstYear - 1) + " '"
25                End If
26                For i = FirstYear To LastYear
27                    Set YCtrl = TempCommandBar.Controls.Add(msoControlPopup)
28                    caption = CStr(i)
29                    caption = Left$(caption, Len(caption) - 1) & "&" & Right$(caption, 1)
30                    YCtrl.caption = caption
31                    YCtrl.Tag = CStr(i)
                      'Trick in line below is explained at:
                      'http://www.tushar-mehta.com/excel/vba/xl%20objects%20and%20procedures%20with%20arguments.htm
32                    YCtrl.OnAction = "'DatePickerAddMonthsToYear " + CStr(i) + " '"
33                Next i
34                If WithLater Then
35                    Set YCtrl = TempCommandBar.Controls.Add(msoControlPopup)
36                    YCtrl.caption = "&Later"
37                    YCtrl.Tag = CStr(LastYear + 1) + "AndLater"
38                    YCtrl.OnAction = "'DatePickerAddLaterYears " + CStr(LastYear + 1) + " '"
39                End If
40            End If
41        End If

42        m_Chosen_Date = 0
43        If AnchorObject Is Nothing Then
44            If Not ActiveCell Is Nothing Then
45                If isFromKeyboard Then
46                    Set AnchorObject = ActiveCell
47                End If
48            End If
49        End If

          'Figure out where the menu should go.
50        XYCoordinatesOfObjectCentre AnchorObject, x, y
51        If x <> 0 And y <> 0 Then
52            TempCommandBar.ShowPopup x, y
53        Else
54            TempCommandBar.ShowPopup
55        End If

          'TempCommandBar.Delete
56        If m_Chosen_Date = 0 Then
57            ChosenDate = "#Cancel!"
58        Else
59            ChosenDate = CDate(m_Chosen_Date)
60        End If
61        Exit Sub
ErrHandler:
62        Throw "#DatePicker (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CommandBarGood
' Author    : Philip
' Date      : 14-Sep-2017
' Purpose   : Check if the existing command bar - from a previous call - is built as we need it to be built...
' -----------------------------------------------------------------------------------------------------------------------
Private Function CommandBarGood(cb As CommandBar, FirstYear As Long, LastYear As Long, WithEarlier As Boolean, WithLater As Boolean) As Boolean
1         On Error GoTo ErrHandler
2         If WithEarlier Then If cb.Controls(1).caption <> "&Earlier" Then Exit Function
3         If WithLater Then If cb.Controls(cb.Controls.Count).caption <> "&Later" Then Exit Function
4         If cb.Controls(IIf(WithEarlier, 2, 1)).caption <> CStr(FirstYear) Then Exit Function
5         If cb.Controls(cb.Controls.Count - IIf(WithLater, 1, 0)).caption <> CStr(LastYear) Then Exit Function
6         CommandBarGood = True
7         Exit Function
ErrHandler:
8         Throw "#CommandBarGood (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DatePickerAddMonthsToYear
' Author    : Philip
' Date      : 20-Jun-2017
' Purpose   : First time the user navigates to a year button, populate its month sub-buttons
' -----------------------------------------------------------------------------------------------------------------------
Public Sub DatePickerAddMonthsToYear(TheYear)
          Dim j As Long
          Dim Mctrl As CommandBarControl
          Dim YCtrl As Object
1         On Error GoTo ErrHandler
          'When FirstYear = LastYear there is no "Year level" in the command bar
2         If Application.CommandBars("DatePicker").Controls.Count = 0 Then
3             Set YCtrl = Application.CommandBars("DatePicker")
4         Else
5             Set YCtrl = Application.CommandBars("DatePicker").FindControl(, , CStr(TheYear), , True)
6         End If

7         If YCtrl.Controls.Count = 0 Then
8             For j = 1 To 12
9                 Set Mctrl = YCtrl.Controls.Add(msoControlPopup)
10                Mctrl.Tag = CStr(TheYear) & "_" & CStr(j)
11                Mctrl.caption = Choose(j, "&Jan", "&Feb", "&Mar", "&Apr", "Ma&y", "J&un", "Ju&l", "Au&g", "&Sep", "&Oct", "&Nov", "&Dec") & "-" & CStr(TheYear)
12                Mctrl.OnAction = "'DatePickerAddDaysToMonth " + CStr(TheYear) + ", " + CStr(j) + "'"
13            Next j
14        End If
15        Exit Sub
ErrHandler:
16        Throw "#DatePickerAddMonthsToYear (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DatePickerAddDaysToMonth
' Author    : Philip
' Date      : 20-Jun-2017
' Purpose   : First time the user navigates to a month button, populate its day sub-buttons
' -----------------------------------------------------------------------------------------------------------------------
Public Sub DatePickerAddDaysToMonth(TheYear As Long, TheMonth As Long)
          Dim Dctrl As CommandBarControl
          Dim k As Long
          Dim Mctrl As CommandBarControl
1         On Error GoTo ErrHandler
          Dim caption As String
2         Set Mctrl = Application.CommandBars("DatePicker").FindControl(, , CStr(TheYear) & "_" & CStr(TheMonth), , True)
3         If Mctrl.Controls.Count = 0 Then
4             For k = DateSerial(TheYear, TheMonth, 1) To DateSerial(TheYear, TheMonth + 1, 1) - 1
5                 Set Dctrl = Mctrl.Controls.Add(msoControlButton)
6                 If True Then
7                     If day(k) < 10 Then
8                         caption = "   &" & Format$(k, "d-mmm-yyyy   ddd")
9                     Else
10                        caption = Format$(k, "d-mmm-yyyy   ddd")
11                        caption = " " & Left$(caption, 1) & "&" & Mid$(caption, 2)
12                    End If
13                Else ' tried another scheme where days of month >= 11 are suffixed by letters but looks clumsy
                      Dim D As Long
                      Dim FirstPart As String
                      Dim LastChar As String
                      Dim Padding As String
14                    D = day(k)
15                    If D <= 10 Then
16                        LastChar = vbNullString
17                    Else
18                        LastChar = "&" & Chr$(D + 54)
19                    End If
20                    If D <= 9 Then
21                        FirstPart = "  &" & Format$(k, "d-mmm-yyyy   ddd")
22                    ElseIf D = 10 Then
23                        FirstPart = "1&0-" & Format$(k, "mmm-yyyy   ddd")
24                    Else
25                        FirstPart = Format$(k, "dd-mmm-yyyy   ddd")
26                    End If
27                    Padding = String(Choose((k Mod 7) + 1, 3, 2, 0, 2, 0, 1, 4) + 5, " ")
28                    caption = FirstPart & Padding & LastChar
29                End If
30                Dctrl.caption = caption
31                Dctrl.OnAction = "'DatePickerSetChosenDate " & CStr(k) & "'"
32                Dctrl.Tag = k
33            Next k
34        End If
35        Exit Sub
ErrHandler:
36        Throw "#DatePickerAddDaysToMonth (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Public Sub DatePickerAddLaterYears(FirstYear As Long)
          Dim Ctrl As CommandBarControl
          Dim i As Long
          Dim YCtrl As CommandBarControl
1         Set Ctrl = Application.CommandBars("DatePicker").FindControl(, , CStr(FirstYear) + "AndLater", , True)
2         If Ctrl.Controls.Count = 0 Then
3             For i = FirstYear To FirstYear + m_ChunkSize - 1
4                 Set YCtrl = Ctrl.Controls.Add(msoControlPopup)
5                 YCtrl.caption = i
6                 YCtrl.Tag = CStr(i)
                  'Trick in line below is explained at:
                  'http://www.tushar-mehta.com/excel/vba/xl%20objects%20and%20procedures%20with%20arguments.htm
7                 YCtrl.OnAction = "'DatePickerAddMonthsToYear " + CStr(i) + " '"
8             Next i
9             Set YCtrl = Ctrl.Controls.Add(msoControlPopup)
10            YCtrl.caption = "&Later"
11            YCtrl.Tag = CStr(FirstYear + m_ChunkSize) + "AndLater"
12            YCtrl.OnAction = "'DatePickerAddLaterYears " + CStr(FirstYear + m_ChunkSize) + " '"
13        End If
End Sub

Public Sub DatePickerAddEarlierYears(FirstYear As Long)
          Dim Ctrl As CommandBarControl
          Dim i As Long
          Dim YCtrl As CommandBarControl
1         Set Ctrl = Application.CommandBars("DatePicker").FindControl(, , CStr(FirstYear) + "AndEarlier", , True)
2         If Ctrl.Controls.Count = 0 Then
3             Set YCtrl = Ctrl.Controls.Add(msoControlPopup)
4             YCtrl.caption = "&Earlier"
5             YCtrl.Tag = CStr(FirstYear - m_ChunkSize) + "AndEarlier"
6             YCtrl.OnAction = "'DatePickerAddEarlierYears " + CStr(FirstYear - m_ChunkSize) + " '"

7             For i = FirstYear - m_ChunkSize + 1 To FirstYear
8                 Set YCtrl = Ctrl.Controls.Add(msoControlPopup)
9                 YCtrl.caption = i
10                YCtrl.Tag = CStr(i)
                  'Trick in line below is explained at:
                  'http://www.tushar-mehta.com/excel/vba/xl%20objects%20and%20procedures%20with%20arguments.htm
11                YCtrl.OnAction = "'DatePickerAddMonthsToYear " + CStr(i) + " '"
12            Next i
13        End If
End Sub

Public Sub DatePickerSetChosenDate(ChosenDate As Long)
1         m_Chosen_Date = ChosenDate
End Sub
