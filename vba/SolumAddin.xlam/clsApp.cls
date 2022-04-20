VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------------------------------------------------------------------
' Module    : clsApp
' Author    : Philip Swannell
' Date      : 07-Nov-2013
' Purpose   : Implements AutoTrace functionality, i.e. precedent arrows are drawn for the active cell whenever the selection changes
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Public WithEvents AppEvents As Application
Attribute AppEvents.VB_VarHelpID = -1
Private BlockSelectionChangeEvent As Boolean
Private m_PreviousSheet As Worksheet
Private m_AutoTraceIsOn As Boolean

Public Property Get AutoTraceIsOn()
1         AutoTraceIsOn = m_AutoTraceIsOn
End Property
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AppEvents_SheetCalculate
' Author    : Philip Swannell
' Date      : 22-May-2016
' Purpose   : Implements the "Refresh pivot tables with sheet calculation" in SolumAddinSettings
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AppEvents_SheetCalculate(ByVal sh As Object)
1         On Error GoTo ErrHandler
          Dim PT As PivotTable
          Dim SPH As clsSheetProtectionHandler
          Static inEvent As Boolean
2         If inEvent Then Exit Sub
3         inEvent = True
4         If sh.PivotTables.Count > 0 Then
5             If GetSetting(gAddinName, "InstallInformation", "RefreshPivotTablesWithSheetCalculation", "False") = "True" Then
6                 Set SPH = CreateSheetProtectionHandler(sh) 'Because pivot tables don't refresh in a protected sheet, attempting to do so would cause an error
7                 For Each PT In sh.PivotTables
8                     PT.RefreshTable
9                 Next PT
10            End If
11        End If
ErrHandler:
12        inEvent = False
End Sub

Public Sub SwitchOnAutoTrace()
1         m_AutoTraceIsOn = True
          Dim rN As Range
2         On Error Resume Next
3         Set rN = ActiveWindow.RangeSelection
4         On Error GoTo 0
5         If Not rN Is Nothing Then
6             UpdateTracerArrows rN
7         End If
End Sub

Public Sub SwitchOffAutoTrace()
1         m_AutoTraceIsOn = False
2         If Not ActiveSheet Is Nothing Then
3             On Error Resume Next
4             If Application.CutCopyMode = False Then m_PreviousSheet.ClearArrows
5             On Error GoTo 0
6         End If
End Sub
Private Function HasPrecedents(TheCell As Range)
          Dim Precedents As Range
1         On Error GoTo ErrHandler
2         If (Not TheCell.HasFormula) Then
3             HasPrecedents = False
4             Exit Function
5         End If
          'Calculating DirectPrecedents fires the selectionChange event so we need a block
6         BlockSelectionChangeEvent = True
7         Set Precedents = TheCell.DirectPrecedents
8         HasPrecedents = True
9         BlockSelectionChangeEvent = False
10        Exit Function
ErrHandler:
11        HasPrecedents = False
12        BlockSelectionChangeEvent = False
End Function

Private Sub AppEvents_SheetBeforeRightClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
1         AmendRightClickCommandBar
End Sub

Private Sub AppEvents_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
1         UpdateTracerArrows Target
End Sub

Private Sub AppEvents_SheetActivate(ByVal sh As Object)
          Dim rN As Range
1         On Error Resume Next
2         Set rN = ActiveWindow.RangeSelection
3         On Error GoTo 0
4         If Not rN Is Nothing Then
5             UpdateTracerArrows rN
6         End If
End Sub
Private Sub AppEvents_SheetDeactivate(ByVal sh As Object)
1         If m_AutoTraceIsOn Then
2             On Error Resume Next
3             If Application.CutCopyMode = False Then m_PreviousSheet.ClearArrows
4             On Error GoTo 0
5         End If
6         RefreshRibbon
End Sub
Private Sub AppEvents_WorkbookActivate(ByVal wb As Excel.Workbook)
1         If m_AutoTraceIsOn Then
              Dim rN As Range
2             On Error Resume Next
3             Set rN = ActiveWindow.RangeSelection
4             On Error GoTo 0
5             If Not rN Is Nothing Then
6                 UpdateTracerArrows rN
7             End If
8         End If
End Sub

Private Sub AppEvents_WorkbookDeactivate(ByVal wb As Excel.Workbook)
1         If m_AutoTraceIsOn Then
2             On Error Resume Next
3             If Application.CutCopyMode = False Then m_PreviousSheet.ClearArrows
4             On Error GoTo 0
5         End If
End Sub
Private Sub AppEvents_WindowActivate(ByVal wb As Excel.Workbook, ByVal wn As Window)
1         If m_AutoTraceIsOn Then
              Dim rN As Range
2             On Error Resume Next
3             Set rN = wn.RangeSelection
4             On Error GoTo 0
5             If Not rN Is Nothing Then
6                 UpdateTracerArrows rN
7             End If
8         End If
9         If NumVisibleWindows = 1 Then
10            RefreshRibbon
11        End If

End Sub
Private Sub AppEvents_WindowDeactivate(ByVal wb As Excel.Workbook, ByVal wn As Window)
1         If m_AutoTraceIsOn Then
2             On Error Resume Next
3             If Application.CutCopyMode = False Then m_PreviousSheet.ClearArrows
4             On Error GoTo 0
5         End If
6         If NumVisibleWindows() = 1 Then
7             RefreshRibbon
8         End If
9         FixFormulaBar
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FixFormulaBar
' Author     : Philip Swannell
' Date       : 01-May-2018
' Purpose    : When there are no visible windows, you can't set the formula bar height manually. That's a little bit annoying.
'              This method fixes it since it's called from AppEvents_WindowDeactivate
' -----------------------------------------------------------------------------------------------------------------------
Private Sub FixFormulaBar()
1         If Not AnyVisibleWindows() Then If Application.DisplayFormulaBar Then If Application.FormulaBarHeight <> 1 Then Application.FormulaBarHeight = 1
End Sub

Private Function AnyVisibleWindows() As Boolean
          Dim w As Window
1         On Error Resume Next
2         For Each w In Application.Windows
3             If w.Visible Then
4                 AnyVisibleWindows = True
5                 On Error GoTo 0
6                 Exit For
7             End If
8         Next
9         On Error GoTo 0
End Function

Private Function NumVisibleWindows() As Long
          Dim N As Long
          Dim w As Window
1         On Error Resume Next
2         For Each w In Application.Windows
3             If w.Visible Then
4                 N = N + 1
5             End If
6         Next
7         NumVisibleWindows = N
8         On Error GoTo 0
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UpdateTracerArrows
' Author    : Philip Swannell
' Date      : 07-Nov-2013
' Purpose   : Put this in a separate method so we can call it from all the events that have
'             the effect of changing the ActiveCell, not just the SelectionChange event
' -----------------------------------------------------------------------------------------------------------------------
Private Sub UpdateTracerArrows(Target As Range)
1         If Not AutoTraceIsOn Then Exit Sub
2         If BlockSelectionChangeEvent Then Exit Sub
3         If TypeName(Target.Parent) <> "Worksheet" Then Exit Sub
4         If Target.Parent.ProtectContents Then Exit Sub

5         If Application.CutCopyMode <> False Then Exit Sub
6         On Error Resume Next
7         m_PreviousSheet.ClearArrows
8         Set m_PreviousSheet = Target.Parent
9         If (HasPrecedents(Target.Cells(1))) Then
10            Target.Cells(1).ShowPrecedents
11        End If
12        On Error GoTo 0
End Sub

Private Sub AppEvents_WorkbookNewSheet(ByVal wb As Excel.Workbook, ByVal sh As Object)
1         RefreshRibbon
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AppEvents_WorkbookOpen
' Author     : Philip Swannell
' Date       : 17-Jan-2018
' Purpose    : I'm really not keen on MS's AutoSave feature, though would be better to have control of whether to
'              switch AutoSave to False when opening workbooks via SolumAddin Settings
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AppEvents_WorkbookOpen(ByVal wb As Excel.Workbook)
1         On Error Resume Next
          Dim Message As String
          Dim Message2 As String
          Const NumSecondsForMessage = 8

2         Select Case LCase$(GetSetting(gAddinName, "InstallInformation", "Application.AutoSave", "LeaveAlone"))
              Case LCase$("SwitchOff")
3                 If wb.AutoSaveOn Then
4                     wb.AutoSaveOn = False
5                     Message = gAddinName & ": AutoSave switched OFF for workbook " + wb.Name + " See ribbon > " & gAddinName & " > Preferences > ""When opening workbooks"""
6                 End If
7             Case LCase$("SwitchOn")
8                 If Not wb.AutoSaveOn Then
9                     wb.AutoSaveOn = True
10                    Message = gAddinName & ": AutoSave switched ON for workbook " + wb.Name + " See ribbon > " & gAddinName & " > Preferences > ""When opening workbooks"""
11                End If
12            Case LCase$("LeaveAlone")
13            Case Else
14        End Select

          Dim LinksChanged As Boolean
          Dim ReferencesChanged As Boolean
15        FixLinks wb, LinksChanged, ReferencesChanged, True
16        If LinksChanged And ReferencesChanged Then
17            Message2 = "Links and VBA References to " & gAddinName & " were automatically corrected"
18        ElseIf LinksChanged Then
19            Message2 = "Links to " & gAddinName & " were automatically corrected"
20        ElseIf ReferencesChanged Then
21            Message2 = "VBA References to " & gAddinName & " were automatically corrected"
22        End If
23        If Message <> vbNullString And Message2 <> vbNullString Then
24            TemporaryMessage Message & ", " & Message2, NumSecondsForMessage
25        ElseIf Message <> vbNullString Then
26            TemporaryMessage Message, NumSecondsForMessage
27        ElseIf Message2 <> vbNullString Then
28            TemporaryMessage Message2, NumSecondsForMessage
29        End If
30        On Error GoTo 0
End Sub

Private Sub Class_Terminate()
1         If m_AutoTraceIsOn Then
2             On Error Resume Next
3             If Application.CutCopyMode = False Then m_PreviousSheet.ClearArrows
4             On Error GoTo 0
5         End If
End Sub
