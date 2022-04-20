VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelStateHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' -----------------------------------------------------------------------------------------------------------------------
' Module    : clsExcelStateHandler
' Author    : Philip Swannell
' Date      : 23-Oct-2013
' Purpose   : Use this class when we want to change properties of the Excel Application during
'             macro execution and be sure that those properties revert to their original state
'             when macro execution halts. As of 23/10/13 we handle EnableEvents and ReferenceStyle
'             but we can add more as required.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private m_OrigReferenceStyle
Private m_origEnableEvents
Private m_origCalculation
Private m_OrigStatusBar As String
Private m_origEditDirectlyInCell As Boolean
Private m_PreserveViewport As Boolean
Private m_origRangeSelectionAddress As String
Private m_origActiveCellAddress As String
Private m_origViewPortAddress As String
Private m_origActiveWorkbookname As String
Private m_origActiveSheetName As String
Private m_origActiveWindowScrollRow As Long
Private m_origActiveWindowScrollColumn As Long
Private m_origActiveWindowName As String

Private Sub Class_Initialize()
1         On Error GoTo ErrHandler

2         m_OrigReferenceStyle = Application.ReferenceStyle
3         m_origEnableEvents = Application.EnableEvents
4         m_origCalculation = Application.Calculation
5         m_OrigStatusBar = Application.StatusBar
6         m_origEditDirectlyInCell = Application.EditDirectlyInCell
7         If Not ActiveWindow Is Nothing Then
8             m_origActiveWindowName = ActiveWindow.caption
9         End If

10        Exit Sub
ErrHandler:
11        Throw "#Class_Initialize (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Class_Terminate()
1         On Error GoTo ErrHandler

2         If Application.ReferenceStyle <> m_OrigReferenceStyle Then
3             Application.ReferenceStyle = m_OrigReferenceStyle
4         End If
5         If Application.EnableEvents <> m_origEnableEvents Then
6             Application.EnableEvents = m_origEnableEvents
7         End If
8         If Application.Calculation <> m_origCalculation Then
9             Application.Calculation = m_origCalculation
10        End If
11        If m_OrigStatusBar = vbNullString Or LCase$(m_OrigStatusBar) = "false" Then
12            If VarType(Application.StatusBar) <> vbBoolean Then
13                Application.StatusBar = False
14            End If
15        Else
16            If CStr(Application.StatusBar) <> m_OrigStatusBar Then
17                Application.StatusBar = m_OrigStatusBar
18            End If
19        End If
20        If Application.EditDirectlyInCell <> m_origEditDirectlyInCell Then
21            Application.EditDirectlyInCell = m_origEditDirectlyInCell
22        End If
23        If m_PreserveViewport Then
24            If IsInCollection(Application.Workbooks, m_origActiveWorkbookname) Then
25                If IsInCollection(Application.Workbooks(m_origActiveWorkbookname).Worksheets, m_origActiveSheetName) Then
26                    If IsInCollection(Application.Windows, m_origActiveWindowName) Then
                          Dim ws As Worksheet
27                        Application.Windows(m_origActiveWindowName).Activate
28                        Set ws = Application.Workbooks(m_origActiveWorkbookname).Worksheets(m_origActiveSheetName)
29                        Application.GoTo ws.Range(m_origViewPortAddress)
30                        ws.Range(m_origRangeSelectionAddress).Select
31                        ws.Range(m_origActiveCellAddress).Activate
32                        ActiveWindow.ScrollRow = m_origActiveWindowScrollRow
33                        ActiveWindow.ScrollColumn = m_origActiveWindowScrollColumn
34                    End If
35                End If
36            End If
37        End If

38        Exit Sub
ErrHandler:
39        Throw "#Class_Terminate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Public Sub Init(Optional Calculation As XlCalculation, _
        Optional ReferenceStyle As XlReferenceStyle, _
        Optional EnableEvents As Variant, _
        Optional ByVal StatusBar As String, _
        Optional EditDirectlyInCell As Variant, _
        Optional PreserveViewport As Boolean)
1         On Error GoTo ErrHandler

2         If Calculation <> 0 Then
3             If Application.Calculation <> Calculation Then
4                 Application.Calculation = Calculation
5             End If
6         End If
7         If ReferenceStyle <> 0 Then
8             If Application.ReferenceStyle <> ReferenceStyle Then
9                 Application.ReferenceStyle = ReferenceStyle
10            End If
11        End If
12        If VarType(EnableEvents) = vbBoolean Then
13            If Application.EnableEvents <> EnableEvents Then
14                Application.EnableEvents = EnableEvents
15            End If
16        End If
17        If StatusBar = vbNullString Or LCase$(StatusBar) = "false" Then
18            If VarType(Application.StatusBar) <> vbBoolean Then
19                Application.StatusBar = False
20            End If
21        Else
22            If CStr(Application.StatusBar) <> StatusBar Then
23                Application.StatusBar = StatusBar
24            End If
25        End If
26        If VarType(EditDirectlyInCell) = vbBoolean Then
27            If Application.EditDirectlyInCell <> EditDirectlyInCell Then
28                Application.EditDirectlyInCell = EditDirectlyInCell
29            End If
30        End If
31        m_PreserveViewport = PreserveViewport
32        If m_PreserveViewport Then
              Dim R As Range
33            On Error Resume Next
34            Set R = Selection
35            On Error GoTo ErrHandler
36            If Not R Is Nothing Then
37                If Not ActiveWindow Is Nothing Then
38                    m_origRangeSelectionAddress = R.address
39                    m_origActiveCellAddress = ActiveCell.address
40                    m_origViewPortAddress = ActiveWindow.VisibleRange.address
41                    m_origActiveWindowScrollRow = ActiveWindow.ScrollRow
42                    m_origActiveWindowScrollColumn = ActiveWindow.ScrollColumn
43                    m_origActiveWorkbookname = R.Parent.Parent.Name
44                    m_origActiveSheetName = R.Parent.Name
45                End If
46            End If
47        End If

48        Exit Sub
ErrHandler:
49        Throw "#Init (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

