VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetProtectionHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' -----------------------------------------------------------------------------------------------------------------------
' Module    : clsSheetProtectionHandler
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Class to simplify the handling of sheet protection. See method CreateSheetProtectionHandler
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private origDrawingObjects As Variant
Private origContents As Variant
Private origScenarios As Variant
Private origUserInterfaceOnly As Variant
Private origAllowFormattingCells As Variant
Private origAllowFormattingColumns As Variant
Private origAllowFormattingRows As Variant
Private origAllowInsertingColumns As Variant
Private origAllowInsertingRows As Variant
Private origAllowInsertingHyperlinks As Variant
Private origAllowDeletingColumns As Variant
Private origAllowDeletingRows As Variant
Private origAllowSorting As Variant
Private origAllowFiltering As Variant
Private origAllowUsingPivotTables As Variant
Private m_ws As Worksheet
Private m_pswd As Variant

Public Sub Init(ws As Worksheet, SetStateTo As Boolean, Optional Password As String)
          Dim pswd As Variant

1         On Error GoTo ErrHandler

2         Set m_ws = ws
3         origDrawingObjects = ws.ProtectDrawingObjects
4         origContents = ws.ProtectContents
5         origScenarios = ws.ProtectScenarios
6         origUserInterfaceOnly = ws.ProtectionMode
7         origAllowFormattingCells = ws.Protection.AllowFormattingCells
8         origAllowFormattingColumns = ws.Protection.AllowFormattingColumns
9         origAllowFormattingRows = ws.Protection.AllowFormattingRows
10        origAllowInsertingColumns = ws.Protection.AllowInsertingColumns
11        origAllowInsertingRows = ws.Protection.AllowInsertingRows
12        origAllowInsertingHyperlinks = ws.Protection.AllowInsertingHyperlinks
13        origAllowDeletingColumns = ws.Protection.AllowDeletingColumns
14        origAllowDeletingRows = ws.Protection.AllowDeletingRows
15        origAllowSorting = ws.Protection.AllowSorting
16        origAllowFiltering = ws.Protection.AllowFiltering
17        origAllowUsingPivotTables = ws.Protection.AllowUsingPivotTables

18        If Password = vbNullString Then
19            pswd = CreateMissing()
20        Else
21            pswd = Password
22        End If
23        m_pswd = pswd

24        If SetStateTo <> origDrawingObjects Or SetStateTo <> origContents Or SetStateTo <> origScenarios Then
25            ws.Protect pswd, SetStateTo, SetStateTo, SetStateTo
26        End If

27        Exit Sub
ErrHandler:
28        Throw "#Init (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Class_Terminate()
1         On Error GoTo ErrHandler

2         m_ws.Protect m_pswd, origDrawingObjects, origContents, origScenarios, origUserInterfaceOnly, origAllowFormattingCells, origAllowFormattingColumns, origAllowFormattingRows, origAllowInsertingColumns, origAllowInsertingRows, origAllowInsertingHyperlinks, origAllowDeletingColumns, origAllowDeletingRows, origAllowSorting, origAllowFiltering, origAllowUsingPivotTables

3         Exit Sub
ErrHandler:
4         Throw "#Class_Terminate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
