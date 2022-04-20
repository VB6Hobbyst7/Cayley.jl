VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shxVADashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
          Dim SPH As SolumAddin.clsSheetProtectionHandler
          Dim TheCpty As Variant
1         On Error GoTo ErrHandler
2         Set SPH = CreateSheetProtectionHandler(Me)
3         If Not Application.Intersect(Target, RangeFromSheet(Me, "BottomHeaderRow").CurrentRegion) Is Nothing Then
4             If Target.Row > RangeFromSheet(Me, "BottomHeaderRow").Row Then
5                 TheCpty = Me.Cells(Target.Row, RangeFromSheet(Me, "BottomHeaderRow").Column)

6                 If VarType(TheCpty) = vbString Then
7                     If Len(TheCpty) > 0 Then
8                         Cancel = True
9                         Application.ScreenUpdating = False
10                        shCounterpartyViewer.Activate
11                        Application.GoTo shCounterpartyViewer.Cells(1, 1)
12                        Application.GoTo RangeFromSheet(shCounterpartyViewer, "SelectedCpty")
13                        RangeFromSheet(shCounterpartyViewer, "SelectedCpty").Value = TheCpty
14                    End If
15                End If
16            End If
17        End If
18        Exit Sub
ErrHandler:
19        SomethingWentWrong "#Worksheet_BeforeDoubleClick (line " & CStr(Erl) + "): " & Err.Description & "!", , MsgBoxTitle()
End Sub
