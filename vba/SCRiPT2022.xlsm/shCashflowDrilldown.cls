VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shCashflowDrilldown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Deactivate()
          Dim oldP

1         Application.OnKey "{ESCAPE}"
2         oldP = ThisWorkbook.ProtectStructure
3         If oldP Then ThisWorkbook.Unprotect
4         Me.Visible = xlSheetHidden
5         If oldP Then ThisWorkbook.Protect

End Sub
