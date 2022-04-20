VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shSpeedGridDrillDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Deactivate()
          Dim oldP

1         oldP = ThisWorkbook.ProtectStructure
2         If oldP Then ThisWorkbook.Unprotect
3         Me.Visible = xlSheetHidden
4         If oldP Then ThisWorkbook.Protect

End Sub

