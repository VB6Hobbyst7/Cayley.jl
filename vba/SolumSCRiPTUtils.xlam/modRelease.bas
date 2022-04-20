Attribute VB_Name = "modRelease"
Option Explicit

Sub AuditMenuSolumSCRiPTUtils()
1         On Error GoTo ErrHandler
2         AuditMenuForAddin ThisWorkbook
3         Exit Sub
ErrHandler:
4         Throw "#AuditMenuSolumSCRiPTUtils (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub
