Attribute VB_Name = "Module1"
Option Explicit

Sub ReleaseCleanup()

          Dim ODA As Boolean
          Dim ws As Worksheet

1         On Error GoTo ErrHandler
2         ODA = Application.DisplayAlerts
3         Application.DisplayAlerts = False
4         For Each ws In ThisWorkbook.Worksheets
5             If ws.Name <> "Audit" Then
6                 ws.Delete
7             End If
8         Next
9         Application.DisplayAlerts = True

10        Exit Sub
ErrHandler:
11        Throw "#ReleaseCleanup (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
