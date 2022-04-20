VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)

          Const Title = "SCRiPT2022"
          Dim MarketBookIsOpen As Boolean, LinesBookIsOpen As Boolean
          Dim Prompt As String
          Dim ThisWorkbookDotSaved As Boolean
          Dim wb As Workbook
          Dim TradesRange As Range
          Dim NumTrades As Long
          
1         On Error GoTo ErrHandler

          If VBProjectIsLocked Then
3             Prompt = "Close " + ThisWorkbook.Name + "?"
4             Set TradesRange = getTradesRange(NumTrades)
5             If NumTrades > 0 Then
6                 Prompt = Prompt + vbLf + vbLf + "Trades are backed on shutdown. Get them back via:" + vbLf + _
                      "Menu..." + vbLf + "Open Trades File" + vbLf + "Restore trades from backups..."
7             End If

8             If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion + vbDefaultButton2, Title) <> vbOK Then
9                 Cancel = True
10                Exit Sub
11            End If
12            ThisWorkbookDotSaved = True
13        Else
14            If Not ThisWorkbook.Saved Then
15                Prompt = "Want to save your changes to '" + ThisWorkbook.Name + "'?"
16                Select Case MsgBoxPlus(Prompt, vbExclamation + vbYesNoCancel, Title, "&Save", "Do&n't Save", "Cancel")
                      Case vbYes
17                        On Error Resume Next
18                        If ThisWorkbook.ReadOnly Then
19                            Prompt = "We can't save '" + ThisWorkbook.Name + "' because the file is read-only." + vbLf + vbLf + _
                                  "to keep your changes, you'll need to save the workbook with a new name or in a different location."
20                            MsgBoxPlus Prompt, vbExclamation, Title
21                            Application.Dialogs(xlDialogSaveAs).Show
22                            Exit Sub
23                        Else
24                            ThisWorkbook.Save
25                            ThisWorkbookDotSaved = True
26                        End If
27                    Case vbNo
28                        ThisWorkbookDotSaved = True
29                    Case vbCancel
30                        Cancel = True
31                        Exit Sub
32                End Select
33            End If
34        End If

35        BackUpTrades True

36        If Not IsInCollection(Application.Workbooks, "Cayley2022.xlsm") Then
37            JuliaEval "exit()"
38            OtherBooksAreOpen MarketBookIsOpen, LinesBookIsOpen

39            If MarketBookIsOpen Then
40                Set wb = OpenMarketWorkbook()
41                If wb.Saved Then wb.Close
42            End If

43            If LinesBookIsOpen Then
44                Set wb = OpenLinesWorkbook()
45                If wb.Saved Then wb.Close
46            End If

47        End If

48        If ThisWorkbookDotSaved Then ThisWorkbook.Saved = True

49        Exit Sub
ErrHandler:
50        SomethingWentWrong "#Workbook_BeforeClose (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
          Dim Prompt As String
          Dim Res  As VbMsgBoxResult

1         On Error GoTo ErrHandler
2         If VBProjectIsLocked() Then
3             Prompt = "Do you really want to save the workbook? You should probably just save the trades."
4             Res = MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion + vbDefaultButton2, MsgBoxTitle(), "Save &Workbook", "Save &Trades")
5             Select Case Res
                  Case vbYes
                      'Nothing to do
6                 Case vbNo
7                     Cancel = True
8                     SaveTradesFile , False, False, False, False
9                 Case vbCancel
10                    Cancel = True
11            End Select
12        End If
13        Exit Sub
ErrHandler:
14        SomethingWentWrong "#Workbook_BeforeSave (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Sub Workbook_Open()
1         On Error GoTo ErrHandler
2         SetCalculationToManual
3         shAudit.Calculate    'ensure the cell that shows the FullName of the workbook is recalulated
4         shConfig.GetFromRegistry
5         shConfig.SaveToRegistry
6         If sEquals(True, ConfigRange("RestoreTradesAtStartup").Value) Then
7             OpenMostRecentBackup
8         End If
9         Exit Sub
ErrHandler:
10        SomethingWentWrong "#Workbook_Open (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Private Function VBProjectIsLocked() As Boolean
1         On Error GoTo ErrHandler
2         If ThisWorkbook.VBProject.Protection = 1 Then
3             VBProjectIsLocked = True
4         End If
5         Exit Function
ErrHandler:
6         VBProjectIsLocked = True
End Function
