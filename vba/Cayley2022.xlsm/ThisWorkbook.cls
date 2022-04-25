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

Private Sub Workbook_Activate()
          'It's possible that we have edited trades in the trades workbook, so clear out cached values
1         FlushStatics
2         SetKeys True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
          Dim LinesBookIsOpen As Boolean
          Dim MarketBookIsOpen As Boolean
          Dim Prompt As String
          Dim ThisWorkbookDotSaved As Boolean
          Dim TradesBookIsOpen As Boolean
          Dim wb As Workbook
          
          'Avoid problem that Workbook_BeforeClose executes even if user hits "Cancel" in the "Want to Save your changes" _
           dialog by emulating that dialog ourselves and exiting early if the user hits cancel.
1         On Error GoTo ErrHandler
2         If Not ThisWorkbook.Saved Then
3             Prompt = "Want to save your changes to '" & ThisWorkbook.Name & "'?"
4             Select Case MsgBoxPlus(Prompt, vbExclamation + vbYesNoCancel, "Cayley2022", "&Save", "Do&n't Save", "Cancel")

                  Case vbYes
5                     On Error Resume Next
6                     If ThisWorkbook.ReadOnly Then
7                         Prompt = "We can't save '" & ThisWorkbook.Name & "' because the file is read-only." & vbLf & vbLf & _
                              "to keep your changes, you'll need to save the workbook with a new name or in a different location."
8                         MsgBoxPlus Prompt, vbExclamation
9                         Application.Dialogs(xlDialogSaveAs).Show
10                        Exit Sub
11                    Else
12                        ThisWorkbook.Save
13                        ThisWorkbookDotSaved = True
14                    End If
15                Case vbNo
16                    ThisWorkbookDotSaved = True
17                Case vbCancel
18                    Cancel = True
19                    Exit Sub
20            End Select
21        End If
          'We are closing the workbook, so shut down Julia (if it's open) and the companion workbooks unless they contain changes
          
22        If Not IsInCollection(Application.Workbooks, "SCRiPT2022.xlsm") Then
23            JuliaEval "exit()"
24            OtherBooksAreOpen MarketBookIsOpen, TradesBookIsOpen, LinesBookIsOpen

25            If MarketBookIsOpen Then
26                Set wb = OpenMarketWorkbook()
27                If wb.Saved Then wb.Close
28            End If

29            If TradesBookIsOpen Then
30                Set wb = OpenTradesWorkbook(True, False)
31                wb.Close
32            End If

33            If LinesBookIsOpen Then
34                Set wb = OpenLinesWorkbook()
35                If wb.Saved Then wb.Close
36            End If

37        End If

38        If ThisWorkbookDotSaved Then
39            ThisWorkbook.Saved = True
40        End If

41        Exit Sub
ErrHandler:
42        SomethingWentWrong "#Workbook_BeforeClose (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub Workbook_Deactivate()
1         SetKeys False
End Sub

Private Sub Workbook_Open()
1         On Error GoTo ErrHandler
2         SetCalculationToManual
3         shConfig.GetFromRegistry
4         shConfig.SaveToRegistry
5         GetETMSFromRegistry
6         SetKeys True
7         Exit Sub
ErrHandler:
8         SomethingWentWrong "#Workbook_Open (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
1         SetKeys True
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
1         SetKeys False
End Sub

