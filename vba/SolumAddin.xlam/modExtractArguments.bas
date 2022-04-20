Attribute VB_Name = "modExtractArguments"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExtractArguments
' Author    : Philip Swannell
' Date      : 14-Dec-2014
' Purpose   : Written at the request of MWB, and hooked through to Ctrl+Shift+G. When executed the
'             formula at the active cell is parsed and the arguments to (the first) function called
'             within that formula are pasted to the range of cells just below the active cell. The built-in
'             "create names" dialog is then displayed, to which the user needs to hit OK.
' -----------------------------------------------------------------------------------------------------------------------
Sub ExtractArguments()
          Dim ArgumentList As Variant
          Dim NumArgs As Long
          Dim TargetCells As Range
          Const MSGBOXTITLE = "Extract Arguments"

1         On Error GoTo ErrHandler
2         If ActiveCell Is Nothing Then Throw "No cell is selected.", True
3         If Not ActiveCell.HasFormula Then Throw "The active cell does not have a formula.", True

4         If Not UnprotectAsk(ActiveSheet, MSGBOXTITLE) Then Exit Sub

5         ArgumentList = sParseArguments(ActiveCell.Formula)
6         ThrowIfError ArgumentList
7         NumArgs = UBound(ArgumentList, 1) - LBound(ArgumentList, 1) + 1
8         Set TargetCells = ActiveCell.Offset(1).Resize(NumArgs, 2)
9         Application.GoTo TargetCells

10        If NumBlanksInRange(TargetCells.Columns(1)) <> NumArgs Then
11            Application.GoTo TargetCells.Columns(1)
12            If MsgBoxPlus("Overwrite these cells?", vbYesNo, MSGBOXTITLE) <> vbYes Then Exit Sub
13            Application.GoTo TargetCells
14        End If

15        TargetCells.Columns(1).Value = ArgumentList
16        Application.Dialogs(xlDialogCreateNames).Show
17        Exit Sub
ErrHandler:
18        SomethingWentWrong "#ExtractArguments (line " & CStr(Erl) + "): " & Err.Description & "!", , MSGBOXTITLE
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sParseArguments
' Author    : Philip Swannell
' Date      : 14-Dec-2014
' Purpose   : A function that accepts the text of a formula such as  =Function(Arg1,Arg2,Arg3) and
'             returns an array consisting of the argument texts, in this case the array: Arg1, Arg2, Arg3
'             sParseArguments is designed to handle quoted text correctly e.g. =Function("A string with commas,","Another string, with commas")
'             is correctly recognised as having two arguments, not four.
'             sParseArguments is also designed to handle nested functions correctly e.g.: =Function1(Function2(Arg1,Arg2),Arg3)
'             is correctly recognised as having two arguments to Function1
'             In outline The function proceeds as follows:
'          1) Take a scratch copy of FormulaText, called FormulaText2
'          2) Within FormulaText2, blank out all characters to the left of which exist an odd number of double quotes
'          3) Within FormulaText2, blank out all characters at which the BracketCount is not positive one.
'          4) Identify the positions of commas in FormulaText2 and use those positions to extract out the arguments from FormulaText
' -----------------------------------------------------------------------------------------------------------------------
Function sParseArguments(ByVal FormulaText As Variant)
          Dim FormulaText2 As String

          Dim BracketCount As Long
          Dim ClosingBracketPos As Long
          Dim DoubleQuoteCount As Long
          Dim i As Long
          Dim OpeningBracketPos As Long

          Const DQ = """"

1         On Error GoTo ErrHandler

          'For convenience
2         If Left(CStr(FormulaText), 1) <> "=" Then
3             If TypeName(FormulaText) = "Range" Then
4                 If FormulaText.Cells(1, 1).HasFormula Then
5                     FormulaText = FormulaText.Cells(1, 1).Formula
6                 End If
7             End If
8         End If

9         FormulaText2 = FormulaText
10        For i = 1 To Len(FormulaText)
11            If Mid$(FormulaText, i, 1) = DQ Then
12                DoubleQuoteCount = DoubleQuoteCount + 1
13            End If
14            If (DoubleQuoteCount Mod 2 = 1) Or (Mid$(FormulaText, i, 1) = DQ And DoubleQuoteCount Mod 2 = 0) Then
15                Mid$(FormulaText2, i, 1) = " "
16            End If
17        Next i

18        For i = 1 To Len(FormulaText2)
19            Select Case Mid$(FormulaText2, i, 1)
                  Case "("
20                    BracketCount = BracketCount + 1
21                Case ")"
22                    BracketCount = BracketCount - 1
23            End Select
24            If BracketCount = 1 Then
25                If OpeningBracketPos = 0 Then
26                    OpeningBracketPos = i
27                End If
28            End If
29            If BracketCount = 0 Then
30                If OpeningBracketPos > 0 Then
31                    ClosingBracketPos = i
32                    Exit For
33                End If
34            End If
35            If BracketCount <> 1 Then
36                Mid$(FormulaText2, i, 1) = " "
37            End If
38        Next i

39        If OpeningBracketPos = 0 Then
40            sParseArguments = "#No opening bracket found in FormulaText!"
41            Exit Function
42        ElseIf ClosingBracketPos = 0 Then
43            sParseArguments = "#No closing bracket found in FormulaText!"
44            Exit Function
45        End If

46        FormulaText = Mid$(FormulaText, OpeningBracketPos + 1, ClosingBracketPos - OpeningBracketPos - 1)
47        FormulaText2 = Mid$(FormulaText2, OpeningBracketPos + 1, ClosingBracketPos - OpeningBracketPos - 1)
          Dim LeftCommaPos As Long
          Dim NumArgs As Long
          Dim ResultArray() As String
          Dim RightCommaPos As Long

48        NumArgs = Len(FormulaText2) - Len(Replace(FormulaText2, ",", vbNullString)) + 1

49        ReDim ResultArray(1 To NumArgs, 1 To 1)

50        For i = 1 To NumArgs
51            LeftCommaPos = RightCommaPos
52            If i = NumArgs Then
53                RightCommaPos = Len(FormulaText) + 1
54            Else
55                RightCommaPos = InStr(RightCommaPos + 1, FormulaText2, ",")
56            End If
57            ResultArray(i, 1) = Mid$(FormulaText, LeftCommaPos + 1, RightCommaPos - LeftCommaPos - 1)
58        Next

59        sParseArguments = ResultArray
60        Exit Function
ErrHandler:
61        sParseArguments = "#sParseArguments (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NumBlanksInRange
' Author    : Philip Swannell
' Date      : 18-Nov-2015 - made method public on this date so can call from other workbooks
' Purpose   : Returns the number of blank (empty) cells in a range
' -----------------------------------------------------------------------------------------------------------------------
Function NumBlanksInRange(TheRange As Range) As Variant
          Dim BlankCells As Range

1         On Error GoTo ErrHandler

2         If TypeName(Application.Caller) = "Range" Then
3             NumBlanksInRange = "#Function NumBlanksInRange cannot be called from a spreadsheet!"        'Since .SpecialCells does not work then
4             Exit Function
5         End If
6         On Error Resume Next
7         Set BlankCells = BlankCellsInRange(TheRange)
8         On Error GoTo ErrHandler
9         If BlankCells Is Nothing Then
10            NumBlanksInRange = 0
11        Else
12            NumBlanksInRange = BlankCells.Cells.CountLarge
13        End If
14        Exit Function
ErrHandler:
15        Throw "#NumBlanksInRange (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
