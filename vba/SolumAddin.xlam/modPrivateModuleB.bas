Attribute VB_Name = "modPrivateModuleB"
Option Explicit
Option Private Module

Function InsertBreaks(ByVal TheString As String)
          Dim i As Long
          Dim LineLength As Long
          Dim Res As String
          Const Width = 90
          Const FirstTab = 0
          Const NextTabs = 13
          Dim DoNewLine As Boolean
          Dim Words
          Dim WordsNLB
1         On Error GoTo ErrHandler
          
2         If InStr(TheString, " ") = 0 Then
3             InsertBreaks = TheString
4             Exit Function
5         End If
          
6         TheString = Replace(TheString, vbLf, vbLf + " ")
7         TheString = Replace(TheString, vbLf + "  ", vbLf + " ")
          
8         Res = String(FirstTab, " ")
9         LineLength = FirstTab

10        Words = sTokeniseString(TheString, " ")
11        WordsNLB = Words
12        For i = 1 To sNRows(Words)
13            WordsNLB(i, 1) = Replace(WordsNLB(i, 1), vbLf, vbNullString)
14        Next

15        For i = 1 To sNRows(Words)
16            DoNewLine = LineLength + Len(WordsNLB(i, 1)) > Width
17            If i > 1 Then
18                If InStr(Words(i - 1, 1), vbLf) > 0 Then
19                    DoNewLine = True
20                End If
21            End If

22            If DoNewLine Then
23                Res = Res + vbLf + "'" + String(NextTabs, " ") + WordsNLB(i, 1)
24                LineLength = 1 + NextTabs + Len(WordsNLB(i, 1))
25            Else
26                Res = Res + " " + WordsNLB(i, 1)
27                LineLength = LineLength + 1 + Len(WordsNLB(i, 1))
28            End If
29        Next
30        InsertBreaks = Res

31        Exit Function
ErrHandler:
32        Throw "#InsertBreaks (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsSortButton
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Decides whether or not a button is a Sort Button
' -----------------------------------------------------------------------------------------------------------------------
Function IsSortButton(b As Button) As Boolean
1         On Error GoTo ErrHandler

2         IsSortButton = False
3         If Len(b.text) = 1 Then
4             If InStr(b.OnAction, "SAISortButtonOnAction") > 0 Then
5                 IsSortButton = True
6             End If
7         End If

8         Exit Function
ErrHandler:
9         Throw "#IsSortButton (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function IsUndoAvailable(SheetToRestoreFrom As Worksheet) As Boolean
1         IsUndoAvailable = Not IsInCollection(SheetToRestoreFrom, "xxx_UndoBufferIsEmpty")
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MappedToUNC
' Author    : Philip
' Date      : 09-Oct-2017
' Purpose   : Morph a file name given using a drive mapping to one given using UNC, wrapped by sFileMappedToUNC
' -----------------------------------------------------------------------------------------------------------------------
Function MappedToUNC(FileName As String)
          Dim D
          Dim FSO As Scripting.FileSystemObject
          Dim SN As String
1         On Error GoTo ErrHandler
2         If Mid$(FileName, 2, 1) = ":" And LCase$(Left$(FileName, 1)) <> "c" Then
3             Set FSO = CreateObject("Scripting.FileSystemObject")
4             On Error Resume Next
5             Set D = FSO.GetDrive(FSO.GetDriveName(FSO.GetAbsolutePathName(FileName)))
6             SN = D.ShareName
7             On Error GoTo ErrHandler
8             If Len(SN) = 0 Then
9                 MappedToUNC = FileName
10            ElseIf InStr(LCase$(SN), "sharepoint") > 0 Then        'don't want to translate addresses mapped to sharepoint since operations (e.g. file reading) won't work on the "raw" web address
11                MappedToUNC = FileName
12            Else
13                MappedToUNC = SN & Mid$(FileName, 3)
14            End If
15        Else
16            MappedToUNC = FileName
17        End If
18        Exit Function
ErrHandler:
19        Throw "#MappedToUNC (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MessagesLogFileName
' Author    : Philip Swannell
' Date      : 05-Nov-2015
' Purpose   : Returns the name of the file to which methods TemporaryMessage and
'             SomethingWentWrong write log data. Files are cleaned up by method
'             CleanTemporaryDirectory. If the file does not exist then this method
'             creates the file and writes a header line to it.
' -----------------------------------------------------------------------------------------------------------------------
Function MessagesLogFileName() As String
          Dim Folder As String
          Dim FSO As Scripting.FileSystemObject
          Dim TS As Scripting.TextStream

1         On Error GoTo ErrHandler

2         Folder = Environ$("Temp")
3         If Right$(Folder, 1) <> "\" Then Folder = Folder + "\"
4         Folder = Folder + "@" + gCompanyName + "Temp\"
5         ThrowIfError sCreateFolder(Folder)

6         MessagesLogFileName = Folder + "MessageLog_" & Format$(Date, "yyyy-mm-dd") + ".txt"

7         If Not sFileExists(MessagesLogFileName) Then
8             Set FSO = New FileSystemObject
9             Set TS = FSO.CreateTextFile(MessagesLogFileName, False, True)
10            TS.WriteLine "Messages written by " & gAddinName & ".xlam " + Format$(Date, "dd-mmm-yyyy")
11            TS.Close
12        End If

13        Exit Function
ErrHandler:
14        Throw "#MessagesLogFileName (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CoreRound
' Author     : Philip Swannell
' Date       : 20-Apr-2018
'              Ties =  0 = Away from zero (like worksheet function ROUND)
'                      1 = Towards zero
'                      2 = Bankers        (like VBA function Round)
'                      3 = Towards plus infinity
'                      4 = Towards minus infinity
' -----------------------------------------------------------------------------------------------------------------------
Function CoreRound(ByVal Number As Double, ByVal NumDigits As Long, Optional Ties = 0)

          Dim RoundRes As Double
          Dim RoundUpRes As Double
          Dim RoundDownRes As Double
          Dim IsTie As Boolean
          Const TiesError = "Ties must be omitted or 0 (ties away from zero), 1 (ties towards zero), 2 (Bankers rounding), 3 (ties towards plus infinity) or 4 (ties towards minus infinity)"

1         On Error GoTo ErrHandler
2         If Not IsNumber(Ties) Then Throw TiesError
3         If Ties = 2 Then
              Dim Exponent As Double
              'VBA function Round implements Bankers rounding for ties, but Round does not support negative NumDigits
4             If NumDigits >= 0 Then
5                 CoreRound = Round(Number, NumDigits)
6             Else
7                 Exponent = 10 ^ -NumDigits
8                 CoreRound = Round(Number / Exponent) * Exponent
9             End If
10            Exit Function
11        End If
          'Excel function ROUND does away-from-zero for ties
12        RoundRes = Application.WorksheetFunction.Round(Number, NumDigits)
13        Select Case Ties
              Case 0
14                CoreRound = RoundRes
15                Exit Function
16            Case 1, 3, 4
                  'This will be rather slow.
17                RoundUpRes = Application.WorksheetFunction.RoundUp(Number, NumDigits)
18                RoundDownRes = Application.WorksheetFunction.RoundDown(Number, NumDigits)
                  
19                IsTie = Number = (RoundUpRes + RoundDownRes) / 2
20                If IsTie Then
21                    If Ties = 1 Then
22                        CoreRound = RoundDownRes
23                    ElseIf Ties = 3 Then
24                        CoreRound = Max(RoundDownRes, RoundUpRes)
25                    Else
26                        CoreRound = Min(RoundDownRes, RoundUpRes)
27                    End If
28                Else
29                    CoreRound = RoundRes 'Not a tie so Excel ROUND function is fine
30                End If
31            Case Else
32                Throw TiesError
33        End Select

34        Exit Function
ErrHandler:
35        CoreRound = "#CoreRound (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ProcessAmpersands
' Author    : Philip Swannell
' Date      : 22-Nov-2015
' Purpose   : For use from forms, decodes a string in which "&" means "next character is an accelerator character,
'             don't display the &" and "&&" means "display a single &" and returns the text to display.
' -----------------------------------------------------------------------------------------------------------------------
Function ProcessAmpersands(ByVal caption As String, ByRef Accelerator As String)
          Dim MatchRes As Long
          Dim Result As String
1         On Error GoTo ErrHandler

          Dim HighAsc As Long        'find a character that's not in the Caption so we can use it as a "placeholder"
2         HighAsc = 37
3         Do While InStr(caption, Chr$(HighAsc)) > 0
4             HighAsc = HighAsc + 1
5         Loop

6         Result = Replace(caption, "&&", Chr$(HighAsc))

7         MatchRes = InStr(Result, "&")
8         If MatchRes > 0 Then
9             Result = Left$(Result, MatchRes - 1) + Mid$(Result, MatchRes + 1)
10            Accelerator = UCase$(Mid$(caption, MatchRes + 1, 1))
11        Else
12            Accelerator = vbNullString
13        End If

14        Result = Replace(Result, Chr$(HighAsc), "&")
15        ProcessAmpersands = Result
16        Exit Function
ErrHandler:
17        Throw "ProcessAmpersands (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function RegExSyntaxValid(RegularExpression As String) As Boolean
          Dim Res As Boolean
          Dim rx As VBScript_RegExp_55.RegExp
1         On Error GoTo ErrHandler
2         Set rx = New RegExp
3         With rx
4             .IgnoreCase = False
5             .Pattern = RegularExpression
6             .Global = False        'Find first match only
7         End With
8         Res = rx.Test("Foo")
9         RegExSyntaxValid = True
10        Exit Function
ErrHandler:
11        RegExSyntaxValid = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sAddinReleaseDate
' Author    : Philip Swannell
' Date      : 24-Dec-2015
' Purpose   : Grabs the DateTime of this version from the Audit sheet
' -----------------------------------------------------------------------------------------------------------------------
Function sAddinReleaseDate()
          Dim TheDate As Long
          Dim TheTime As Double

1         On Error GoTo ErrHandler
2         TheDate = shAudit.Range("Headers").Cells(2, 2).Value2
3         TheTime = shAudit.Range("Headers").Cells(2, 3).Value2
4         TheTime = TheTime - Application.WorksheetFunction.RoundDown(TheTime, 0)

5         sAddinReleaseDate = TheDate + TheTime
6         Exit Function
ErrHandler:
7         sAddinReleaseDate = "#sAddinReleaseDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sAddinVersionNumber
' Author    : Philip Swannell
' Date      : 24-Dec-2015
' Purpose   : Grabs the version number from the Audit sheet
' -----------------------------------------------------------------------------------------------------------------------
Function sAddinVersionNumber()
1         On Error GoTo ErrHandler
2         sAddinVersionNumber = shAudit.Range("Headers").Cells(2, 1).Value
3         Exit Function
ErrHandler:
4         sAddinVersionNumber = "#sAddinVersionNumber (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function SafeAbs(Number)
1         On Error GoTo ErrHandler
2         If IsNumberOrDate(Number) Then
3             SafeAbs = Abs(Number)
4         Else
5             SafeAbs = "#Non-number found!"
6         End If
7         Exit Function
ErrHandler:
8         SafeAbs = "#" + Err.Description + "!"
End Function

Function SafeAdd(x, y)
1         On Error GoTo ErrHandler
2         If Not IsNumberOrDate(x) Then
3             SafeAdd = "#Non-number found!"
4         ElseIf Not IsNumberOrDate(y) Then
5             SafeAdd = "#Non-number found!"
6         Else
7             SafeAdd = x + y
8         End If
9         Exit Function
ErrHandler:
10        SafeAdd = "#" + Err.Description & "!"
End Function

Function SafeAnd(a, b)
1         On Error GoTo ErrHandler
2         If VarType(a) <> vbBoolean Or VarType(b) <> vbBoolean Then
3             SafeAnd = "#Non-Boolean found!"
4         Else
5             SafeAnd = a And b
6         End If
7         Exit Function
ErrHandler:
8         SafeAnd = "#" & Err.Description & "!"
End Function

Function SafeConcatenate(a, b)
1         On Error GoTo ErrHandler
2         If IsError(a) Then
3             SafeConcatenate = a
4         ElseIf IsError(b) Then
5             SafeConcatenate = b
6         Else
7             SafeConcatenate = CStr(a) & CStr(b)
8         End If
9         Exit Function
ErrHandler:
10        SafeConcatenate = "#" + Err.Description + "!"
End Function

Function SafeCStr(x As Variant)
1         On Error GoTo ErrHandler
2         SafeCStr = CStr(x)
3         Exit Function
ErrHandler:
4         SafeCStr = "#Cannot represent " + TypeName(x) + "!"
End Function

Function SafeDivide(a, b)
1         On Error GoTo ErrHandler

2         If Not IsNumberOrDate(a) Then
3             SafeDivide = "#Non-number found!"
4         ElseIf Not IsNumberOrDate(b) Then
5             SafeDivide = "#Non-number found!"
6         Else
7             SafeDivide = a / b
8         End If

9         Exit Function
ErrHandler:
10        SafeDivide = "#" + Err.Description + "!"
11        Exit Function
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedures : Safe... functions: SafeAbs, SafeAdd, SafeAnd, SafeConcatenate, SafeCStr,
'              SafeDivide, SafeExp, SafeIf, SafeIfErrorString, SafeInterp, SafeLeft, SafeLog,
'              SafeMax, SafeMin, SafeMultiply, SafeOr, SafePower, SafeRight, CoreStringBetweenStrings,
'              SafeSubtract
' Author    : Philip Swannell
' Date      : 30-Apr-2015
' Purpose   : So that we can get element-wise error handling in sArrayAdd etc. Also this module
'             is Option Private Module to stop the functions being visible themselves in Excel
' -----------------------------------------------------------------------------------------------------------------------
Function SafeExp(Number)
1         On Error GoTo ErrHandler
2         If IsNumberOrDate(Number) Then
3             SafeExp = Exp(Number)
4         Else
5             SafeExp = "#Non-number found!"
6         End If
7         Exit Function
ErrHandler:
8         SafeExp = "#" + Err.Description + "!"
End Function

Function SafeIf(IfCondition, ValueIfTrue, ValueIfFalse)
1         On Error GoTo ErrHandler
2         If VarType(IfCondition) <> vbBoolean Then Throw "Non Boolean detected"
3         SafeIf = IIf(IfCondition, ValueIfTrue, ValueIfFalse)
4         Exit Function
ErrHandler:
5         SafeIf = "#" & Err.Description & "!"
End Function

Function SafeIfErrorString(x, y)
1         On Error GoTo ErrHandler
2         SafeIfErrorString = x
3         If VarType(x) = vbString Then
4             If Left$(x, 1) = "#" Then
5                 If Right$(x, 1) = "!" Then
6                     SafeIfErrorString = y
7                 End If
8             End If
9         End If
10        Exit Function
ErrHandler:
11        Throw SafeIfErrorString = "#" + Err.Description + "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SafeInterp
' Author    : Philip Swannell
' Date      : 28-May-2015
' Purpose   : sub of sInterp
' -----------------------------------------------------------------------------------------------------------------------
Function SafeInterp(x, x1, X2, y1, y2)
1         On Error GoTo ErrHandler
2         SafeInterp = (y1 * (X2 - x) + y2 * (x - x1)) / (X2 - x1)
3         Exit Function
ErrHandler:
4         SafeInterp = "#" & Err.Description & "!"
End Function

Function SafeLeft(AString, NumChars)
1         On Error GoTo ErrHandler
2         If NumChars < 0 Then
3             SafeLeft = Left$(AString, Len(AString) + NumChars)
4         Else
5             SafeLeft = Left$(AString, NumChars)
6         End If
7         Exit Function
ErrHandler:
8         SafeLeft = "#" & Err.Description & "!"
End Function

Function SafeLog(Number)
1         On Error GoTo ErrHandler
2         If IsNumberOrDate(Number) Then
3             SafeLog = Log(Number)
4         Else
5             SafeLog = "#Non-number found!"
6         End If

7         Exit Function
ErrHandler:
8         SafeLog = "#" + Err.Description + "!"
End Function

Function SafeMax(a, b)
1         On Error GoTo ErrHandler
2         If Not IsNumberOrDate(a) Then
3             SafeMax = "#Non-number found!"
4         ElseIf Not IsNumberOrDate(b) Then
5             SafeMax = "#Non-number found!"
6         ElseIf a > b Then
7             SafeMax = a
8         Else
9             SafeMax = b
10        End If
11        Exit Function
ErrHandler:
12        SafeMax = "#" & Err.Description & "!"
End Function

Function SafeMedian2(Numbers)
1         On Error GoTo ErrHandler
2         SafeMedian2 = Application.WorksheetFunction.Median(Numbers)
3         Exit Function
ErrHandler:
4         SafeMedian2 = CVErr(xlErrNA)
End Function

Function SafeMedian(Numbers)
1         On Error GoTo ErrHandler
2         SafeMedian = Application.WorksheetFunction.Median(Numbers)
3         Exit Function
ErrHandler:
4         SafeMedian = "#SafeMedian (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function SafeMin(a, b)
1         On Error GoTo ErrHandler
2         If Not IsNumberOrDate(a) Then
3             SafeMin = "#Non-number found!"
4         ElseIf Not IsNumberOrDate(b) Then
5             SafeMin = "#Non-number found!"
6         ElseIf a > b Then
7             SafeMin = b
8         Else
9             SafeMin = a
10        End If
11        Exit Function
ErrHandler:
12        SafeMin = "#" & Err.Description & "!"
End Function

Function SafeMultiply(a, b)
1         On Error GoTo ErrHandler
2         If Not IsNumberOrDate(a) Then
3             SafeMultiply = "#Non-number found!"
4         ElseIf Not IsNumberOrDate(b) Then
5             SafeMultiply = "#Non-number found!"
6         Else
7             SafeMultiply = a * b
8         End If
9         Exit Function
ErrHandler:
10        SafeMultiply = "#" & Err.Description + "!"
End Function

Function SafeOr(a, b)
1         On Error GoTo ErrHOrler
2         If VarType(a) <> vbBoolean Or VarType(b) <> vbBoolean Then
3             SafeOr = "#Non Boolean found!"
4         Else
5             SafeOr = a Or b
6         End If
7         Exit Function
ErrHOrler:
8         SafeOr = "#" & Err.Description & "!"
End Function

Function SafePower(Base, Exponent)
1         On Error GoTo ErrHandler
2         SafePower = Base ^ Exponent
3         Exit Function
ErrHandler:
4         SafePower = "#" & Err.Description & "!"
End Function

Function SafeRight(AString, NumChars)
1         On Error GoTo ErrHandler
2         If NumChars < 0 Then
3             SafeRight = Right$(AString, Len(AString) + NumChars)
4         Else
5             SafeRight = Right$(AString, NumChars)
6         End If
7         Exit Function
ErrHandler:
8         SafeRight = "#" & Err.Description & "!"
End Function

Function CoreStringBetweenStrings(TheString, LeftString, RightString, Optional IncludeLeftString As Boolean, Optional IncludeRightString As Boolean)
          Dim MatchPoint1 As Long        ' the position of the first character to return
          Dim MatchPoint2 As Long        ' the position of the last character to return
          Dim FoundLeft As Boolean
          Dim FoundRight As Boolean
          
1         On Error GoTo ErrHandler
          
2         If VarType(TheString) <> vbString Or VarType(LeftString) <> vbString Or VarType(RightString) <> vbString Then Throw "Inputs must be strings"
3         If LeftString = vbNullString Then
4             MatchPoint1 = 0
5         Else
6             MatchPoint1 = InStr(1, TheString, LeftString, vbTextCompare)
7         End If

8         If MatchPoint1 = 0 Then
9             FoundLeft = False
10            MatchPoint1 = 1
11        Else
12            FoundLeft = True
13        End If

14        If RightString = vbNullString Then
15            MatchPoint2 = 0
16        ElseIf FoundLeft Then
17            MatchPoint2 = InStr(MatchPoint1 + Len(LeftString), TheString, RightString, vbTextCompare)
18        Else
19            MatchPoint2 = InStr(1, TheString, RightString, vbTextCompare)
20        End If

21        If MatchPoint2 = 0 Then
22            FoundRight = False
23            MatchPoint2 = Len(TheString)
24        Else
25            FoundRight = True
26            MatchPoint2 = MatchPoint2 - 1
27        End If

28        If Not IncludeLeftString Then
29            If FoundLeft Then
30                MatchPoint1 = MatchPoint1 + Len(LeftString)
31            End If
32        End If

33        If IncludeRightString Then
34            If FoundRight Then
35                MatchPoint2 = MatchPoint2 + Len(RightString)
36            End If
37        End If

38        CoreStringBetweenStrings = Mid$(TheString, MatchPoint1, MatchPoint2 - MatchPoint1 + 1)

39        Exit Function
ErrHandler:
40        CoreStringBetweenStrings = "#CoreStringBetweenStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SafeSubtract
' Author    : Philip Swannell
' Date      : 22-Jun-2016
' Purpose   : low-level subtraction with error handling
' -----------------------------------------------------------------------------------------------------------------------
Function SafeSubtract(a, b)
1         On Error GoTo ErrHandler
2         If Not IsNumberOrDate(a) Then
3             SafeSubtract = "#Non-number found!"
4         ElseIf Not IsNumberOrDate(b) Then
5             SafeSubtract = "#Non-number found!"
6         Else
7             SafeSubtract = a - b
8         End If
9         Exit Function
ErrHandler:
10        SafeSubtract = "#" & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : sPromote
' Author     : Philip Swannell
' Date       : 03-May-2018
' Purpose    : Called from frmMultipleChoice, and I can't imagine I would have other need for it, hence in this Private Module.
'              Returns a re-ordered array in which the chosen elements have been moved up or down by Steps places
' Parameters :
'  TheArray    : A (2d) columm array of arbitrary values, typically strings
'  ChooseVector: A (2d) columm array of Booleans, indicate which are the elements of TheArray that should be moved up or down.
'                Note that this argument is ByRef and is amended so that it indicates the new positions of the chosen elements.
'  Steps       : The number of places that each of the chosen elements should be moved up (Steps negative) or down (Steps positive)
' -----------------------------------------------------------------------------------------------------------------------
Function sPromote(TheArray As Variant, ByRef ChooseVector As Variant, Steps As Long)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim NumIn As Long
          Dim NumOut As Long
          Dim x As Variant

          Dim FinalOrder As Variant
          Dim InIndexes As Variant
          Dim LookupTable As Variant
          Dim NewInIndexes As Variant
          Dim NewOutIndexes As Variant
          Dim OutIndexes As Variant
          Dim Result As Variant

1         On Error GoTo ErrHandler
2         NC = sNCols(TheArray)
3         NR = sNRows(TheArray)
4         If sNRows(ChooseVector) <> NR Then Throw "TheArray and ChooseVector must have the same number of rows"
5         For Each x In ChooseVector
6             If VarType(x) <> vbBoolean Then Throw "ChooseVector must be 1-column array of logical values"
7         Next
8         NumIn = sArrayCount(ChooseVector)
9         NumOut = NR - NumIn

10        If NumIn = 0 Or NumOut = 0 Or Steps = 0 Then
11            sPromote = TheArray
12            Exit Function
13        End If

14        InIndexes = sMChoose(sIntegers(NR), ChooseVector)
15        OutIndexes = sMChoose(sIntegers(NR), sArrayNot(ChooseVector))
16        NewInIndexes = sArrayAdd(InIndexes, Steps)

          'Avoid "dropping off the end"
17        If Steps < 0 Then
18            NewInIndexes(1, 1) = SafeMax(1, NewInIndexes(1, 1))
19            For i = 2 To NumIn
20                NewInIndexes(i, 1) = SafeMax(NewInIndexes(i - 1, 1) + 1, NewInIndexes(i, 1))
21            Next
22        Else
23            NewInIndexes(NumIn, 1) = SafeMin(NR, NewInIndexes(NumIn, 1))
24            For i = NumIn - 1 To 1 Step -1
25                NewInIndexes(i, 1) = SafeMin(NewInIndexes(i + 1, 1) - 1, NewInIndexes(i, 1))
26            Next
27        End If

28        NewOutIndexes = sSubArray(sCompareTwoArrays(sIntegers(NR), NewInIndexes, "In1AndNotIn2"), 2)

29        LookupTable = sArraySquare(InIndexes, NewInIndexes, OutIndexes, NewOutIndexes)
30        FinalOrder = sVLookup(sIntegers(NR), LookupTable)

31        Result = sReshape(vbNullString, NR, NC)
32        For i = 1 To NR
33            For j = 1 To NC
34                Result(FinalOrder(i, 1), j) = TheArray(i, j)
35            Next j
36        Next i

          'ByRef argument - set it to the new indicators for the selected items after promotion
37        ChooseVector = sArrayIsNumber(sMatch(sIntegers(NR), NewInIndexes))
38        sPromote = Result

39        Exit Function
ErrHandler:
40        sPromote = "#sPromote (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : StringLessThan
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : String comparison with capital letters being treated as adjacent to their
'             lower case equivalents and accented characters being adjacent to their
'             non-accented equivalents. We have for example e < E < é < É < f < F
'             The sort-order is taken from that implemented by Microsoft in Excel's in-built
'             sorting (in case-sensitive mode).
'             Returns True if String A is "less than" string B. Returns False otherwise and in
'             particular returns False when the strings are the same.
' -----------------------------------------------------------------------------------------------------------------------
Function StringLessThan(ByVal a As String, ByVal b As String, CaseSensitive As Boolean) As Boolean
          Dim i As Long
          Dim LA As Long
          Dim LB As Long
          Dim soA As Long
          Dim soB As Long
          Static SortOrder As Variant

1         On Error GoTo ErrHandler

2         If IsEmpty(SortOrder) Then

3             SortOrder = VBA.Array(0, 2, 3, 4, 5, 6, 7, 8, 9, 35, 36, 37, 38, 39, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                  21, 22, 23, 24, 25, 26, 27, 33, 40, 41, 42, 43, 44, 45, 29, 46, 47, 48, 88, 49, 30, 50, _
                  51, 115, 119, 121, 123, 125, 126, 127, 128, 129, 130, 52, 53, 89, 90, 91, 54, 55, 132, _
                  149, 151, 155, 159, 169, 172, 174, 176, 186, 188, 190, 192, 194, 198, 215, 217, 219, 221, _
                  226, 231, 241, 243, 245, 247, 253, 56, 57, 58, 59, 61, 62, 131, 148, 150, 154, 158, 168, _
                  171, 173, 175, 185, 187, 189, 191, 193, 197, 214, 216, 218, 220, 225, 230, 240, 242, 244, _
                  246, 252, 63, 64, 65, 66, 28, 87, 110, 77, 170, 80, 105, 106, 107, 60, 109, 223, 81, 213, _
                  111, 255, 112, 113, 75, 76, 78, 79, 108, 31, 32, 74, 229, 222, 82, 212, 114, 254, 251, 34, _
                  67, 83, 84, 85, 86, 68, 97, 69, 98, 133, 93, 99, 1, 100, 70, 101, 92, 122, 124, 71, 102, _
                  103, 104, 72, 120, 199, 94, 116, 117, 118, 73, 137, 135, 139, 143, 141, 145, 147, 153, 163, _
                  161, 165, 167, 180, 178, 182, 184, 157, 196, 203, 201, 205, 209, 207, 95, 211, 235, 233, _
                  237, 239, 249, 228, 224, 136, 134, 138, 142, 140, 144, 146, 152, 162, 160, 164, 166, 179, _
                  177, 181, 183, 156, 195, 202, 200, 204, 208, 206, 96, 210, 234, 232, 236, 238, 248, 227, 250)
4         End If

5         If Not CaseSensitive Then a = LCase$(a): b = LCase$(b)        'Might be faster to handle case insensitive comparison with a different sort order that maps lower and upper case version of the same letter to the same position...

6         LA = Len(a)
7         LB = Len(b)
8         If LA = 0 Then
9             StringLessThan = (LB > 0)
10        ElseIf LB = 0 Then
11            StringLessThan = False
12        Else
13            i = 1
14            Do While Mid$(a, i, 1) = Mid$(b, i, 1) And i < LA And i < LB
15                i = i + 1
16            Loop
17            soA = SortOrder(Asc(Mid$(a, i, 1)))
18            soB = SortOrder(Asc(Mid$(b, i, 1)))

19            If soA = soB Then
20                StringLessThan = LA < LB
21            Else
22                StringLessThan = soA < soB
23            End If
24        End If

25        Exit Function
ErrHandler:
26        Throw "#StringLessThan (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StripNonAscii
' Author     : Philip Swannell
' Date       : 30-Apr-2019
' Purpose    : Removes all characters with AscW > 255 from string x
' -----------------------------------------------------------------------------------------------------------------------
Function StripNonAscii(x As String)
          Dim AnyFound As Boolean
          Dim i As Long
          Dim Res As String
          Dim StartFrom As Long

1         On Error GoTo ErrHandler
2         For i = 1 To Len(x)
3             If AscW(Mid$(x, i, 1)) > 255 Then
4                 AnyFound = True
5                 StartFrom = i
6                 Exit For
7             End If
8         Next i

9         If AnyFound Then
10            Res = Left$(x, StartFrom - 1)
11            For i = StartFrom To Len(x)
12                If AscW(Mid$(x, i, 1)) <= 255 Then
13                    Res = Res & Mid$(x, i, 1)
14                End If
15            Next i
16            StripNonAscii = Res
17        Else
18            StripNonAscii = x
19        End If

20        Exit Function
ErrHandler:
21        Throw "#StripNonAscii (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : UnprotectAsk
' Author    : Philip Swannell
' Date      : 09-Oct-2018
' Purpose   : Common code to use in methods that need the user's approval to make changes to the sheet.
' -----------------------------------------------------------------------------------------------------------------------
Function UnprotectAsk(TargetSheet As Worksheet, Optional Title As String = gAddinName, Optional TargetRange As Range) As Boolean
          Dim Prompt As String
          Dim Res As VbMsgBoxResult
1         On Error GoTo ErrHandler
2         Prompt = "You cannot use this command on a protected sheet. To use this" & _
              " command, you must first unprotect the sheet (Review tab, Changes group, Unprotect Sheet button)." & _
              " You may be prompted for a password."

3         If Not TargetRange Is Nothing Then
4             If Not TargetRange.Parent Is TargetSheet Then
5                 Throw "TargetRange must be a range of TargetSheet"
6             End If
7             If Not IsNull(TargetRange.Locked) Then
8                 If TargetRange.Locked = False Then
9                     UnprotectAsk = True
10                    Exit Function
11                End If
12            End If
13        End If

14        If SheetIsProtectedWithPassword(TargetSheet) Then
15            Prompt = "You cannot use this command on a protected sheet. To use this" & _
                  " command, you must first unprotect the sheet (Review tab, Changes group, Unprotect Sheet button)." & _
                  " You will be prompted for a password."

16            Res = MsgBoxPlus(Prompt, vbOKOnly + vbInformation, Title)
17            UnprotectAsk = False
18            Exit Function
19        ElseIf TargetSheet.ProtectContents Then
20            Prompt = "You cannot use this command on a protected sheet. To use this" & _
                  " command, you must first unprotect the sheet (Review tab, Changes group, Unprotect Sheet button)."
21            Res = MsgBoxPlus(Prompt, vbOKCancel + vbQuestion + vbDefaultButton2, Title, "Unprotect Now!", "OK")
22            If Res = vbOK Then
23                TargetSheet.Protect , False
24                UnprotectAsk = True
25            Else
26                UnprotectAsk = False
27            End If
28        Else
29            UnprotectAsk = True
30        End If

31        Exit Function
ErrHandler:
32        Throw "#UnprotectAsk (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : VariantLessThan
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Compare two values which may be of types Empty, Double, String, Boolean, Error, Byte, Single, Currency, Decimal
'             Returns True if A < B or False otherwise (A >= B)
'             Ordering is Numbers < Strings < Booleans < Errors, with Dates treated as
'             the number to which they cast and Empty being treated as zero. Ordering of
'             two strings is handled by the method StringLessThan.
' -----------------------------------------------------------------------------------------------------------------------
Function VariantLessThan(a As Variant, b As Variant, CaseSensitive As Boolean)
          Dim VTA As Long
          Dim VTB As Long
1         On Error GoTo ErrHandler

2         VTA = VarType(a)
3         VTB = VarType(b)
4         If VTA = VTB Then
5             Select Case VTA
                  Case vbString
6                     VariantLessThan = StringLessThan(CStr(a), CStr(b), CaseSensitive)
7                 Case vbBoolean
                      'We have False < True which is the same as Excel's sorting but different _
                       from native VBA which has True < False
8                     VariantLessThan = (Not a) And b
9                 Case Else
10                    VariantLessThan = a < b
                      'For all other types we are happy with the native VBA sorting. In particular _
                       for error values  coming from worksheet cells the order appears to be: _
                            #NULL! < #DIV/0! < #VALUE! < #REF! < #NAME? < #NUM! < #N/A. Excel sorting _
                          treats all these error values as the same as one-another and their relative _
                          position in a sort remains the same.
11                    VariantLessThan = a < b
12            End Select
13        Else
              'We want Numbers < Strings < Booleans < Errors, with Dates treated as the number to which they cast
              'Numbers can be Byte, Integer, Long, Currency, Single, Decimal or Double
14            If VTA = vbError Then
15                VariantLessThan = False
16            ElseIf VTB = vbError Then
17                VariantLessThan = True
18            ElseIf VTA = vbBoolean Then
19                VariantLessThan = False
20            ElseIf VTB = vbBoolean Then
21                VariantLessThan = True
22            ElseIf VTA = vbString Then
23                VariantLessThan = False
24            ElseIf VTB = vbString Then
25                VariantLessThan = True
26            Else
                  'Both A & B must be Date or Number (i.e. Byte, Integer, Long, Single, Double, Decimal) _
                   for which it should A < B should yield the result we want...
27                VariantLessThan = a < b
28            End If
29        End If

30        Exit Function
ErrHandler:
31        Throw "#VariantLessThan (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function VariantLessThanOrEqual(a, b, CaseSensitive As Boolean) As Boolean
1         If VariantLessThan(a, b, CaseSensitive) Then
2             VariantLessThanOrEqual = True
3         ElseIf sEquals(a, b, CaseSensitive) Then
4             VariantLessThanOrEqual = True
5         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : WorkbookAndAddInList
' Author    : Philip Swannell
' Date      : 22-Oct-2013
' Purpose   : Emulates Excel4 function DOCUMENTS(), but no longer (as of 3 Dec 2018) uses
'             any Excel4 code since that was causing bizarre bugs when called from Ribbon callbacks.
' DocumentType = 1 for Workbooks, 2 for AddIns, 3 for both workbooks and AddIns
'             Returns a column array or an Empty variable if there are no documents open
' -----------------------------------------------------------------------------------------------------------------------
Function WorkbookAndAddInList(DocumentType As Long, Optional App As Application)
          Dim FoundSome As Boolean
          Dim i As Long
          Dim st As clsStacker
          Dim wb As Excel.Workbook

1         If App Is Nothing Then Set App = Excel.Application

2         Set st = CreateStacker()
3         i = 1

4         If DocumentType = 1 Or DocumentType = 3 Then
5             For Each wb In App.Workbooks
6                 st.Stack0D wb.Name
7                 FoundSome = True
8             Next
9         End If

10        If DocumentType = 2 Or DocumentType = 3 Then
              Dim a As Object
11            For Each a In Application.AddIns2
12                If IsInCollection(Application.Workbooks, a.Name) Then
13                    If Application.Workbooks(a.Name).isAddin Then
14                        FoundSome = True
15                        st.Stack0D a.Name
16                    End If
17                End If
18            Next a
19        End If

20        If FoundSome Then 'Have seen cases where the Addins2 collection contains repeats, hence remove duplicates
21            WorkbookAndAddInList = sRemoveDuplicates(st.Report, True)
22        Else
23            WorkbookAndAddInList = Empty
24        End If
End Function

Function RegExpFromLiteral(StringLiteral As String, Invert As Boolean)
          Dim i As Long
          Dim Res As String
          Const TheChars = ".$^{}[]()|*+?"
1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler

3         If Not Invert Then
4             Res = Replace(StringLiteral, "\", "\\")
5             For i = 1 To Len(TheChars)
6                 Res = Replace(Res, Mid$(TheChars, i, 1), "\" + Mid$(TheChars, i, 1))
7             Next i
8         Else
9             Res = StringLiteral
10            For i = Len(TheChars) To 1 Step -1
11                Res = Replace(Res, "\" + Mid$(TheChars, i, 1), Mid$(TheChars, i, 1))
12            Next i
13            Res = Replace(Res, "\\", "\")
14        End If

15        RegExpFromLiteral = Res
16        Exit Function
ErrHandler:
17        Throw "#RegExpFromLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ExcelSupportsSpill
' Author     : Philip Swannell
' Date       : 04-Jan-2020
' Purpose    : Returns TRUE if current installation of Excel supports dynamic array formulas.
' -----------------------------------------------------------------------------------------------------------------------
Function ExcelSupportsSpill() As Boolean
1         On Error GoTo ErrHandler
2         ExcelSupportsSpill = Application.WorksheetFunction.Sequence(2)(2, 1) = 2
3         Exit Function
ErrHandler:
4         ExcelSupportsSpill = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExpandRangeToIncludeEntireArrayFormulas
' Author    : Philip Swannell
' Date      : 03-Nov-2015
' Purpose   : Returns a range expanded so that any array formulas within the input TheRange
'             are entirely included in the output. Handles both old-style CSE array formulas
'             and new-style dynamic array formulas.
' -----------------------------------------------------------------------------------------------------------------------
Function ExpandRangeToIncludeEntireArrayFormulas(TheRange As Range) As Range
          Dim c As Range
          Dim CurrentArray As Range
          Dim Exists As Boolean
          Dim RangeToReturn As Range
          Dim RangeToSearch As Range
          Dim SPH As clsSheetProtectionHandler
          Dim TempRange As Range

1         On Error GoTo ErrHandler
2         If Not ExcelSupportsSpill() Then
3             If sEquals(EdgeCells(TheRange).HasArray, False) Then
4                 Set ExpandRangeToIncludeEntireArrayFormulas = TheRange
5                 Exit Function
6             End If
7         End If

8         Set RangeToReturn = TheRange
9         Set RangeToSearch = Application.Intersect(TheRange, TheRange.Parent.UsedRange)

10        If RangeToSearch Is Nothing Then
11            Set ExpandRangeToIncludeEntireArrayFormulas = TheRange
12            Exit Function
13        Else
              'No point in searching the interior of the areas of RangeToSearch
14            Set RangeToSearch = EdgeCells(RangeToSearch)
15        End If

16        Set SPH = CreateSheetProtectionHandler(TheRange.Parent)

17        For Each c In RangeToSearch.Cells
18            Set CurrentArray = CurrentArray2(c, Exists) 'Cope with both old-style CSE array formulas and new-style dynamic arrays
19            If Exists Then
20                Set TempRange = Application.Intersect(CurrentArray, RangeToReturn)
21                If c.address = TempRange.Cells(1, 1).address Then
22                    If TempRange.Cells.CountLarge < CurrentArray.Cells.CountLarge Then
23                        Set RangeToReturn = Application.Union(RangeToReturn, CurrentArray)
24                    End If
25                End If
26            End If
27        Next c

28        Set ExpandRangeToIncludeEntireArrayFormulas = RangeToReturn

29        Exit Function
ErrHandler:
30        Throw "#ExpandRangeToIncludeEntireArrayFormulas (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CurrentArray2
' Author     : Philip Swannell
' Date       : 31-Jul-2019
' Purpose    : Returns either the old-style "CSE" R.CurrentArray or the new-style R.SpillParent.SpillingToRange as appropriate.
'              NB R is assumed to be one-cell range
' -----------------------------------------------------------------------------------------------------------------------
Private Function CurrentArray2(R As Range, retExists As Boolean) As Range

1         On Error GoTo ErrHandler
2         retExists = False
3         If R.HasArray Then
4             Set CurrentArray2 = R.CurrentArray
5             retExists = True
6         ElseIf ExcelSupportsSpill() Then
7             If Not R.SpillParent Is Nothing Then
8                 retExists = True
9                 Set CurrentArray2 = R.SpillParent.SpillingToRange
10            ElseIf IsError(R.Value) Then
11                If CStr(R.Value) = "Error 2045" Then '"#SPILL!"
12                    retExists = True
13                    Set CurrentArray2 = R
14                End If
15            End If
16        End If
17        Exit Function
ErrHandler:
18        Throw "#CurrentArray2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : EdgeCells
' Author     : Philip Swannell
' Date       : 31-Jul-2019
' Purpose    : Returns the "edge cells" of a range, i.e. those in top row, bottom row, left col and right col.
' -----------------------------------------------------------------------------------------------------------------------
Private Function EdgeCells(R As Range) As Range
          Dim a As Range, TheseEdges As Range
          Dim Ret As Range
1         On Error GoTo ErrHandler
2         For Each a In R.Areas
3             If a.Rows.Count > 2 And a.Columns.Count > 2 Then
4                 Set TheseEdges = Application.Union(a.Columns(1), a.Columns(a.Columns.Count), _
                      Range(a.Cells(1, 2), a.Cells(1, a.Columns.Count - 1)), _
                      Range(a.Cells(a.Rows.Count, 2), a.Cells(a.Rows.Count, a.Columns.Count - 1)))
5             Else
6                 Set TheseEdges = a
7             End If
8             If Ret Is Nothing Then
9                 Set Ret = TheseEdges
10            Else
11                Set Ret = Application.Union(Ret, TheseEdges)
12            End If
13        Next a

14        Set EdgeCells = Ret
15        Exit Function
ErrHandler:
16        Throw "#EdgeCells (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function InDeveloperMode() As Boolean
1         On Error GoTo ErrHandler
2         InDeveloperMode = GetSetting(gAddinName, "InstallInformation", "DeveloperMode", "Standard") = "Developer"
3         Exit Function
ErrHandler:
4         Throw "#InDeveloperMode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LocalWorkbookName
' Author     : Adapted from https://stackoverflow.com/questions/46346567/thisworkbook-fullname-returns-a-url-after-syncing-with-onedrive-i-want-the-file/67697487#67697487
' Date       : 22-Apr-2022
' Purpose    : Return the address on the local PC of a file that may be on OneDrive
' -----------------------------------------------------------------------------------------------------------------------
Function LocalWorkbookName(wb As Workbook) As String
          ' Set default return
1         On Error GoTo ErrHandler
          Static Memoize As Dictionary

2         LocalWorkbookName = wb.FullName
3         If InStr(1, LocalWorkbookName, "https://", vbTextCompare) = 0 Then
4             Exit Function
5         End If
          'Use a dictionary to memoize the result since execution takes 0.1 to 0.2 seconds
6         If Memoize Is Nothing Then Set Memoize = New Dictionary
7         If Memoize.Exists(LocalWorkbookName) Then
8             LocalWorkbookName = Memoize(LocalWorkbookName)
9             Exit Function
10        End If
          
          Const HKEY_CURRENT_USER = &H80000001

          Dim strValue As String
          
11        Dim objReg As Object: Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
12        Dim strRegPath As String: strRegPath = "Software\SyncEngines\Providers\OneDrive\"
          Dim arrSubKeys() As Variant
13        objReg.EnumKey HKEY_CURRENT_USER, strRegPath, arrSubKeys
          
          Dim varKey As Variant
14        For Each varKey In arrSubKeys
              ' check if this key has a value named "UrlNamespace", and save the value to strValue
15            objReg.getStringValue HKEY_CURRENT_USER, strRegPath & varKey, "UrlNamespace", strValue
          
              ' If the namespace is in FullName, then we know we have a URL and need to get the path on disk
16            If InStr(wb.FullName, strValue) > 0 Then
                  Dim strTemp As String
                  Dim strCID As String
                  Dim strMountpoint As String
                  
                  ' Get the mount point for OneDrive
17                objReg.getStringValue HKEY_CURRENT_USER, strRegPath & varKey, "MountPoint", strMountpoint
                  
                  ' Get the CID
18                objReg.getStringValue HKEY_CURRENT_USER, strRegPath & varKey, "CID", strCID
                  
                  ' Add a slash, if the CID returned something
19                If strCID <> vbNullString Then
20                    strCID = "/" & strCID
21                End If

                  ' strip off the namespace and CID
22                strTemp = Right(wb.FullName, Len(wb.FullName) - Len(strValue & strCID))
                  
                  ' replace all forward slashes with backslashes
23                LocalWorkbookName = strMountpoint & Replace(strTemp, "/", "\")

24                Memoize.Add wb.FullName, LocalWorkbookName

25                Exit Function
26            End If
27        Next

28        Exit Function
ErrHandler:
29        Throw "#LocalWorkbookName (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'For use by sFileSave and other functions
Sub CheckFileNameIsAbsolute(ByVal FileName As String)
1         FileName = Replace(FileName, "/", "\") 'flip unix to windows
2         If InStr(FileName, "\") = 0 Then
3             Throw "FileName must include a path"
4         End If
5         If Left(FileName, 2) = "\\" Then Exit Sub
6         If Mid(FileName, 2, 2) = ":\" Then Exit Sub
7         Throw "File '" + FileName + "' must be provided with an absolute path - i.e. start with 'X:\' for some drive letter X or start with '\\' for UNC paths"
End Sub

Function CoreEDate(start_date As Long, months As Long)
          Dim y As Long, M As Long, D As Long, upperbound As Long

1         On Error GoTo ErrHandler
2         y = Year(start_date)
3         M = Month(start_date)
4         D = day(start_date)

5         CoreEDate = DateSerial(y, M + months, D)
6         upperbound = DateSerial(y, M + months + 1, 1) - 1
7         If CoreEDate > upperbound Then CoreEDate = upperbound

8         Exit Function
ErrHandler:
9         Throw "#CoreEDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFileTranspose
' Author    : Philip Swannell
' Date      : 27-May-2019
' Purpose   : wrapped by sFileTranspose
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileTranspose(InputFile As String, OutputFile As String, Delimiter As String)
          Dim Contents As Variant

1         Contents = ThrowIfError(sFileShow(InputFile, Delimiter, False, False, False))
2         Contents = ThrowIfError(sArrayTranspose(Contents))
3         CoreFileTranspose = ThrowIfError(sFileSave(OutputFile, Contents, Delimiter))

4         On Error GoTo ErrHandler

5         Exit Function
ErrHandler:
6         CoreFileTranspose = "#CoreFileTranspose (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'It's a PITA that VBA does not have native Min or max functions, so add in this private module
Function Max(x, y)
1         If x > y Then
2             Max = x
3         Else
4             Max = y
5         End If
End Function

Function Min(x, y)
1         If x < y Then
2             Min = x
3         Else
4             Min = y
5         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CoreJoinPath
' Author     : Philip Swannell
' Date       : 07-Apr-2020
' Purpose    :
' Parameters :
'  PathsToJoin:
' -----------------------------------------------------------------------------------------------------------------------
Function CoreJoinPath(ParamArray PathsToJoin())
          Dim i As Long
          Dim LB As Long, UB As Long, ChooseVector
          Dim Result As String
1         On Error GoTo ErrHandler
2         LB = LBound(PathsToJoin)
3         UB = UBound(PathsToJoin)

4         Result = Replace(CStr(PathsToJoin(LB)), "/", "\") 'handle unix convention as well as windows
5         For i = LB + 1 To UB
              Dim ThisPart As String
6             ThisPart = Replace(PathsToJoin(i), "/", "\")
7             If Len(ThisPart) > 0 Then
8                 If Mid(ThisPart, 2, 2) = ":\" Or Left(ThisPart, 2) = "\\" Then
9                     Result = ThisPart
10                Else
11                    If Right(Result, 1) <> "\" And Len(Result) > 0 Then
12                        Result = Result + "\"
13                    End If
14                    If Left(ThisPart, 1) = "\" Then
15                        ThisPart = Mid(ThisPart, 2)
16                    End If
17                    Result = Result + ThisPart
18                End If
19            End If
20        Next i

          'Deal with \..\ elements, which mean "pop a level"
21        If InStr(Result, "\..\") > 0 Or Right(Result, 3) = "\.." Then
              Dim parts, FirstDD
22            parts = sTokeniseString(Result, "\")
TryAgain:
23            FirstDD = sMatch("..", parts)
24            If IsNumber(FirstDD) Then
25                If FirstDD > 1 Then
26                    If sNRows(parts) = 2 Then
27                        Result = ""
28                    Else
29                        ChooseVector = sReshape(True, sNRows(parts), 1)
30                        ChooseVector(FirstDD, 1) = False
31                        ChooseVector(FirstDD - 1, 1) = False
32                        parts = sMChoose(parts, ChooseVector)
33                        Result = sConcatenateStrings(parts, "\")
34                        GoTo TryAgain
35                    End If
36                End If
37            End If
38        End If

39        CoreJoinPath = Result

40        Exit Function
ErrHandler:
41        Throw "#CoreJoinPath (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function CoreRelativePath(ByVal FullPath As String, ByVal BasePath As String)
          Dim i As Long
          Dim N As Long
          Dim NumCommon As Long
          Dim NumPartsA As Long
          Dim NumPartsB As Long
          Dim origRelativeTo As String
          Dim PartsA As Variant
          Dim PartsB As Variant
          Dim Result As String

1         On Error GoTo ErrHandler

2         origRelativeTo = BasePath
3         If Right(BasePath, 1) = "\" Then
4             BasePath = Left(BasePath, Len(BasePath) - 1)
5         End If

6         PartsA = VBA.Split(FullPath, "\")
7         PartsB = VBA.Split(BasePath, "\")

8         NumPartsA = UBound(PartsA) - LBound(PartsA) + 1
9         NumPartsB = UBound(PartsB) - LBound(PartsB) + 1

10        N = IIf(NumPartsA < NumPartsB, NumPartsA, NumPartsB)

11        For i = 1 To N
12            If LCase(PartsA(i - 1)) = LCase(PartsB(i - 1)) Then
13                NumCommon = NumCommon + 1
14            Else
15                Exit For
16            End If
17        Next

18        If NumCommon > 1 Then
19            For i = 1 To NumPartsB - NumCommon
20                Result = Result & "..\"
21            Next i
22            For i = NumCommon + 1 To NumPartsA
23                If Right(Result, 1) = "\" Then
24                    Result = Result & PartsA(i - 1)
25                ElseIf Result = "" Then
26                    Result = PartsA(i - 1)
27                Else
28                    Result = Result & "\" & PartsA(i - 1)
29                End If
30            Next i
              'For safety, check that the logic has worked
31            If LCase(CoreJoinPath(origRelativeTo, Result)) = LCase(FullPath) Then
32                CoreRelativePath = Result
33            Else
34                CoreRelativePath = FullPath
35            End If
36        Else
37            CoreRelativePath = FullPath
38        End If

39        Exit Function
ErrHandler:
40        Throw "#CoreRelativePath (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ArrayToRLiteral
' Author    : Philip Swannell
' Date      : 12-Nov-2017
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function ArrayToRLiteral(Optional Data, Optional MissingBecomes As String)
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim NC As Long
          Dim NR As Long
          Dim Strings() As String

1         On Error GoTo ErrHandler

2         If IsMissing(Data) Or IsEmpty(Data) Then
3             ArrayToRLiteral = MissingBecomes
4             Exit Function
5         End If

6         Force2DArrayR Data, NR, NC

7         ReDim Strings(1 To NR * NC)

8         k = 1
9         For i = 1 To NR
10            For j = 1 To NC
11                Select Case VarType(Data(i, j))
                      Case vbString
12                        Strings(k) = RLiteralFromString(CStr(Data(i, j)))
13                    Case vbBoolean
14                        Strings(k) = UCase$(Data(i, j))
15                    Case Else
16                        Strings(k) = CStr(Data(i, j))
17                End Select
18                k = k + 1
19            Next j
20        Next i

21        If NR * NC = 1 Then
22            ArrayToRLiteral = Strings(1)
23        Else
24            ArrayToRLiteral = "c(" + VBA.Join(Strings, ", ") + ")"
25        End If

26        Exit Function
ErrHandler:
27        Throw "#ArrayToRLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

