Attribute VB_Name = "modR"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sExecuteRCode
' Author    : Philip
' Date      : 20-Sep-2017
' Purpose   : Wrap all calls to the mechanism we are using for for R <-> Excel communication.
'             Whatever method we use - Paul's C# code, BERT, or RExcel then this method
'             must ensure returns are either singletons (Double, String Boolean etc) or
'             1-based two dimensional arrays where R-vectors become 2-d arrays with a single column.
'             This is the behaviour of BERT.
' -----------------------------------------------------------------------------------------------------------------------
Function sExecuteRCode(ByVal RCode As String, Optional ErrorLevel As Long = 2)
          '2 = Trap errors and warnings: if an error or a warning happens return from function will be an error string #the message of the error or warning!
          '1 = Trap errors: warnings ignored, if an error happens return from function will be an error string #the message of the error!
          '0 = trap neither errors nor warnings, warnings ignored, if an error happens the return will be #NULL!

1         On Error GoTo ErrHandler
          Dim Result As Variant

2         Select Case ErrorLevel
              Case 1, 2
                  Const TrapErrors = "TrapErrors<-function(x){" & vbLf & _
                      "  out <- tryCatch({x" & vbLf & _
                      "  }, error = function(e) {" & vbLf & _
                      "    paste0(""#"",e$message,""!"")" & vbLf & _
                      "  }," & vbLf & _
                      "  warning = function(e) {" & vbLf & _
                      "    paste0(""#"",e$message,""!"")}" & vbLf & _
                      "  )" & vbLf & _
                      "return(out)" & vbLf & _
                      "}"
3                 If Not Application.Run("BERT.Exec", "exists(""TrapErrors"")") Then
4                     Application.Run "BERT.Exec", TrapErrors
5                 End If
6                 If ErrorLevel = 2 Then
7                     RCode = "TrapErrors({" + RCode + "})"    'Use of curly braces means that multiple line R expressions work (delimited by semi colons)
8                 Else
9                     RCode = "TrapErrors(suppressWarnings({" + RCode + "}))"
10                End If
11            Case 0
                  'nothing to do
12            Case Else
13                Throw "Unrecognised value for ErrorLevel. Must be 0,1 or 2"
14        End Select
15        Result = Application.Run("BERT.Exec", RCode)
16        sExecuteRCode = Result

17        Exit Function
ErrHandler:
18        sExecuteRCode = "#sExecuteRCode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub TestGetDataframe()
1         g GetDataframe("data.frame(x=1:5,y=letters[1:5])", False, False)
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetDataframe
' Author    : Philip
' Date      : 05-Oct-2017
' Purpose   : Gets the contents of an R dataframe, with control over whether row names and column names are returned.
'             Note that Excel will crash if BERT 1.62.1 is used. Duncan Werner fixed that with BERT 1.62.2
' -----------------------------------------------------------------------------------------------------------------------
Function GetDataframe(ByVal Expression As String, Optional ColNamesAtTop As Boolean, Optional RowNamesOnLeft As Boolean)
          Dim RawData

1         On Error GoTo ErrHandler
2         RawData = ThrowIfError(sExecuteRCode(Expression))
3         If ColNamesAtTop And RowNamesOnLeft Then
4             GetDataframe = RawData
5         ElseIf Not ColNamesAtTop And RowNamesOnLeft Then
6             GetDataframe = sSubArray(RawData, 2, 1)
7         ElseIf ColNamesAtTop And Not RowNamesOnLeft Then
8             GetDataframe = sSubArray(RawData, 1, 2)
9         ElseIf Not ColNamesAtTop And Not RowNamesOnLeft Then
10            GetDataframe = sSubArray(RawData, 2, 2)
11        End If

12        Exit Function
ErrHandler:
13        GetDataframe = "#GetDataframe (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sNearestCorrelationMatrix
' Author    : Philip Swannell
' Date      : 08-Jun-2017
' Purpose   : Wrapper to R function nearPD, but for correlation matrices. Requires package
'             Matrix to have been loaded in the R instance. See https://stat.ethz.ch/R-manual/R-devel/library/Matrix/html/nearPD.html
' -----------------------------------------------------------------------------------------------------------------------
Function sNearestCorrelationMatrix(Matrix, Optional EigenTolerance As Double = 0.000001, Optional Method As String = "Higham")
          Dim a As Variant
          Dim Expression As String
          Dim i As Long
          Dim j As Long
          Dim p As Long

1         On Error GoTo ErrHandler
2         Force2DArrayR Matrix

3         a = Matrix

4         If LBound(a, 1) <> 1 Or LBound(a, 2) <> 1 Then Throw "Matrix must have lower bounds of 1"
5         If UBound(a, 1) <> UBound(a, 2) Then Throw "Matrix must be square with the same number of rows as columns"
6         p = UBound(a, 1)

7         For i = 1 To p
8             For j = 1 To i
9                 If Not IsNumber(a(i, j)) Then Throw "All elements of Matrix must be numbers but element " + CStr(i) + "," + CStr(j) + " is not"
10                If Not IsNumber(a(j, i)) Then Throw "All elements of Matrix must be numbers but element " + CStr(j) + "," + CStr(i) + " is not"
11                If a(j, i) <> a(i, j) Then Throw "Matrix must be a symmetric but isn't since element " + CStr(i) + "," + CStr(j) + " is not a equal to element " + CStr(j) + "," + CStr(i)
12                If Abs(a(i, j)) > 1 Then Throw "All elements of Matrix must be in the range -1 to 1 but element " + CStr(i) + "," + CStr(j) + " is not"
13            Next j
14        Next i

15        If LCase$(Method) = "higham" Then
              'Call R method nearPD that implements Nick Higham's algorithm
16            SaveArrayToR a, "TempSNCM", 2, False, False
17            Expression = "as.matrix(Matrix::nearPD(TempSNCM, corr = TRUE,eig.tol = " + CStr(EigenTolerance) + ")$mat)"
18            sNearestCorrelationMatrix = ThrowIfError(sExecuteRCode(Expression))
19        ElseIf LCase$(Method) = "minimumeigenvalue" Then
              'Calculate EigenValues, floor them, construct the positive definite matrix with floored eigenvalues and same eigen vectors, renormalise
              Dim EigenValues
              Dim EigenVectors
              Dim FlooredEigenValues
              Dim Matrix2
              Dim Matrix3
              Dim MaxEigenValue As Double
              Dim MinEigenValue As Double
              Dim Normalise
              Dim Res
20            Res = ThrowIfError(sEigen(Matrix))
21            EigenValues = sSubArray(Res, 1, 1, , 1)
22            EigenVectors = sSubArray(Res, 1, 2)
23            MaxEigenValue = sColumnMax(EigenValues)(1, 1)
24            MinEigenValue = sColumnMin(EigenValues)(1, 1)
25            If MinEigenValue > 0 Then
26                If MinEigenValue >= MaxEigenValue * EigenTolerance Then
27                    sNearestCorrelationMatrix = Matrix
28                    Exit Function
29                End If
30            End If
31            FlooredEigenValues = sArrayMax(EigenValues, Abs(MaxEigenValue) * EigenTolerance)
32            Matrix2 = Application.WorksheetFunction.MMult(EigenVectors, sDiagonal(FlooredEigenValues))
33            Matrix2 = Application.WorksheetFunction.MMult(Matrix2, Application.WorksheetFunction.MInverse(EigenVectors))

34            Normalise = sDiagonal(Matrix2)
35            Normalise = sArrayMultiply(Normalise, sArrayTranspose(Normalise))
36            Normalise = sArrayPower(Normalise, -0.5)

37            Matrix3 = sArrayMultiply(Matrix2, Normalise)
              'may not be quite symmetric thanks to numerical errors, so fix up here...
              Dim x As Double
38            For i = 2 To p
39                For j = 1 To i - 1
40                    x = (Matrix3(i, j) + Matrix3(j, i)) / 2
41                    Matrix3(i, j) = x
42                    Matrix3(j, i) = x
43                Next j
44            Next i
45            For i = 1 To p
46                Matrix3(i, i) = 1    ' was seeing very small differences from 1 eg 1e-16
47            Next
48            sNearestCorrelationMatrix = Matrix3
49        Else
50            Throw "Method must be 'Higham' or 'MinimumEigenValue'"
51        End If

52        Exit Function
ErrHandler:
53        sNearestCorrelationMatrix = "#sNearestCorrelationMatrix (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveDataframe
' Author    : Hermione Glyn
' Date      : 27-Sep-2016
' Purpose   : Replacement for PB's C# function sfns.scSaveDataFrameWithTypeNameHeader
'             since the latter has sub-linear performance in the number of rows in the data passed in.
'             First row of the data gives "type" information for the column, second row
'             gives the names of the columns when created as a dataframe in R. Subsequent rows are the data.
' -----------------------------------------------------------------------------------------------------------------------
Function SaveDataframe(TheData, NameInR As String)
          Dim DateExpression As String
          Dim Expression1 As String
          Dim Expression2 As String
          Dim FileName As String
          Dim HeaderName As String
          Dim i As Long
          Dim Res As String
          Dim ResFromR

1         On Error GoTo ErrHandler

2         If Not ValidRVariableName(NameInR) Then Throw "'" + NameInR + "' is not valid as a name in R"

3         For i = 1 To sNCols(TheData)
4             HeaderName = TheData(2, i)
5             Select Case TheData(1, i)
                  Case "CHAR"
6                     Res = """character"""
7                 Case "BOOL"
8                     Res = """logical"""
9                 Case "DOUBLE", "INT", "DATESTR"
10                    Res = """numeric"""
11                Case Else
12                    Throw "Unrecognised header in first row of trade data: Column types can be BOOL, CHAR, INT, DOUBLE or DATESTR"
13            End Select
14            If i = 1 Then
15                Expression1 = Res
16            Else
17                Expression1 = Expression1 & "," & Res
18            End If
19        Next i

20        FileName = Environ$("Temp") & "\" & Replace(CStr(sElapsedTime()), ".", vbNullString)
21        ThrowIfError sFileSaveCSV(FileName, TheData)

22        Expression2 = NameInR & " <- read.csv(""" & Replace(FileName, "\", "/") & """, stringsAsFactors = FALSE, skip = 1, na.strings =" & """Empty""" & ", colClasses = c(" & Expression1 & "))"
          'Line below is to work-around a crash-bug in BERT 1.62.1 that is not present in BERT 1.61
          'Excel crashes if when BERT is returning a dataframe (perhaps only if the dataframe contains character columns?)
23        Expression2 = "(" + Expression2 + ")[1,1]"
24        ResFromR = sExecuteRCode(Expression2)
25        ThrowIfError ResFromR
26        For i = 1 To sNCols(TheData)        'We pass dates as numbers then have to flip to R-style dates inside R!
27            If TheData(1, i) = "DATESTR" Then
                  Dim CheckedFunction As Boolean
                  Dim ColName As String
                  Dim Expression3 As String
28                If Not CheckedFunction Then
29                    If Not sExecuteRCode("exists(""exceldatetodate"")") Then
30                        Expression3 = "source(""" + gRSourcePath + "UtilsDate.R"", encoding = ""Windows-1252"")"
31                        ThrowIfError sExecuteRCode(Expression3)
32                        CheckedFunction = True
33                    End If
34                End If
35                ColName = TheData(2, i)
36                DateExpression = NameInR & "$" + ColName + " <- as.Date(exceldatetodate(" & NameInR & "$" + ColName + "), origin = " & """1970-01-01""" & ")"
37                ThrowIfError (sExecuteRCode(DateExpression))
38            End If
39        Next

40        ThrowIfError sFileDelete(FileName)
41        SaveDataframe = NameInR

42        Exit Function
ErrHandler:
43        Dim TheErr As String: TheErr = "#SaveDataframe (line " & CStr(Erl) + "): " & Err.Description & "!"
44        sFileDelete FileName
45        If TypeName(Application.Caller) = "Range" Then SaveDataframe = TheErr Else Throw TheErr
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RLiteralFromString
' Author    : Philip Swannell
' Date      : 03-May-2017
' Purpose   : Take care when passing strings to R which contain double quote or backslash characters
' -----------------------------------------------------------------------------------------------------------------------
Function RLiteralFromString(InputString As String)
          Const DQ = """"
1         RLiteralFromString = DQ + Replace(Replace(InputString, "\", "\\"), DQ, "\" + DQ) + DQ
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RLiteralsFromStrings
' Author    : Philip Swannell
' Date      : 03-May-2017
' Purpose   : Multi-call version of RLiteralFromString
' -----------------------------------------------------------------------------------------------------------------------
Function RLiteralsFromStrings(ByVal InputArray)
          Dim i As Long
          Dim j As Long
          Dim Res() As String

1         On Error GoTo ErrHandler
2         If TypeName(InputArray) = "Range" Then InputArray = InputArray.Value
3         Select Case NumDimensions(InputArray)

              Case 0
4                 RLiteralsFromStrings = RLiteralFromString(CStr(InputArray))
5             Case 1
6                 ReDim Res(LBound(InputArray) To UBound(InputArray))
7                 For i = LBound(InputArray) To UBound(InputArray)
8                     Res(i) = RLiteralFromString(CStr(InputArray(i)))
9                 Next i
10                RLiteralsFromStrings = Res
11            Case 2
12                ReDim Res(LBound(InputArray, 1) To UBound(InputArray, 1), LBound(InputArray, 2) To UBound(InputArray, 2))
13                For i = LBound(InputArray, 1) To UBound(InputArray, 1)
14                    For j = LBound(InputArray, 2) To UBound(InputArray, 2)
15                        Res(i, j) = RLiteralFromString(CStr(InputArray(i, j)))
16                    Next j
17                Next i
18                RLiteralsFromStrings = Res
19        End Select
20        Exit Function
ErrHandler:
21        Throw "#RLiteralsFromStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Sub TestSaveArrayToR()
          Dim Data
1         Data = sArraySquare(100, 200, 300, 4001)

2         Debug.Print SaveArrayToR(Data, "TestVar", 2, False, False, False)

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveArrayToR
' Author     : Philip Swannell
' Date       : 11-Dec-2017
' Purpose    : Uses BERT.Call to save an array to R. If the data to be saved is of mixed type then the array in R will be an
'              array of single-element lists, unless Argument SaveNumbersOnly is TRUE in which case non-numbers become NA within
'              a numeric array
' Parameters :
'  Data             : 2D array of arbitrary data
'  NameInR          :
'  NumDimsInR       :
'  TopRowIsColNames : Boolean
'  LeftColIsRowNames: Boolean
'  SaveNumbersOnly  :
' PGS 23-March-2019
' Found a change in behaviour of BERT.Call when using this function as part of ISDA work. Must have worked approx 1 year ago. Missing arguments to the call to BERT.Call
'appear in R as Boolean FALSE, rather than defaulting according to the R function's declaration. Fix is (in this case) is to pass CVErr(xlErrNA) which gets converted to the R's NA
' -----------------------------------------------------------------------------------------------------------------------
Function SaveArrayToR(ByVal Data As Variant, NameInR As String, NumDimsInR As Long, Optional TopRowIsColNames As Boolean, Optional LeftColIsRowNames As Boolean, Optional SaveNumbersOnly As Boolean)

1         On Error GoTo ErrHandler
          Static HaveChecked As Boolean
2         Force2DArrayR Data

3         If Not HaveChecked Then
4             CheckR "SaveArrayToR", gPackagesSAI, gRSourcePath + "UtilsPGS.R", "SourceAllFiles()"
5             HaveChecked = True
6         End If

7         If Not ValidRVariableName(NameInR) Then Throw "'" + NameInR + "' is not valid as a name in R"
8         If NumDimsInR <> 1 And NumDimsInR <> 2 Then Throw "NumDimsInR must be 1 or 2"
          Dim LeftColRowNames As Variant
          Dim TopRowColNames As Variant

9         If TopRowIsColNames And LeftColIsRowNames Then
10            If NumDimsInR = 1 Then Throw "NumDimsInR must be 2 if both TopRowIsColNames and LeftColIsRowNames are TRUE"
11            TopRowColNames = sSubArray(Data, 1, 2, 1)
12            LeftColRowNames = sSubArray(Data, 2, 1, , 1)
13            Data = sSubArray(Data, 2, 2)
14            SaveArrayToR = Application.Run("BERT.Call", "SaveArrayToR", Data, NameInR, NumDimsInR, TopRowColNames, LeftColRowNames, SaveNumbersOnly)
15        ElseIf TopRowIsColNames Then
16            If NumDimsInR = 1 Then If sNRows(Data) <> 2 Then Throw "Data must have two rows when NumDimsInR is 1 and TopRowIsColNames is TRUE"
17            TopRowColNames = sSubArray(Data, 1, 1, 1)
18            Data = sSubArray(Data, 2)
19            SaveArrayToR = Application.Run("BERT.Call", "SaveArrayToR", Data, NameInR, NumDimsInR, TopRowColNames, , SaveNumbersOnly)
20        ElseIf LeftColIsRowNames Then
21            If NumDimsInR = 1 Then If sNCols(Data) <> 2 Then Throw "Data must have two columns when NumDimsInR is 1 and LeftColIsRowNames is TRUE"
22            LeftColRowNames = sSubArray(Data, 1, 1, , 1)
23            Data = sSubArray(Data, 1, 2)
24            SaveArrayToR = Application.Run("BERT.Call", "SaveArrayToR", Data, NameInR, NumDimsInR, CVErr(xlErrNA), LeftColRowNames, SaveNumbersOnly)
25        Else
26            SaveArrayToR = Application.Run("BERT.Call", "SaveArrayToR", Data, NameInR, NumDimsInR, CVErr(xlErrNA), CVErr(xlErrNA), SaveNumbersOnly)
27        End If

28        Exit Function
ErrHandler:
29        SaveArrayToR = "#SaveArrayToR (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveSingletonToR
' Author    : Philip Swannell
' Date      : 20-Apr-2016
' Purpose   : Save a singleton down to R. Could alternatively use sFns.Providers.Default.SetPropertyDouble and similar
' -----------------------------------------------------------------------------------------------------------------------
Function SaveSingletonToR(NameInR As String, Value As Variant)
          Dim Expression As String

1         On Error GoTo ErrHandler
2         If Not ValidRVariableName(NameInR) Then Throw "'" + NameInR + "' is not valid as a name in R"
3         Select Case VarType(Value)

              Case vbDouble, vbInteger, vbSingle, vbLong
4                 Expression = NameInR & " <- " & CStr(Value)
5             Case vbString
6                 Expression = NameInR & " <- " & RLiteralFromString(CStr(Value))
7             Case vbBoolean
8                 Expression = NameInR & " <- " & UCase$(CStr(Value))
9             Case vbDate
10                Expression = NameInR & " <- as.Date(""" + Format$(Value, "yyyy-mm-dd") & """,""%Y-%m-%d"")"
11            Case Else
12                Throw "Unsupported Type of argument Value"
13        End Select

14        ThrowIfError sExecuteRCode(Expression)

15        SaveSingletonToR = NameInR
16        Exit Function
ErrHandler:
17        Dim TheErr As String: TheErr = "#SaveSingletonToR (line " & CStr(Erl) + "): " & Err.Description & "!"
18        If TypeName(Application.Caller) = "Range" Then SaveSingletonToR = TheErr Else Throw TheErr
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ValidRVariableName
' Author    : Philip Swannell
' Date      : 11-May-2017
' Purpose   : Codes R's rules for valid variable names - see https://cran.r-project.org/doc/manuals/R-intro.html#R-commands_003b-case-sensitivity-etc
' -----------------------------------------------------------------------------------------------------------------------
Function ValidRVariableName(NameInR As String) As Boolean
          Static rx As VBScript_RegExp_55.RegExp

1         On Error GoTo ErrHandler
2         If rx Is Nothing Then
3             Set rx = New RegExp
4             With rx
5                 .IgnoreCase = False
6                 .Pattern = "^[a-zA-Z0-9_\.]*$"
7                 .Global = False        'Find first match only
8             End With
9         End If

10        ValidRVariableName = True
11        If Not rx.Test(NameInR) Then
12            ValidRVariableName = False
13        Else
14            If Left$(NameInR, 1) = "." Then
15                Select Case Asc(Mid$(NameInR, 2, 1))
                      Case 48 To 57    '"0"-"9"
16                        ValidRVariableName = False
17                End Select
18            End If
19        End If
20        Exit Function
ErrHandler:
21        Throw "#ValidRVariableName (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveLabelledDataToRList
' Author    : Philip Swannell
' Date      : 11-May-2017
' Purpose   : Simple way to save two column label-value pairs as a list in R
' -----------------------------------------------------------------------------------------------------------------------
Function SaveLabelledDataToRList(ByVal LabelledData As Variant, NameInR As String)
          Dim Expression As String
          Dim i As Long
          Dim ThisExpression As String

          Dim NR As Long
          Dim Strings() As String

1         On Error GoTo ErrHandler
2         Force2DArrayR LabelledData
3         NR = sNRows(LabelledData)
4         If Not ValidRVariableName(NameInR) Then Throw "'" + NameInR + "' is not valid as a name in R"
5         If sNCols(LabelledData) <> 2 Then Throw "LabelledData must have two columns"

6         ReDim Strings(1 To NR)

7         For i = 1 To NR
8             If VarType(LabelledData(i, 1)) <> vbString Then Throw "Left column of LabelledData must contain strings, but element at row " + CStr(i) + " does not"
9             If Not ValidRVariableName(CStr(LabelledData(i, 1))) Then Throw "Invalid label '" + LabelledData(i, 1) + "' in LabelledData at row " + CStr(i)

10            ThisExpression = LabelledData(i, 1) & " = "

11            Select Case VarType(LabelledData(i, 2))
                  Case vbDouble, vbInteger, vbSingle, vbLong
12                    ThisExpression = ThisExpression + CStr(LabelledData(i, 2))
13                Case vbBoolean
14                    ThisExpression = ThisExpression + UCase$(CStr(LabelledData(i, 2)))
15                Case vbString
16                    ThisExpression = ThisExpression + RLiteralFromString(CStr(LabelledData(i, 2)))
17                Case Else
18                    Throw "Value of unrecognized type in right column of LabelledData at row " + CStr(i)
19            End Select
20            Strings(i) = ThisExpression
21        Next i

22        Expression = NameInR + " <- list(" + VBA.Join(Strings, ", ") + ")"

23        ThrowIfError sExecuteRCode(Expression)

24        SaveLabelledDataToRList = "OK"

25        Exit Function
ErrHandler:
26        SaveLabelledDataToRList = "#SaveLabelledDataToRList (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : To1Based2D
' Author    : Philip Swannell
' Date      : 09-Sep-2016
' Purpose   : Returns from R are zero-based, but functions in SolumAddin are not well
'             tested against zero-based arrays (e.g. sArrayDivide fails). So convert to 1-based 2 dimensional array
'             If InputArray is not an array output is 2d array whose element(1,1) is InputArray
'             If InputArray is 1-d then output is 1-based 2-dimensional COLUMN array
'             If inputArray is 2-d then output is transformed (if necessary) to be 1-based in both dimensions
' -----------------------------------------------------------------------------------------------------------------------
Function To1Based2D(InputArray)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim ND As Long
          Dim NR As Long
          Dim OffsetC As Long
          Dim OffsetR As Long
          Dim Res() As Variant

1         On Error GoTo ErrHandler

2         ND = NumDimensions(InputArray)
3         Select Case ND

              Case 0
4                 ReDim Res(1 To 1, 1 To 1)
5                 Res(1, 1) = InputArray
6             Case 1
7                 NR = UBound(InputArray) - LBound(InputArray) + 1
8                 OffsetR = 1 - LBound(InputArray)
9                 ReDim Res(1 To NR, 1 To 1)
10                For i = 1 To NR
11                    Res(i, 1) = InputArray(i - OffsetR)
12                Next i
13            Case 2
14                If LBound(InputArray, 1) = 1 Then
15                    If LBound(InputArray, 2) = 1 Then
16                        To1Based2D = InputArray
17                        Exit Function
18                    End If
19                End If
20                NR = UBound(InputArray, 1) - LBound(InputArray, 1) + 1
21                NC = UBound(InputArray, 2) - LBound(InputArray, 2) + 1
22                OffsetR = 1 - LBound(InputArray, 1)
23                OffsetC = 1 - LBound(InputArray, 2)
24                If OffsetR = 0 And OffsetC = 0 Then
25                    To1Based2D = InputArray
26                    Exit Function
27                End If
28                ReDim Res(1 To NR, 1 To NC)
29                For i = 1 To NR
30                    For j = 1 To NC
31                        Res(i, j) = InputArray(i - OffsetR, j - OffsetC)
32                    Next j
33                Next i
34            Case Else
35                Throw "Cannot convert array with " + CStr(ND) + " dimensions to 2 dimensions"
36        End Select
37        To1Based2D = Res

38        Exit Function
ErrHandler:
39        Throw "#To1Based2D (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CheckR
' Author    : Philip Swannell
' Date      : 12-Nov-2017
' Purpose   : Routine to call once, the first time a VBA wrapper to an R function is called
'    FunctionName - the name of the VBAWrapper, used only in error construction
'    Packages - comma delimited string listing the necessary packages used by the underlying R function
'    SourceFile - the location of the file containing the R code
' -----------------------------------------------------------------------------------------------------------------------
Function CheckR(Optional FunctionName As String, Optional Packages As String, Optional SourceFile As String, Optional ExtraRCodeToExecute As String, Optional ThrowErrors As Boolean = True)
          Dim aPackages As Variant
          Dim ChooseVector As Variant
          Dim CopyOfErr As String
          Dim Expression As String
          Dim i As Long
          Dim MissingPackages
          Dim Prompt As String
          Dim Res
          Dim Result As Variant
          Dim ThisPackage As String
          Const Title = "Check R Installation"

1         On Error Resume Next
2         Res = Application.Run("BERT.Exec", "1+1")
3         On Error GoTo ErrHandler
4         If Not sEquals(Res, 2) Then
5             Throw "BERT (Basic Excel R Toolkit) is not installed, or has been disabled."
6         End If
7         If Len(Packages) > 0 Then
8             aPackages = sTokeniseString(Packages)
9             ChooseVector = sReshape(True, sNRows(aPackages), 1)
10            For i = 1 To sNRows(aPackages)
11                ThisPackage = aPackages(i, 1)
12                Expression = "suppressWarnings(library(" + ThisPackage + "))"
13                Result = sExecuteRCode(Expression, 1)    'Note suppressing warnings here
14                ChooseVector(i, 1) = sIsErrorString(Result)
15            Next i
16            If sColumnOr(ChooseVector)(1, 1) Then
17                MissingPackages = sMChoose(aPackages, ChooseVector)
18                Expression = "install.packages(" + ArrayToRLiteral(MissingPackages) + ", repos = ""http://cran.rstudio.com/"", method = ""wininet"")"
19                Prompt = "The R environment requires the packages " + sConcatenateStrings(MissingPackages, ", ") + " to be installed." + vbLf + vbLf + "Click OK to install"
20                If FunctionName <> vbNullString Then
21                    Prompt = "Failure in function " + FunctionName + ". " & Prompt
22                End If
23                MsgBoxPlus Prompt, vbOKOnly + vbExclamation, Title, , , , , , , , 120, vbOK
24                sExecuteRCode Expression
25                Throw "Packages have been installed"
26            End If
27        End If
28        If Len(SourceFile) > 1 Then
29            Expression = "suppressWarnings(source(""" & Replace(SourceFile, "\", "/") + """, encoding = ""Windows-1252""))"
30            Result = sExecuteRCode(Expression)
31            If sIsErrorString(Result) Then
32                Prompt = "Error when sourcing R file '" + SourceFile + "': " + Result
33                If TypeName(Application.Caller) <> "Range" Then
34                    MsgBoxPlus Prompt, , Title, , , , , , , , 120, vbOK
35                End If
36                Throw Result
37            End If
38        End If
39        If ExtraRCodeToExecute <> vbNullString Then
40            Result = sExecuteRCode(ExtraRCodeToExecute)
41            If sIsErrorString(Result) Then
42                Prompt = "Error when executing R command: '" + ExtraRCodeToExecute + "':" + vbLf + Result
43                If TypeName(Application.Caller) <> "Range" Then
44                    MsgBoxPlus Prompt, , Title, , , , , , , , 120, vbOK
45                End If
46                Throw Result
47            End If
48        End If
49        CheckR = "OK"
50        Exit Function
ErrHandler:
51        CopyOfErr = "#CheckR (line " & CStr(Erl) + "): " & Err.Description & "!"
52        If ThrowErrors Then
53            Throw CopyOfErr
54        Else
55            CheckR = CopyOfErr
56        End If
End Function
