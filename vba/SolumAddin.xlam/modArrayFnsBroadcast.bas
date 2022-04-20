Attribute VB_Name = "modArrayFnsBroadcast"
Option Explicit

'To put these definitions into alphabetical order see helper workbook https://d.docs.live.net/4251b448d4115355/Excel Sheets/FuncIDDefinitions.xlsm
Enum BroadcastFuncID
    FuncIdAdd = 1
    FuncIdAnd = 2
    FuncIdBarrierOption = 3
    FuncIdBlackScholes = 4
    FuncIdConcatenate = 5
    FuncIdCreateFolder = 6
    FuncIdDeleteFolder = 7
    FuncIdDivide = 8
    FuncIDEDate = 9
    FuncIdEquals = 10
    FuncIdFileCopy = 11
    FuncIdFileCopySkip = 12
    FuncIdFileDelete = 13
    FuncIdFileExif = 14
    FuncIdFileExists = 15
    FuncIdFileInfo = 16
    FuncIdFileIsUnicode = 17
    FuncIdFileLastModifiedDate = 18
    FuncIdFileMove = 19
    FuncIdFileNumLines = 20
    FuncIdFileRegExReplace = 21
    FuncIdFileRename = 22
    FuncIDFileTranspose = 23
    FuncIdFileUnblock = 24
    FuncIdFolderCopy = 25
    FuncIdFolderExists = 26
    FuncIdFolderIsWritable = 27
    FuncIdFolderMove = 28
    FuncIdFolderRename = 29
    FuncIdIf = 30
    FuncIdIfErrorString = 31
    FuncIdIndex = 32
    FuncIDISDASIMMMakeID = 33
    FuncIdLeft = 34
    FuncIdLessThan = 35
    FuncIdlessThanOrEqual = 36
    FuncIdLike = 37
    FuncIdMax = 38
    FuncIdMean = 39
    FuncIdMin = 40
    FuncIdMultiply = 41
    FuncIdNearlyEquals = 42
    FuncIdNormOpt = 43
    FuncIdOptSolveVol = 44
    FuncIdor = 45
    FuncIdPower = 46
    FuncIdRight = 47
    FuncIDRound = 48
    FuncIdRoundSF = 49
    FuncIdStringBetweenStrings = 50
    FuncIdSubtract = 51
    FuncIdURLDownloadToFile = 52
    FuncIdISDASIMMApplyRounding2022 = 53
    FuncIdRelativePath = 54
End Enum

Public Const cEPSILON As Double = 0.000000000000001        'Used by sNearlyEquals, sArraysNearlyEqual, sArraysNearlyIdentical

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sNearlyEquals
' Author    : Philip Swannell
' Date      : 16-May-2015
' Purpose   : Test if two values are "nearly equal". If A and B are both numbers, then they are nearly
'             equal if and only if:
'             Abs(A-B) < Epsilon * Max(1,Abs(A),Abs(B))
'             If A and B are of different type (e.g. Number versus String) then the
'             function returns FALSE.
' Arguments
' A         : Any value (but not an array of values).
' B         : Any value (but not an array of values).
' CaseSensitive: TRUE for case sensitive comparison of strings. FALSE or omitted for case insensitive
'             comparison.
' Epsilon   : Epsilon determines the tolerance for comparison of two number via the formula:
'             A Nearly Equals B iff Abs(A-B) < Epsilon * Max(1,Abs(A),Abs(B))
'             If omitted Epsilon defaults to 0.000000000000001 (i.e. 10^-15)
'
' -----------------------------------------------------------------------------------------------------------------------
Function sNearlyEquals(ByVal a, ByVal b, Optional CaseSensitive As Boolean = False, Optional Epsilon As Double = cEPSILON)
Attribute sNearlyEquals.VB_Description = "Test if two values are ""nearly equal"". If A and B are both numbers, then they are nearly equal if and only if:\nAbs(A-B) < Epsilon * Max(1,Abs(A),Abs(B))\nIf A and B are of different type (e.g. Number versus String) then the function returns FALSE."
Attribute sNearlyEquals.VB_ProcData.VB_Invoke_Func = " \n27"
          'Const Epsilon As Double = 0.000000000000001
          Dim CompareTo As Double
          Dim VTA As Long
          Dim VTB As Long

1         On Error GoTo ErrHandler

2         VTA = VarType(a)
3         VTB = VarType(b)
4         If VTA >= vbArray Or VTB >= vbArray Then
5             sNearlyEquals = "#sNearlyEquals: Function does not handle arrays. Use sArrayNearlyEquals or sArraysNearlyIdentical instead!"
6             Exit Function
7         End If

          'Both numbers...
8         If IsNumber(a) Then
9             If IsNumber(b) Then
10                If a = b Then
11                    sNearlyEquals = True
12                    Exit Function
13                End If
14                a = CDbl(a): b = CDbl(b)
15                CompareTo = Abs(a)
16                If Abs(b) > Abs(a) Then
17                    CompareTo = Abs(b)
18                End If
19                If 1 > CompareTo Then
20                    CompareTo = 1
21                End If
22                CompareTo = Epsilon * CompareTo
23                sNearlyEquals = Abs(a - b) < CompareTo
24                Exit Function
25            End If
26        End If
          'At least one is not a number...
27        If VTA = VTB Then
28            If VTA = vbString And Not CaseSensitive Then
29                If Len(a) = Len(b) Then
30                    sNearlyEquals = UCase$(a) = UCase$(b)
31                Else
32                    sNearlyEquals = False
33                End If
34            Else
35                sNearlyEquals = (a = b)
36            End If
37        Else
38            If VTA = vbBoolean Or VTB = vbBoolean Or VTA = vbString Or VTB = vbString Then
39                sNearlyEquals = False
40            Else
41                sNearlyEquals = (a = b)
42            End If
43        End If

44        Exit Function
ErrHandler:
45        sNearlyEquals = False
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayDivide
' Author    : Philip Swannell
' Date      : 02-May-2015 <-- developed at home in ATigerLib.xlam, to be transferred to SolumAddin
' Purpose   : Element-wise division of arrays of numbers. Non-numbers in the input arguments will
'             produce an error string (e.g. ""#Type mismatch!"") in the corresponding element
'             of the output array.
' Arguments
' Array1    : Any number or array of numbers.
' Array2    : Any number or array of numbers.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayDivide(Array1 As Variant, Array2 As Variant)
Attribute sArrayDivide.VB_Description = "Element-wise division of arrays of numbers. Non-numbers in the input arguments will produce an error string (e.g. ""#Type mismatch!"") in the corresponding element of the output array."
Attribute sArrayDivide.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         If VarType(Array1) < vbArray And VarType(Array2) < vbArray Then
3             sArrayDivide = SafeDivide(Array1, Array2)
4         Else
5             sArrayDivide = Broadcast(FuncIdDivide, Array1, Array2)
6         End If
7         Exit Function
ErrHandler:
8         sArrayDivide = "#sArrayDivide (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayLike
' Author    : Philip Swannell
' Date      : 04-Sep-2015
' Purpose   : Pattern matching. Returns TRUE if TheString matches Pattern.
'             ? - any single character
'             * - zero or more characters
'             # - any digit (0-9)
'             [charlist] - any single character in charlist
'             [!charlist] - any single character not in charlist
'             Case insensitive.
' Arguments
' TheString : A string or array of strings to be matched against Pattern. Non-string input is cast to a
'             string.
' Pattern   : A string expression conforming to the pattern-matching conventions described above or in
'             more detail at
'             https://msdn.microsoft.com/en-us/library/office/gg251796(v=office.15).aspx
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayLike(TheString As Variant, Pattern As Variant)
Attribute sArrayLike.VB_Description = "Pattern matching. Returns TRUE if TheString matches Pattern.\n? - any single character\n* - zero or more characters\n# - any digit (0-9)\n[charlist] - any single character in charlist\n[!charlist] - any single character not in charlist\nCase insensitive."
Attribute sArrayLike.VB_ProcData.VB_Invoke_Func = " \n25"
1         On Error GoTo ErrHandler
2         If VarType(TheString) < vbArray And VarType(Pattern) < vbArray Then
3             sArrayLike = SafeLike(CStr(TheString), CStr(Pattern))
4         Else
5             sArrayLike = Broadcast(FuncIdLike, TheString, Pattern)
6         End If
7         Exit Function
ErrHandler:
8         sArrayLike = "#sArrayLike (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayEquals
' Author    : Philip Swannell
' Date      : 03-May-2015 <-- written at home in ATigerLib.xlam, ported to SolumAddin.xlam
' Purpose   : Element-wise testing for equality of two arrays - the array version of sEquals. Like the =
'             operator in Excel array formulas, but capable of comparing error values, so
'             always returns an array of logicals. See also sArraysIdentical.
' Arguments
' Array1    : The first array to compare, with arbitrary values - numbers, text, errors, logicals etc.
' Array2    : The second array to compare, with arbitrary values - numbers, text, errors, logicals etc.
' CaseSensitive: Determines if comparison of strings is done in a case sensitive manner. If omitted
'             defaults to FALSE (case insensitive matching).
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayEquals(Array1 As Variant, Array2 As Variant, Optional CaseSensitive As Variant = False)
Attribute sArrayEquals.VB_Description = "Element-wise testing for equality of two arrays - the array version of sEquals. Like the = operator in Excel array formulas, but capable of comparing error values, so always returns an array of logicals. See also sArraysIdentical."
Attribute sArrayEquals.VB_ProcData.VB_Invoke_Func = " \n27"
1         On Error GoTo ErrHandler
2         If VarType(Array1) < vbArray And VarType(Array2) < vbArray And VarType(CaseSensitive) = vbBoolean Then
3             sArrayEquals = sEquals(Array1, Array2, CBool(CaseSensitive))
4         Else
5             sArrayEquals = Broadcast(FuncIdEquals, Array1, Array2, CaseSensitive)
6         End If
7         Exit Function
ErrHandler:
8         sArrayEquals = "#sArrayEquals (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRoundSF
' Author    : Philip Swannell
' Date      : 22-Jun-2021
' Purpose   : Rounds a number to an input number of significant figures (base 10). All arguments may be
'             arrays.
' Arguments
' Number    : An array of arbitrary values, but non-numbers in the input yield error strings in the
'             output.
' NumSFs    : A whole number greater than or equal to 1.
' Ties      : Controls how ties (halves) are rounded. Omitted or zero = Away from zero, 1 = Towards
'             zero, 2 = Bankers rounding, 3 = Towards plus infinity, 4 = Towards minus
'             infinity.
' -----------------------------------------------------------------------------------------------------------------------
Function sRoundSF(Number As Variant, NumSFs As Variant, Ties As Variant)
Attribute sRoundSF.VB_Description = "Rounds a number to an input number of significant figures (base 10). All arguments may be arrays."
Attribute sRoundSF.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         If VarType(Number) < vbArray And VarType(NumSFs) < vbArray And VarType(Ties) < vbArray Then
3             sRoundSF = CoreRoundSF(Number, NumSFs, CLng(Ties))
4         Else
5             sRoundSF = Broadcast(FuncIdRoundSF, Number, NumSFs, Ties)
6         End If
7         Exit Function
ErrHandler:
8         sRoundSF = "#sRoundSF (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRound
' Author    : Philip Swannell
' Date      : 22-Jun-2021
' Purpose   : Rounds a number to an input number of digits (base 10). All arguments may be arrays.
' Arguments
' Number    : An array of arbitrary values, but non-numbers in the input yield error strings in the
'             output.
' NumDigits : A whole number. The function rounds to this number of decimal places. Negative numbers are
'             supported.
' Ties      : Controls how ties (halves) are rounded. Omitted or zero = Away from zero, 1 = Towards
'             zero, 2 = Bankers rounding, 3 = Towards plus infinity, 4 = Towards minus
'             infinity.
' -----------------------------------------------------------------------------------------------------------------------
Function sRound(Number As Variant, NumDigits As Variant, Optional Ties As Variant = 0)
Attribute sRound.VB_Description = "Rounds a number to an input number of digits (base 10). All arguments may be arrays."
Attribute sRound.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         If VarType(Number) < vbArray And VarType(NumDigits) < vbArray And VarType(Ties) < vbArray Then
3             sRound = CoreRound(CDbl(Number), CLng(NumDigits), CLng(Ties))
4         Else
5             sRound = Broadcast(FuncIDRound, Number, NumDigits, Ties)
6         End If
7         Exit Function
ErrHandler:
8         sRound = "#sRound (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileInfo
' Author    : Philip Swannell
' Date      : 01-May-2018
' Purpose   : Returns information about a file such as its size, last modified date, or type.
' Arguments
' FileName  : The full name of the file, including the path. May be an array.
' Info      : String (may be array). Allowed: 'Attributes', 'DateCreated' or 'C', 'DateLastAccessed' or
'             'A', 'DateLastModified' or 'M', 'Drive', 'FullName' or 'F', 'Name' or 'N',
'             'ParentFolder', 'ShortName', 'ShortPath', 'Size' or 'S', 'Type' or 'T',
'             'MD5', 'NumLines'
' -----------------------------------------------------------------------------------------------------------------------
Function sFileInfo(FileName As Variant, Optional Info As Variant)
Attribute sFileInfo.VB_Description = "Returns information about a file such as its size, last modified date, or type."
Attribute sFileInfo.VB_ProcData.VB_Invoke_Func = " \n26"
1         On Error GoTo ErrHandler
          Const allInfos = "Size,DateCreated,DateLastAccessed,DateLastModified,MD5"
          Dim AddLabelsInCol As Boolean
          Dim AddLabelsInRow As Boolean
          
2         If IsEmpty(Info) Or IsMissing(Info) Then
3             If sNRows(FileName) = 1 Then
4                 AddLabelsInCol = True
5                 Info = sTokeniseString(allInfos)
6             ElseIf sNCols(FileName) = 1 Then
7                 Info = sArrayTranspose(sTokeniseString(allInfos))
8                 AddLabelsInRow = True
9             End If
10        End If

11        If VarType(FileName) < vbArray And VarType(Info) < vbArray Then
12            sFileInfo = CoreFileInfo(CStr(FileName), CStr(Info))
13            Exit Function
14        Else
15            sFileInfo = Broadcast(FuncIdFileInfo, FileName, Info)
16            If AddLabelsInCol Then
17                sFileInfo = sArrayRange(Info, sFileInfo)
18            ElseIf AddLabelsInRow Then
19                sFileInfo = sArrayStack(Info, sFileInfo)
20            End If

21        End If
22        Exit Function
ErrHandler:
23        sFileInfo = "#sFileInfo (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayNearlyEquals
' Author    : Philip Swannell
' Date      : 16-May-2015
' Purpose   : Element-wise test if arrays of values are "nearly equal". If corresponding elements A and
'             B are both numbers, then they are nearly equal if and only if:
'             Abs(A-B) < Epsilon * Max(1,Abs(A),Abs(B))
' Arguments
' Array1    : An array of arbitrary values.
' Array2    : An array of arbitrary values.
' CaseSensitive: TRUE for case sensitive comparison of strings. FALSE or omitted for case insensitive
'             comparison.
' Epsilon   : Epsilon determines the tolerance for comparison of two number via the formula:
'             A Nearly Equals B iff Abs(A-B) < Epsilon * Max(1,Abs(A),Abs(B))
'             If omitted Epsilon defaults to 0.000000000000001 (i.e. 10^-15)
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayNearlyEquals(Array1 As Variant, Array2 As Variant, Optional CaseSensitive As Variant = False, Optional Epsilon As Variant = cEPSILON)
Attribute sArrayNearlyEquals.VB_Description = "Element-wise test if arrays of values are ""nearly equal"". If corresponding elements A and B are both numbers, then they are nearly equal if and only if:\nAbs(A-B) < Epsilon * Max(1,Abs(A),Abs(B))"
Attribute sArrayNearlyEquals.VB_ProcData.VB_Invoke_Func = " \n24"
1         On Error GoTo ErrHandler
2         If VarType(Array1) < vbArray And VarType(Array2) < vbArray And VarType(CaseSensitive) = vbBoolean And VarType(Epsilon) < vbArray Then
3             sArrayNearlyEquals = sNearlyEquals(Array1, Array2, CBool(CaseSensitive), CDbl(Epsilon))
4         Else
5             sArrayNearlyEquals = Broadcast(FuncIdNearlyEquals, Array1, Array2, CaseSensitive, Epsilon)
6         End If
7         Exit Function
ErrHandler:
8         sArrayNearlyEquals = "#sArrayNearlyEquals (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayLeft
' Author    : Philip Swannell
' Date      : 15-May-2015
' Purpose   : Returns the first NumChars characters of the string Strings, which may be an array. If
'             Strings has fewer characters, the whole string is returned. If NumChars is
'             negative, the final -NumChars characters are removed.
' Arguments
' Strings   : String to be truncated. Can be array.
' NumChars  : Integer number of characters to take. If negative, then that many characters are dropped
'             from the right-hand side. If zero, an empty string is returned. Can be array.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayLeft(Strings, NumChars)
Attribute sArrayLeft.VB_Description = "Returns the first NumChars characters of the string Strings, which may be an array. If Strings has fewer characters, the whole string is returned. If NumChars is negative, the final -NumChars characters are removed."
Attribute sArrayLeft.VB_ProcData.VB_Invoke_Func = " \n25"
1         On Error GoTo ErrHandler
2         If VarType(Strings) < vbArray And VarType(NumChars) < vbArray Then
3             sArrayLeft = SafeLeft(Strings, NumChars)
4         Else
5             sArrayLeft = Broadcast(FuncIdLeft, Strings, NumChars)
6         End If
7         Exit Function
ErrHandler:
8         sArrayLeft = "#sArrayLeft (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayRight
' Author    : Philip Swannell
' Date      : 15-May-2015
' Purpose   : Returns the last NumChars characters of the string Strings, which may be an array. If
'             Strings has fewer characters, the whole string is returned. If NumChars is
'             negative, the first -NumChars characters are removed.
' Arguments
' Strings   : String to be truncated. Can be array.
' NumChars  : Integer number of characters to take. If negative, then that many characters are dropped
'             from the left-hand side. If zero, an empty string is returned. Can be array.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayRight(Strings, NumChars)
Attribute sArrayRight.VB_Description = "Returns the last NumChars characters of the string Strings, which may be an array. If Strings has fewer characters, the whole string is returned. If NumChars is negative, the first -NumChars characters are removed."
Attribute sArrayRight.VB_ProcData.VB_Invoke_Func = " \n25"
1         On Error GoTo ErrHandler
2         If VarType(Strings) < vbArray And VarType(NumChars) < vbArray Then
3             sArrayRight = SafeRight(Strings, NumChars)
4         Else
5             sArrayRight = Broadcast(FuncIdRight, Strings, NumChars)
6         End If
7         Exit Function
ErrHandler:
8         sArrayRight = "#sArrayRight (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sIfErrorString
' Author    : Philip Swannell
' Date      : 21-May-2015
' Purpose   : If Value is an error string (a string starting with # and ending with !) then the function
'             returns ValueIfErrorString. Otherwise the function returns Value unchanged.
'             The function is analagous to the native Excel function IFERROR.
' Arguments
' Value     : Any value or array of values
' ValueIfErrorString: Any value or array of values
' -----------------------------------------------------------------------------------------------------------------------
Function sIfErrorString(Value, ValueIfErrorString)
Attribute sIfErrorString.VB_Description = "If Value is an error string (a string starting with # and ending with !) then the function returns ValueIfErrorString. Otherwise the function returns Value unchanged. The function is analagous to the native Excel function IFERROR."
Attribute sIfErrorString.VB_ProcData.VB_Invoke_Func = " \n24"
1         On Error GoTo ErrHandler
2         If VarType(Value) < vbArray And VarType(ValueIfErrorString) < vbArray Then
3             sIfErrorString = SafeIfErrorString(Value, ValueIfErrorString)
4         Else
5             sIfErrorString = Broadcast(FuncIdIfErrorString, Value, ValueIfErrorString)
6         End If
7         Exit Function
ErrHandler:
8         sIfErrorString = "#sIfErrorString (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayGreaterThan
' Author    : Philip Swannell
' Date      : 07-May-2015
' Purpose   : Element-wise comparison of arrays. TRUE if Array1 > Array2. The function always returns
'             logicals even when elements of the inputs are of different type or are error
'             values.
' Arguments
' Array1    : The first array to compare
' Array2    : The second array to compare
'
' Notes     : Rules:
'             TRUE > FALSE
'             Error > Logical > String > Number
'             #N/A! > #NUM! > #NAME? > #REF! > #VALUE! > #DIV/0! > #NULL!
'             Dates are treated as the number to which they cast.
'             Strings are compared in a way that pays attention to accented characters e.g.
'             F > f > É > é > E > e
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayGreaterThan(Array1 As Variant, Array2 As Variant)
Attribute sArrayGreaterThan.VB_Description = "Element-wise evaluation of Array1 > Array2. Returns logicals whatever the type of the input elements."
Attribute sArrayGreaterThan.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         If VarType(Array1) < vbArray And VarType(Array2) < vbArray Then
3             sArrayGreaterThan = VariantLessThan(Array2, Array1, True)        'Note argument switch here
4         Else
5             sArrayGreaterThan = Broadcast(FuncIdLessThan, Array2, Array1)        'Note argument switch here
6         End If
7         Exit Function
ErrHandler:
8         sArrayGreaterThan = "#sArrayGreaterThan (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayGreaterThanOrEqual
' Author    : Philip Swannell
' Purpose   : Element-wise comparison of arrays. TRUE if Array1 >= Array2. The function always returns
'             logicals even when elements of the inputs are of different type or are error
'             values.
'
'             Rules:
'             TRUE > FALSE
'             Error > Logical > String > Number
'             F > f > É > é > E > e
' Arguments
' Array1    : The first array to compare
' Array2    : The second array to compare
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayGreaterThanOrEqual(Array1 As Variant, Array2 As Variant)
Attribute sArrayGreaterThanOrEqual.VB_Description = "Element-wise evaluation of Array1 >= Array2. Returns logicals whatever the type of the input elements."
Attribute sArrayGreaterThanOrEqual.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         If VarType(Array1) < vbArray And VarType(Array2) < vbArray Then
3             sArrayGreaterThanOrEqual = VariantLessThanOrEqual(Array2, Array1, True)        'Note argument switch here
4         Else
5             sArrayGreaterThanOrEqual = Broadcast(FuncIdlessThanOrEqual, Array2, Array1)        'Note argument switch here
6         End If
7         Exit Function
ErrHandler:
8         sArrayGreaterThanOrEqual = "#sArrayGreaterThanOrEqual (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayIf
' Author    : Philip Swannell
' Purpose   : Performs elementwise test of logical value IfCondition, and returns the corresponding
'             element of ValueIfTrue if the condition holds, and the element
'             of ValueIfFalse if the condition does not hold. All arguments can be arrays.
' Arguments
' IfCondition: Array of logicals controlling which return should be given.
' ValueIfTrue: Array of any type, giving the values to be returned if TRUE.
' ValueIfFalse: Array of any type, giving the values to be returned if FALSE.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayIf(IfCondition As Variant, ValueIfTrue As Variant, ValueIfFalse As Variant)
Attribute sArrayIf.VB_Description = "Performs element-wise test of logical value IfCondition, and returns the corresponding element of ValueIfTrue if the condition holds, and the element of ValueIfFalse if the condition does not hold. All arguments can be arrays."
Attribute sArrayIf.VB_ProcData.VB_Invoke_Func = " \n24"
1         On Error GoTo ErrHandler
2         If VarType(IfCondition) < vbArray And VarType(ValueIfTrue) < vbArray And VarType(ValueIfFalse) < vbArray Then
3             sArrayIf = SafeIf(IfCondition, ValueIfTrue, ValueIfFalse)
4         Else
5             sArrayIf = Broadcast(FuncIdIf, IfCondition, ValueIfTrue, ValueIfFalse)
6         End If
7         Exit Function
ErrHandler:
8         sArrayIf = "#sArrayIf (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayLessThan
' Author    : Philip Swannell
' Date      : 07-May-2015
' Purpose   : Element-wise comparison of arrays. TRUE if Array1 < Array2. The function always returns
'             logicals even when elements of the inputs are of different type or are error
'             values.
' Arguments
' Array1    : The first array to compare
' Array2    : The second array to compare
'
' Notes     : Rules:
'             FALSE < TRUE
'             Number < String < Logical < Error
'             #NULL! < #DIV/0! < #VALUE! < #REF! < #NAME? < #NUM! < #N/A!
'             Dates are treated as the number to which they cast.
'             Strings are compared in a way that pays attention to accented characters e.g.
'             e < E < é < É < f < F
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayLessThan(Array1 As Variant, Array2 As Variant)
Attribute sArrayLessThan.VB_Description = "Element-wise evaluation of Array1 < Array2. Returns logicals whatever the type of the input elements."
Attribute sArrayLessThan.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         If VarType(Array1) < vbArray And VarType(Array2) < vbArray Then
3             sArrayLessThan = VariantLessThan(Array1, Array2, True)
4         Else
5             sArrayLessThan = Broadcast(FuncIdLessThan, Array1, Array2)
6         End If
7         Exit Function
ErrHandler:
8         sArrayLessThan = "#sArrayLessThan (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayLessThanOrEqual
' Author    : Philip Swannell
' Purpose   : Element-wise comparison of arrays. TRUE if Array1 <= Array2. The function always returns
'             logicals even when elements of the inputs are of different type or are error
'             values.
'
'             Rules:
'             FALSE < TRUE
'             Number < String < Logical < Error
'             e < E < é < É < f < F
' Arguments
' Array1    : The first array to compare
' Array2    : The second array to compare
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayLessThanOrEqual(Array1 As Variant, Array2 As Variant)
Attribute sArrayLessThanOrEqual.VB_Description = "Element-wise evaluation of Array1 <= Array2. Returns logicals whatever the type of the input elements."
Attribute sArrayLessThanOrEqual.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         If VarType(Array1) < vbArray And VarType(Array2) < vbArray Then
3             sArrayLessThanOrEqual = VariantLessThanOrEqual(Array1, Array2, True)
4         Else
5             sArrayLessThanOrEqual = Broadcast(FuncIdlessThanOrEqual, Array1, Array2)
6         End If
7         Exit Function
ErrHandler:
8         sArrayLessThanOrEqual = "#sArrayLessThanOrEqual (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayPower
' Author    : Philip Swannell
' Purpose   : Returns the value of Base raised to the power of Exponent. Both arguments may be arrays.
'             The function will return 1 if both inputs are zero.
'
'             Returns an error if:
'
'             Base is zero and Exponent is negative, or
'             Base is negative and Exponent is non-integral.
' Arguments
' Base      : Number whose power will be taken. Can be any sign. Can be an array of numbers.
' Exponent  : Numerical value of the exponent (power). Can be any sign. Can be an array of numbers.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayPower(Base As Variant, Exponent As Variant)
Attribute sArrayPower.VB_Description = "Returns the value of Base raised to the power of Exponent. Both arguments may be arrays. The function will return 1 if both inputs are zero.\n\nReturns an error if:\n\nBase is zero and Exponent is negative, or\nBase is negative and Exponent is non-integral."
Attribute sArrayPower.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         If VarType(Base) < vbArray And VarType(Exponent) < vbArray Then
3             sArrayPower = SafePower(Base, Exponent)
4         Else
5             sArrayPower = Broadcast(FuncIdPower, Base, Exponent)
6         End If
7         Exit Function
ErrHandler:
8         sArrayPower = "#sArrayPower (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArraySubtract
' Author    : Philip Swannell
' Purpose   : Element-wise subtraction of arrays of numbers for use from VBA. Replicates the - operator
'             in Excel array formulas.
' Arguments
' Array1    : An array of numbers
' Array2    : An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sArraySubtract(Array1 As Variant, Array2 As Variant)
Attribute sArraySubtract.VB_Description = "Element-wise subtraction of arrays of numbers for use from VBA. Replicates the - operator in Excel array formulas."
Attribute sArraySubtract.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         If VarType(Array1) < vbArray And VarType(Array2) < vbArray Then
3             sArraySubtract = SafeSubtract(Array1, Array2)
4         Else
5             sArraySubtract = Broadcast(FuncIdSubtract, Array1, Array2)
6         End If

7         Exit Function
ErrHandler:
8         sArraySubtract = "#sArraySubtract (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sStringBetweenStrings
' Author    : Philip Swannell
' Date      : 07-May-2015
' Purpose   : The function returns the substring of the input TheString which lies between LeftString
'             and RightString. All arguments can be arrays of strings.
' Arguments
' TheString : The input string to be searched.
' LeftString: The returned string will start immediately after the first occurrence of LeftString in
'             TheString. If LeftString is not found or is the null string or missing, then
'             the return will start at the first character of TheString.
' RightString: The return will stop immediately before the first subsequent occurrence of RightString. If
'             such occurrrence is not found or if RightString is the null string or
'             missing, then the return will stop at the last character of TheString.
' IncludeLeftString: If TRUE, then if LeftString appears in TheString, the return will include LeftString. This
'             argument is optional and defaults to FALSE.
' IncludeRightString: If TRUE, then if RightString appears in TheString (and appears after the first occurance
'             of LeftString) then the return will include RightString. This argument is
'             optional and defaults to FALSE.
' -----------------------------------------------------------------------------------------------------------------------
Function sStringBetweenStrings(TheString, Optional ByVal LeftString, Optional ByVal RightString, Optional IncludeLeftString As Boolean, Optional IncludeRightString As Boolean)
Attribute sStringBetweenStrings.VB_Description = "The function returns the substring of the input TheString which lies between LeftString and RightString. All arguments can be arrays of strings."
Attribute sStringBetweenStrings.VB_ProcData.VB_Invoke_Func = " \n25"
1         If IsEmpty(LeftString) Or IsMissing(LeftString) Then LeftString = vbNullString
2         If IsEmpty(RightString) Or IsMissing(RightString) Then RightString = vbNullString
3         On Error GoTo ErrHandler
4         If VarType(TheString) < vbArray And VarType(LeftString) < vbArray And VarType(RightString) < vbArray Then
5             sStringBetweenStrings = CoreStringBetweenStrings(TheString, LeftString, RightString, IncludeLeftString, IncludeRightString)
6         Else
7             sStringBetweenStrings = Broadcast(FuncIdStringBetweenStrings, TheString, LeftString, RightString, IncludeLeftString, IncludeRightString)
8         End If
9         Exit Function
ErrHandler:
10        sStringBetweenStrings = "#sStringBetweenStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayAdd
' Author    : Philip Swannell
' Date      : 02-May-2015 <-- developed at home in ATigerLib.xlam, to be transferred to SolumAddin
' Purpose   : Element-wise addition of 2-dimensional arrays. Mimics addition operator in Excel and other
'             array-processing languages(so has no use from Excel, only from VBA). Handles matrix to matrix
'             addition, matrix to vector, scalar to vector etc.
'             Height (no of rows) of output is minimum of heights of input arrays ignoring those of height 1.
'             Width of output is minimum of widths of inputs, ignoring those of width 1.
'             Where elements of the input arrays are not numbers then the corresponding element of the output
'             will be an error string, such as #Type mismatch! (Excel addition would give an excel error value such as #VALUE!)
'             Warning for VBA users: If the inputs are 0 or 1 dimensional then they are converted to
'             2 dimensional array and that change will be visible to the calling method (since ParamArray is always ByRef)
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayAdd(ParamArray ArraysToAdd())
Attribute sArrayAdd.VB_Description = "Element-wise addition of arrays of numbers for use from VBA. Replicates the + operator in Excel array formulas."
Attribute sArrayAdd.VB_ProcData.VB_Invoke_Func = " \n30"
1         sArrayAdd = BroadcastAssociative(FuncIdAdd, ArraysToAdd)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayAnd
' Author    : Philip Swannell
' Date      : 09-May-2015
' Purpose   : Takes the logical 'or' of any number of arrays. Will return TRUE for any array element if
'             any input array has TRUE in the corresponding element.
' Arguments
' ArraysToAnd:
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayAnd(ParamArray ArraysToAnd())
Attribute sArrayAnd.VB_Description = "Element-wise 'and' of any number of arrays of logical values."
Attribute sArrayAnd.VB_ProcData.VB_Invoke_Func = " \n24"
1         sArrayAnd = BroadcastAssociative(FuncIdAnd, ArraysToAnd)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayConcatenate
' Author    : Philip Swannell
' Date      : 07-May-2015
' Purpose   : Concatenate arrays of strings. Error values (e.g. #NAME?, #DIV0!)  are passed through to
'             the output. Other non strings (numbers, logicals) are cast to strings before
'             concatenation.
' Arguments
' ArraysToConcatenate: An array of values. Numbers and strings will be cast to strings before concatenation.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayConcatenate(ParamArray ArraysToConcatenate())
Attribute sArrayConcatenate.VB_Description = "Concatenate arrays of strings. Error values (e.g. #NAME?, #DIV0!)  are passed through to the output. Other non strings (numbers, logicals) are cast to strings before concatenation."
Attribute sArrayConcatenate.VB_ProcData.VB_Invoke_Func = " \n24"
1         sArrayConcatenate = BroadcastAssociative(FuncIdConcatenate, ArraysToConcatenate)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayMax
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : Returns the element-wise maximum of an arbitrary number of arrays. If non-numbers appear
'             in the  arrays then an error string will appear in the corresponding
'             element(s) of the output. If all arrays are missing returns an error.
' Arguments
' Arrays    :
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayMax(ParamArray Arrays())
Attribute sArrayMax.VB_Description = "Returns the element-wise maximum of an arbitrary number of arrays. If non-numbers appear in the  arrays then an error string will appear in the corresponding element(s) of the output. If all arrays are missing returns an error."
Attribute sArrayMax.VB_ProcData.VB_Invoke_Func = " \n30"
1         sArrayMax = BroadcastAssociative(FuncIdMax, Arrays)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayMin
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : Returns the element-wise minimum of an arbitrary number of arrays. If non-numbers appear
'             in the  arrays then an error string will appear in the corresponding
'             element(s) of the output. If all arrays are missing returns an error.
' Arguments
' Arrays    :
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayMin(ParamArray Arrays())
Attribute sArrayMin.VB_Description = "Returns the element-wise minimum of an arbitrary number of arrays. If non-numbers appear in the  arrays then an error string will appear in the corresponding element(s) of the output. If all arrays are missing returns an error."
Attribute sArrayMin.VB_ProcData.VB_Invoke_Func = " \n30"
1         sArrayMin = BroadcastAssociative(FuncIdMin, Arrays)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMaxOfArray
' Author    : Philip Swannell
' Date      : 07-Mar-2016
' Purpose   : Returns the maximum of the elements of the input TheArray. If any elements are not
'             numbers, an error is returned.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sMaxOfArray(ByVal TheArray)
Attribute sMaxOfArray.VB_Description = "Returns the maximum of the elements of the input TheArray. If any elements are not numbers, an error is returned."
Attribute sMaxOfArray.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim c As Variant
          Dim Res As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray
3         Res = TheArray(1, 1)

4         For Each c In TheArray
5             If Not IsNumberOrDate(c) Then Throw "Non-number encountered"
6             Res = SafeMax(Res, c)
7         Next c
8         sMaxOfArray = Res

9         Exit Function
ErrHandler:
10        sMaxOfArray = "#sMaxOfArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sMinOfArray
' Author    : Philip Swannell
' Date      : 07-Mar-2016
' Purpose   : Returns the minimum of the elements of the input TheArray. If any elements are not
'             numbers, an error is returned.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sMinOfArray(ByVal TheArray)
Attribute sMinOfArray.VB_Description = "Returns the minimum of the elements of the input TheArray. If any elements are not numbers, an error is returned."
Attribute sMinOfArray.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim c As Variant
          Dim Res As Variant

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray
3         Res = TheArray(1, 1)

4         For Each c In TheArray
5             If Not IsNumberOrDate(c) Then Throw "Non-number encountered"
6             Res = SafeMin(Res, c)
7         Next c
8         sMinOfArray = Res

9         Exit Function
ErrHandler:
10        sMinOfArray = "#sMinOfArray (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sAny
' Author    : Philip Swannell
' Date      : 07-Mar-2016
' Purpose   : Returns TRUE if all the elements of TheArray are logicals and any of them are TRUE; FALSE
'             if all the element of TheArray are FALSE, and an error if any element of
'             TheArray is not a logical.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sAny(ByVal TheArray)
Attribute sAny.VB_Description = "Returns TRUE if all the elements of TheArray are logicals and any are TRUE; FALSE if all are FALSE, and an error if any is not a logical."
Attribute sAny.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim c As Variant
          Dim Res As Boolean

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray
3         Res = False

4         For Each c In TheArray
5             If VarType(c) <> vbBoolean Then Throw "Non-logical encountered"
6             Res = Res Or c
7         Next c
8         sAny = Res
9         Exit Function
ErrHandler:
10        sAny = "#sAny (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sAll
' Author    : Philip Swannell
' Date      : 07-Mar-2016
' Purpose   : Returns TRUE if all the elements of TheArray are TRUE; FALSE if all the elements of
'             TheArray are logicals and any is FALSE; and an error if any element of
'             TheArray is not a logical.
' Arguments
' TheArray  : An array of arbitrary values.
' -----------------------------------------------------------------------------------------------------------------------
Function sAll(ByVal TheArray)
Attribute sAll.VB_Description = "Returns TRUE if all the elements of TheArray are TRUE; FALSE if all are logical and any is FALSE; and an error is any is not a logical."
Attribute sAll.VB_ProcData.VB_Invoke_Func = " \n24"
          Dim c As Variant
          Dim Res As Boolean

1         On Error GoTo ErrHandler
2         Force2DArrayR TheArray
3         Res = True
4         For Each c In TheArray
5             If VarType(c) <> vbBoolean Then Throw "Non-logical encountered"
6             Res = Res And c
7         Next c
8         sAll = Res
9         Exit Function
ErrHandler:
10        sAll = "#sAll (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayMultiply
' Author    : Philip Swannell
' Date      : 02-May-2015 <-- developed at home in ATigerLib.xlam, to be transferred to SolumAddin
' Purpose   : Element-wise multiplication of 2-dimensional arrays. Mimics multiply operator in Excel and other
'             array-processing languages(so has no use from Excel, only from VBA). Handles matrix to matrix
'             multiplication, matrix to vector, scalar to vector etc.
'             Height (no of rows) of output is minimum of heights of input arrays ignoring those of height 1.
'             Width of output is minimum of widths of inputs, ignoring those of width 1.
'             Where elements of the input arrays are not numbers then the corresponding element of the output
'             will be an error string, such as #Type mismatch! (Excel addition would give an excel error value such as #VALUE!)
'             Warning for VBA users: If the inputs are 0 or 1 dimensional then they are converted to
'             2 dimensional array and that change will be visible to the calling method (since ParamArray is always ByRef)
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayMultiply(ParamArray ArraysToMultiply())
Attribute sArrayMultiply.VB_Description = "Element-wise addition of arrays of numbers for use from VBA. Replicates the * operator in Excel array formulas."
Attribute sArrayMultiply.VB_ProcData.VB_Invoke_Func = " \n30"
1         sArrayMultiply = BroadcastAssociative(FuncIdMultiply, ArraysToMultiply)
End Function
' Procedure : sArrayOr
' Author    : Philip Swannell
' Date      : 09-May-2015
' Purpose   : Multi-call Or
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayOr(ParamArray ArraysToOr())
Attribute sArrayOr.VB_Description = "Element-wise 'or' of any number of arrays of logical values."
Attribute sArrayOr.VB_ProcData.VB_Invoke_Func = " \n24"
1         sArrayOr = BroadcastAssociative(FuncIdor, ArraysToOr)
End Function

'===========================================================================================END OF BROADCASTASSOCIATIVE FUNCTIONS
'BROADCASTCOLUMN FUNCTIONS
' -----------------------------------------------------------------------------------------------------------------------
' Procedures : sColumnAnd, sColumnMax, sColumnMean, sColumnMin, sColumnOr, sColumnProduct, sColumnStDev, sColumnSum
' Author    : Philip Swannell
' Date      : 06-May-2015
' Purpose   : Element-wise processing of the columns of an array to yield a row vector
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnAnd(ByVal ArrayOfLogicals)
Attribute sColumnAnd.VB_Description = "For an array of logicals returns a one row array of logicals where each element of the output is TRUE only if all elements in the corresponding column of the input is TRUE. Non logicals in the input yield error strings within the output."
Attribute sColumnAnd.VB_ProcData.VB_Invoke_Func = " \n24"
1         sColumnAnd = BroadcastColumn(FuncIdAnd, ArrayOfLogicals)
End Function

Function sColumnMax(ByVal ArrayOfNumbers)
Attribute sColumnMax.VB_Description = "For an array of numbers returns a one row array of numbers where each element of the output is the maximum of the corresponding column of the input. Non-numbers in the input yield error strings within the output."
Attribute sColumnMax.VB_ProcData.VB_Invoke_Func = " \n30"
1         sColumnMax = BroadcastColumn(FuncIdMax, ArrayOfNumbers)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnMean
' Author    : Philip Swannell
' Date      : 15-Jul-2017
' Purpose   : For an array of numbers returns a one row array of numbers where each element of the
'             return is the mean of the corresponding column of the input.
' Arguments
' ArrayOfNumbers: An array of numbers
' IgnoreNonNumbers: If FALSE (the default) then non-numbers in ArrayOfNumbers yield error strings in the
'             corresponding element of the return. If TRUE the non-numbers are excluded
'             from calculation of the column means.
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnMean(ByVal ArrayOfNumbers, Optional IgnoreNonNumbers As Boolean)
Attribute sColumnMean.VB_Description = "For an array of numbers returns a one row array of numbers where each element of the return is the mean of the corresponding column of the input."
Attribute sColumnMean.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
          Dim NC As Long
          Dim NR As Long
2         Force2DArrayR ArrayOfNumbers, NR, NC
3         If Not IgnoreNonNumbers Then
4             sColumnMean = sArrayDivide(sColumnSum(ArrayOfNumbers), NR)
5         Else
              Dim Denom As Double
              Dim i As Long
              Dim j As Long
              Dim NUM As Double
              Dim Result() As Variant
6             ReDim Result(1 To 1, 1 To NC)
7             For j = 1 To NC
8                 NUM = 0: Denom = 0
9                 For i = 1 To NR
10                    If IsNumberOrDate(ArrayOfNumbers(i, j)) Then
11                        NUM = NUM + ArrayOfNumbers(i, j)
12                        Denom = Denom + 1
13                    End If
14                Next i
15                If Denom = 0 Then
16                    Result(1, j) = "#No numbers found!"
17                Else
18                    Result(1, j) = NUM / Denom
19                End If
20            Next j
21            sColumnMean = Result
22        End If

23        Exit Function
ErrHandler:
24        sColumnMean = "#sColumnMean (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnMin
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one row array of numbers where each element of the
'             output is the minimum of the corresponding column of the input. Non-numbers
'             in the input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnMin(ByVal ArrayOfNumbers)
Attribute sColumnMin.VB_Description = "For an array of numbers returns a one row array of numbers where each element of the output is the minimum of the corresponding column of the input. Non-numbers in the input yield error strings within the output."
Attribute sColumnMin.VB_ProcData.VB_Invoke_Func = " \n30"
1         sColumnMin = BroadcastColumn(FuncIdMin, ArrayOfNumbers)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnOr
' Author    : Philip Swannell
' Purpose   : For an array of logicals returns a one row array of logicals where each element of the
'             output is TRUE only if any element in the corresponding column of the input
'             is TRUE. Non logicals in the input yield error strings within the output.
' Arguments
' ArrayOfLogicals: An array of logicals
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnOr(ByVal ArrayOfLogicals)
Attribute sColumnOr.VB_Description = "For an array of logicals returns a one row array of logicals where each element of the output is TRUE only if any element in the corresponding column of the input is TRUE. Non logicals in the input yield error strings within the output."
Attribute sColumnOr.VB_ProcData.VB_Invoke_Func = " \n24"
1         sColumnOr = BroadcastColumn(FuncIdor, ArrayOfLogicals)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnProduct
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one row array of numbers where each element of the
'             output is the product of all elements in the corresponding column of the
'             input. Non-numbers in the input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnProduct(ByVal ArrayOfNumbers)
Attribute sColumnProduct.VB_Description = "For an array of numbers returns a one row array of numbers where each element of the output is the product of all elements in the corresponding column of the input. Non-numbers in the input yield error strings within the output."
Attribute sColumnProduct.VB_ProcData.VB_Invoke_Func = " \n30"
1         sColumnProduct = BroadcastColumn(FuncIdMultiply, ArrayOfNumbers)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnStDev
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one row array of numbers where each element of the
'             output is the population standard deviation of the corresponding column of
'             the input. Non-numbers in the input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnStDev(ByVal ArrayOfNumbers)
Attribute sColumnStDev.VB_Description = "For an array of numbers returns a one row array of numbers where each element of the output is the population standard deviation of the corresponding column of the input. Non-numbers in the input yield error strings within the output."
Attribute sColumnStDev.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim Res
1         Force2DArrayR ArrayOfNumbers
2         Res = sArraySubtract(ArrayOfNumbers, sColumnMean(ArrayOfNumbers))
3         sColumnStDev = sArrayPower(sArrayDivide(sColumnSum(sArrayMultiply(Res, Res)), sNRows(ArrayOfNumbers)), 0.5)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnSum
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one row array of numbers where each element of the
'             output is the sum of the corresponding column of the input. Non-numbers in
'             the input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnSum(ByVal ArrayOfNumbers)
Attribute sColumnSum.VB_Description = "For an array of numbers returns a one row array of numbers where each element of the output is the sum of the corresponding column of the input. Non-numbers in the input yield error strings within the output."
Attribute sColumnSum.VB_ProcData.VB_Invoke_Func = " \n30"
1         sColumnSum = BroadcastColumn(FuncIdAdd, ArrayOfNumbers)
End Function

'===========================================================================================END OF BROADCASTCOLUMN FUNCTIONS
'BROADCASTROW FUNCTIONS
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowAnd
' Author    : Philip Swannell
' Purpose   : For an array of logicals returns a one column array of logicals where each element of the
'             output is TRUE only if all elements in the corresponding row of the input is
'             TRUE. Non logicals in the input yield error strings within the output.
' Arguments
' ArrayOfLogicals: An array of logicals
' -----------------------------------------------------------------------------------------------------------------------
Function sRowAnd(ByVal ArrayOfLogicals)
Attribute sRowAnd.VB_Description = "For an array of logicals returns a one column array of logicals where each element of the output is TRUE only if all elements in the corresponding row of the input is TRUE. Non logicals in the input yield error strings within the output."
Attribute sRowAnd.VB_ProcData.VB_Invoke_Func = " \n24"
1         sRowAnd = BroadcastRow(FuncIdAnd, ArrayOfLogicals)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowMax
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one column array of numbers where each element of the
'             output is the maximum of the corresponding row of the input. Non-numbers in
'             the input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sRowMax(ByVal ArrayOfNumbers)
Attribute sRowMax.VB_Description = "For an array of numbers returns a one column array of numbers where each element of the output is the maximum of the corresponding row of the input. Non-numbers in the input yield error strings within the output."
Attribute sRowMax.VB_ProcData.VB_Invoke_Func = " \n30"
1         sRowMax = BroadcastRow(FuncIdMax, ArrayOfNumbers)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowMean
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one column array of numbers where each element of the
'             output is the mean of the corresponding row of the input. Non-numbers in the
'             input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sRowMean(ByVal ArrayOfNumbers)
Attribute sRowMean.VB_Description = "For an array of numbers returns a one column array of numbers where each element of the output is the mean of the corresponding row of the input. Non-numbers in the input yield error strings within the output."
Attribute sRowMean.VB_ProcData.VB_Invoke_Func = " \n30"
1         On Error GoTo ErrHandler
2         Force2DArrayR ArrayOfNumbers
3         sRowMean = sArrayDivide(sRowSum(ArrayOfNumbers), sNCols(ArrayOfNumbers))
4         Exit Function
ErrHandler:
5         sRowMean = "#sRowMean (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowMin
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one column array of numbers where each element of the
'             output is the minimum of the corresponding row of the input. Non-numbers in
'             the input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sRowMin(ByVal ArrayOfNumbers)
Attribute sRowMin.VB_Description = "For an array of numbers returns a one column array of numbers where each element of the output is the minimum of the corresponding row of the input. Non-numbers in the input yield error strings within the output."
Attribute sRowMin.VB_ProcData.VB_Invoke_Func = " \n30"
1         sRowMin = BroadcastRow(FuncIdMin, ArrayOfNumbers)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowOr
' Author    : Philip Swannell
' Purpose   : For an array of logicals returns a one column array of logicals where each element of the
'             output is TRUE only if any element in the corresponding row of the input is
'             TRUE. Non logicals in the input yield error strings within the output.
' Arguments
' ArrayOfLogicals: An array of logicals
' -----------------------------------------------------------------------------------------------------------------------
Function sRowOr(ByVal ArrayOfLogicals)
Attribute sRowOr.VB_Description = "For an array of logicals returns a one column array of logicals where each element of the output is TRUE only if any element in the corresponding row of the input is TRUE. Non logicals in the input yield error strings within the output."
Attribute sRowOr.VB_ProcData.VB_Invoke_Func = " \n24"
1         sRowOr = BroadcastRow(FuncIdor, ArrayOfLogicals)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowProduct
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one column array of numbers where each element of the
'             output is the product of all elements in the corresponding row of the input.
'             Non-numbers in the input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sRowProduct(ByVal ArrayOfNumbers)
Attribute sRowProduct.VB_Description = "For an array of numbers returns a one column array of numbers where each element of the output is the product of all elements in the corresponding row of the input. Non-numbers in the input yield error strings within the output."
Attribute sRowProduct.VB_ProcData.VB_Invoke_Func = " \n30"
1         sRowProduct = BroadcastRow(FuncIdMultiply, ArrayOfNumbers)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowStDev
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one column array of numbers where each element of the
'             output is the population standard deviation of the corresponding row of the
'             input. Non-numbers in the input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sRowStDev(ByVal ArrayOfNumbers)
Attribute sRowStDev.VB_Description = "For an array of numbers returns a one column array of numbers where each element of the output is the population standard deviation of the corresponding row of the input. Non-numbers in the input yield error strings within the output."
Attribute sRowStDev.VB_ProcData.VB_Invoke_Func = " \n30"
          Dim Res
1         Force2DArrayR ArrayOfNumbers
2         Res = sArraySubtract(ArrayOfNumbers, sRowMean(ArrayOfNumbers))
3         sRowStDev = sArrayPower(sArrayDivide(sRowSum(sArrayMultiply(Res, Res)), sNCols(ArrayOfNumbers)), 0.5)
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRowSum
' Author    : Philip Swannell
' Purpose   : For an array of numbers returns a one column array of numbers where each element of the
'             output is the sum of the corresponding row of the input. Non-numbers in the
'             input yield error strings within the output.
' Arguments
' ArrayOfNumbers: An array of numbers
' -----------------------------------------------------------------------------------------------------------------------
Function sRowSum(ByVal ArrayOfNumbers)
Attribute sRowSum.VB_Description = "For an array of numbers returns a one column array of numbers where each element of the output is the sum of the corresponding row of the input. Non-numbers in the input yield error strings within the output."
Attribute sRowSum.VB_ProcData.VB_Invoke_Func = " \n30"
1         sRowSum = BroadcastRow(FuncIdAdd, ArrayOfNumbers)
End Function
'===========================================================================================END OF BROADCASTROW FUNCTIONS
