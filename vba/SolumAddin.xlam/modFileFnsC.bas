Attribute VB_Name = "modFileFnsC"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileScale
' Author    : Philip Swannell
' Date      : 05-Jun-2019
' Purpose   : Creates a "scaled" copy of a text file in which numeric elements are multiplied by Factor.
'             Non-numeric elements and top and left headers are unchanged.
' Arguments
' SourceFile: The full name of the source file, including the path. An array of file names is not
'             currently supported.
' TargetFile: The full name of the target file, including the path. An array of file names is not
'             currently supported.
' Factor    : The multiplicative factor.
' Delimiter : The delimiter character. If omitted the function guesses the delimiter from the contents
'             of the first line of SourceFile.
' NumTopHeaders: The number of header rows at the top of the file. If omitted defaults to 1. Header rows
'             are copied unchanged to TargetFile.
' NumLeftHeaders: The number of header columns at the left of the file. If omitted defaults to 1. Header
'             columns are copied unchanged to TargetFile.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileScale(SourceFile As String, TargetFile As String, Factor As Double, Delimiter As String, Optional NumTopHeaders As Long = 1, Optional NumLeftHeaders As Long = 1)
Attribute sFileScale.VB_Description = "Creates a ""scaled"" copy of a text file in which numeric elements are multiplied by Factor. Non-numeric elements and top and left headers are unchanged."
Attribute sFileScale.VB_ProcData.VB_Invoke_Func = " \n26"

          Dim F As Scripting.file
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim j As Long
          Dim RawLine As String
          Dim SplitLine() As String
          Dim Tin As TextStream
          Dim Tout As TextStream

1         On Error GoTo ErrHandler
          
2         If Not sFileExists(SourceFile) Then Throw "Cannot find file '" + SourceFile + "'"
3         If LCase$(SourceFile) = LCase$(TargetFile) Then Throw "SourceFile and TargetFile must be different"
4         CheckFileNameIsAbsolute TargetFile
          
5         If Delimiter = "" Then Throw "Delimiter must be provided"

6         If NumTopHeaders < 0 Then Throw "NumTopHeaders must not be negative"
7         If NumLeftHeaders < 0 Then Throw "NumLeftHeaders must not be negative"

8         If Factor = 1 Then
9             sFileScale = sFileCopy(SourceFile, TargetFile)
10            Exit Function
11        End If

12        Set FSO = New FileSystemObject
13        Set F = FSO.GetFile(SourceFile)

14        Set Tin = F.OpenAsTextStream(ForReading)
15        Set Tout = FSO.OpenTextFile(TargetFile, ForWriting, True, TristateFalse)

16        For i = 1 To NumTopHeaders
17            If Not Tin.atEndOfStream Then
18                RawLine = Tin.ReadLine
19                Tout.WriteLine RawLine
20            End If
21        Next

22        Do While Not Tin.atEndOfStream
23            RawLine = Tin.ReadLine
24            SplitLine = VBA.Split(RawLine, Delimiter)
25            On Error Resume Next
26            For j = NumLeftHeaders To UBound(SplitLine) 'Split returns zero-based array
27                SplitLine(j) = CStr(CDbl(SplitLine(j) * Factor))
28            Next j
29            On Error GoTo ErrHandler
30            Tout.WriteLine VBA.Join(SplitLine, Delimiter)
31        Loop

32        Tin.Close: Tout.Close: Set Tin = Nothing: Set Tout = Nothing: Set F = Nothing: Set FSO = Nothing

33        sFileScale = TargetFile

34        Exit Function
ErrHandler:
35        Throw "#sFileScale (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileSuppressZeros
' Author    : Philip Swannell
' Date      : 06-Jun-2019
' Purpose   : Creates a copy of a text file in which zero elements are replaced by ValueIfZero. Zero
'             elements within top and left headers are unchanged.
' Arguments
' SourceFile: The full name of the source file, including the path. An array of file names is not
'             currently supported.
' TargetFile: The full name of the target file, including the path. An array of file names is not
'             currently supported.
' ValueIfZero: The value which replaces the zero elements in the file.
' Delimiter : The delimiter character. If omitted the function guesses the delimiter from the contents
'             of the first line of SourceFile.
' NumTopHeaders: The number of header rows at the top of the file. If omitted defaults to 1. Header rows
'             are copied unchanged to TargetFile.
' NumLeftHeaders: The number of header columns at the left of the file. If omitted defaults to 1. Header
'             columns are copied unchanged to TargetFile.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileSuppressZeros(SourceFile As String, TargetFile As String, ValueIfZero As Variant, Delimiter As String, Optional NumTopHeaders As Long = 1, Optional NumLeftHeaders As Long = 1)
Attribute sFileSuppressZeros.VB_Description = "Creates a copy of a text file in which zero elements are replaced by ValueIfZero. Zero elements within top and left headers are unchanged."
Attribute sFileSuppressZeros.VB_ProcData.VB_Invoke_Func = " \n26"

          Dim ErrNum As Long
          Dim F As Scripting.file
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim j As Long
          Dim RawLine As String
          Dim SplitLine() As String
          Dim strValueIfZero As String
          Dim Tin As TextStream
          Dim Tout As TextStream
          Const IsUnicode = False

1         On Error GoTo ErrHandler
          
2         If Not sFileExists(SourceFile) Then Throw "Cannot find file '" + SourceFile + "'"
3         If LCase$(SourceFile) = LCase$(TargetFile) Then Throw "SourceFile and TargetFile must be different"
          
4         If Delimiter = "" Then Throw "Delimiter must be provided"

5         If NumTopHeaders < 0 Then Throw "NumTopHeaders must not be negative"
6         If NumLeftHeaders < 0 Then Throw "NumLeftHeaders must not be negative"

7         If VarType(ValueIfZero) >= vbArray Then Throw "ValueIfZero cannot be an array"
8         On Error Resume Next
9         strValueIfZero = CStr(ValueIfZero)
10        ErrNum = Err.Number
11        On Error GoTo ErrHandler
12        If ErrNum <> 0 Then Throw "That ValueIfZero is not valid because it cannot be cast to a string"

13        If ValueIfZero = 0 Then
14            sFileSuppressZeros = sFileCopy(SourceFile, TargetFile)
15            Exit Function
16        End If

17        Set FSO = New FileSystemObject
18        Set F = FSO.GetFile(SourceFile)

19        Set Tin = F.OpenAsTextStream(ForReading, IIf(IsUnicode, TristateTrue, TristateFalse))
20        Set Tout = FSO.OpenTextFile(TargetFile, ForWriting, True, TristateFalse)

21        For i = 1 To NumTopHeaders
22            If Not Tin.atEndOfStream Then
23                RawLine = Tin.ReadLine
24                Tout.WriteLine RawLine
25            End If
26        Next

27        Do While Not Tin.atEndOfStream
28            RawLine = Tin.ReadLine
29            SplitLine = VBA.Split(RawLine, Delimiter)
30            On Error Resume Next
31            For j = NumLeftHeaders To UBound(SplitLine) 'Split returns zero-based array
32                If SplitLine(j) = "0" Then SplitLine(j) = strValueIfZero
33            Next j
34            On Error GoTo ErrHandler
35            Tout.WriteLine VBA.Join(SplitLine, Delimiter)
36        Loop

37        Tin.Close: Tout.Close: Set Tin = Nothing: Set Tout = Nothing: Set F = Nothing: Set FSO = Nothing

38        sFileSuppressZeros = TargetFile

39        Exit Function
ErrHandler:
40        Throw "#sFileSuppressZeros (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

