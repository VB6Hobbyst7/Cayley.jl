Attribute VB_Name = "modMoreFunctions"
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCellInfo
' Author    : Philip Swannell
' Date      : 11-Nov-2015
' Purpose   : Returns information about a cell or range of cells.
' Arguments
' info_type : A string that can take one of the following values: Address, BookName, FileName, FontName,
'             FontSize, Formula, NumberFormat, SheetName.
' Reference : Optional. A cell or range of cells. If omitted, then information is returned about the
'             cell in which the formula resides.
' -----------------------------------------------------------------------------------------------------------------------
Function sCellInfo(info_type As String, Optional Reference As Range)
Attribute sCellInfo.VB_Description = "Returns information about a cell or range of cells."
Attribute sCellInfo.VB_ProcData.VB_Invoke_Func = " \n28"
          'When adding extra supported info_types don't forget to change the error message _
           returned in case Else and also to update the help on the Help sheet of this workbook.

1         On Error GoTo ErrHandler
2         Application.Volatile

3         If Reference Is Nothing Then Set Reference = Application.Caller

4         Select Case LCase$(info_type)
              Case "address"
5                 sCellInfo = Reference.address
6             Case "sheetname"
7                 sCellInfo = Reference.Worksheet.Name
8             Case "bookname"
9                 sCellInfo = Reference.Worksheet.Parent.Name
10            Case "filename"
11                sCellInfo = sFileMappedToUNC(Reference.Worksheet.Parent.FullName)
12            Case "formula"
13                sCellInfo = IfNull(Reference.Formula)
14            Case "numberformat"
15                sCellInfo = IfNull(Reference.NumberFormat)
16            Case "fontname"
17                sCellInfo = IfNull(Reference.Font.Name)
18            Case "fontsize"
19                sCellInfo = IfNull(Reference.Font.Size)
20            Case Else
21                sCellInfo = "#info_type not recognised. Allowed values are: Address, BookName, FileName, FontName, FontSize, Formula, NumberFormat, SheetName!"
22        End Select
23        Exit Function
ErrHandler:
24        sCellInfo = "#sCellInfo (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
Private Function IfNull(v As Variant, Optional ValueIfNull = "#Null!")
1         If IsNull(v) Then
2             IfNull = ValueIfNull
3         Else
4             IfNull = v
5         End If
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sEnvironmentVariable
' Author    : Philip Swannell
' Date      : 3-May-2015
' Purpose   : Returns the String associated with an operating system environment variable VariableName.
'             If VariableName is omitted, then the function returns a two-column list of
'             all operating system environment variables names and values.
'
' Arguments
' VariableName: The variable name for which an associated value is sought. For example "Path",
'             "Computername". If omitted a list of allowed VariableNames is returned. May
'             be an array.
' -----------------------------------------------------------------------------------------------------------------------
Function sEnvironmentVariable(Optional ByVal VariableName As Variant) As Variant
Attribute sEnvironmentVariable.VB_Description = "Returns the String associated with an operating system environment variable VariableName. If VariableName is omitted, then the function returns a two-column list of all operating system environment variables names and values.\n"
Attribute sEnvironmentVariable.VB_ProcData.VB_Invoke_Func = " \n28"
          Dim EqPos As Long
          Dim i As Long
          Dim STK As clsStacker
          Dim thisName As String
          Dim ThisPair
1         On Error GoTo ErrHandler
2         Set STK = CreateStacker()
3         If IsMissing(VariableName) Then
4             ThisPair = sArrayRange(vbNullString, vbNullString)
5             i = 1
6             Do While Len(Environ$(i)) > 0
7                 thisName = Environ$(i)
8                 EqPos = InStr(thisName, "=")
9                 If EqPos > 0 Then
10                    ThisPair(1, 1) = Left$(thisName, EqPos - 1)
11                    ThisPair(1, 2) = Mid$(thisName, EqPos + 1)
12                    STK.Stack2D ThisPair
13                End If
14                i = i + 1
15            Loop
16            sEnvironmentVariable = STK.Report
17        ElseIf VarType(VariableName) < vbArray Then
18            thisName = Environ$(CStr(VariableName))
19            If thisName = vbNullString Then
20                sEnvironmentVariable = "#Environment VariableName '" + VariableName + "' not found. Call sEnvironmentVariable with no arguments to get a list of environment VariableName names!"
21            Else
22                sEnvironmentVariable = thisName
23            End If
24        ElseIf VarType(VariableName) >= vbArray Then
              Dim j As Long
              Dim NC As Long
              Dim NR As Long
              Dim Result
25            Force2DArrayR VariableName, NR, NC
26            Result = sReshape(vbNullString, NR, NC)
27            For i = 1 To NR
28                For j = 1 To NC
29                    Result(i, j) = Environ$(CStr(VariableName(i, j)))
30                Next j
31            Next i
32            sEnvironmentVariable = Result
33        End If
34        Exit Function
ErrHandler:
35        sEnvironmentVariable = "#sEnvironmentVariable (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sBookName
' Author    : Philip Swannell
' Date      : 6-Jul-2015
' Purpose   : Returns the name of the workbook containing the range Reference
' Arguments
' Reference : A cell or range of cells. The return from the function is the name of the workbook that
'             contains this cell. If this argument is omitted, then the return from the
'             function is the name of the workbook in which the formula is entered.
' WithPath  : If FALSE or omitted the function returns the name of the workbook. If TRUE, then the
'             function returns the name of the workbook including the path.
' -----------------------------------------------------------------------------------------------------------------------
Function sBookName(Optional Reference As Range, Optional WithPath As Boolean, Optional LocalPathForOneDriveFiles As Boolean = True)
Attribute sBookName.VB_Description = "Returns the name of the workbook containing the range Reference."
Attribute sBookName.VB_ProcData.VB_Invoke_Func = " \n28"
1         Application.Volatile
2         On Error GoTo ErrHandler
3         If Reference Is Nothing Then Set Reference = Application.Caller
4         If WithPath Then
5             If LocalPathForOneDriveFiles Then
6                 sBookName = sFileMappedToUNC(LocalWorkbookName(Reference.Parent.Parent))
7             Else
8                 sBookName = sFileMappedToUNC(Reference.Parent.Parent.FullName)
9             End If
10        Else
11            sBookName = Reference.Parent.Parent.Name
12        End If
13        Exit Function
ErrHandler:
14        sBookName = "#sBookName (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileMappedToUNC
' Author    : Philip Swannell
' Date      : 09-Oct-2017
' Purpose   : Converts a file name given using a drive letter to a file name in the form
'             \\host-name\share-name\file-path.
' Arguments
' FileNames : Full name of a file, or an array of file names.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileMappedToUNC(FileNames As Variant)
Attribute sFileMappedToUNC.VB_Description = "Converts a file name given using a drive letter to a file name in the form \\\\host-name\\share-name\\file-path."
Attribute sFileMappedToUNC.VB_ProcData.VB_Invoke_Func = " \n26"
          Dim CachedRes As String
          Dim DoCaching As Boolean
          Dim First2 As String
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Res() As String

1         On Error GoTo ErrHandler
2         If VarType(FileNames) < vbArray Then
3             sFileMappedToUNC = MappedToUNC(CStr(FileNames))
4             Exit Function
5         End If

          'Quite common for all the files to be on the same drive, so cache the result of the first file processed
6         Force2DArrayR FileNames, NR, NC
7         ReDim Res(1 To NR, 1 To NC)

8         For i = 1 To NR
9             For j = 1 To NC
10                If i = 1 And j = 1 Then
11                    Res(i, j) = MappedToUNC(CStr(FileNames(i, j)))
12                    If Len(Res(i, j)) <> Len(FileNames(i, j)) Then
13                        DoCaching = True
14                        First2 = Left$(FileNames(i, j), 2)
15                        CachedRes = Left$(Res(i, j), Len(Res(i, j)) - Len(FileNames(i, j)) + 2)
16                    End If
17                Else
18                    If DoCaching And Left$(FileNames(i, j), 2) = First2 Then
19                        Res(i, j) = CachedRes + Mid$(FileNames(i, j), 3)
20                    Else
21                        Res(i, j) = MappedToUNC(CStr(FileNames(i, j)))
22                    End If
23                End If
24            Next j
25        Next i

26        sFileMappedToUNC = Res

27        Exit Function
ErrHandler:
28        sFileMappedToUNC = "#sFileMappedToUNC (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sColumnFromTable
' Author    : Philip Swannell
' Date      : 20-Oct-2016
' Purpose   : Returns a column (without header) from a table with headers. The column returned is that
'             whose header matches Header.
' Arguments
' Table     : A range whose top row is header labels, or (for use from VBA) a ListObject (aka Table) or
'             an array. The return from the function is a Range in the first two cases and
'             an array in the third case.
' Header    : The header string to match.
' -----------------------------------------------------------------------------------------------------------------------
Function sColumnFromTable(Table As Variant, header As String) As Variant
Attribute sColumnFromTable.VB_Description = "Returns a column (without header) from a table with headers. The column returned is that whose header matches Header."
Attribute sColumnFromTable.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim ColNo As Variant
          Dim CopyOfErr As String
          Dim ThrowErrors As Boolean
1         On Error GoTo ErrHandler

2         Select Case TypeName(Table)
              Case "Range"
3                 ColNo = sMatch(header, sArrayTranspose(Table.Rows(1).Value))
4                 If Not IsNumber(ColNo) Then
5                     Throw "Cannot find header '" + header + "' in top row of range " + Replace(Table.address, "$", vbNullString) + " on sheet " + Table.Parent.Name
6                 Else
7                     With Table
8                         Set sColumnFromTable = .Columns(ColNo).Offset(1).Resize(.Rows.Count - 1)
9                     End With
10                End If
11            Case "ListObject"
12                ThrowErrors = True        ' we can't be being called from a sheet
13                ColNo = sMatch(header, sArrayTranspose(Table.HeaderRowRange.Value))
14                If Not IsNumber(ColNo) Then
15                    Throw "Cannot find header '" + header + "' in range " + AddressND(Table.HeaderRowRange) + " on sheet " + Table.Parent.Name
16                Else
17                    Set sColumnFromTable = Table.DataBodyRange.Columns(ColNo)
18                End If
19            Case Else
20                ColNo = sMatch(header, sArrayTranspose(sSubArray(Table, 1, 1, 1)))
21                If Not IsNumber(ColNo) Then
22                    Throw "Cannot find header '" + header + "' in top row of Table"
23                Else
24                    sColumnFromTable = sSubArray(Table, 2, ColNo, , 1)
25                End If
26        End Select

27        Exit Function
ErrHandler:
28        CopyOfErr = "#sColumnFromTable (line " & CStr(Erl) + "): " & Err.Description & "!"
29        If ThrowErrors Then
30            Throw CopyOfErr
31        Else
32            sColumnFromTable = CopyOfErr
33        End If
End Function

Function sColumnsFromTable(ByVal Table As Variant, ByVal Headers As Variant, Optional WithTopRow As Boolean)
Attribute sColumnsFromTable.VB_Description = "Returns an array from a table with headers. The columns returned are those whose headers match the elements of Headers."
Attribute sColumnsFromTable.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim TableHeaders
          Dim i As Long
          Dim ColNumbers, RowNumbers

1         On Error GoTo ErrHandler
2         Force2DArrayRMulti Table, Headers
3         If sNRows(Table) <= 1 Then Throw "Table must have at least two rows, with the top row being headers"

4         If sNCols(Headers) <> 1 Then
5             Headers = sReshape(Headers, sNRows(Headers) * sNCols(Headers), 1)
6         End If
7         TableHeaders = sArrayTranspose(sSubArray(Table, 1, 1, 1))
8         ColNumbers = sArrayTranspose(sMatch(Headers, TableHeaders))
9         For i = 1 To sNCols(ColNumbers)
10            If Not (IsNumber(ColNumbers(1, i))) Then
11                Throw "Cannot find header '" + Headers(i, 1) + "' in top row of Table"
12            End If
13        Next
14        If WithTopRow Then
15            RowNumbers = CreateMissing()
16        Else
17            RowNumbers = sGrid(2, sNRows(Table), sNRows(Table) - 1)
18        End If
19        sColumnsFromTable = sIndex(Table, RowNumbers, ColNumbers)

20        Exit Function
ErrHandler:
21        sColumnsFromTable = "#sColumnsFromTable (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
