Attribute VB_Name = "modISDASIMME"
Option Explicit

'Work for 2021 ISDA

'Emulates EDATE, but handles arrays
Function sEDate(start_date As Variant, months As Variant)
Attribute sEDate.VB_Description = "Returns the serial number that represents the date that is the indicated number of months before or after a specified date (the start_date). Like Excel's EDATE but also works for array inputs."
Attribute sEDate.VB_ProcData.VB_Invoke_Func = " \n28"
1         On Error GoTo ErrHandler
2         If VarType(start_date) < vbArray And VarType(months) < vbArray Then
3             sEDate = SafeSubtract(start_date, months)
4         Else
5             sEDate = Broadcast(FuncIDEDate, start_date, months)
6         End If

7         Exit Function
ErrHandler:
8         sEDate = "#sEDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMM_HStackFiles
' Author     : Philip Swannell
' Date       : 02-May-2021
' Purpose    : "Stack two files side-by side" needed in 2021 project for Goldman EQ vega files, which were provided in
'              "two parts", function will not scale well  but should work for the case at hand.
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMM_HStackFiles(File1 As String, File2 As String, ResultFile As String)

          Dim File1Contents, File2Contents, File1LeftCol, File2LeftCol
          Dim File1Headers, File2Headers
          Dim ResultFileContents
          'Check left columns identical
1         File1Contents = ThrowIfError(sFileShow(File1))
2         File2Contents = ThrowIfError(sFileShow(File2))
3         File1LeftCol = sSubArray(File1Contents, 1, 1, , 1)
4         File2LeftCol = sSubArray(File2Contents, 1, 1, , 1)
5         If Not sArraysIdentical(File1LeftCol, File2LeftCol) Then
6             Throw "left columns of files not identical"
7         End If
          'check no headers appear in both files
8         File1Headers = sArrayTranspose(sSubArray(File1Contents, 1, 2, 1))
9         File2Headers = sArrayTranspose(sSubArray(File2Contents, 1, 2, 1))
10        If sNRows(sCompareTwoArrays(File1Headers, File2Headers, "Common")) > 1 Then
11            Throw "File1 headers and File2 headers are not disjoint"
12        End If
          
13        ResultFileContents = sArrayRange(File1Contents, sSubArray(File2Contents, 1, 2))
14        ISDASIMM_HStackFiles = sFileSaveCSV(ResultFile, ResultFileContents)

15        On Error GoTo ErrHandler
          
16        Exit Function
ErrHandler:
17        ISDASIMM_HStackFiles = "#ISDASIMM_HStackFiles (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMCompareIRWs
' Author     : Philip Swannell
' Date       : 24-May-2021
' Purpose    : For use from Analyse Data Series workbook to compare two sets of individual risk weights by looking at the
'              standard deviation of differences between the two sets
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMCompareIRWs(SheetName As String, Header1 As String, Header2 As String, Optional BookName As String)

          Dim wb As Workbook, ws As Worksheet
          Dim R As Range
          Const RangeName = "IvSDataWithHeaders"
          Dim Values1, Values2, ChooseVector, Differences

1         On Error GoTo ErrHandler
2         If BookName = "" Then
3             Set wb = Application.Caller.Parent.Parent
4         Else
5             If Not IsInCollection(Application.Workbooks, BookName) Then Throw "Workbook '" + BookName + "' is not open"
6             Set wb = Application.Workbooks(BookName)
7         End If

8         If Not IsInCollection(wb.Worksheets, SheetName) Then Throw "Workbook '" + BookName + "' does not have worksheet '" + SheetName + "'"
9         Set ws = wb.Worksheets(SheetName)

10        If Not IsInCollection(ws.Names, RangeName) Then Throw "Worksheet '" + SheetName + "' does not have range named '" + RangeName + "'"

11        Set R = ws.Range(RangeName)

12        Values1 = ThrowIfError(sColumnFromTable(R, Header1))
13        Values2 = ThrowIfError(sColumnFromTable(R, Header2))
14        ChooseVector = sArrayAnd(sArrayIsNumber(Values1), sArrayIsNumber(Values2))
15        Values1 = sMChoose(Values1, ChooseVector)
16        Values2 = sMChoose(Values2, ChooseVector)
17        Differences = sArraySubtract(Values1, Values2)
18        ISDASIMMCompareIRWs = Application.WorksheetFunction.StDev_P(Differences)

19        Exit Function
ErrHandler:
20        ISDASIMMCompareIRWs = "#ISDASIMMCompareIRWs (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

