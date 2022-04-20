Attribute VB_Name = "modTradeBackups"
'---------------------------------------------------------------------------------------
' Module    : modTradeBackups
' Author    : Philip Swannell
' Date      : 13-Jan-2017
' Purpose   : Code relating to automatic saving of trades to a back up directory
'---------------------------------------------------------------------------------------
Option Explicit

Public g_LastTradeBackUpTime As Double

'---------------------------------------------------------------------------------------
' Procedure : BackUpDirectory
' Author    : Philip Swannell
' Date      : 13-Jan-2017
' Purpose   : Returns location of directory to save backups to
'---------------------------------------------------------------------------------------
Function BackUpDirectory()
1         On Error GoTo ErrHandler
2         BackUpDirectory = Environ("temp")
3         If Right(BackUpDirectory, 1) <> "\" Then BackUpDirectory = BackUpDirectory + "\"
4         BackUpDirectory = BackUpDirectory + gProjectName & "TradeBackups\"
5         ThrowIfError sCreateFolder(BackUpDirectory)
6         Exit Function
ErrHandler:
7         Throw "#BackUpDirectory (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'---------------------------------------------------------------------------------------
' Procedure : BackUpTrades
' Author    : Philip Swannell
' Date      : 13-Jan-2017
' Purpose   : Backs up trades on the portfolio sheet.
'---------------------------------------------------------------------------------------
Sub BackUpTrades(Optional EvenWhenZeroTrades As Boolean, Optional ByRef FileName As String)
1         On Error GoTo ErrHandler
          Dim NumTrades As Long
          Dim R As Range
          Static LastDataSaved As Variant, LastFileNameSaved As String
          Dim DoSave As Boolean
          Dim ThisDataSaved As Variant

2         Set R = getTradesRange(NumTrades, False)
3         If NumTrades > 0 Or EvenWhenZeroTrades Then
4             DoSave = True
5             FileName = BackUpDirectory + "Trades " + Format(Now(), " yyyy-mm-dd hh-mm-ss") + "(" + CStr(NumTrades) + ").stf"

              'If the data saved is the same as the previous backup, rename the previous backup to reflect current time
6             ThisDataSaved = R.Value2
7             If LastFileNameSaved <> "" Then
8                 If sArraysIdentical(ThisDataSaved, LastDataSaved) Then
9                     If sFileExists(LastFileNameSaved) Then
10                        sFileRename LastFileNameSaved, FileName    'does not throw an error if fails
11                        If sFileExists(FileName) Then
12                            DoSave = False
13                        End If
14                    End If
15                End If
16            End If
17            If DoSave Then
18                SaveTradesFile FileName, False, False, False, EvenWhenZeroTrades
19            End If

20            LastDataSaved = ThisDataSaved
21            LastFileNameSaved = FileName
22            g_LastTradeBackUpTime = Now()
23            CleanUpBackUps
24        End If
25        Exit Sub
ErrHandler:
26        Throw "#BackUpTrades (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CleanUpBackUps
' Author    : Philip Swannell
' Date      : 13-Jan-2017
' Purpose   : Removes backups. Strategy is to detect if older backups duplicate a younger
'             backup if so delete older. Then keep no more than NumToKeep. Note that we use
'             the "checksum" in the file to determine duplicates and checksums only exist
'             for files saved with version of SCRiPT dated 6-Feb-2019 or later.
'---------------------------------------------------------------------------------------
Sub CleanUpBackUps()

          Dim anyDuplicates As Boolean
          Dim DirlistRet
          Dim FileNames
          Dim i As Long
          Dim MatchIDs
          Dim NumBackups
          Const NumToKeep = 40
          Dim SPH As clsScreenUpdateHandler
          Dim XSH As clsExcelStateHandler
          
1         Set SPH = CreateScreenUpdateHandler()
2         Set XSH = CreateExcelStateHandler(, , False)
          
3         On Error GoTo ErrHandler

4         DirlistRet = sDirList(BackUpDirectory, False, False, "FC", "F", "*Trades*.stf")
5         If sIsErrorString(DirlistRet) Then Exit Sub

6         DirlistRet = sSortedArray(DirlistRet, 2, , , False)
7         FileNames = sSubArray(DirlistRet, 1, 1, , 1)
8         NumBackups = sNRows(DirlistRet)
          Dim Hashes
9         Hashes = sReshape("", sNRows(DirlistRet), 1)
10        For i = 1 To sNRows(Hashes)
11            Hashes(i, 1) = TradeFileInfo(CStr(DirlistRet(i, 1)), "CheckSum")
12            If sIsErrorString(Hashes(i, 1)) Then Hashes(i, 1) = CStr(sElapsedTime)
13        Next i
14        MatchIDs = sMatch(Hashes, Hashes)
15        Force2DArray MatchIDs
          
          'Delete backups that duplicate a younger backup
16        For i = 1 To sNRows(DirlistRet)
17            If MatchIDs(i, 1) <> i Then
18                sFileDelete DirlistRet(i, 1)
19                anyDuplicates = True
20            End If
21        Next
22        If anyDuplicates Then
23            DirlistRet = sDirList(BackUpDirectory, False, False, "FC", "F", "*Trades*.stf")
24        End If

25        If sNRows(DirlistRet) > NumToKeep Then
26            For i = NumToKeep + 1 To NumBackups
27                sFileDelete DirlistRet(i, 1)
28            Next i
29        End If

30        Exit Sub
ErrHandler:
31        Throw "#CleanUpBackUps (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Sub unhandlederror()

          Dim foo
1         foo = "x" + 1

End Sub

'---------------------------------------------------------------------------------------
' Procedure : OpenBackUps
' Author    : Philip Swannell
' Date      : 13-Jan-2017
' Purpose   : Presents the user with a list of available backups, so that they can choose one to restore.
'---------------------------------------------------------------------------------------
Sub OpenBackUps()
          Dim ChooseVector
          Dim Chosen
          Dim DateDesc As String
          Dim Descriptions
          Dim DirlistRet
          Dim Explanations
          Dim FileNames
          Dim FileToOpen As String
          Dim FormatString As String
          Dim i As Long
          Dim NT As String

1         On Error GoTo ErrHandler
2         DirlistRet = sDirList(BackUpDirectory(), False, False, "FCN", "F", "*Trades*(*).stf")
3         If sIsErrorString(DirlistRet) Then Throw "No backups of trades found", True
4         DirlistRet = sSortedArray(DirlistRet, 2, , , False)
5         FileNames = sSubArray(DirlistRet, 1, 3, , 1)

6         ChooseVector = sIsRegMatch("^((?!\(0\)).)*$", FileNames) 'Does not contain "(0)" - i.e. a backup with no trades in it, as may be saved by the workbook close event, but not otherwise
7         If Not sColumnOr(ChooseVector)(1, 1) Then Throw "No backups of trades found", True
8         If Not sColumnAnd(ChooseVector)(1, 1) Then
9             DirlistRet = sMChoose(DirlistRet, ChooseVector)
10            FileNames = sMChoose(FileNames, ChooseVector)
11        End If

12        Descriptions = sReshape(0, sNRows(DirlistRet), 1)
13        Explanations = sReshape("", sNRows(DirlistRet), 1)
14        StatusBarWrap "Scanning Backup files"
15        For i = 1 To sNRows(DirlistRet)
16            NT = sStringBetweenStrings(DirlistRet(i, 3), "(", ")")
17            If Format(DirlistRet(i, 2), "dd-mm-yy") = Format(Date, "dd-mm-yy") Then
18                DateDesc = "today at "
19                FormatString = "hh:mm"
20            ElseIf Format(DirlistRet(i, 2), "dd-mm-yy") = Format(Date - 1, "dd-mm-yy") Then
21                DateDesc = "yesterday at "
22                FormatString = "hh:mm"
23            ElseIf Format(DirlistRet(i, 2), "yyyy") = Format(Date, "yyyy") Then
24                DateDesc = " "
25                FormatString = "d-mmm at hh:mm"
26            Else
27                DateDesc = " "
28                FormatString = "d-mmm-yy at hh:mm"
29            End If

30            Descriptions(i, 1) = Format(NT, "###,##0") + " trade" + IIf(NT = "1", "", "s") + " backed up " + DateDesc + _
                  Format(DirlistRet(i, 2), FormatString) + " " + _
                  sDescribeTime(Now() - DirlistRet(i, 2))
31            Descriptions(i, 1) = Replace(Descriptions(i, 1), "  ", " ")
                  
32            Explanations(i, 1) = TradeFileInfo(CStr(DirlistRet(i, 1)), "TradesSummary")
33        Next i
34        StatusBarWrap False

35        Chosen = ShowSingleChoiceDialog(Descriptions, Explanations, , , , "Open trade backup", "Select trade backup to restore")
36        If IsEmpty(Chosen) Then Exit Sub

37        FileToOpen = DirlistRet(sMatch(Chosen, Descriptions), 1)

38        OpenTradesFile FileToOpen, True
39        Exit Sub
ErrHandler:
40        Throw "#OpenBackUps (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' ----------------------------------------------------------------
' Procedure Name: OpenMostRecentBackup
' Purpose: Opens the most recently saved trade backup, for use from the workbook open
' Author: Philip Swannell
' Date: 01-Dec-2017
' ----------------------------------------------------------------
Sub OpenMostRecentBackup()
          Dim DirlistRet
          Dim FileToOpen As String
1         On Error GoTo ErrHandler
2         DirlistRet = sDirList(BackUpDirectory, False, False, "FCN", "F", "Trades*(*).stf")
3         If sIsErrorString(DirlistRet) Then Exit Sub
4         DirlistRet = sSortedArray(DirlistRet, 2, , , False)
5         FileToOpen = DirlistRet(1, 1)
6         If InStr(FileToOpen, "(0)") > 0 Then ' most recent backup contained no trades...
7             ClearPortfolioSheet
8         Else
9             OpenTradesFile FileToOpen, True, True

10        End If
11        Exit Sub
ErrHandler:
12        Throw "#OpenMostRecentBackup (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

