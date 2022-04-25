Attribute VB_Name = "modPrivateModuleA"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : modPrivateModuleA
' Author    : Philip Swannell
' Date      : 03-Dec-2015
' Purpose   : "Low-level" functions that we don't want to expose directly to Excel, but
'             are used from functions in more than one module, so can't be made Private.
'             Instead put them in this Option Private Module
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Option Private Module
Private m_OSV_OptStyle As EnmOptStyle
Private m_OSV_Value As Double
Private m_OSV_Forward As Double
Private m_OSV_Strike As Double
Private m_OSV_Time As Double
Private m_OSV_logNormal As Boolean

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As Long

Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindow Lib "USER32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As Long

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Const GW_HWNDNEXT = 2

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FunctionWizardActive
' Author     : Philip Swannell
' Date       : 13-Jul-2021
' Purpose    : Test if the Fn wizard is active to allow early exit in slow functions.
' Posted to: https://stackoverflow.com/questions/20866484/can-i-disable-a-vba-udf-calculation-when-the-insert-function-function-arguments
' -----------------------------------------------------------------------------------------------------------------------
Function FunctionWizardActive() As Boolean

          Dim ExcelPID As Long
          Dim lhWndP As LongPtr
          Dim WindowTitle As String
          Dim WindowPID As Long
          Const FunctionWizardCaption = "Function Arguments" 'This won't work for non English-language Excel
          
1         On Error GoTo ErrHandler
2         If TypeName(Application.Caller) = "Range" Then
              'The "CommandBars test" below is usually sufficient to determine that the Function Wizard is active,
              'but can sometimes give a false positive. Example: When a csv file is opened (via File Open) then all
              'active workbooks are calculated (even if calculation is set to manual!) with
              'Application.CommandBars("Standard").Controls(1).Enabled being False.
              'So apply a further test using Windows API to loop over all windows checking for a window with title
              '"Function  Arguments", checking also the process id.
3             If Not Application.CommandBars("Standard").Controls(1).Enabled Then
4                 ExcelPID = GetCurrentProcessId()
5                 lhWndP = FindWindow(vbNullString, vbNullString) 'PARENT WINDOW
6                 Do While lhWndP <> 0
7                     WindowTitle = String(GetWindowTextLength(lhWndP) + 1, Chr$(0))
8                     GetWindowText lhWndP, WindowTitle, Len(WindowTitle)
9                     WindowTitle = Left$(WindowTitle, Len(WindowTitle) - 1)
10                    If WindowTitle = FunctionWizardCaption Then
11                        GetWindowThreadProcessId lhWndP, WindowPID
12                        If WindowPID = ExcelPID Then
13                            FunctionWizardActive = True
14                            Exit Function
15                        End If
16                    End If
17                    lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
18                Loop
19            End If
20        End If

21        Exit Function
ErrHandler:
22        Throw "#FunctionWizardActive (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AddressND
' Author    : Philip
' Date      : 09-Oct-2017
' Purpose   : Address of a range without those pesky $s
' -----------------------------------------------------------------------------------------------------------------------
Function AddressND(R As Range)
1         AddressND = Replace(R.address, "$", vbNullString)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sCreateFolder
' Author    : Philip Swannell
' Date      : 29-Jun-2018
' Purpose   : Creates a folder on disk. FolderPath can be passed in as C:\This\That\TheOther even if the
'             folder C:\This does not yet exist. If successful returns the name of the
'             folder. If not successful returns an error string.
' Arguments
' FolderPath: Path of the folder to be created. For example C:\temp\My_New_Folder. It does not matter if
'             this path has a terminating backslash or not. This argument may be an array.
' -----------------------------------------------------------------------------------------------------------------------
Function CoreCreateFolder(ByVal FolderPath As String)
          Dim F As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim isUNC As Boolean
          Dim ParentFolderName

1         On Error GoTo ErrHandler

2         If Left$(FolderPath, 2) = "\\" Then
3             isUNC = True
4         ElseIf Mid$(FolderPath, 2, 2) <> ":\" Or Asc(UCase$(Left$(FolderPath, 1))) < 65 Or Asc(UCase$(Left$(FolderPath, 1))) > 90 Then
5             Throw "First three characters of FolderPath must give drive letter followed by "":\"" or else be""\\"" for " & _
                  "UNC folder name"
6         End If

7         FolderPath = Replace(FolderPath, "/", "\")

8         If Right$(FolderPath, 1) <> "\" Then
9             FolderPath = FolderPath + "\"
10        End If

11        Set FSO = New FileSystemObject
12        If CoreFolderExists(FolderPath) Then
13            GoTo EarlyExit
14        End If

          'Go back until we find parent folder that does exist
15        For i = Len(FolderPath) - 1 To 3 Step -1
16            If Mid$(FolderPath, i, 1) = "\" Then
17                If CoreFolderExists(Left$(FolderPath, i)) Then
18                    Set F = FSO.GetFolder(Left$(FolderPath, i))
19                    ParentFolderName = Left$(FolderPath, i)
20                    Exit For
21                End If
22            End If
23        Next i

24        If F Is Nothing Then Throw "Cannot create folder " + Left$(FolderPath, 3)

          'now add folders one level at a time
25        For i = Len(ParentFolderName) + 1 To Len(FolderPath)
26            If Mid$(FolderPath, i, 1) = "\" Then
                  Dim ThisFolderName As String
27                ThisFolderName = Mid$(FolderPath, InStrRev(FolderPath, "\", i - 1) + 1, i - 1 - InStrRev(FolderPath, "\", i - 1))
28                F.SubFolders.Add ThisFolderName
29                Set F = FSO.GetFolder(Left$(FolderPath, i))
30            End If
31        Next i

EarlyExit:
32        Set F = FSO.GetFolder(FolderPath)
33        CoreCreateFolder = F.Path
34        Set F = Nothing: Set FSO = Nothing

35        Exit Function
ErrHandler:
36        CoreCreateFolder = "#CoreCreateFolder (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreDeleteFolder
' Author    : Philip
' Date      : 13-Jun-2017
' Purpose   : Delete a folder and all sub-folders and files. Use with care!
' -----------------------------------------------------------------------------------------------------------------------
Function CoreDeleteFolder(FolderPath As String)
          Dim fo As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set fo = FSO.GetFolder(FolderPath)
4         fo.Delete True
5         Set fo = Nothing: Set FSO = Nothing

6         Exit Function
ErrHandler:
7         CoreDeleteFolder = "#CoreDeleteFolder (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFileCopy
' Author    : Philip Swannell
' Date      : 14-Dec-2015
' Purpose   : Wrapped by sFileCopy. Handles URL for Source file, but not for target file.
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileCopy(SourceFile As String, TargetFile As String, Optional OverwriteReadOnlyFiles As Boolean)
          Dim SF As Scripting.file
          Dim FSO As Scripting.FileSystemObject
          Dim EN As Long
          Dim TF As Scripting.file
          Dim TargetExists As Boolean
          Dim TargetIsReadOnly
                        
1         On Error GoTo ErrHandler
2         CheckFileNameIsAbsolute TargetFile
3         If LCase$(Left$(SourceFile, 4)) <> "http" Then
4             CheckFileNameIsAbsolute SourceFile
5         End If

6         Set FSO = New FileSystemObject

7         On Error Resume Next
8         Set TF = FSO.GetFile(TargetFile)
9         EN = Err.Number
10        On Error GoTo ErrHandler
11        TargetExists = EN = 0
12        If TargetExists Then
13            TargetIsReadOnly = TF.Attributes And 1
14        End If
              
15        If TargetIsReadOnly Then
16            If OverwriteReadOnlyFiles Then
17                TF.Attributes = TF.Attributes - 1
18            Else
19                Throw "Copy failed because TargetFile '" & TargetFile & "' already exists and is read only. Consider setting argument OverwriteReadOnlyFiles to True"
20            End If
21        End If

22        If LCase$(Left$(SourceFile, 4)) = "http" Then
23            CoreFileCopy = CoreURLDownloadToFile(SourceFile, TargetFile)
24        Else
25            Set SF = FSO.GetFile(SourceFile)

26            SF.Copy TargetFile, True
27            CoreFileCopy = TargetFile
28            Set FSO = Nothing: Set SF = Nothing
29        End If
30        Exit Function
ErrHandler:
31        CoreFileCopy = "#" & Err.Description & "!"
32        Set FSO = Nothing: Set SF = Nothing
End Function


'http://www.vbforums.com/showthread.php?412514-Binary-copy-file
' but with bug fixes...
Function BinaryCopyFile(Source As String, Dest As String)
          Dim intFF1 As Integer, intFF2 As Integer
          Dim lngFilesize As Long
          Const CHUNK_SIZE = 8192
          Dim Buffer(0 To CHUNK_SIZE - 1) As Byte 'This line amended
          Dim BufferRemain() As Byte
1         On Error GoTo ErrHandler
2         intFF1 = FreeFile
3         Open Source For Binary Access Read As #intFF1
4         intFF2 = FreeFile
5         Open Dest For Binary Access Write As #intFF2
6         lngFilesize = LOF(intFF1)
7         Do While lngFilesize >= CHUNK_SIZE
8             Get #intFF1, , Buffer
9             Put #intFF2, , Buffer
10            lngFilesize = lngFilesize - CHUNK_SIZE
11        Loop
12        If lngFilesize >= 1 Then 'This line amended
13            ReDim BufferRemain(0 To lngFilesize - 1)
14            Get #intFF1, , BufferRemain
15            Put #intFF2, , BufferRemain
16        End If
17        Close #intFF1
18        Close #intFF2
19        BinaryCopyFile = True

20        Exit Function
ErrHandler:
21        Throw "#BinaryCopyFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function CoreFileUnblock(SourceFile As String)
          Dim OldCheckSum As String
          Dim NewCheckSum As String
          Dim TempFile As String

1         On Error GoTo ErrHandler

TryAgain:
2         TempFile = Environ$("Temp") & "\" & Replace(CStr(sElapsedTime), ".", "") & ".tmp"
3         If sFileExists(TempFile) Then GoTo TryAgain
4         BinaryCopyFile SourceFile, TempFile
5         OldCheckSum = ThrowIfError(sFileCheckSum(SourceFile)) 'This is quite "belt-and-braces" since not sure if I trust method BinaryFileCopy!
6         NewCheckSum = ThrowIfError(sFileCheckSum(TempFile))
7         If OldCheckSum <> NewCheckSum Then Throw "Assertion failed. Temporary file does not have the same check sum as SourceFile"
8         ThrowIfError CoreFileDelete(SourceFile)
9         ThrowIfError CoreFileMove(TempFile, SourceFile)
10        CoreFileUnblock = SourceFile

11        Exit Function
ErrHandler:
12        CoreFileUnblock = "#CoreFileUnblock (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFileDelete
' Author    : Philip Swannell
' Date      : 14-Dec-2015
' Purpose   : wrapped by sFileDelete for array-processing
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileDelete(FileName As String) As Variant
          Dim F As Scripting.file
          Dim FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         CheckFileNameIsAbsolute FileName
3         Set FSO = New FileSystemObject
4         Set F = FSO.GetFile(FileName)
5         F.Delete
6         CoreFileDelete = True
7         Set F = Nothing: Set FSO = Nothing
8         Exit Function
ErrHandler:
9         CoreFileDelete = "#" + Err.Description + "!"
End Function

Function CoreFileExif(FileName As String, PropertyNumber As Long)
          Dim GoodNumber As Boolean
          Dim intPos As Long
          Dim strName As String
          Dim ThisPath As Variant
          Static LastFileName As String
          Static LastPath As String
          Static objFolder As Object
          Static objFolderItem As Object
          Static objShell As Object

1         On Error GoTo ErrHandler

2         If IsNumber(PropertyNumber) Then If PropertyNumber >= -1 Then If PropertyNumber <= 312 Then If PropertyNumber = CLng(PropertyNumber) Then GoodNumber = True
3         If Not GoodNumber Then
4             CoreFileExif = "#Illegal PropertyNumber. Call CoreFileExif with no arguments for a table of allowed PropertyNumbers!"
5             Exit Function
6         End If

7         intPos = InStrRev(FileName, "\")
8         ThisPath = Left$(FileName, intPos)
9         strName = Mid$(FileName, intPos + 1)
10        If objShell Is Nothing Then
11            Set objShell = CreateObject("Shell.Application")
12        End If
13        If ThisPath <> LastPath Then
14            Set objFolder = objShell.Namespace(ThisPath)
15        End If
16        If FileName <> LastFileName Then
17            Set objFolderItem = objFolder.ParseName(strName)
18        End If
19        If Not objFolderItem Is Nothing Then
20            CoreFileExif = objFolder.GetDetailsOf(objFolderItem, PropertyNumber)
21            If VarType(CoreFileExif) = vbString Then
                  'Strange - was finding that return for e.g. "Dimensions" included non-printing wide characters
22                If PropertyNumber <> 0 Then 'PropertyNum 0 is file name, which can legitimately contain characters > 255
23                    CoreFileExif = StripNonAscii(CStr(CoreFileExif))
24                End If
25            End If
              '  CoreFileExif = objFolder.GetDetailsOf(Null, n) ' pass null and the function returns the name of the number n e.g. 13 -> "Contributing artists"
26        End If
27        LastPath = ThisPath
28        LastFileName = FileName

29        Exit Function
ErrHandler:
30        CoreFileExif = "#CoreFileExif (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFileExists
' Author    : Philip Swannell
' Date      : 14-Dec-2015
' Purpose   : Wrapped by sFileExists for multi-call
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileExists(FileName As String) As Variant
          Dim F As Scripting.file
          Static FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         If FSO Is Nothing Then Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(FileName)
4         CoreFileExists = True
5         Exit Function
ErrHandler:
6         CoreFileExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CoreFileInfo
' Author     : Philip Swannell
' Date       : 01-May-2018
' Purpose    : Wrapped by sFileInfo
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileInfo(FileName As String, Info As String)
          Static FSO As Scripting.FileSystemObject
          Dim F As Scripting.file
1         On Error GoTo ErrHandler
2         If FSO Is Nothing Then Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(FileName)
4         Select Case LCase$(Info)
              Case "attributes", "attrib"
5                 CoreFileInfo = DecodeFileAttributes(F.Attributes)
6             Case "datecreated", "c", "createddate"
7                 CoreFileInfo = F.DateCreated
8             Case "datelastaccessed", "a", "lastaccesseddate"
9                 CoreFileInfo = F.DateLastAccessed
10            Case "datelastmodified", "m", "lastmodifieddate"
11                CoreFileInfo = F.DateLastModified
12            Case "drive", "d"
13                CoreFileInfo = F.Drive
14            Case "md5", "#"
                  Static Hasher As Object
                  Static ZeroByteHashMD5 As String
15                If Hasher Is Nothing Then Set Hasher = CreateHasher(True, ZeroByteHashMD5)
                  Dim TheHash As String
16                HashFromFileName Hasher, F.Path, TheHash, ZeroByteHashMD5
17                CoreFileInfo = TheHash
18            Case "sha1"
                  Static Hasher2 As Object
                  Static ZeroByteHashSHA1 As String
19                If Hasher2 Is Nothing Then Set Hasher2 = CreateHasher(False, ZeroByteHashSHA1)
20                HashFromFileName Hasher2, F.Path, TheHash, ZeroByteHashSHA1
21                CoreFileInfo = TheHash
22            Case "name", "n"
23                CoreFileInfo = F.Name
24            Case "numlines"
25                CoreFileInfo = CoreFileNumLines(F.Path)
26            Case "parentfolder"
27                CoreFileInfo = F.ParentFolder
28            Case "path", "fullname", "f", "fullfilename"
29                CoreFileInfo = F.Path
30            Case "shortname"
31                CoreFileInfo = F.ShortName
32            Case "shortpath"
33                CoreFileInfo = F.ShortPath
34            Case "size", "s"
35                CoreFileInfo = F.Size
36            Case "type", "t"
37                CoreFileInfo = F.Type
38            Case "uncname"
39                If Mid$(FileName, 2, 1) = ":" Then
                      Dim strDrive As String
                      Dim strShare As String
40                    strDrive = FSO.GetDriveName(FileName)
41                    strShare = FSO.Drives(strDrive & "\").ShareName
42                    If Len(strShare) = 0 Then
43                        CoreFileInfo = FileName
44                    Else
45                        CoreFileInfo = strShare & Mid$(FileName, 3)
46                    End If
47                Else
48                    CoreFileInfo = FileName
49                End If
50            Case Else
51                CoreFileInfo = "Info not recognised. Must be one of 'Attributes', 'DateCreated' or 'C', 'DateLastAccessed' or 'A', 'DateLastModified' or 'M', 'Drive', 'FullName' or 'F', 'MD5', 'SHA1', 'Name' or 'N', 'NumLines', 'ParentFolder', 'ShortName', 'ShortPath', 'Size' or 'S', 'Type' or 'T', 'UNCName'"
52        End Select

53        Exit Function
ErrHandler:
54        CoreFileInfo = "#" + Err.Description + "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFileLastModifiedDate
' Author    : Philip Swannell
' Date      : 14-Dec-2015
' Purpose   : Wrapped by sFileLastModifiedDate
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileLastModifiedDate(FileName As String)
          Dim F As Scripting.file
          Dim FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(FileName)
4         CoreFileLastModifiedDate = F.DateLastModified
5         Set FSO = Nothing: Set F = Nothing
6         Exit Function
ErrHandler:
7         CoreFileLastModifiedDate = "#" + Err.Description + "!"
8         Set FSO = Nothing: Set F = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFileMove
' Author    : Philip Swannell
' Date      : 01-Mar-2016
' Purpose   : Wrapped by sMoveFile. NB overwrites target and creates any necessary folders
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileMove(FromFile As String, ToFile As String)
          Dim ErrDesc As String
          Dim ErrNumber As Long
          Dim FSO As Scripting.FileSystemObject

          Dim F As Scripting.file
1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(FromFile)
4         CheckFileNameIsAbsolute ToFile

5         On Error Resume Next
6         F.Move ToFile
7         ErrNumber = Err.Number: ErrDesc = Err.Description
8         On Error GoTo ErrHandler
9         If ErrNumber = 76 Then        'Path not found for ToFile
10            ThrowIfError sCreateFolder(sSplitPath(ToFile, False))
11            F.Move ToFile
12        ElseIf ErrNumber = 58 Then        'ToFile already exists
13            ThrowIfError sFileDelete(ToFile)
14            F.Move ToFile
15        ElseIf ErrNumber <> 0 Then
16            Throw ErrDesc
17        End If

18        CoreFileMove = True
19        Set FSO = Nothing: Set F = Nothing
20        Exit Function
ErrHandler:
21        CoreFileMove = "#" + Err.Description + "!"
22        Set FSO = Nothing: Set F = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CoreFileNumLines
' Author     : Philip Swannell
' Date       : 29-Jun-2018
' Purpose    : Wrapped by sFileNumLines
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileNumLines(FileName)
          Dim F As Scripting.file
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim t As TextStream

1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(FileName)
4         If F.Size = 0 Then
5             CoreFileNumLines = 0
6         Else
7             Set t = F.OpenAsTextStream(ForReading)
8             Do While Not t.atEndOfStream
9                 i = i + 1
10                t.ReadLine
11            Loop
12            CoreFileNumLines = i
13            t.Close: Set t = Nothing
14        End If
15        Set F = Nothing: Set FSO = Nothing
16        Exit Function
ErrHandler:
17        CoreFileNumLines = "#" & Err.Description & "!"
End Function

Function CoreFileCopySkip(SourceFile, TargetFile, NumLinesToSkip)
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim Unicode As Boolean
          
          Dim sts As Scripting.TextStream
          Dim tts As Scripting.TextStream
1         On Error GoTo ErrHandler

2         If VarType(SourceFile) <> vbString Then Throw "SourceFile must be a string"
3         If VarType(TargetFile) <> vbString Then Throw "TargetFile must be a string"
          'not a very sophisticated check on whether they point to the same file...
4         If LCase(SourceFile) = LCase(TargetFile) Then Throw "SourceFile and TargetFile must be different"
5         If Not IsNumber(NumLinesToSkip) Then Throw "NumLinesToSkip must be a whole number"
6         If NumLinesToSkip <> CLng(NumLinesToSkip) Then Throw "NumLinesToSkip must be a whole number"
7         If NumLinesToSkip < 0 Then Throw "NumLinesToSkip must be zero or positive"
8         Unicode = ThrowIfError(IsUnicodeFile(CStr(SourceFile)))

9         If NumLinesToSkip = 0 Then
10            CoreFileCopySkip = CoreFileCopy(CStr(SourceFile), CStr(TargetFile))
11            Exit Function
12        End If

13        CheckFileNameIsAbsolute SourceFile
14        CheckFileNameIsAbsolute TargetFile

15        Set FSO = New FileSystemObject

16        Set sts = FSO.OpenTextFile(SourceFile, ForReading, , IIf(Unicode, TristateTrue, TristateFalse))
17        Set tts = FSO.CreateTextFile(TargetFile, , Unicode)

18        For i = 1 To NumLinesToSkip
19            If sts.atEndOfStream Then Throw "Cannot skip that many lines"
20            sts.SkipLine
21        Next
22        While Not sts.atEndOfStream
23            tts.WriteLine sts.ReadLine
24        Wend
25        sts.Close
26        tts.Close
27        CoreFileCopySkip = TargetFile

28        Exit Function
ErrHandler:
29        CoreFileCopySkip = "#" & Err.Description & "!"
30        If Not sts Is Nothing Then sts.Close
31        If Not tts Is Nothing Then tts.Close
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFileRegExReplace
' Author    : Philip Swannell
' Date      : 13-Jan-2016
' Purpose   : Do regular expression replacement in a file...
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileRegExReplace(SourceFile As String, TargetFile As String, RegularExpression As String, Replacement As String, CaseSensitive As Boolean)

          Dim FileContents As String
          Dim FSO As Scripting.FileSystemObject
          Dim t1 As TextStream
          Dim t2 As TextStream

1         On Error GoTo ErrHandler
2         CheckFileNameIsAbsolute SourceFile
3         CheckFileNameIsAbsolute TargetFile

4         If Not RegExSyntaxValid(RegularExpression) Then
5             CoreFileRegExReplace = "#Invalid syntax for RegularExpression!"
6             Exit Function
7         End If

8         Set FSO = New FileSystemObject
9         Set t1 = FSO.OpenTextFile(SourceFile, ForReading)
10        FileContents = t1.ReadAll
11        FileContents = sRegExReplace(FileContents, RegularExpression, Replacement, CaseSensitive)
12        t1.Close
13        Set t2 = FSO.CreateTextFile(TargetFile, True)
14        t2.Write FileContents
15        t2.Close
16        CoreFileRegExReplace = TargetFile
17        Set FSO = Nothing
18        Set t1 = Nothing
19        Set t2 = Nothing
20        Exit Function
ErrHandler:
21        CoreFileRegExReplace = "#CoreFileRegExReplace (line " & CStr(Erl) + "): " & Err.Description & "!"
22        If Not t1 Is Nothing Then
23            t1.Close
24            Set t1 = Nothing
25        End If

26        If Not t2 Is Nothing Then
27            t2.Close
28            Set t2 = Nothing
29        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CoreFolderRename
' Author     : Philip Swannell
' Date       : 12-Jun-2019
' Purpose    : Wrapped by sFolderRename
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFolderRename(ByVal OldFolderPath As String, NewFolderName As String)
          Dim F As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         Do While Right$(OldFolderPath, 1) = "\"
3             OldFolderPath = Left$(OldFolderPath, Len(OldFolderPath) - 1)
4         Loop
5         If InStr(NewFolderName, "\") > 0 Then Throw "NewFolderName cannot contain backslash character"
6         Set FSO = New FileSystemObject
7         Set F = FSO.GetFolder(OldFolderPath)
8         F.Name = NewFolderName
9         CoreFolderRename = F.Path
10        Exit Function
ErrHandler:
11        CoreFolderRename = "#CoreFolderRename (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFileRename
' Author    : Philip Swannell
' Date      : 01-Mar-2016
' Purpose   : Wrapped by sFileRename.
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFileRename(OldFileName As String, NewFileName As String)
          Dim FSO As Scripting.FileSystemObject

          Dim F As Scripting.file
1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(OldFileName)

4         If InStr(NewFileName, "\") > 0 Then
5             If LCase$(sSplitPath(OldFileName, False)) <> LCase$(sSplitPath(NewFileName, False)) Then
6                 Throw "Cannot rename file to a different folder"
7             End If
8         End If

9         F.Name = sSplitPath(NewFileName, True)

10        CoreFileRename = F.Path
11        Set FSO = Nothing: Set F = Nothing
12        Exit Function
ErrHandler:
13        CoreFileRename = "#" + Err.Description + "!"
14        Set FSO = Nothing: Set F = Nothing
End Function

Function CoreFolderCopy(ByVal SourceFolder As String, ByVal TargetFolder As String)
          Dim fl As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         If Right$(SourceFolder, 1) = "\" Then SourceFolder = Left$(SourceFolder, Len(SourceFolder) - 1)
3         If Right$(TargetFolder, 1) = "\" Then TargetFolder = Left$(TargetFolder, Len(TargetFolder) - 1)
4         ThrowIfError sCreateFolder(TargetFolder)        'since otherwise the parent folder of target folder must already exist
5         Set FSO = New FileSystemObject
6         Set fl = FSO.GetFolder(SourceFolder)
7         fl.Copy TargetFolder
8         CoreFolderCopy = True
9         Set FSO = Nothing: Set fl = Nothing

10        Exit Function
ErrHandler:
11        CoreFolderCopy = "#" + Err.Description + "!"
12        Set FSO = Nothing: Set fl = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFolderExists
' Author    : Philip Swannell
' Date      : 07-Oct-2013
' Purpose   : Returns True or False. Does not matter if FolderPath has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFolderExists(ByVal FolderPath As String)
          Dim F As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFolder(FolderPath)
4         CoreFolderExists = True
5         Exit Function
ErrHandler:
6         CoreFolderExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CoreFolderIsWritable
' Author     : Philip Swannell
' Date       : 25-Jun-2019
' Purpose    : Returns true if a folder exists and can be written to, false otherwise. Uses FSO hence supports unicode characters.
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFolderIsWritable(ByVal FolderPath As String)
          Dim FName As String
          Static FSO As Scripting.FileSystemObject
          Dim Counter As Long
          Dim EN As Long
          Dim t As Scripting.TextStream

1         On Error GoTo ErrHandler
2         If (Right$(FolderPath, 1) <> "\") Then FolderPath = FolderPath & "\"
3         If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject
4         If Not FSO.FolderExists(FolderPath) Then
5             CoreFolderIsWritable = False
6         Else
7             Do
8                 FName = FolderPath & "TempFile" & Counter & ".tmp"
9                 Counter = Counter + 1
10            Loop Until Not CoreFileExists(FName)
11            On Error Resume Next
12            Set t = FSO.OpenTextFile(FName, ForWriting, True)
13            EN = Err.Number
14            On Error GoTo ErrHandler
15            If EN = 0 Then
16                t.Close
17                FSO.GetFile(FName).Delete
18                CoreFolderIsWritable = True
19            Else
20                CoreFolderIsWritable = False
21            End If
22        End If

23        Exit Function
ErrHandler:
24        Throw "#CoreFolderIsWritable (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function CoreFolderMove(ByVal FromFolder As String, ByVal ToFolder As String)
          Dim fl As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
          Dim ParentOfTarget As String
1         On Error GoTo ErrHandler
2         If Right$(FromFolder, 1) = "\" Then FromFolder = Left$(FromFolder, Len(FromFolder) - 1)
3         If Right$(ToFolder, 1) = "\" Then ToFolder = Left$(ToFolder, Len(ToFolder) - 1)
4         If Not sFolderExists(FromFolder) Then Throw "Cannot find folder '" + FromFolder + "'"
5         If sFolderExists(ToFolder) Then Throw "Folder '" + ToFolder + "' already exists"
6         If LCase$(Left$(ToFolder, Len(FromFolder) + 1)) = LCase$(FromFolder) + "\" Then Throw "Cannot move folder to a sub-folder of itself"
7         ParentOfTarget = sSplitPath(ToFolder, False)
8         If InStr(ParentOfTarget, "\") > 0 Then
9             ThrowIfError sCreateFolder(ParentOfTarget)        'since otherwise the parent folder of target folder must already exist
10        End If
11        Set FSO = New FileSystemObject
12        Set fl = FSO.GetFolder(FromFolder)
13        fl.Move ToFolder
14        CoreFolderMove = True
15        Set FSO = Nothing: Set fl = Nothing

16        Exit Function
ErrHandler:
17        CoreFolderMove = "#" + Err.Description + "!"
18        Set FSO = Nothing: Set fl = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreFontIsInstalled
' Author    : Philip Swannell
' Date      : 14-Oct-2013
' Purpose   : Tests if a font is installed, presumably quite slow so we use static variables
'             in method SortButtonSetDirection. This code from JW at http://j-walk.com/ss/excel/tips/tip79.htm
' -----------------------------------------------------------------------------------------------------------------------
Function CoreFontIsInstalled(FontName) As Boolean
          '   Returns True if FontName is installed
1         On Error GoTo ErrHandler

2         CoreFontIsInstalled = False
          Dim FontList As CommandBarComboBox
          Dim i As Long
          Dim TempBar As CommandBar

3         Set FontList = Application.CommandBars("Formatting").FindControl(ID:=1728)

          '   If Font control is missing, create a temp CommandBar
4         If FontList Is Nothing Then
5             Set TempBar = Application.CommandBars.Add
6             Set FontList = TempBar.Controls.Add(ID:=1728)
7         End If

8         For i = 0 To FontList.ListCount - 1
9             If LCase$(FontList.List(i + 1)) = LCase$(FontName) Then
10                CoreFontIsInstalled = True
11                On Error Resume Next
12                TempBar.Delete
13                On Error GoTo ErrHandler
14                Exit Function
15            End If
16        Next i

          '   Delete temp CommandBar if it exists
17        On Error Resume Next
18        TempBar.Delete
19        On Error GoTo ErrHandler

20        Exit Function
ErrHandler:
21        Throw "#CoreFontIsInstalled (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreNormOpt
' Author    : Philip Swannell
' Date      : 27-Nov-2015
' Purpose   : Valuation of option in "Normal" model, returns undiscounted value
' -----------------------------------------------------------------------------------------------------------------------
Function CoreNormOpt(OptStyle As EnmOptStyle, Forward, Strike, Volatility, Time)
1         On Error GoTo ErrHandler
          Dim D As Double
          Dim VolRootT

2         If Not IsNumber(Forward) Then Throw "Forward must be a number"
3         If Not IsNumber(Strike) Then Throw "Strike must be a number"
4         If Not IsNumber(Volatility) Then Throw "Volatility must be a number"
5         If Not IsNumber(Time) Then Throw "Time must be a number"
6         If Time < 0 Then Throw "Time must be positive or zero"

7         Select Case OptStyle
              Case OptStyleBuy
8                 CoreNormOpt = Forward - Strike
9                 Exit Function
10            Case OptStyleSell
11                CoreNormOpt = Strike - Forward
12                Exit Function
13        End Select

14        If Time = 0 Then
15            Select Case OptStyle
                  Case OptStyleCall
16                    CoreNormOpt = IIf(Forward > Strike, Forward - Strike, 0)
17                Case OptStylePut
18                    CoreNormOpt = IIf(Forward > Strike, 0, Strike - Forward)
19                Case optStyleUpDigital
20                    CoreNormOpt = IIf(Forward > Strike, 1, 0)        '> or >=
21                Case optStyleDownDigital
22                    CoreNormOpt = IIf(Forward > Strike, 0, 1)        '> or >=
23                Case Else
24                    Throw "Unhandled OptionStyle"
25            End Select
26        Else
27            VolRootT = Volatility * Sqr(Time)
28            D = (Forward - Strike) / VolRootT
29            Select Case OptStyle
                  Case OptStyleCall
30                    CoreNormOpt = (Forward - Strike) * func_normsdist(D) + VolRootT * func_normsdense(-D)
31                Case OptStylePut
32                    CoreNormOpt = (Strike - Forward) * func_normsdist(-D) + VolRootT * func_normsdense(D)
33                Case optStyleUpDigital
34                    CoreNormOpt = func_normsdist(D)
35                Case optStyleDownDigital
36                    CoreNormOpt = func_normsdist(-D)
37                Case Else
38                    Throw "Unhandled OptionStyle"
39            End Select
40        End If

41        Exit Function
ErrHandler:
42        CoreNormOpt = "#CoreNormOpt (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function CoreOptSolveVol(OptStyle As EnmOptStyle, Value, Forward, Strike, Time, LogNormal)
1         On Error GoTo ErrHandler
          Dim Res
2         If Not IsNumber(Forward) Then Throw "Forward must be a number"
3         If Not IsNumber(Strike) Then Throw "Strike must be a number"
4         If Not IsNumber(Value) Then Throw "Value must be a positive number"
5         If Value < 0 Then Throw "Value must be positive"
6         If Not IsNumber(Time) Then Throw "Time must be a number"
7         If VarType(LogNormal) <> vbBoolean Then Throw "LogNormal must be True for Log-Normal vol or False for Normal vol"
8         If Time < 0 Then Throw "Time must be positive or zero"

9         m_OSV_OptStyle = OptStyle
10        m_OSV_Value = Value
11        m_OSV_Forward = Forward
12        m_OSV_Strike = Strike
13        m_OSV_Time = Time
14        m_OSV_logNormal = LogNormal

15        Res = fsolve("OptSolveVolObjectiveFn", 0.05)
16        If VarType(Res) >= vbArray Then Res = Res(1)
17        CoreOptSolveVol = Res

18        Exit Function
ErrHandler:
19        CoreOptSolveVol = "#CoreOptSolveVol (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CoreParseDate
' Author     : Philip Swannell
' Date       : 13-Dec-2017
' Purpose    : At heart of sParseDate, also used by sFileShow
' Parameters :
'  DateString   :
'  DateOrder    :Order of date elements:0 = month-day-year 1 = day-month-year 2 = year-month-day
'  DateSeparator: character string typically "/" or "-", " " or ",". Defaults from Application.International(xlDateSeparator)
' -----------------------------------------------------------------------------------------------------------------------
Function CoreParseDate(DateString As String, DateOrder As Long, DateSeparator As String, ThrowOnError As Boolean)
          Dim CopyOfErr As String
          Dim D As String
          Dim M As String
          Dim pos1 As Long
          Dim pos2 As Long
          Dim y As String
          Static SysDateOrder As Variant
          Static SysDateSeparator As String
          
1         On Error GoTo ErrHandler
2         If IsEmpty(SysDateOrder) Then SysDateOrder = Application.International(xlDateOrder)
3         If Len(SysDateSeparator) = 0 Then SysDateSeparator = Application.International(xlDateSeparator)

4         If Len(DateSeparator) = 0 Then
5             If Len(DateString) <> 8 Then Throw "When there is no date separator, DateString must have eight characters"
6             If DateOrder = 0 Then
7                 M = Left$(DateString, 2)
8                 D = Mid$(DateString, 3, 2)
9                 y = Right$(DateString, 4)
10            ElseIf DateOrder = 1 Then
11                D = Left$(DateString, 2)
12                M = Mid$(DateString, 3, 2)
13                y = Right$(DateString, 4)
14            ElseIf DateOrder = 2 Then
15                y = Left$(DateString, 4)
16                M = Mid$(DateString, 5, 2)
17                D = Right$(DateString, 2)
18            Else
19                Throw "DateOrder must be 0,1, or 2"
20            End If
21        Else
22            pos1 = InStr(DateString, DateSeparator)
23            pos2 = InStr(pos1 + 1, DateString, DateSeparator)
24            If pos1 = 0 Or pos2 = 0 Then Throw "DateString not valid - must contain two instances of '" + DateSeparator + "'"
25            If DateOrder = 0 Then
26                M = Left$(DateString, pos1 - 1)
27                D = Mid$(DateString, pos1 + 1, pos2 - pos1 - 1)
28                y = Mid$(DateString, pos2 + 1)
29            ElseIf DateOrder = 1 Then
30                D = Left$(DateString, pos1 - 1)
31                M = Mid$(DateString, pos1 + 1, pos2 - pos1 - 1)
32                y = Mid$(DateString, pos2 + 1)
33            ElseIf DateOrder = 2 Then
34                y = Left$(DateString, pos1 - 1)
35                M = Mid$(DateString, pos1 + 1, pos2 - pos1 - 1)
36                D = Mid$(DateString, pos2 + 1)
37            Else
38                Throw "DateOrder must be 0,1, or 2"
39            End If
40        End If

41        If Not IsNumeric(y) Then Throw "Year invalid"
42        If Not IsNumeric(M) Then
43            Throw "Month invalid"
44        ElseIf (CLng(M) > 12 Or CLng(M) < 1) Then
45            Throw "Month out of range"
46        End If

47        If Not IsNumeric(D) Then
48            Throw "Day invalid"
49        Else
50            Select Case CLng(D)
                  Case Is < 1, Is > 31
51                    Throw "Day out of range"
52                Case 29, 30, 31
53                    If CLng(D) > day(DateSerial(y, CLng(M) + 1, 0)) Then
54                        Throw "Day out of range"
55                    End If
56            End Select
57        End If

58        If SysDateOrder = 0 Then
59            CoreParseDate = CLng(CDate(M + SysDateSeparator + D + SysDateSeparator + y))
60        ElseIf SysDateOrder = 1 Then
61            CoreParseDate = CLng(CDate(D + SysDateSeparator + M + SysDateSeparator + y))
62        ElseIf SysDateOrder = 2 Then
63            CoreParseDate = CLng(CDate(y + SysDateSeparator + M + SysDateSeparator + D))
64        End If
65        Exit Function
ErrHandler:
66        CopyOfErr = "#CoreParseDate (line " & CStr(Erl) + "): " & Err.Description & "!"
67        If ThrowOnError Then
68            Throw CopyOfErr
69        Else
70            CoreParseDate = CopyOfErr
71        End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CoreRoundSF
' Author     : Philip Swannell
' Date       : 20-Apr-2018
' Purpose    : Round to significant figures, wrapped by sRoundSF
' -----------------------------------------------------------------------------------------------------------------------
Function CoreRoundSF(Number As Variant, NumSigFigs As Variant, Ties As Variant)
          Dim Exponent As Double
1         On Error GoTo ErrHandler
2         If IsNumberOrDate(Number) And IsNumber(NumSigFigs) Then
3             If NumSigFigs < 1 Then
4                 CoreRoundSF = "#NumSigFigs must be at least 1!"
5                 Exit Function
6             ElseIf (NumSigFigs) <> CLng(NumSigFigs) Then
7                 CoreRoundSF = "#NumSigFigs must a whole number!"
8             Else
9                 If Number <> 0 Then
10                    Exponent = Int(Log(Abs(Number)) / Log(10#))
11                Else
12                    Exponent = 0
13                End If
14                CoreRoundSF = CoreRound(Number, _
                      NumSigFigs - (1 + Exponent), Ties)
15            End If
16        Else
17            CoreRoundSF = "#Non number found!"
18        End If

19        Exit Function
ErrHandler:
20        CoreRoundSF = "#CoreRoundSF (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function CoreSplitPath(FullFileName As Variant, Optional ReturnFileName As Boolean = True) As Variant
1         On Error GoTo ErrHandler
          Dim SlashPos As Long
          Dim SlashPos2 As Long
2         If VarType(FullFileName) = vbString Then
3             SlashPos = InStrRev(FullFileName, "\")
4             SlashPos2 = InStrRev(FullFileName, "/")
5             If SlashPos2 > SlashPos Then SlashPos = SlashPos2
6             If SlashPos = 0 Then Throw "Neither '\' nor '/' found"

7             If ReturnFileName Then
8                 CoreSplitPath = Mid$(FullFileName, SlashPos + 1)
9             Else
10                CoreSplitPath = Left$(FullFileName, SlashPos - 1)
11            End If
12        Else
13            Throw "FullFileName must be a string"
14        End If

15        Exit Function
ErrHandler:
16        CoreSplitPath = "#" & Err.Description & "!"
End Function

Function CoreSuppressNAs(x, y)
1         CoreSuppressNAs = x
2         If IsError(x) Then If CStr(x) = "Error 2042" Then CoreSuppressNAs = y
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CoreURLDownloadToFile
' Author    : Philip Swannell
' Date      : 14-Dec-2015
' Purpose   : Downloads bits from the Internet and saves them to a file.
'             See https://msdn.microsoft.com/en-us/library/ms775123(v=vs.85).aspx
'             This function is wrapped by sURLDownloadToFile for broadcast behaviour
' -----------------------------------------------------------------------------------------------------------------------
Function CoreURLDownloadToFile(URLAddress As String, ByVal FileName As String)
          Dim Res
          Dim TargetFolder As String

1         On Error GoTo ErrHandler
2         TargetFolder = sSplitPath(CStr(FileName), False)
3         If Left$(TargetFolder, 1) = "#" Then Throw "Invalid FileName"
4         If Not sFolderExists(TargetFolder) Then Throw "Folder " + TargetFolder + " does not exist"
5         If CoreFileExists(FileName) Then
6             Res = CStr(CoreFileDelete(FileName))
7             If Left$(Res, 1) = "#" Then Throw "Cannot overwrite " + FileName
8         End If
9         Res = URLDownloadToFile(0, URLAddress, FileName, 0, 0)
10        If Res <> 0 Then Throw "Error with error code: " + CStr(Res)
11        If Not CoreFileExists(FileName) Then Throw "Unknown error"
12        CoreURLDownloadToFile = True

13        Exit Function
ErrHandler:
14        CoreURLDownloadToFile = "#CoreURLDownloadToFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : DataFromAuditSheet
' Author    : Philip Swannell
' Date      : 26-Apr-2016
' Purpose   : Grab location of various "release folders" from the Audit sheet, taking
'             different values when working from home, since not then on the solum network.
' -----------------------------------------------------------------------------------------------------------------------
Function DataFromAuditSheet(Description As String)
1         On Error GoTo ErrHandler

2         DataFromAuditSheet = RangeFromSheet(shAudit, Description).Value
3         DataFromAuditSheet = Replace(DataFromAuditSheet, "%OneDriveUserFolder%", sRegistryRead("HKCU\SOFTWARE\Microsoft\OneDrive\UserFolder"))

4         Exit Function
ErrHandler:
5         Throw "#DataFromAuditSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DecodeFileAttributes
' Author     : Philip Swannell
' Date       : 04-May-2018
' Purpose    : Translate file attributes into string. See https://superuser.com/questions/44812/windows-explorers-file-attribute-column-values
'              or (more authoratative) https://msdn.microsoft.com/en-us/library/windows/desktop/gg258117%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
' Parameters :
'  x:
' -----------------------------------------------------------------------------------------------------------------------
Private Function DecodeFileAttributes(x As Long)

          Dim Res As String
1         If x And 1 Then Res = Res & "R"    'Read-only
2         If x And 2 Then Res = Res & "H"    'Hidden
3         If x And 4 Then Res = Res & "S"    'System
4         If x And 8 Then Res = Res & "V"    'Volume label (obsolete in NTFS and must not be set)
5         If x And 16 Then Res = Res & "D"    'Directory
6         If x And 32 Then Res = Res & "A"    'Archive
7         If x And 64 Then Res = Res & "X"    'Device (reserved by system and must not be set)
8         If x And 128 Then Res = Res & "N"    'Normal (i.e. no other attributes set)
9         If x And 256 Then Res = Res & "T"    'Temporary
10        If x And 512 Then Res = Res & "P"    'Sparse file
11        If x And 1024 Then Res = Res & "L"    'Symbolic link / Junction / Mount point / has a reparse point
12        If x And 2048 Then Res = Res & "C"    'Compressed
13        If x And 4096 Then Res = Res & "O"    'Offline
14        If x And 8192 Then Res = Res & "I"    'Not content indexed (shown as 'N' in Explorer in Windows Vista)
15        If x And 16384 Then Res = Res & "E"    'Encrypted
16        DecodeFileAttributes = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ExecuteExcel4Macro2
' Author    : Philip Swannell
' Date      : 09-Mar-2017
' Purpose   : A bit of a nightmare. It appears that certain releases of Excel 2016 crash
'             on any use of Application.ExecuteExcel4Macro
'             (example 16.0.7766.7080 - see File > Account > About Excel)
'             This can be used as a replacement to avoid such crashes.
' -----------------------------------------------------------------------------------------------------------------------
Function ExecuteExcel4Macro2(TheString)
          Dim c As Range
          Dim oldADA
1         On Error GoTo ErrHandler

2         If TypeName(Application.Caller) = "Range" Then Throw "Function cannot be called from worksheet formula, only from VBA"

          '8 Feb 2019. Try switching back to Application.ExecuteExcel4Macro and see if we encounter crashes
          'Has the advantage that can be called from spreadsheet, i.e. when TypeName(Application.Caller) = "Range"
3         ExecuteExcel4Macro2 = Application.ExecuteExcel4Macro(TheString)
4         Exit Function

5         Set c = ThisWorkbook.Excel4MacroSheets("Macro1").Range("A1")
6         oldADA = Application.DisplayAlerts
7         If oldADA Then Application.DisplayAlerts = False
8         c.Formula = "=" + TheString
9         If oldADA Then Application.DisplayAlerts = True
10        Application.Run c
11        ExecuteExcel4Macro2 = c.Value
12        c.Clear
13        Exit Function
ErrHandler:
14        Throw "#ExecuteExcel4Macro2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: GetHelpData
' Purpose: Gets the help data from the Help sheet of this workbook and the Help sheet of SolumSCRiPTUtils.xlam, if that's open
' Author: Philip Swannell
' Date: 29-Nov-2017
' -----------------------------------------------------------------------------------------------------------------------
Function GetHelpData()
          Dim r1 As Range
          Dim r2 As Range
          Dim Result
          Dim wb As Excel.Workbook
1         On Error GoTo ErrHandler
2         Set r1 = shHelp.Range("TheData")
3         Result = HelpDataFillIn(r1)
4         If IsInCollection(Application.Workbooks, gAddinName2 & ".xlam") Then
5             Set wb = Application.Workbooks(gAddinName2 & ".xlam")
6             If IsInCollection(wb.Worksheets, "Help") Then
7                 Set r2 = RangeFromSheet(wb.Worksheets("Help"), "TheData")
8                 Result = sArrayStack(Result, HelpDataFillIn(r2))
9             End If
10        End If
11        GetHelpData = sSortedArray(Result)
12        Exit Function
ErrHandler:
13        Throw "#GetHelpData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function HelpDataFillIn(R As Range, Optional NumCols As Long)
          Dim ArgumentDescriptions As Variant
          Dim ExtraHelp As String
          Dim FunctionDescription As String
          Dim i As Long
          Dim LongFunctionName As String
          Dim NumArgs As Long
          Dim Result
1         On Error GoTo ErrHandler
2         If NumCols > 0 Then
3             Result = R.Resize(, 4).Value
4         Else
5             Result = R.Value
6         End If
7         For i = 1 To sNRows(Result)
8             If IsEmpty(Result(i, 3)) Then
9                 NumArgs = R.Cells(i, 5).Value
10                If NumArgs < 1 Then
11                    ArgumentDescriptions = vbNullString
12                Else
13                    ArgumentDescriptions = R.Cells(i, 8).Resize(1, NumArgs).Value
14                End If
15                ExtraHelp = R.Cells(i, 6)
16                LongFunctionName = Result(i, 2)
17                FunctionDescription = R.Cells(i, 7).Value
18                Result(i, 3) = HelpFromFunctionAndArgDescriptions(LongFunctionName, FunctionDescription, ArgumentDescriptions, ExtraHelp)
19            End If
20        Next
21        HelpDataFillIn = Result
22        Exit Function
ErrHandler:
23        Throw "#HelpDataFillIn (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HelpFromFunctionAndArgDescriptions
' Author    : Philip Swannell
' Date      : 07-May-2015
' Purpose   : Helps in the construction of the "Help" sheet. Use to keep Help text that appears in the
'             SOLUM > Help dialog in sync with the help that appears in the Excel function wizard.
' -----------------------------------------------------------------------------------------------------------------------
Function HelpFromFunctionAndArgDescriptions(LongFunctionName As String, FunctionDescription, ByVal ArgumentDescriptions, ByVal ExtraHelp As String)
          Dim ArgNames
          Dim i As Long
          Dim NumArgs As Long
          Dim NumDescriptions As Long
          Dim ReturnString
1         Force2DArrayR ArgumentDescriptions

2         ArgNames = sStringBetweenStrings(LongFunctionName, "(", ")")
3         ArgNames = Replace(ArgNames, ",...", vbNullString)
4         ReturnString = FunctionDescription
5         If Len(ArgNames) > 0 Then
6             ArgNames = sTokeniseString(CStr(ArgNames))
7             NumArgs = sNRows(ArgNames)
8             NumDescriptions = sNCols(ArgumentDescriptions)
9             ReturnString = ReturnString + vbLf + vbLf + "Arguments" + vbLf
10        End If
11        For i = 1 To NumArgs
12            ReturnString = ReturnString + ArgNames(i, 1)
13            If i <= NumDescriptions Then ReturnString = ReturnString + ": " + vbLf + CStr(ArgumentDescriptions(1, i))
14            If i < NumArgs Then ReturnString = ReturnString + vbLf
15        Next i
16        If Len(ExtraHelp) > 0 Then
17            Do While (Left$(ExtraHelp, 1) = vbCr Or Left$(ExtraHelp, 1) = vbLf)
18                ExtraHelp = Right$(ExtraHelp, Len(ExtraHelp) - 1)
19            Loop
20            ReturnString = ReturnString + vbLf + vbLf + ExtraHelp
21        End If
22        HelpFromFunctionAndArgDescriptions = ReturnString
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HelpVBE
' Author     : Philip Swannell
' Date       : 30-Nov-2018
' Purpose    : Returns the help for a function (read from the Help worksheet) in a Format suitable to become the function's
'              header.
' -----------------------------------------------------------------------------------------------------------------------
Function HelpVBE(FunctionName As String, Optional TheDate As Variant)
          Dim ArgNames
          Dim Col_ArgumentDescription As Variant
          Dim Col_Description As Long
          Dim Col_Extra As Long
          Dim Col_FnName As Long
          Dim Col_ParseArgs As Long
          Dim ColumnHeaders As Variant
          Dim ExtraHelp As String
          Dim i As Long
          Dim NumArgs As Long
          Dim R As Range
          Dim RowNum As Variant
          Dim SA As clsStringAppend
          
1         On Error GoTo ErrHandler

2         Set SA = New clsStringAppend
3         Set R = shHelp.Range("TheData")

4         ColumnHeaders = Application.WorksheetFunction.Transpose(shHelp.Range("TheData").Rows(0).Value)
5         Col_FnName = ThrowIfError(sMatch("FunctionName", ColumnHeaders))
6         Col_Description = ThrowIfError(sMatch("FunctionWizard Description", ColumnHeaders))
7         Col_ArgumentDescription = ThrowIfError(sMatch("Argument 1", ColumnHeaders))
8         Col_ParseArgs = ThrowIfError(sMatch("TheExplanationTitles", ColumnHeaders))
9         Col_Extra = ThrowIfError(sMatch("Extra Description", ColumnHeaders))

10        RowNum = sMatch(FunctionName, R.Columns(Col_FnName).Value)
11        If Not (IsNumber(RowNum)) Then Throw "Cannot find help for function '" + FunctionName + "' on Help sheet"
12        SA.Append "'" & String(105, "-") & vbLf
13        SA.Append "' Procedure : " & FunctionName & vbLf
14        SA.Append "' Author    : Philip Swannell" & vbLf

15        If IsMissing(TheDate) Then TheDate = Date
16        SA.Append "' Date      : "
17        If VarType(TheDate) = vbString Then
18            SA.Append TheDate & vbLf
19        Else
20            SA.Append Format$(TheDate, "dd-mmm-yyyy") & vbLf
21        End If
22        SA.Append "' Purpose   :" & InsertBreaks(R.Cells(RowNum, Col_Description)) & vbLf
23        SA.Append "' Arguments" & vbLf

24        ArgNames = R.Cells(RowNum, Col_ParseArgs)
25        ArgNames = sStringBetweenStrings(ArgNames, "(", ")")
26        ArgNames = Replace(ArgNames, ",...", vbNullString)
27        If Len(ArgNames) > 0 Then
28            ArgNames = sTokeniseString(CStr(ArgNames))
29            NumArgs = sNRows(ArgNames)
30            For i = 1 To NumArgs
31                SA.Append "' " & ArgNames(i, 1)
32                If Len(ArgNames(i, 1)) < 10 Then SA.Append String(10 - Len(ArgNames(i, 1)), " ")
33                SA.Append ":" & InsertBreaks(CStr(R.Cells(RowNum, Col_ArgumentDescription - 1 + i).Value)) + vbLf
34            Next
35        End If
36        ExtraHelp = R.Cells(RowNum, Col_Extra).Value
37        If Len(ExtraHelp) > 0 Then
38            Do While (Left$(ExtraHelp, 1) = vbLf Or Left$(ExtraHelp, 1) = vbCr)
39                ExtraHelp = Right$(ExtraHelp, Len(ExtraHelp) - 1)
40            Loop
41            SA.Append ("'" & vbLf)
42            SA.Append "' Notes     :"
43            SA.Append InsertBreaks(ExtraHelp)
44            SA.Append vbLf
45        End If
46        SA.Append "'" & String(105, "-")
47        CopyStringToClipboard SA.Report
48        HelpVBE = SA.Report
49        Exit Function
ErrHandler:
50        HelpVBE = "#HelpVBE (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : OptSolveVolObjectiveFn
' Author    : Philip Swannell
' Date      : 03-Dec-2015
' Purpose   : Called by: sOptSolveVol > CoreOptSolveVol > fsolve >...> OptSolveVolObjectiveFn
' -----------------------------------------------------------------------------------------------------------------------
Function OptSolveVolObjectiveFn(VolGuess)
          Const Multiplier = 1000
1         If m_OSV_logNormal Then
2             OptSolveVolObjectiveFn = (bsCore(m_OSV_OptStyle, m_OSV_Forward, m_OSV_Strike, VolGuess, m_OSV_Time) - m_OSV_Value) * Multiplier
3         Else
4             OptSolveVolObjectiveFn = (CoreNormOpt(m_OSV_OptStyle, m_OSV_Forward, m_OSV_Strike, VolGuess, m_OSV_Time) - m_OSV_Value) * Multiplier
5         End If
End Function
