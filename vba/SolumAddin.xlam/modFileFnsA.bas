Attribute VB_Name = "modFileFnsA"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CreateHasher
' Author     : Philip Swannell
' Date       : 27-Jul-2020
' Purpose    : Wrap call to CreateObject so that error can mention probable cause - lack of .net framework.
' Parameters :
'  MD5:
' -----------------------------------------------------------------------------------------------------------------------
Function CreateHasher(MD5 As Boolean, ByRef ZeroByteHash As String) As Object

1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler
          Dim ErrorMessage As String
          Dim ObjectName As String
          
3         If MD5 Then
4             ObjectName = "System.Security.Cryptography.MD5CryptoServiceProvider"
5             ZeroByteHash = "d41d8cd98f00b204e9800998ecf8427e"
6         Else
7             ObjectName = "System.Security.Cryptography.SHA1CryptoServiceProvider"
8             ZeroByteHash = "da39a3ee5e6b4b0d3255bfef95601890afd80709"
9         End If
          
10        Set CreateHasher = CreateObject(ObjectName)
          
11        Exit Function
ErrHandler:
12        ErrorMessage = "Cannot create object '" + ObjectName + "' Possible cause Microsoft .NETFramework 3.5 is not installed"

13        Throw "#CreateHasher (line " & CStr(Erl) + "): " & ErrorMessage & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sArrayCheckSum
' Author    : Philip Swannell
' Date      : 05-Feb-2019
' Purpose   : Returns a crytographic hash of an array.
' Arguments
' TheArray  : The array. Must be 0, 1 or 2 dimensional. Elements must strings, numbers, Booleans or
'             error values.
'
' Notes     : Implementation:
'             The array is converted to a string using sMakeArrayString and then the string
'             is hashed using the "System.Security.Cryptography.MD5CryptoServiceProvider"
'             built into Windows.
' -----------------------------------------------------------------------------------------------------------------------
Function sArrayCheckSum(ByVal TheArray As Variant)
Attribute sArrayCheckSum.VB_Description = "Returns a crytographic hash of an array."
Attribute sArrayCheckSum.VB_ProcData.VB_Invoke_Func = " \n27"
          Dim b() As Byte
          Dim Hasher As Object
          Dim HashRes As Variant
          Dim i As Long
          Dim Result As String
          Dim ZeroByteHash As String

1         On Error GoTo ErrHandler
2         b = ThrowIfError(sMakeArrayString(TheArray))
3         Set Hasher = CreateHasher(True, ZeroByteHash)
4         HashRes = Hasher.ComputeHash_2(b)
5         Result = vbNullString
6         For i = 1 To LenB(HashRes)
7             Result = Result & UCase$(Right$("0" & Hex$(AscB(MidB$(HashRes, i, 1))), 2))
8         Next
9         sArrayCheckSum = Result

10        Exit Function
ErrHandler:
11        sArrayCheckSum = "#sArrayCheckSum (line " & CStr(Erl) + "): " & Err.Description & "!"
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
Function sCreateFolder(FolderPath)
Attribute sCreateFolder.VB_Description = "Creates a folder on disk. FolderPath can be passed in as C:\\This\\That\\TheOther even if the folder C:\\This does not yet exist. If successful, returns the name of the folder. If not successful, returns an error string."
Attribute sCreateFolder.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sCreateFolder = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sCreateFolder = Broadcast1Arg(FuncIdCreateFolder, FolderPath)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sDeleteAllFilesInFolder
' Author    : Philip Swannell
' Date      : 07-Sep-2013
' Purpose   : Use this function VERY CAREFULLY - there is no "Are you sure?" dialog. Deletes all files
'             in the given FolderPath, even those with the ReadOnly attribute set. The
'             folder itself is not deleted.
' Arguments
' FolderPath: Path of the folder for which all files in the folder are to be deleted. For example
'             C:\temp. It does not matter if this path has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Function sDeleteAllFilesInFolder(FolderPath As String)
Attribute sDeleteAllFilesInFolder.VB_Description = "Use this function VERY CAREFULLY - there is no ""Are you sure?"" dialog. Deletes all files in the given FolderPath, even those with the ReadOnly attribute set. The folder itself is not deleted."
Attribute sDeleteAllFilesInFolder.VB_ProcData.VB_Invoke_Func = " \n26"

          Dim fi As Scripting.file
          Dim fo As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
          Dim CopyOfErr As String
1         On Error GoTo ErrHandler

2         If FunctionWizardActive() Then
3             sDeleteAllFilesInFolder = "#Disabled in Function Dialog!"
4             Exit Function
5         End If

6         Set FSO = New FileSystemObject
7         Set fo = FSO.GetFolder(FolderPath)
8         For Each fi In fo.Files
9             fi.Delete True
10        Next fi
11        Set fo = Nothing: Set FSO = Nothing

12        Exit Function
ErrHandler:
13        CopyOfErr = "#sDeleteAllFilesInFolder (line " & CStr(Erl) + "): " & Err.Description & "!"
14        Set fi = Nothing: Set fo = Nothing: Set FSO = Nothing
15        sDeleteAllFilesInFolder = CopyOfErr
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sDeleteFolder
' Author    : Philip Swannell
' Purpose   : Use this function VERY CAREFULLY - there is no "Are you sure?" dialog. Deletes the given
'             folder and all files and sub-folders within it, even those with the ReadOnly
'             attribute set.
' Arguments
' FolderPath: Path of the folder to be deleted. For example C:\temp. It does not matter if this path has
'             a terminating backslash or not. This argument may be an array.
' -----------------------------------------------------------------------------------------------------------------------
Function sDeleteFolder(FolderPath)
Attribute sDeleteFolder.VB_Description = "Use this function VERY CAREFULLY - there is no ""Are you sure?"" dialog. Deletes the given folder and all files and sub-folders within it, even those with the ReadOnly attribute set."
Attribute sDeleteFolder.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sDeleteFolder = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sDeleteFolder = Broadcast1Arg(FuncIdDeleteFolder, FolderPath)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileCheckSum
' Author    : Philip Swannell
' Date      : 15-Dec-2015
' Purpose   : Returns a cryptographic hash of a file, using either of the MD5 or SHA1 algorithms. Useful
'             for checking if two files have the same contents. See
'             https://en.wikipedia.org/wiki/MD5 for more information.
' Arguments
' FileName  : Full name (with path) of the file, or an array of file names.
' Method    : Optional argument. Omit or "MD5" for MD5 algorithm, "SHA1" for SHA1 algorithm.
' See http://stackoverflow.com/questions/2826302/how-to-get-the-md5-hex-hash-for-a-file-using-vba
' -----------------------------------------------------------------------------------------------------------------------
Function sFileCheckSum(ByVal FileName As Variant, Optional Method As String = "MD5")
Attribute sFileCheckSum.VB_Description = "Returns a cryptographic hash of a file, using either of the MD5 or SHA1 algorithms. Useful for checking if two files have the same contents. See https://en.wikipedia.org/wiki/MD5 for more information."
Attribute sFileCheckSum.VB_ProcData.VB_Invoke_Func = " \n26"
          Dim Hasher As Object
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result() As String
          Dim TheHash As String
          Dim ZeroByteHas As String
1         On Error GoTo ErrHandler
2         Select Case UCase$(Method)
              Case "MD5"
3                 Set Hasher = CreateHasher(True, ZeroByteHas)
4             Case "SHA1"
5                 Set Hasher = CreateHasher(False, ZeroByteHas)
6             Case Else
7                 sFileCheckSum = "#Method not recognised, allowed values are MD5 and SHA1!"
8                 Exit Function
9         End Select

10        If VarType(FileName) < vbArray Then
11            HashFromFileName Hasher, FileName, TheHash, ZeroByteHas
12            sFileCheckSum = TheHash
13        Else
14            Force2DArrayR FileName, NR, NC
15            ReDim Result(1 To NR, 1 To NC)
16            For i = 1 To NR
17                For j = 1 To NC
18                    HashFromFileName Hasher, FileName(i, j), TheHash, ZeroByteHas
19                    Result(i, j) = TheHash
20                Next j
21            Next i
22            sFileCheckSum = Result
23        End If

24        Set Hasher = Nothing
25        Exit Function
ErrHandler:
26        sFileCheckSum = "#sFileCheckSum (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : GetFileBytes
' Author    : Philip Swannell
' Date      : 16-Dec-2015
' Purpose   : Read a file into a zero-based one dimensional byte array. Uses quite old-school file handling but is fast.
'             From http://stackoverflow.com/questions/2826302/how-to-get-the-md5-hex-hash-for-a-file-using-vba
' -----------------------------------------------------------------------------------------------------------------------
Private Function GetFileBytes(ByVal FileName As String) As Byte()
          Dim bytRtnVal() As Byte
          Dim lngFileNum As Long
          Dim CopyOfErr As String
          Dim FileLength
1         On Error GoTo ErrHandler
2         lngFileNum = FreeFile
          'Note that the calling function sFileCheckSum checks that the file exists...
3         Open FileName For Binary Access Read As lngFileNum
4         FileLength = LOF(lngFileNum)
5         If FileLength < 0 Then
6             FileLength = 4294967296# + FileLength
7         End If

8         ReDim bytRtnVal(0& To FileLength - 1&) As Byte
9         Get lngFileNum, , bytRtnVal
10        Close lngFileNum
11        GetFileBytes = bytRtnVal
12        Erase bytRtnVal
13        Exit Function
ErrHandler:
14        CopyOfErr = Err.Description
15        If lngFileNum > 0 Then Close lngFileNum
16        Throw "#GetFileBytes (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : HashFromFileName
' Author    : Philip Swannell
' Date      : 17-Dec-2015
' Purpose   : Sub-routine of sFileCheckSum, needed to get element-wise error handling
'             Note this is a Sub to avoid exposing directly to Excel - tried moving to modPrivateModule but got strange errors
'NB. Must not call Dir$ in this function as it is a subroutine of sDirList that's using Dir to iterate files
' -----------------------------------------------------------------------------------------------------------------------
Sub HashFromFileName(Hasher As Object, FileName As Variant, ByRef Result As String, ZeroByteHash As String)
          Dim Bytes As Variant
          Dim i As Long
1         On Error GoTo ErrHandler

2         If VarType(FileName) <> vbString Then
3             Result = "#FileName must be a string!"
4         ElseIf Not CoreFileExists(CStr(FileName)) Then
5             Result = "#File not found!"
6         Else
7             On Error Resume Next

8             Bytes = Hasher.ComputeHash_2(GetFileBytes(FileName))
9             If Err.Number = 0 Then
10                On Error GoTo ErrHandler
11                Result = vbNullString
12                For i = 1 To LenB(Bytes)
13                    Result = Result & LCase$(Right$("0" & Hex$(AscB(MidB$(Bytes, i, 1))), 2))
14                Next
15            Else
16                Result = Err.Description
17                On Error GoTo ErrHandler
18                If FileLen(FileName) = 0 Then
19                    Result = ZeroByteHash
20                    Exit Sub
21                End If
22            End If
23        End If
24        Exit Sub
ErrHandler:
25        Result = "#HashFromFileName (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileCopy
' Author    : Philip Swannell
' Date      : 17-Jun-2013
' Purpose   : Copies a named file (or files) from one location to another, overwriting if the target
'             already exists. If the copy process fails, then an error string is returned.
' Arguments
' SourceFile: Full name (with path) of the source file, or URL address of the file (http or https). Can
'             be an array, in which case TargetFile must be an array of the same
'             dimensions.
' TargetFile: Full name (with path) of the target destination.  Can be an array, in which case
'             SourceFile must be an array of the same dimensions.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileCopy(ByVal SourceFile As Variant, ByVal TargetFile As Variant)
Attribute sFileCopy.VB_Description = "Copies a named file (or files) from one location to another, overwriting if the target already exists. If the copy process fails, then an error string is returned."
Attribute sFileCopy.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFileCopy = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFileCopy = Broadcast2Args(FuncIdFileCopy, SourceFile, TargetFile)
End Function

Function sFileCopySkip(ByVal SourceFile As Variant, ByVal TargetFile As Variant, NumLinesToSkip, Optional Unicode = False)
Attribute sFileCopySkip.VB_Description = "Copies a text file from one location to another, but with the first NumLinesToSkip of the file not copied to the target file."
Attribute sFileCopySkip.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFileCopySkip = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFileCopySkip = Broadcast(FuncIdFileCopySkip, SourceFile, TargetFile, NumLinesToSkip, Unicode)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileDelete
' Author    : Philip Swannell
' Date      : 14-Dec-2015
' Purpose   : Use with care! The function deletes the file with the given FileName from the disk. The
'             function returns the value TRUE if successful. It does not support wildcards
'             and does not handle folders.
' Arguments
' FileName  : Full filename of the file to delete, including path. Can be an array to delete many files.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileDelete(ByVal FileName As Variant) As Variant
Attribute sFileDelete.VB_Description = "Use with care! The function deletes the file with the given FileName from the disk. The function returns the value TRUE if successful. It does not support wildcards and does not handle folders."
Attribute sFileDelete.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFileDelete = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFileDelete = Broadcast1Arg(FuncIdFileDelete, FileName)
End Function

Function sFileUnblock(ByVal FileName As Variant) As Variant
Attribute sFileUnblock.VB_Description = "Unblocks a file which may have been downloaded from the internet."
Attribute sFileUnblock.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFileUnblock = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFileUnblock = Broadcast1Arg(FuncIdFileUnblock, FileName)
End Function

Function sFileIsUnicode(ByVal FileName As Variant)
Attribute sFileIsUnicode.VB_Description = "Returns TRUE if a file is a text file encoded in unicode rather than ascii."
Attribute sFileIsUnicode.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFileIsUnicode = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFileIsUnicode = Broadcast1Arg(FuncIdFileIsUnicode, FileName)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileExif
' Author    : Philip Swannell
' Date      : 29-Oct-2017
' Purpose   : Returns "Exif" data from a file, such as the name of the artist (for a music file) or the
'             size of the file in pixels (for an image file).
' Arguments
' FileName  : Full filename, including the path. Does not yet support arrays of file names.
' PropertyNumber: A number in the range -1 to 312. For example 13 gives the "Contributing artists" for a
'             music file. Call the function with no arguments to see a table of the
'             meanings of all allowed PropertyNumbers.
'             Adapted from https://windowssecrets.com/forums/showthread.php/103921-Picture-Property-(VB-VBA-2003
' -----------------------------------------------------------------------------------------------------------------------
Function sFileExif(Optional FileName As Variant, Optional PropertyNumber As Variant)
Attribute sFileExif.VB_Description = "Returns ""Exif"" data from a file, such as the name of the artist (for a music file) or the size of the file in pixels (for an image file)."
Attribute sFileExif.VB_ProcData.VB_Invoke_Func = " \n26"

1         On Error GoTo ErrHandler

2         If IsMissing(FileName) Then
3             sFileExif = sArrayRange(sGrid(-1, 312, 314), sTokeniseString("Info tip,Name,Size,Item type,Date modified,Date created,Date accessed,Attributes,Offline status,Availability,Perceived type,Owner,Kind,Date taken,Contributing artists,Album,Year,Genre," & _
                  "Conductors,Tags,Rating,Authors,Title,Subject,Categories,Comments,Copyright,#,Length,Bit rate,Protected,Camera model,Dimensions,Camera maker,Company,File description," & _
                  "Program name,Duration,Is online,Is recurring,Location,Optional attendee addresses,Optional attendees,Organiser address,Organiser name,Reminder time,Required attendee addresses," & _
                  "Required attendees,Resources,Meeting status,Free/busy status,Total size,Account name,,Task status,Computer,Anniversary,Assistant's name,Assistant's phone,Birthday,Business address," & _
                  "Business city,Business country/region,Business P.O. box,Business postcode,Business county/region,Business street,Business fax,Business home page,Business phone,Call-back number," & _
                  "Car phone,Children,Company main phone,Department,Email address,Email2,Email3,Email list,Email display name,File as,First name,Full name,Gender,Given name,Hobbies,Home address," & _
                  "Home city,Home country/region,Home P.O. box,Home postcode,Home county/region,Home street,Home fax,Home phone,IM addresses,Initials,Job title,Label,Surname,Postal address,Middle name," & _
                  "Mobile phone,Nickname,Office location,Other address,Other city,Other country/region,Other P.O. box,Other postcode,Other county/region,Other street,Pager,Personal title,City," & _
                  "Country/region,P.O. box,Postcode,County/Region,Street,Primary email,Primary phone,Profession,Spouse/Partner,Suffix,TTY/TTD phone,Telex,Web page,Content status,Content type," & _
                  "Date acquired,Date archived,Date completed,Device category,Connected,Discovery method,Friendly name,Local computer,Manufacturer,Model,Paired,Classification,Status,Status,Client ID," & _
                  "Contributors,Content created,Last printed,Date last saved,Division,Document ID,Pages,Slides,Total editing time,Word count,Due date,End date,File count,File extension,Filename," & _
                  "File version,Flag colour,Flag status,Space free,,,Group,Sharing type,Bit depth,Horizontal resolution,Width,Vertical resolution,Height,Importance,Is attachment,Is deleted," & _
                  "Encryption status,Has flag,Is completed,Incomplete,Read status,Shared,Creators,Date,Folder name,Folder path,Folder,Participants,Path,By location,Type,Contact names,Entry type,Language," & _
                  "Date visited,Description,Link status,Link target,URL,,,,Media created,Date released,Encoded by,Episode number,Producers,Publisher,Season number,Subtitle,User web URL,Writers,," & _
                  "Attachments,Bcc addresses,Bcc,Cc addresses,Cc,Conversation ID,Date received,Date sent,From addresses,From,Has attachments,Sender address,Sender name,Store,To addresses,To do title," & _
                  "To,Mileage,Album artist,Sort album artist,Album ID,Sort album,Sort contributing artists,Beats-per-minute,Composers,Sort composer,Disc,Initial key,Part of a compilation,Mood,Part of set," & _
                  "Full stop,Colour,Parental rating,Parental rating reason,Space used,EXIF version,Event,Exposure bias,Exposure program,Exposure time,F-stop,Flash mode,Focal length,35mm focal length," & _
                  "ISO speed,Lens maker,Lens model,Light source,Max aperture,Metering mode,Orientation,People,Program mode,Saturation,Subject distance,White balance,Priority,Project,Channel number," & _
                  "Episode name,Closed captioning,Rerun,SAP,Broadcast date,Program description,Recording time,Station call sign,Station name,Summary,Snippets,Auto summary,Relevance,File ownership," & _
                  "Sensitivity,Shared with,Sharing status,,Product name,Product version,Support link,Source,Start date,Sharing,Sync status,Billing information,Complete,Task owner,Sort title,Total file size," & _
                  "Legal trademarks,Video compression,Directors,Data rate,Frame height,Frame rate,Frame width,Video orientation,Total bitrate,Masters Keywords (debug),Masters Keywords (debug)"))
4             Exit Function
5         End If

6         If VarType(FileName) < vbArray And VarType(PropertyNumber) < vbArray Then
7             If Not IsNumber(PropertyNumber) Then Throw "PropertNumber must be a number"
8             sFileExif = CoreFileExif(CStr(FileName), CDbl(PropertyNumber))
9         Else
10            sFileExif = Broadcast(FuncIdFileExif, FileName, PropertyNumber)
11        End If

12        Exit Function
ErrHandler:
13        sFileExif = "#sFileExif (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileExists
' Author    : Philip Swannell
' Date      : 08-Sep-2013
' Purpose   : Returns TRUE if a file of the given FileName exists on disk or FALSE otherwise.
' Arguments
' FileName  : The full name of the file, including the path. Can be an array of file names.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileExists(ByVal FileName As Variant) As Variant
Attribute sFileExists.VB_Description = "Returns TRUE if a file of the given FileName exists on disk or FALSE otherwise."
Attribute sFileExists.VB_ProcData.VB_Invoke_Func = " \n26"
1         sFileExists = Broadcast1Arg(FuncIdFileExists, FileName)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileHeaders
' Author    : Philip Swannell
' Date      : 15-Jul-2017
' Purpose   : Returns the top row of a text file, as an array. The file is assumed to be a Windows file,
'             Unix files are not currently supported.
' Arguments
' FileName  : The full name of the file, including the path.
' Delimiter : Enter the delimiter character or omit for the function to guess the delimiter as the first
'             occurrence of comma, tab, semi-colon, colon or vertical bar (|).
' LineNumber: The line in the file at which the headers are found, defaults to 1 for the first line.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileHeaders(FileName As String, Optional Delimiter As Variant, Optional LineNumber As Long = 1)
Attribute sFileHeaders.VB_Description = "Returns the top row of a text file, as an array. The file is assumed to be a Windows file, Unix files are not currently supported."
Attribute sFileHeaders.VB_ProcData.VB_Invoke_Func = " \n26"

1         On Error GoTo ErrHandler
2         sFileHeaders = ThrowIfError(sFileShow(FileName, Delimiter, , , , , , , , , LineNumber, 1, 1))

3         Exit Function
ErrHandler:
4         sFileHeaders = "#sFileHeaders (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileLastModifiedDate
' Author    : Philip Swannell
' Date      : 18-Jun-2013
' Purpose   : Returns the DateLastModified property of a file on disk.
' Arguments
' FileName  : The full name of the file, including the path. Can be an array of file names.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileLastModifiedDate(FileName As Variant)
Attribute sFileLastModifiedDate.VB_Description = "Returns the DateLastModified property of a file on disk."
Attribute sFileLastModifiedDate.VB_ProcData.VB_Invoke_Func = " \n26"
1         sFileLastModifiedDate = Broadcast1Arg(FuncIdFileLastModifiedDate, FileName)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileMove
' Author    : Philip Swannell
' Date      : 01-Mar-2016
' Purpose   : Moves a named file (or files) from one location to another. If the ToFile already exists
'             it is overwritten. If the folder for the ToFile does not yet exist it is
'             created.  If the process fails, then an error string is returned.
' Arguments
' FromFile  : Full name (with path) of the current location of the file. Can be an array, in which case
'             ToFile must be an array of the same dimensions.
' ToFile    : Full name (with path) of the new location of the file. Can be an array, in which case
'             FromFile must be an array of the same dimensions.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileMove(ByVal FromFile As Variant, ByVal ToFile As Variant)
Attribute sFileMove.VB_Description = "Moves a named file (or files) from one location to another. If the ToFile already exists it is overwritten. If the folder for the ToFile does not yet exist it is created.  If the process fails, then an error string is returned."
Attribute sFileMove.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFileMove = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFileMove = Broadcast2Args(FuncIdFileMove, FromFile, ToFile)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileNumLines
' Author    : Philip Swannell
' Date      : 23-Feb-2018
' Purpose   : Counts the number of lines in a text file.
' Arguments
' FileName  : Full name (with path) of the file. May also be an array of file names.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileNumLines(ByVal FileName As Variant) As Variant
Attribute sFileNumLines.VB_Description = "Counts the number of lines in a text file."
Attribute sFileNumLines.VB_ProcData.VB_Invoke_Func = " \n26"
1         sFileNumLines = Broadcast1Arg(FuncIdFileNumLines, FileName)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileRename
' Author    : Philip Swannell
' Purpose   : Renames an existing file. If a file of the new name already exists it is overwritten. If
'             the process fails an error string is returned.
' Arguments
' OldFileName: Full name (with path) of the current file name. Can be an array, in which case NewFileName
'             must be an array of the same dimensions.
' NewFileName: New name for the file, with or without the path, but if path is given it must be the same
'             as the path to OldFileName. Can be an array, in which case OldFileName must
'             be an array of the same dimensions.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileRename(ByVal OldFileName As Variant, NewFileName As Variant)
Attribute sFileRename.VB_Description = "Renames an existing file. If a file of the new name already exists it is overwritten. If the process fails an error string is returned."
Attribute sFileRename.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFileRename = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFileRename = Broadcast2Args(FuncIdFileRename, OldFileName, NewFileName)
End Function

Function sFolderRename(ByVal OldFolderPath As Variant, NewFolderName As Variant)
Attribute sFolderRename.VB_Description = "Renames a folder, or array of folders."
Attribute sFolderRename.VB_ProcData.VB_Invoke_Func = " \n26"
1         If FunctionWizardActive() Then
2             sFolderRename = "#Disabled in Function Dialog!"
3             Exit Function
4         End If
5         sFolderRename = Broadcast2Args(FuncIdFolderRename, OldFolderPath, NewFolderName)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sFileSave
' Author    : Philip Swannell
' Date      : 18-Dec-2015
' Purpose   : Creates a text file on disk containing the data in the array DataToWrite. Any existing
'             file of the same name is overwritten. If successful the function returns the
'             name of the file written, otherwise an error string.
' Arguments
' FileName  : The full name of the file, including the path.
' DataToWrite: An array of arbitrary data.
' Delimiter : The delimiter character(s)
' EscapeCharacter: If EscapeCharacter is passed, then any instances of Delimiter within elements of
'             DataToWrite are first replaced by EscapeCharacter.
' ReplaceCRLF: If TRUE, then any carriage return and line feed characters within elements of DataToWrite
'             are first replaced by <CRLF> (carriage return-line feed pair) <CR> (carriage
'             return) or <LF> (line feed).
' LineEndings: Specifies the line endings of the file written. Allowed values 'Windows' or 'CRLF', 'Unix'
'             or 'LF', 'Macintosh' or 'CR'. This argument is optional and defaults to
'             'Windows'.
' LastLineHasLineEnding: Specifies whether the last line in the file ends with a line ending. This argument is
'             optional and defaults to FALSE.
' -----------------------------------------------------------------------------------------------------------------------
Function sFileSave(FileName As String, ByVal DataToWrite As Variant, Optional Delimiter As String, Optional EscapeCharacter As String, Optional ReplaceCRLF As Boolean, Optional LineEndings As String = "Windows", Optional LastLineHasLineEnding As Boolean = False, Optional Unicode As Boolean)
Attribute sFileSave.VB_Description = "Creates a text file on disk containing the data in the array DataToWrite. Any existing file of the same name is overwritten. If successful, the function returns the name of the file written, otherwise an error string."
Attribute sFileSave.VB_ProcData.VB_Invoke_Func = " \n26"

          Dim DataToWrite1Col
          Dim FSO As Scripting.FileSystemObject
          Dim i As Long
          Dim j As Long
          Dim LineEnding As String
          Dim NC As Long
          Dim NR As Long
          Dim t As TextStream
          Dim CopyOfErr As String

1         On Error GoTo ErrHandler
2         CheckFileNameIsAbsolute FileName

3         Force2DArrayR DataToWrite, NR, NC

4         Select Case LCase$(LineEndings)
              Case vbNullString, "windows", "crlf", vbCrLf
5                 LineEnding = vbCrLf
6             Case "unix", "lf", vbLf
7                 LineEnding = vbLf
8             Case "macintosh", "mac", "cr", vbCr
9                 LineEnding = vbCr
10            Case Else
11                Throw "LineEndings must be 'Windows', 'Unix' or 'Macintosh'"
12        End Select

13        If EscapeCharacter <> vbNullString Then
14            For i = 1 To NR
15                For j = 1 To NC
16                    If VarType(DataToWrite(i, j)) = vbString Then
17                        If InStr(DataToWrite(i, j), Delimiter) > 0 Then
18                            DataToWrite(i, j) = Replace(DataToWrite(i, j), Delimiter, EscapeCharacter)
19                        End If
20                    End If
21                Next j
22            Next i
23        End If

24        If ReplaceCRLF Then
25            For i = 1 To NR
26                For j = 1 To NC
27                    If VarType(DataToWrite(i, j)) = vbString Then
28                        DataToWrite(i, j) = Replace(DataToWrite(i, j), vbCrLf, "<CRLF>")
29                        DataToWrite(i, j) = Replace(DataToWrite(i, j), vbCr, "<CR>")
30                        DataToWrite(i, j) = Replace(DataToWrite(i, j), vbLf, "<LF>")
31                    End If
32                Next j
33            Next i
34        End If

35        DataToWrite1Col = sRowConcatenateStrings(DataToWrite, Delimiter)
36        Force2DArray DataToWrite1Col
37        Set FSO = New FileSystemObject
38        Set t = FSO.CreateTextFile(FileName, True, Unicode)
39        NR = sNRows(DataToWrite1Col)
40        For i = 1 To NR - 1
41            t.Write DataToWrite1Col(i, 1) + LineEnding
42        Next i
43        t.Write DataToWrite1Col(NR, 1)
44        If LastLineHasLineEnding Then t.Write LineEnding

45        t.Close: Set t = Nothing: Set FSO = Nothing
46        sFileSave = FileName
47        Exit Function
ErrHandler:
48        CopyOfErr = "#sFileSave (line " & CStr(Erl) + "): " & Err.Description & "!"
49        If Not t Is Nothing Then t.Close: Set t = Nothing: Set FSO = Nothing
50        sFileSave = CopyOfErr
End Function

