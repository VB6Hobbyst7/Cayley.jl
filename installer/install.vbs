' Installer for Cayley2022
' Philip Swannell 12 March 2022

'To Debug this file, install visual studio set up for debugging 
'Then run from a command prompt (in the appropriate folder)
'cscript.exe /x Install.vbs
'https://www.codeproject.com/Tips/864659/How-to-Debug-Visual-Basic-Script-with-Visual-Studi

Option Explicit

Const Website = "https://github.com/SolumCayley/Cayley.jl"
Const GIFRecordingFlagFile = "C:\Temp\RecordingGIF.tmp"
Const MsgBoxTitle = "Install Cayley 2022"
Const MsgBoxTitleBad = "Install Cayley 2022 - Error Encountered"
Const ElevateToAdmin = True
Const IntellisenseName32 = "ExcelDna.IntelliSense.xll"
Const IntellisenseName64 = "ExcelDna.IntelliSense64.xll"
Const JuliaExcelName = "JuliaExcel.xlam"
Const JuliaExcelFolder = "C:\Users\Public\JuliaExcel" 'Copied here by JuliaExcel installer
Const JuliaExcelWebsite = "https://github.com/PGS62/JuliaExcel.jl"
Dim JuliaExcelFullName
Dim JuliaExcelIsInstalled

Const SnakeTail = "C:\Program Files\SnakeTail\SnakeTail.exe"
Const SnakeTail_B = "C:\Program Files (x86)\SnakeTail\SnakeTail.exe"
Dim SnakeTailIsInstalled
Const SnakeTailWebsite = "http://snakenest.com/snaketail/"

'Folders to delete. Cayley2017 files were installed here
Const OldFolder1 = "C:\Program Files\Solum\Addins"
Const OldFolder2 = "C:\Program Files\Solum\CSharp"
Const OldFolder3 = "C:\Program Files\Solum"

'Addin to uninstall
Const OldAddInName1 = "SolumAddin.xlam"
Const OldAddInName2 = "SolumSCRiPTExcel-Addin.xll"

Dim AddinsDest
AddinsDest = "C:\ProgramData\Solum\Addins"'Will be changed later if PC has a different AltStartUp

'PGS 14/4/2022 Try no longer installing Intellisense (in fact uninstall it) thanks to it being a suspect
'for causing Airbus's anti-virus (Netscope) to "take a disliking" to the software, i.e. cause it to hang randomly.
Dim InstallIntellisense
InstallIntellisense = False

Const IntellisenseDest  = "C:\ProgramData\Solum\ExcelDNA"
Const WorkbooksDest = "C:\ProgramData\Solum\Workbooks"
Const DataDest = "C:\ProgramData\Solum\Data\Trades"
Const MarketDataDest = "C:\ProgramData\Solum\Data\Market"

Const WorkbookNames = "Cayley2022.xlsm,CayleyMarketData.xlsm,CayleyTradesTemplate.xlsm,CayleyLines.xlsm,SCRiPT2022.xlsm"
Const DataFileNames = "ExampleFxTrades.csv,ExampleRatesTrades.CSV,ExampleAmortisation.csv"
Const MarketDataFileNames = "20220228_solum.out,20211028solum.out,ExampleSolum.out"
Const AddinName1 = "SolumAddin.xlam"
Const AddinName2 = "SolumSCRiPTUtils.xlam"

Dim AddinsSource
Dim AltStartupPath
Dim BaseFolder
Dim DataSource
Dim MarketDataSource
Dim gErrorsEncountered
Dim GIFRecordingMode

Dim IntellisenseName
Dim IntellisenseSource
Dim myWS
Dim Prompt
Dim WorkbooksSource

Function IsProcessRunning(strComputer, strProcess)
    Dim Process, strObject
    IsProcessRunning = False
    strObject = "winmgmts://" & strComputer
    For Each Process In GetObject(strObject).InstancesOf("win32_process")
        If UCase(Process.Name) = UCase(strProcess) Then
            IsProcessRunning = True
            Exit Function
        End If
    Next
End Function

Function CheckProcess(TheProcessName)
    Dim exc, result
    exc = IsProcessRunning(".", TheProcessName)
    If (exc = True) Then
        result = MsgBox(TheProcessName & _
        " is running. Please close it and then click OK to continue.", _
        vbOKOnly + vbExclamation, MsgBoxTitle)
        exc = IsProcessRunning(".", TheProcessName)
        If (exc = True) Then
            result = MsgBox(TheProcessName & " is still running. Please close the " & _
                    "program and restart the installation." & vbLf & vbLf & _
                    "Can't see " & TheProcessName & "?" & vbLf & "Use Windows Task " & _
                    "Manager to check for a ""ghost"" process." & vblf & vbLf & _
                    "Also check that no other user of this PC is logged in and using Excel.", _
                    vbOKOnly + vbExclamation, MsgBoxTitle)
            WScript.Quit
        End If
    End If
End Function

Function FolderExists(TheFolderName)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(TheFolderName)
End Function

Function FolderIsWritable(FolderPath)
          Dim FName
          Dim fso
          Dim Counter
          Dim EN
          Dim T

         If (Right(FolderPath, 1) <> "\") Then FolderPath = FolderPath & "\"
         Set fso = CreateObject("Scripting.FileSystemObject")
         If Not fso.FolderExists(FolderPath) Then
             FolderIsWritable = False
         Else
            Counter = 0
            Do
                FName = FolderPath & "TempFile" & Counter & ".tmp"
                Counter = Counter + 1
            Loop Until Not FileExists(FName)
            On Error Resume Next
            Set T = fso.OpenTextFile(FName, 2, True)
            EN = Err.Number
            On Error GoTo 0
            If EN = 0 Then
                T.Close
                fso.GetFile(FName).Delete
                FolderIsWritable = True
            Else
                FolderIsWritable = False
            End If
        End If
End Function

Function DeleteFolder(TheFolderName)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    f = fso.DeleteFolder(TheFolderName, True)
    If Err.Number <> 0 Then
        gErrorsEncountered = True
        MsgBox "Failed to delete folder '" & TheFolderName & "'" & vbLf & _
            Err.Description, vbExclamation, MsgBoxTitleBad
    End If
End Function

Function DeleteFile(FileName)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.GetFile(FileName).Delete
    If Err.Number <> 0 Then
        gErrorsEncountered = True
        MsgBox "Failed to delete file '" & FileName & "'" & vbLf & _
            Err.Description, vbExclamation, MsgBoxTitleBad
    End If
End Function

Function FileExists(FileName)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fso.GetFile(FileName)
    On Error GoTo 0
    FileExists = TypeName(f) <> "Empty"
    Exit Function
End Function

'Pass FileNames as a string, comma-delimited for multiple files
Function CopyNamedFiles(ByVal TheSourceFolder, ByVal TheDestinationFolder, _
                        ByVal FileNames, ThrowErrorIfNoSourceFile)
    Dim fso
    Dim FileNamesArray, i, ErrorMessage
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (Right(TheSourceFolder, 1) <> "\") Then
        TheSourceFolder = TheSourceFolder & "\"
    End If
    If (Right(TheDestinationFolder, 1) <> "\") Then
        TheDestinationFolder = TheDestinationFolder & "\"
    End If

    FileNamesArray = Split(FileNames, ",")
    For i = LBound(FileNamesArray) To UBound(FileNamesArray)
        If Not (FileExists(TheSourceFolder & FileNamesArray(i))) Then
            If ThrowErrorIfNoSourceFile Then
                gErrorsEncountered = True
                ErrorMessage = "Cannot find file: " & TheSourceFolder & FileNamesArray(i)
                MsgBox ErrorMessage, vbOKOnly + vbExclamation, MsgBoxTitleBad
            End If
        Else
            If FileExists(TheDestinationFolder & FileNamesArray(i)) Then
                On Error Resume Next
                MakeFileWritable TheDestinationFolder & FileNamesArray(i)
            End If
            On Error Resume Next
            fso.CopyFile TheSourceFolder & FileNamesArray(i), _
                         TheDestinationFolder & FileNamesArray(i), True
            If Err.Number <> 0 Then
                gErrorsEncountered = True
                ErrorMessage = "Failed to copy from: " & _
                    TheSourceFolder & FileNamesArray(i) & vbLf & _
                    "to: " & TheDestinationFolder & FileNamesArray(i) & vbLf & _
                    "Error: " & Err.Description
                    If FileExists(TheSourceFolder & FileNamesArray(i)) Then
                        If FileExists(TheDestinationFolder & FileNamesArray(i)) Then
                            ErrorMessage = ErrorMessage & vblf & vbLf & _
                                "Does another user of this PC have the file open in Excel? Check that no other users of the PC are logged in"
                        End If
                    End If
                MsgBox ErrorMessage, vbOKOnly + vbExclamation, MsgBoxTitleBad
            End If
        End If
    Next
End Function

Function MakeFileWritable(FileName)
    Const ReadOnly = 1
    Dim fso
    Dim f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(FileName)
    If f.Attributes And ReadOnly Then
       f.Attributes = f.Attributes XOR ReadOnly 
    End If
End Function

Function MakeFileReadOnly(FileName)
    Const ReadOnly = 1
    Dim fso
    Dim f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(FileName)
    If Not (f.Attributes And ReadOnly) Then
       f.Attributes = f.Attributes XOR ReadOnly 
    End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CreatePath
' Purpose   : Creates a folder on disk. Path can be passed in as C:\This\That\TheOther even if the
'             folder C:\This does not yet exist. If successful returns the name of the
'             folder. If not successful raises an error.
' Arguments
' Path:       Path of the folder to be created. For example C:\temp\My_New_Folder. It does not matter if
'             this path has a terminating backslash or not.
' -----------------------------------------------------------------------------------------------------------------------
Function CreatePath(ByVal Path)
    Dim F 'As Scripting.Folder
    Dim FSO 'As Scripting.FileSystemObject
    Dim i 'As Long
    Dim ParentFolderName 'As String
    Dim ThisFolderName 'As String

    If Left(Path, 2) = "\\" Then
        'it's a UNC path
    ElseIf Mid(Path, 2, 2) <> ":\" Or Asc(UCase(Left(Path, 1))) < 65 Or Asc(UCase(Left(Path, 1))) > 90 Then
        Err.Raise vbObjectError + 1, , "First three characters of Path must give drive letter followed by "":\"" or else be""\\"" for " & _
            "UNC folder name"
    End If

    Path = Replace(Path, "/", "\")

    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FolderExists(Path) Then
        'Go back until we find parent folder that does exist
        For i = Len(Path) - 1 To 3 Step -1
            If Mid(Path, i, 1) = "\" Then
                If FolderExists(Left(Path, i)) Then
                    Set F = FSO.GetFolder(Left(Path, i))
                    ParentFolderName = Left(Path, i)
                    Exit For
                End If
            End If
        Next

        If F Is Nothing Then Err.Raise vbObjectError + 1, , "Cannot create folder " & Left(Path, 3)

        'Add folders one level at a time
        For i = Len(ParentFolderName) + 1 To Len(Path)
            If Mid(Path, i, 1) = "\" Then
                ThisFolderName = Mid(Path, InStrRev(Path, "\", i - 1) + 1, i - 1 - InStrRev(Path, "\", i - 1))
                F.SubFolders.Add ThisFolderName
                Set F = FSO.GetFolder(Left(Path, i))
            End If
        Next

    End If
    
    Set F = FSO.GetFolder(Path)
    CreatePath = F.Path
    Set F = Nothing: Set FSO = Nothing
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetAltStartupPath
' Purpose   : Gets the AltStartupPath, by looking in the Registry
'---------------------------------------------------------------------------------------
Function GetAltStartupPath()
    GetAltStartupPath = RegistryRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
    gOfficeVersion & "\Excel\Options\AltStartup", "Not found")
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetAltStartupPath
' Purpose   : Sets the AltStartupPath, by writing to in the Registry.
'---------------------------------------------------------------------------------------
Function SetAltStartupPath(Path) '(App,Path)
    RegistryWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & gOfficeVersion & _
    "\Excel\Options\AltStartup", Path
End Function

'---------------------------------------------------------------------------------------
' Procedure : RegistryRead
' Purpose   : Read a value from the Registry
' https://msdn.microsoft.com/en-us/library/x05fawxd(v=vs.84).aspx
'---------------------------------------------------------------------------------------
Function RegistryRead(RegKey, DefaultValue)
    Dim myWS
    RegistryRead = DefaultValue
    Set myWS = CreateObject("WScript.Shell")
    On Error Resume Next
    RegistryRead = myWS.RegRead(RegKey) 
    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : RegistryWrite
' Purpose   : Write to the Registry
' https://msdn.microsoft.com/en-us/library/yfdfhz1b(v=vs.84).aspx
'---------------------------------------------------------------------------------------
Function RegistryWrite(RegKey, NewValue)
    Dim myWS
    Set myWS = CreateObject("WScript.Shell")
    myWS.RegWrite RegKey, NewValue, "REG_SZ"
End Function

'---------------------------------------------------------------------------------------
' Procedure : sRegistryKeyExists
' Purpose   : Returns True or False according to whether a RegKey exists in the Registry
'---------------------------------------------------------------------------------------
Function RegistryKeyExists(RegKey)
    Dim myWS, Res
    Res = Empty
    Set myWS = CreateObject("WScript.Shell")
    On Error Resume Next
    Res = myWS.RegRead(RegKey)
    On Error GoTo 0
    RegistryKeyExists = Not (IsEmpty(Res))
End Function

Function RegistryDelete(RegKey)
    Dim myWS
    Set myWS = CreateObject("WScript.Shell")
    myWS.regDelete RegKey
End Function

'Apparently VBScript has no in-line if. So create one, but note that unlike
'VB6/VBA's Iif this one does not evaluate both truepart and falsepart.
Function IIf( expr, truepart, falsepart )
    If expr Then
        IIf = truepart
    Else
        IIf = falsepart
    End If
End Function

Function Environ(Expression)
	Dim WshShell
	Set WshShell = CreateObject("WScript.Shell")
	Environ = WshShell.ExpandEnvironmentStrings("%" & Expression & "%")
End Function

Function IsAddinInstalled(AddinFullName)
    Dim RegKeyBranch
    Dim RegKeyLeaf
    Dim i
    Dim Found
    Dim RegValue

    RegKeyBranch = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
                    gOfficeVersion & "\Excel\Options\"
    i = 0
    Do
        i = i + 1
        RegKeyLeaf = "OPEN" & IIf(i > 1, CStr(i - 1), "")
        If RegistryKeyExists(RegKeyBranch & RegKeyLeaf) Then
            RegValue = RegistryRead(RegKeyBranch & RegKeyLeaf, "")
            Found = InStr(LCase(RegValue), LCase(AddinFullName)) > 0
            If Found Then
                IsAddinInstalled = True
                Exit Function
            End If
        Else
            Exit Do
        End If
    Loop

    IsAddinInstalled = False

End Function


Sub InstallExcelAddin(AddinFullName, WithSlashR)
    Dim RegKeyBranch
    Dim RegKeyLeaf
    Dim i
    Dim Found
    Dim NumAddins
    Dim RegValue

    if Not FileExists(AddinFullName) Then
        gErrorsEncountered = True
        MsgBox "Cannot install addin '" & AddinFullName & "` because the file is not found",vbCritical
        Exit Sub
    End If

    RegKeyBranch = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
                    gOfficeVersion & "\Excel\Options\"
    i = 0
    Do
        i = i + 1
        RegKeyLeaf = "OPEN" & IIf(i > 1, CStr(i - 1), "")
        If RegistryKeyExists(RegKeyBranch & RegKeyLeaf) Then
            NumAddins = NumAddins + 1
            RegValue = RegistryRead(RegKeyBranch & RegKeyLeaf, "")
            Found = InStr(LCase(RegValue), LCase(AddinFullName)) > 0
            If Found Then Exit Sub
        Else
            Exit Do
        End If
    Loop

    RegKeyLeaf = "OPEN" & IIf(NumAddins > 0, CStr(NumAddins), "")
    'I can't discover what is the significance of the /R that appears in the Registry for
    'some addins but not for others...
    If WithSlashR Then
        RegValue = "/R """ & AddinFullName & """"
    Else
        RegValue = AddinFullName
    End If
    RegistryWrite RegKeyBranch & RegKeyLeaf, RegValue
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetOfficeVersionAndBitness
' Author     : Philip Swannell
' Date       : 14-Dec-2021
' Notes      : Previously was trying to determine office version and bitness by reading the registry, which turns out to
'              be hard to do, for example when a PC has had various versions of Office installed. So reverted to 
'              launching Excel via CreateObject.
'              I posted something along these lines at
'              https://stackoverflow.com/questions/2203980/detect-whether-office-is-32bit-or-64bit-via-the-registry
' -----------------------------------------------------------------------------------------------------------------------
Function GetOfficeVersionAndBitness(OfficeVersion,OfficeBitness)
    Dim Excel, EN

    On Error Resume Next
    Set Excel = CreateObject("Excel.Application")
    EN = Err.Number
    Excel.Visible = False
    On Error GoTo 0

    If EN = 0 Then
        If InStr(Excel.OperatingSystem,"64") > 0 Then
            OfficeBitness = 64
            OfficeVersion = Excel.Version
        Else
            OfficeBitness = 32
            OfficeVersion = Excel.Version
        End if
        Excel.Quit
    Else
        OfficeBitness = 0
        OfficeVersion = "Office Not found"
    End If

    Set Excel = Nothing
End Function

Function FriendlyOfficeVersion(OfficeVersion)
    Dim OV, ReleaseID

    ReleaseID = RegistryRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\ProductReleaseIDs","Not found")

    Select Case OfficeVersion
        Case "16.0"
            If ReleaseID = "Not found" Then
                FriendlyOfficeVersion = "Office 2016"
            ElseIf instr(ReleaseID,"365") >0 Then
                FriendlyOfficeVersion = "Office 365"
            ElseIf instr(ReleaseID,"2019") >0 Then
                FriendlyOfficeVersion = "Office 2019"
            ElseIf instr(ReleaseID,"2021") >0 Then
                FriendlyOfficeVersion = "Office 2021"
            Else
                FriendlyOfficeVersion = ReleaseID
            End If
        Case "15.0"
            FriendlyOfficeVersion = "Office 2013"
        Case "14.0"
            FriendlyOfficeVersion = "Office 2010"
        Case "12.0"
            FriendlyOfficeVersion = "Office 2007"
        Case "11.0"
            FriendlyOfficeVersion = "Office 2003"
        Case "10.0"
            FriendlyOfficeVersion = "Office XP"
        Case "9.0"
            FriendlyOfficeVersion = "Office 2000"
        Case "8.0"
            FriendlyOfficeVersion = "Office 98"
        Case "7.0"
            FriendlyOfficeVersion = "Office 97"            
        Case Else
            FriendlyOfficeVersion = OV
    End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DeleteExcelAddinFromRegistry
' Purpose    : Edits the Windows Registry to ensure that excel does not load a particular addin. Will not work if the addin
'              is located in the AltStartUp path
' Parameters :
'  AddinName:  The file name of the addin e.g. "ExcelDna.IntelliSense64.xll" can include the path if we want to remove an
'              addin only if it's currently being loaded from the "wrong" location.
' -----------------------------------------------------------------------------------------------------------------------
Sub DeleteExcelAddinFromRegistry(AddinName)
    Dim RegKey
    Dim AllKeys()
    Dim i, j
    Dim RegKeyLeaf
    Dim NumAddins
    Dim Found

    RegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & gOfficeVersion & "\Excel\Options\"
    i = 0
    Do
        i = i + 1
        RegKeyLeaf = "OPEN" & IIf(i > 1, CStr(i - 1), "")
        If RegistryKeyExists(RegKey & RegKeyLeaf) Then
            NumAddins = NumAddins + 1
        Else
            Exit Do
        End If
    Loop

    Found = False
    
    ReDim AllKeys(NumAddins - 1, 1) 'VBScript has base 0 so that's two columns
    For i = 0 To NumAddins - 1
        RegKeyLeaf = "OPEN" & IIf(i > 0, CStr(i), "")
        AllKeys(i, 0) = RegKeyLeaf
        AllKeys(i, 1) = RegistryRead(RegKey & RegKeyLeaf, "")
        If InStr(LCase(AllKeys(i, 1)), LCase(AddinName)) > 0 Then
            Found = True
        End If
    Next

    If Not Found Then Exit Sub

    For i = 0 To NumAddins - 1
        RegistryDelete RegKey & AllKeys(i, 0)
    Next

    j = 0
    For i = 0 To NumAddins - 1
        If InStr(LCase(AllKeys(i, 1)), LCase(AddinName)) = 0 Then
            j = j + 1
            RegKeyLeaf = "OPEN" & IIf(j > 1, CStr(j - 1), "")
            RegistryWrite RegKey & RegKeyLeaf, AllKeys(i, 1)
        End If
    Next
End Sub

'*******************************************************************************************
'Effective start of this VBScript. Note elevating to admin as per 
'http://www.winhelponline.com/blog/vbscripts-and-uac-elevation/
'We install to C:\ProgramData see
'https://stackoverflow.com/questions/22107812/privileges-owner-issue-when-writing-in-c-programdata
'*******************************************************************************************
If (WScript.Arguments.length = 0) And ElevateToAdmin Then
   Dim objShell, ThisFileName
   Set objShell = CreateObject("Shell.Application")
   'Pass a bogus argument, say [ uac]
   ThisFileName = WScript.ScriptFullName
   objShell.ShellExecute "wscript.exe", Chr(34) & _
      ThisFileName & Chr(34) & " uac", "", "runas", 1
Else
    Dim gOfficeVersion, gOfficeBitness
    Set myWS = CreateObject("WScript.Shell")

    gErrorsEncountered = False
    If Not GIFRecordingMode Then
        'CheckProcess must be called BEFORE GetOfficeVersionAndBitness
        CheckProcess "Excel.exe"
    End If

    GetOfficeVersionAndBitness gOfficeVersion, gOfficeBitness

    GIFRecordingMode = FileExists(GIFRecordingFlagFile)

    If gOfficeVersion = "Office Not found" Then
    Prompt = "Installation cannot proceed because no version of Microsoft Office has " & _
               "been detected on this PC." & vblf  & vblf & _
               "The script attempts to detect the installed versions of Office by " & _
               "executing the code `CreateObject(""Excel.Application"")` which should " & _
               "launch Excel so that its version can be determined." & _ 
               vblf & vblf & "However, that didn't work. So it seems you need to " & _
               "install Microsoft Office before installing Cayley."

        MsgBox Prompt,vbCritical,MsgBoxTitleBad
        WScript.Quit
    End If    

    JuliaExcelFullName = JuliaExcelFolder & "\" & JuliaExcelName
    JuliaExcelIsInstalled = False
    If IsAddinInstalled(JuliaExcelName) Then'Note using JuliaExcelName, not JuliaExcelFullName, i.e. it's OK if the addin is on the PC in a different location
        JuliaExcelIsInstalled = True
    ElseIf FileExists(GetAltStartupPath() & "\" & JuliaExcelName) Then
        JuliaExcelIsInstalled = True
    ElseIf FileExists("C:\Projects\JuliaExcel\workbooks\" & JuliaExcelName) Then
        JuliaExcelIsInstalled = True'On development PC
    End If

    If Not JuliaExcelIsInstalled Then
        Prompt = "It seems that JuliaExcel is not installed on this PC. The Cayley software " & _
                "will not work correctly until JuliaExcel is also installed." & _
                vbLF & vbLF & _
                "After installing the Cayley software, please install JuliaExcel." & _
                vblf & vblf & JuliaExcelWebsite
        MsgBox Prompt,vbInformation,MsgBoxTitleBad
    End If    

    BaseFolder = WScript.ScriptFullName
    BaseFolder = Left(BaseFolder, InStrRev(BaseFolder, "\") - 1)
    BaseFolder = Left(BaseFolder, InStrRev(BaseFolder, "\"))
    AddinsSource = BaseFolder & "addins\"
    IntellisenseSource = AddinsSource
    WorkbooksSource = BaseFolder & "workbooks\"
    DataSource = BaseFolder & "Data\Trades"
    MarketDataSource = BaseFolder & "Data\Market"

	AltStartupPath = GetAltStartupPath()
	If AltStartupPath = "" Or AltStartupPath = "Not found" Then
	    ' Leave AddinsDest as set at top of this file
    Else
	    'NB IF USER HAS ALREADY GOT AltStartupPath set then we use that location...
	    AddinsDest = GetAltStartupPath()
	End If

    Select Case gOfficeBitness
    Case 32
        IntellisenseName = IntellisenseName32
    Case 64
        IntellisenseName = IntellisenseName64
    Case Else
        InstallIntellisense = False
    End Select

    Prompt = "This will install the 2022 version of 'Cayley', software provided by Solum Financial Limited for use by Airbus Treasury." & _
        vbLf & vblf & _
        "If the 2017 version of Cayley is detected, it will be removed " & _
        "because the 2017 and 2022 versions cannot co-exist on the same PC." & vbLF & vbLf & _
        "Do you wish to continue?" & vblf  & vblf & _
        "More information at:" & vblf & Website & string(2,vblf) & _
        "Details:" & vblf & _
        "Office version detected: " & FriendlyOfficeVersion(gOfficeVersion) & " " & gOfficeBitness & "bit" & vblf & vbLf & _
        "Files will be copied from:" & vblf & _
        BaseFolder & vblf & vblf & _
        "Files will be copied to:" & vblf & _
        WorkbooksDest & vblf & _
        AddinsDest

    If InstallIntellisense Then
        Prompt = Prompt & vbLf & IntellisenseDest
    End If

    If MsgBox(Prompt, vbYesNo + vbQuestion, MsgBoxTitle) <> vbYes Then WScript.Quit

    If not GIFRecordingMode Then
        'Delete the old...
        If FolderExists(OldFolder1) Then
            DeleteFolder OldFolder1
        End If

        If FolderExists(OldFolder2) Then
            DeleteFolder OldFolder2
        End If

        If FolderExists(OldFolder3) Then
            DeleteFolder OldFolder3
        End If

        'Because Cayley 2017 installed here, and we don't want users inadvertantly
        'opening old files which will get compile errors against the new addins.
        If FolderExists(WorkbooksDest) Then
            DeleteFolder WorkbooksDest
        End If

        DeleteExcelAddinFromRegistry OldAddInName1
        DeleteExcelAddinFromRegistry OldAddInName2

        SetAltStartupPath AddinsDest

        'Install the new...
        CreatePath WorkbooksDest
        CopyNamedFiles WorkbooksSource, WorkBooksDest, WorkbookNames, True

        CreatePath AddinsDest
        CopyNamedFiles AddinsSource, AddinsDest, AddinName1, True
        CopyNamedFiles AddinsSource, AddinsDest, AddinName2, True
        MakeFileReadOnly AddinsDest & "\" & AddinName1
        MakeFileReadOnly AddinsDest & "\" & AddinName2
        'There are two ways that Excel loads addins, it loads all addins in the AltStartUp
        'path and separately loads all addins that are listed in the registry key that we
        'manipulate via InstallExcelAddin. For SolumAddin.xlam and SolumSCRiPTIUtils.xlam
        'we use AltStartUp, but for Intellisense we use the Registry.
        DeleteExcelAddinFromRegistry AddinName1
        DeleteExcelAddinFromRegistry AddinName2
        CreatePath DataDest
        CopyNamedFiles DataSource, DataDest, DataFileNames, True
        CreatePath MarketDataDest
        CopyNamedFiles MarketDataSource, MarketDataDest, MarketDataFileNames, True

        If InstallIntellisense Then
            CreatePath IntellisenseDest
            CopyNamedFiles IntellisenseSource, IntellisenseDest, IntellisenseName, True
            'Good to delete both then install the correct one, e.g. when a PC previously had
            '32 bit Office, but now has 64 bit or vice versa.
            DeleteExcelAddinFromRegistry IntellisenseName32
            DeleteExcelAddinFromRegistry IntellisenseName64
            InstallExcelAddin IntellisenseDest & "\" & IntellisenseName, True
        Else
            DeleteExcelAddinFromRegistry IntellisenseName32
            DeleteExcelAddinFromRegistry IntellisenseName64
            If FolderExists(IntellisenseDest) Then
                DeleteFolder IntellisenseDest
            End If

        End If
    End If

    SnakeTailIsInstalled = FileExists(SnakeTail) or FileExists(SnakeTail_B)

    If gErrorsEncountered Then
        Prompt = "The install script has finished, but errors were encountered, " & _
                 "which may mean the software will not work correctly." & vblf & vblf & _
                 Website
        MsgBox Prompt, vbOKOnly + vbCritical, MsgBoxTitleBad
    Else
        Prompt = "Cayley2022 installed correctly." & vblf & _
                 vblf & Website
        if Not SnakeTailIsInstalled Then
            Prompt = Prompt & vblf & vblf & _
                "SnakeTail is not installed on this PC, but it needs to be." & vblf & vbLf & _
                "Click OK to go to the SnakeTail website at" & vblf & _
                SnakeTailWebsite
            MsgBox Prompt, vbOKOnly + vbExclamation, MsgBoxTitle
            myWS.run SnakeTailWebsite
        Else
            MsgBox Prompt, vbOKOnly + vbInformation, MsgBoxTitle
        End If
        
    End If

    'Don't record this bit. Have this warning after forgetting about the flag file!
    If GIFRecordingMode Then 
        MsgBox "That previous message was false. The installation was blocked by the " & _
                "existence of file '" & GIFRecordingFlagFile & "'",vbOKOnly + vbCritical, _
                MsgBoxTitleBad
    End If

    WScript.Quit
End If
