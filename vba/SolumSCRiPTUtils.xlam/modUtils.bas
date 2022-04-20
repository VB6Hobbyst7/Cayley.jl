Attribute VB_Name = "modUtils"
Option Explicit
Public Const gWHATIF = "WHATIF"
Public Const gSELF = "SELF"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ArePackagesMissing
' Author    : Philip Swannell
' Date      : 16-Oct-2016
' Purpose   : Are some of the required packages missing. This method has undesirable
'             hard-coding of the list of packages required...
' -----------------------------------------------------------------------------------------------------------------------
Function ArePackagesMissing(Optional Packages As String = gPackages) As Boolean
          Dim InstalledPackages
          Dim RequiredPackages
1         On Error GoTo ErrHandler
2         InstalledPackages = ThrowIfError(sExecuteRCode("installed.packages()[,1]"))
3         RequiredPackages = sTokeniseString(Packages)
4         ArePackagesMissing = sNRows(sCompareTwoArrays(InstalledPackages, RequiredPackages, "In2AndNotIn1")) > 1
5         Exit Function
ErrHandler:
6         Throw "#ArePackagesMissing (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Sub TestInstallPackages()
    On Error GoTo ErrHandler
    InstallPackages True

    Exit Sub
ErrHandler:
    SomethingWentWrong "#TestInstallPackages (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : InstallPackages
' Author     : Philip Swannell
' Date       : 30-Jul-2020 - re-written to use "versions" package https://cran.r-project.org/web/packages/versions/versions.pdf
' Purpose    : Installs all packages required by either SCRiPT or SolumAddin. Uses R package "versions" to get versions as of a particular date,
'              but that date is set to the day before the code executes.
' Parameters :
'  ForceInstall: If True then ALL required packages are installed, if false then all required packages are installed only if one or more are not currently installed
' -----------------------------------------------------------------------------------------------------------------------
Sub InstallPackages(ForceInstall As Boolean)

          Dim AsOfDate As String
          Dim CodetoExecute
          Dim ErrorLevel As Long
          Dim PackagesAfter
          Dim PackagesBefore
          Dim PackagesFailedToInstall
          Dim PackagesRequired As String
          Dim Prompt As String
          Dim Result
          Const Title = "Install Packages"
          
          Const dq = """"
          
1         On Error GoTo ErrHandler

2         AsOfDate = Format(Date - 1, "yyyy-mm-dd")
          
3         PackagesRequired = sConcatenateStrings(sRemoveDuplicates(sTokeniseString(gPackages + "," + gPackagesSAI)))

4         If ForceInstall Then        'calling from menu
5             Prompt = "The R code used for trade valuation and PFE calculation relies on a number of ""Packages"" that need to be downloaded." + vbLf + vbLf + _
                  "Would you like to install recent versions of these packages now? This requires an internet connection and may take a few minutes. You might see ""Progress Bars"" on the screen."

6             If MsgBoxPlus(Prompt, vbQuestion + vbOKCancel, Title, "Yes, install", "No, do nothing") <> vbOK Then Exit Sub
7         ElseIf ArePackagesMissing(PackagesRequired) Then
8             Prompt = "The R code used for trade valuation and PFE calculation relies on a number of ""Packages"" that need to be downloaded." + vbLf + vbLf + _
                  "It appears that some required packages are not installed. You should install them now and that requires an internet connection and may take a few minutes. You might see ""Progress Bars"" on the screen." + vbLf + vbLf + "Proceed?"
9             If MsgBoxPlus(Prompt, vbQuestion + vbOKCancel, Title, "Yes, install", "No, do nothing") <> vbOK Then
10                Throw "This workbook will not function correctly until you install the necessary R packages." + vbLf + vbLf + "See Menu > R Environment > Install Packages...", True
11            End If
12        Else
13            Exit Sub
14        End If
          
15        PackagesBefore = sExecuteRCode("installed.packages()[, 1]")
16        ErrorLevel = 1 'ignore warnings
          
17        If Not IsNumber(sMatch("versions", PackagesBefore)) Then
18            ThrowIfError sExecuteRCode("install.packages(""versions"", repos = ""http://cran.rstudio.com/"", method = ""wininet"", quiet = TRUE)", ErrorLevel)
19        End If

20        ThrowIfError sExecuteRCode("library(versions)", ErrorLevel)
21        CodetoExecute = "install.dates(c(" + sConcatenateStrings(sArrayConcatenate(dq, sTokeniseString(gPackages), dq)) + "),as.Date(" + dq + AsOfDate + dq + "))"
22        sExecuteRCode CodetoExecute, ErrorLevel
          
23        PackagesAfter = sExecuteRCode("installed.packages()[, 1]")

24        PackagesFailedToInstall = sCompareTwoArrays(sTokeniseString(PackagesRequired), PackagesAfter, "12")

25        Result = sExecuteRCode("installed.packages()[,1:3]")
26        Result = sSubArray(Result, , 2)

27        If sNRows(PackagesFailedToInstall) > 1 Then
28            PackagesFailedToInstall = sSubArray(PackagesFailedToInstall, 2)
29            Prompt = "Something went wrong. The following packages failed to install:" + vbLf + _
                  sConcatenateStrings(PackagesFailedToInstall, vbLf) + vbLf + vbLf + _
                  "Here is a list of all the R packages currently installed" + vbLf + vbLf + _
                  sConcatenateStrings(sJustifyArrayOfStrings(sArrayMakeText(Result), "Segoe UI", 9, vbTab), vbLf)
30            MsgBoxPlus Prompt, vbOKOnly + vbCritical, Title, , , , , 2000
31        ElseIf ForceInstall Then
32            Prompt = "All done." + vbLf + vbLf + _
                  "Here is a list of all the R packages currently installed" + vbLf + vbLf + _
                  sConcatenateStrings(sJustifyArrayOfStrings(sArrayMakeText(Result), "Segoe UI", 9, vbTab), vbLf)
33            MsgBoxPlus Prompt, vbInformation, Title, , , , , 2000
34        End If

35        Exit Sub
ErrHandler:
36        Throw "#InstallPackages (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : LoggingIsOn
' Author    : Philip Swannell
' Date      : 12-Sep-2016
' Purpose   : Is Logging switched on in R?
' -----------------------------------------------------------------------------------------------------------------------
Function LoggingIsOn() As Boolean
          Dim Expression As String
1         On Error GoTo ErrHandler
2         Expression = "if(!exists(""gDoLogging"")){FALSE}else{gDoLogging}"
3         LoggingIsOn = ThrowIfError(sExecuteRCode(Expression))
4         Exit Function
ErrHandler:
5         Throw "#LoggingIsOn (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : LoggingSwitchOn
' Author    : Philip Swannell
' Date      : 29-Sep-2016
' Purpose   : Switch logging in R on or off
' -----------------------------------------------------------------------------------------------------------------------
Sub LoggingSwitchOn(SwitchOn As Boolean)
          Dim Expression As String
1         On Error GoTo ErrHandler
2         Expression = "gDoLogging <- " + UCase(SwitchOn)
3         ThrowIfError sExecuteRCode(Expression)
4         Exit Sub
ErrHandler:
5         Throw "#LoggingSwitchOn (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Sub RefreshIntellisenseSheetInSCRiPTUtils()
    RefreshIntellisenseSheet ThisWorkbook.Worksheets("Help").Range("TheData"), ThisWorkbook.Worksheets("_IntelliSense_")
    UninstallIntellisense
    InstallIntellisense
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ResetREnvironment
' Author    : Philip Swannell
' Date      : 06-Jan-2016
' Purpose   : Remove all variables from R environment...
' -----------------------------------------------------------------------------------------------------------------------
Public Sub ResetREnvironment(WithDialog As Boolean, WithStatusBarMessage As Boolean)
          Dim NumObjects0
          Dim NumObjects1
          Dim NumRemoved
1         On Error GoTo ErrHandler
          Dim Prompt As String
          Const Title = "Reset R Environment"

2         If WithDialog Then
3             Prompt = "This method removes all data currently held in memory in the R environment in which trade valuation and PFE calculation takes place. The method is designed to be used only by developers of the R code." + vbLf + vbLf + _
                       "Proceed?"
4             If MsgBoxPlus(Prompt, vbQuestion + vbOKCancel + vbDefaultButton2, Title, "Yes,remove data", "No, do nothing") <> vbOK Then
5                 Exit Sub
6             End If
7         End If
8         NumObjects0 = ThrowIfError(sExecuteRCode("length(ls())"))
9         ThrowIfError sExecuteRCode("remove(list = setdiff(ls(), c(""BERT.Version"", ""gDoLogging"")))") 'Must not remove variable BERT.version or else method TestInstallation fails.
10        NumObjects1 = ThrowIfError(sExecuteRCode("length(ls())"))
11        NumRemoved = NumObjects0 - NumObjects1
12        If WithStatusBarMessage Then
13            Prompt = "R Environment reset. " + CStr(NumRemoved) + " object" + IIf(NumRemoved = 1, "", "s") + " removed."
14            MsgBoxPlus Prompt, vbInformation, Title
15        End If
16        Exit Sub
ErrHandler:
17        SomethingWentWrong "#ResetREnvironment (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Sub SaveMe()

ThisWorkbook.IsAddin = True
ThisWorkbook.SaveAs "c:\ProgramData\Solum\Addins\SolumSCRiPTUtils.xlam", xlOpenXMLAddIn

End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SaveREnvironment
' Author    : Philip Swannell
' Date      : 03-Nov-2015
' Purpose   : Save the R environment - very useful debugging tool
' -----------------------------------------------------------------------------------------------------------------------
Public Sub SaveREnvironment(FileNameStub As String)
          Dim Expression As String
          Dim FileName As String
          Dim FileNameUnix As String
          Dim Message As String
          Dim Prompt As String
          Const CheckBoxCaption = "Don't show this message again"
          Static CheckBoxValue As Boolean

1         On Error GoTo ErrHandler

          Const Title = "Save R environment"

2         If sFolderExists("c:\temp") Then
3             If sFolderIsWritable("c:\temp") Then
4                 FileName = "c:\temp"
5             End If
6         End If
7         If FileName = "" Then
8             FileName = sEnvironmentVariable("temp")
9         End If
10        FileName = FileName + "\" + FileNameStub + "-" + Format(Now, "yyyy-mm-dd-hh-mm-ss") & ".rdata"
11        FileNameUnix = Replace(FileName, "\", "/")

12        Prompt = "This method saves the R environment to file to enable debugging of the code against the current trade and market data. The ""load"" command (for use in the debugging environment) is copied to the Windows clipboard." + vbLf + vbLf + _
                   "Proceed?"
13        If Not CheckBoxValue Then
14            CheckBoxValue = True    'i.e. We assume the user needs to see the dialog only once
15            If MsgBoxPlus(Prompt, vbQuestion + vbOKCancel + vbDefaultButton2, Title, "Yes, save", "No do nothing", , , , CheckBoxCaption, CheckBoxValue) <> vbOK Then Exit Sub
16        End If

17        CopyStringToClipboard "load(""" + FileNameUnix + """)"
18        Expression = "save.image(""" & FileNameUnix & """)"
19        ThrowIfError sExecuteRCode(Expression)

20        Message = "R environment saved to: " + FileName
21        TemporaryMessage Replace(Message, vbLf, " "), , ""
22        Exit Sub
ErrHandler:
23        SomethingWentWrong "#SaveREnvironment (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SCRiPTLogFileName
' Author     : Philip Swannell
' Date       : 15-Oct-2018
' Purpose    : Returns the name of the log file that SCRiPT will use. If such file does not exist then this method creates it.
' Parameters :
'  Reset:      If true then all existing contents are deleted
' -----------------------------------------------------------------------------------------------------------------------
Function SCRiPTLogFileName(Optional Reset As Boolean) As String
          Dim FileName As String
          Dim FilePath As String
1         On Error GoTo ErrHandler
2         FileName = "SCRiPTLog-" + Format(Date, "yyyy-mm-dd") + ".log"
3         FilePath = Environ("Temp") & "\" & FileName
4         If Not sFileExists(FilePath) Or Reset Then
5             Open FilePath For Output As #1
6             Close #1
7         End If
8         SCRiPTLogFileName = FilePath

9         Exit Function
ErrHandler:
10        Throw "#SCRiPTLogFileName (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ShowLogFiles
' Author    : Philip Swannell
' Date      : 09-Nov-2015
' Purpose   : Display The log file created by R code. Attempt to not open a second instance
'             of SnakeTail, but if the log file is open in an instance of snaketail without
'             being the top-most tab then a second instance will open. SnakeTail appears not
'             to have much in the way of command line options.
' Reset     : Deletes contents of log file
' -----------------------------------------------------------------------------------------------------------------------
Sub ShowLogFiles(Optional Reset As Boolean)
          Dim EN As Long
          Dim FileName As String
          Dim FilePath As String
          Dim Title As String
1         On Error GoTo ErrHandler
            
2         FilePath = SCRiPTLogFileName(Reset)
3         FileName = sSplitPath(FilePath)

4         Title = "SnakeTail - [" & FileName
5         On Error Resume Next
6         AppActivate Title, False
7         EN = Err.Number
8         On Error GoTo ErrHandler
9         If EN <> 0 Then ShowFileInSnakeTail FilePath
10        Exit Sub
ErrHandler:
11        Throw "#ShowLogFiles (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SourceRCode
' Author    : Philip Swannell
' Date      : 14-Sep-2016
' Purpose   : Resource the R code, useful when debugging
' -----------------------------------------------------------------------------------------------------------------------
Sub SourceRCode(Optional Silent As Boolean = False, Optional WithCayley As Boolean = False, Optional ByVal Path As String)
1         On Error GoTo ErrHandler
          Dim Prompt As String
          Const Title = "Source R code"
          Const CheckBoxCaption = "Don't show this message again"
          Static CheckBoxValue As Boolean
          
2         If Path = "" Then Path = gRSourcePath
3         If Not sFolderExists(Path) Then Throw ("Cannot find folder '" + Path + "'")
4         Path = Replace(Path, "\", "/")
5         If Right(Path, 1) <> "/" Then Path = Path + "/"

6         If Not Silent Then
7             If Not CheckBoxValue Then
8                 CheckBoxValue = True
9                 Prompt = "This method ""sources"" the R code used for trade valuation and PFE calculation. The method is designed for use by developers of the R code." + vbLf + vbLf + _
                      "The R code is in folder" + vbLf + Replace(Path, "/", "\") + vbLf + vbLf + _
                      "Proceed?"
10                If MsgBoxPlus(Prompt, vbQuestion + vbOKCancel + vbDefaultButton2, Title, "Yes, source the code", "No do nothing", , , , CheckBoxCaption, CheckBoxValue) <> vbOK Then Exit Sub
11            End If
12        End If

13        ThrowIfError sExecuteRCode("suppressWarnings(source(""" + Path + "UtilsPGS.R""))")
14        ThrowIfError sExecuteRCode("suppressWarnings(source(""" + Path + "SCRiPTMain.R""))")
15        ThrowIfError sExecuteRCode("suppressWarnings(SourceAllFiles(" + UCase(WithCayley) + "))")
16        Exit Sub
ErrHandler:
17        SomethingWentWrong "#SourceRCode (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : stopImplicitCluster
' Author    : Philip Swannell
' Date      : 27-Jan-2016
' Purpose   : Call R function stopImplicitCluster. Called from main menu and also (inSilentMode)
'             from the workbook close event
' -----------------------------------------------------------------------------------------------------------------------
Sub stopImplicitCluster(SilentMode As Boolean)
          Dim NumProcesses As Long
          Dim Process As Object
          Const Title = "Shut down R processes"
          Const ProcessName = "Rscript.exe"

          Dim Message As String

1         If SilentMode Then
2             On Error Resume Next
3             sExecuteRCode "stopImplicitCluster()"
4             Exit Sub
5         Else
6             On Error GoTo ErrHandler
7             Message = "The R code used for trade valuation and PFE calculation uses parallel processes, which " + _
                        "can be seen in the Windows Task Manager as multiple instances of " + ProcessName + "." & vbLf + vbLf + _
                        "The processes start the first time you value trades after opening this workbook and are shut down " + _
                        "when you close the workbook."
8             For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & ProcessName & "'")
9                 NumProcesses = NumProcesses + 1
10            Next
11            If NumProcesses = 0 Then
12                Message = Message + vbLf + vbLf + "No parallel processes are currently running."
13                MsgBoxPlus Message, vbOKOnly + vbInformation, Title
14                Exit Sub
15            Else
16                Message = Message + vbLf + vbLf + "There " + IIf(NumProcesses > 1, "are", "is") + " currently " + CStr(NumProcesses) + _
                          " parallel processes" + IIf(NumProcesses > 1, "s", "") + " running. Would you like to shut them down now?"
17                If MsgBoxPlus(Message, vbYesNo + vbQuestion + vbDefaultButton2, Title, "Yes, shut down", "No, do nothing") <> vbYes Then Exit Sub
18                sExecuteRCode "stopImplicitCluster()"
19                NumProcesses = 0
20                For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & ProcessName & "'")
21                    NumProcesses = NumProcesses + 1
22                Next
23            End If
24        End If
25        Exit Sub
ErrHandler:
26        SomethingWentWrong "#stopImplicitCluster (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : TestInstallation
' Author    : Philip Swannell
' Date      : 10-Apr-2016
' Purpose   : Test if installation of software looks good for running SCRiPT. Updated
'             22 Sep 2017 for use of BERT rather than SolumSCRiPTExcel-AddIn
' -----------------------------------------------------------------------------------------------------------------------
Function TestInstallation() As Boolean
          Dim a As AddIn
          Dim Res As Variant
          Dim wb As Workbook
          Const SSEFileName = "SolumSCRiPTExcel-AddIn.xll"    'We no longer use this add-in so this method uninstalls it if it's still installed...
          Dim BERTInstalled As Boolean
          Dim DLLs As Variant
          Dim ErrorMessage As String
          Dim MainDotRLocation As String
          Const BERTURL = "https://github.com/sdllc/Basic-Excel-R-Toolkit/releases/tag/2.4.4"
          Const VersionRequired = "BERT Version 2.4.4" 'This method can allow two versions of BERT as OK. Want to allow only one version? then set these two constants to the same value.
          Const VersionRequired_2 = "BERT Version 2.4.4"

1         On Error Resume Next
2         Set wb = Application.Workbooks("SolumAddin.xlam")
3         On Error GoTo ErrHandler
4         If wb Is Nothing Then Err.Raise vbObjectError + 1, , "SolumAddin is not installed"    'but if this is the case none of the code in this addin will compile so there's not really much point to this test...

          'If the user has launched, (say) SCRiPT.xlsm (which calls this method from its workbook open event), simultaneously with launching Excel _
           (e.g. by doubleclicking on the file in explorer) then it appears that BERT is not loaded at the time this _
           code runs, so best to exit early. Test for that condition is quite crude...
5         If Now - shAudit.Range("LaunchTime").Value2 < 5 / 24 / 60 / 60 Then '5 seconds since this workbook opened
6             TestInstallation = False
7             Exit Function
8         End If

          'Many things don't work if the Decimal separator and list separator are "wrong". And it's not worth trying to make SCRiPT work reliably in such cases...
9         If Application.International(xlDecimalSeparator) + Application.International(xlListSeparator) <> ".," Then
10            ErrorMessage = "SCRiPT requires the Windows Decimal Separator be a dot (.) and the Windows List Separator be a comma (,)." + vbLf + vbLf + "To run SCRiPT please change these two settings via 'Windows Control Panel' > 'Change date, time or number formats' > 'Additional settings'."
11            GoTo PostErrorMessage
12        End If

13        For Each a In Application.AddIns
14            If InStr(1, a.FullName, SSEFileName, vbTextCompare) > 0 Then
15                If a.Installed Then
16                    a.Installed = False
17                End If
18            End If
19        Next a

20        If Not IsNull(Application.RegisteredFunctions) Then
21            DLLs = sSplitPath(sSubArray(Application.RegisteredFunctions, 1, 1, , 1), True)
22            If IsNumber(sMatch("BERT32.xll", DLLs)) Then
23                BERTInstalled = True
24            ElseIf IsNumber(sMatch("BERT64.xll", DLLs)) Then
25                BERTInstalled = True
26            End If
27        End If

28        If Not BERTInstalled Then
29            ErrorMessage = "BERT (Basic Excel R Toolkit), is not installed. Please install it by downloading from:" + vbLf + BERTURL + vbLf + vbLf + _
                  "Hint:" + vbLf + "It's possible that BERT has been installed but has been disabled, in which case you need to enable it. See: File > Options > Add-Ins > Manage > Disabled Items > Go"
30            GoTo PostErrorMessage
31        End If

32        Res = 0
33        On Error Resume Next
34        Res = Application.Run("BERT.Exec", "1+1")
35        On Error GoTo ErrHandler
36        If Res <> 2 Then
37            ErrorMessage = "There is a problem. BERT (Basic Excel R Toolkit) seems to be installed but does not seem to be working." + vbLf + vbLf + "Please run the BERT installer again, which is available from " + BERTURL
38            GoTo PostErrorMessage
39        End If

          Dim VersionInstalled As String
40        On Error Resume Next
          'Some versions of BERT lack the BERT.version list (or the list may have been removed from memory _
           See https://github.com/sdllc/Basic-Excel-R-Toolkit/issues/103
          '        VersionInstalled = CStr(Application.Run("BERT.Exec", "BERT.version$version.string"))
41        VersionInstalled = CStr(Application.Run("BERT.Exec", "paste0(""Bert Version "",Sys.getenv(""BERT_VERSION""))"))
42        On Error GoTo ErrHandler
43        If VersionInstalled = "" Or InStr(VersionInstalled, "Error") > 0 Then
44            If VersionInstalled = "" Or InStr(VersionInstalled, "Error") > 0 Then VersionInstalled = "Cannot determine BERT version"
45        End If

46        If (LCase(VersionInstalled) <> LCase(VersionRequired)) And (LCase(VersionInstalled) <> LCase(VersionRequired_2)) Then
47            If VersionRequired = VersionRequired_2 Then
48                ErrorMessage = "There is a problem. BERT (Basic Excel R Toolkit) is installed, but the version is '" + VersionInstalled + "' instead of '" + VersionRequired + "'" + vbLf + vbLf + "Please run the BERT installer again, which is available from " + BERTURL
49            Else
50                ErrorMessage = "There is a problem. BERT (Basic Excel R Toolkit) is installed, but the version is '" + VersionInstalled + "' instead of either '" + VersionRequired + "' or '" + VersionRequired_2 + "'" + vbLf + vbLf + "Please run the BERT installer again, which is available from " + BERTURL
51            End If
52            GoTo PostErrorMessage
53        End If

54        MainDotRLocation = Replace(gRSourcePath, "/", "\") + "SCRiPTMain.R"

55        If Not sFileExists(MainDotRLocation) Then
56            ErrorMessage = "There seems to be a problem with the installation of the Solum software," + vbLf + "for example, the file '" + MainDotRLocation + "' cannot be found." + vbLf + vbLf + _
                  "Please run the installation script (install.vbs) again to correct this."
57            Throw ErrorMessage, True
58        End If

59        TestInstallation = True
60        Exit Function

PostErrorMessage:
          Dim MsgBoxRes As VbMsgBoxResult
          
61        If InStr(ErrorMessage, BERTURL) > 0 Then
62            MsgBoxRes = MsgBoxPlus(ErrorMessage, _
                  vbInformation + vbYesNo, "Solum", "OK, I'll do that later", "Go to website", , , 350)
63        Else
64            MsgBoxRes = MsgBoxPlus(ErrorMessage, _
                  vbInformation + vbOKOnly, "Solum", , , , , 350)
65        End If
              
66        Select Case MsgBoxRes
              Case vbYes, vbOK
67                Exit Function
68            Case vbNo
69                ThisWorkbook.FollowHyperlink Address:=BERTURL, NewWindow:=True
70        End Select

71        Exit Function
ErrHandler:
72        SomethingWentWrong "#TestInstallation (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

