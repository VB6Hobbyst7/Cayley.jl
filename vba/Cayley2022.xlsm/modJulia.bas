Attribute VB_Name = "modJulia"
Option Explicit
'Dec 2021, adding code using Julia via JuliaExcel

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaEvalWrapper
' Author     : Philip Swannell
' Date       : 14-Feb-2022
' Purpose    : It's possible that the model does not exist in Julia when the VBA code expected that it does.
'              Example: The user shuts down Julia executable, which will then get automatically relaunched, but without
'              the model being brought back into existence. This method ensures a friendly error message in that case.
'              would be better to reinstantiate the model on-the-fly but that would involve passing the fxshock and
'              fxvolshock down the call stack.
' Parameters :
'  EvaluateThis:
'  ModelName   :
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaEvalWrapper(EvaluateThis As String, ModelName As String)

          Dim Expression As String
          Dim FriendlyError As String
1         On Error GoTo ErrHandler

2         MessageLogWrite "Calling Julia: " & EvaluateThis

3         Expression = "if(@isdefined(" & ModelName & "));return(" & EvaluateThis & ");else;return(""ModelNotDefined"");end"

4         Assign JuliaEvalWrapper, JuliaEvalVBA(Expression)

5         If VarType(JuliaEvalWrapper) = vbString Then
6             If JuliaEvalWrapper = "ModelNotDefined" Then
7                 FriendlyError = "The Hull-White model '" & ModelName & _
                      "' is not defined in Julia. Please use 'Menu' -> 'Build Hull-White Model' to recreate it."
8                 Throw FriendlyError, True
9             ElseIf Left(JuliaEvalWrapper, 1) = "#" Then
10                If Right(JuliaEvalWrapper, 1) = "!" Then
11                    Throw JuliaEvalWrapper
12                End If
13            End If
14        End If

15        Exit Function
ErrHandler:
16        Throw "#JuliaEvalWrapper (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Sub LaunchJuliaWithoutSystemImage()
          Const TimeOut = 90

1         On Error GoTo ErrHandler
2         If JuliaIsRunning Then
3             JuliaEval "exit()"
4         End If
5         StatusBarWrap "Launching Julia with timeout of " & CStr(TimeOut) & " seconds"
6         ThrowIfError JuliaLaunch(UseLinux(), False, " --threads=auto", "XVA,Cayley", , TimeOut)
7         StatusBarWrap False
8         Exit Sub
ErrHandler:
9         Throw "#LaunchJuliaWithoutSystemImage (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

Function UseLinux()
1         On Error GoTo ErrHandler
          'Airbus unable to install Linux under WSL, so stubbed this function to False
2         UseLinux = False
3         Exit Function
ErrHandler:
4         Throw "#UseLinux (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaLaunchForCayley
' Author     : Philip Swannell
' Date       : 21-Dec-2021
' Purpose    : Launch Julia with a system image in which the XVA package is pre-compiled.
' -----------------------------------------------------------------------------------------------------------------------
Sub JuliaLaunchForCayley()

          Dim CommandLineOptions As String
          Dim OS As String
          Dim Prompt As String
          Dim SystemImage As String
          Dim SystemImageX As String
          Dim XSH As clsExcelStateHandler
          Const TimeOut As Long = 60
          Dim SheetToActivate As Worksheet

1         On Error GoTo ErrHandler

2         If Not JuliaIsRunning() Then
              Set SheetToActivate = ActiveSheet
3             SystemImage = IIf(UseLinux(), gSysImageXVALinux, gSysImageXVAWindows)
4             SystemImageX = MorphSlashes(SystemImage, UseLinux())

5             Set XSH = CreateExcelStateHandler(, , , "Waiting for Julia to launch")
6             OS = IIf(UseLinux(), "Linux", "Windows")
7             If sFileExists(SystemImage) Then
8                 CommandLineOptions = " --threads auto --sysimage " & SystemImageX
9             Else
10                Prompt = "There is no ""system image"" available for Julia on " & OS & _
                      ". What would you like to do?" & vbLf & vbLf & _
                      "Recommended:" & vbLf & _
                      "Create a system image now. This will take about 15 minutes and will save time " & _
                      "in the future, by doing ""ahead-of-time"" compilation of the Julia code." & vbLf & vbLf & _
                      "Not recommended, except for developer use:" & vbLf & _
                      "Run Julia without a system image."
11                Select Case MsgBoxPlus(Prompt, vbYesNoCancel + vbQuestion, , _
                      "Create a system image now", "Run Julia without system image")
                      Case vbYes
12                        JuliaCreateSystemImage True, UseLinux()
13                        Exit Sub
14                    Case vbNo
15                        CommandLineOptions = " --threads auto"
16                    Case vbCancel
17                        Throw "User Cancelled", True
18                End Select
19            End If
20            StatusBarWrap "Launching Julia with a timeout of " & CStr(TimeOut) & " seconds"
21            ThrowIfError JuliaLaunch(UseLinux(), True, CommandLineOptions, "XVA,Cayley")
22            StatusBarWrap False
23            SafeAppActivate SheetToActivate
24        End If

25        Exit Sub
ErrHandler:
26        Throw "#JuliaLaunchForCayley (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PathInJulia
' Author     : Philip Swannell
' Date       : 17-Dec-2021
' Purpose    : Convert a file's Windows-style address for use in Julia, which may be running (on the local PC) on either
'              Linux under WSL or on Windows.
' Parameters :
'  WindowsAddress:
'  OnWSL         :
' -----------------------------------------------------------------------------------------------------------------------
Function PathInJulia(WindowsAddress As String, OnWSL As Boolean)
1         On Error GoTo ErrHandler
2         Select Case Mid(WindowsAddress, 2, 2)
              Case ":/", ":\"
3                 If OnWSL Then
4                     PathInJulia = "/mnt/" & LCase(Left(WindowsAddress, 1)) & Replace(Mid(WindowsAddress, 3), "\", "/")
5                 Else
6                     PathInJulia = Replace(WindowsAddress, "\", "/")
7                 End If
8             Case Else
9                 Throw "WindowsAddress must start with characters ""x:\"" for some drive-letter x"
10        End Select
11        Exit Function
ErrHandler:
12        Throw "#PathInJulia (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaCreateSystemImage
' Author     : Philip Swannell
' Purpose    : Create a system image including compiled package XVA.
' N.B. Requires that a compiler has been installed for use from WSL.
' See https://linuxize.com/post/how-to-install-gcc-compiler-on-ubuntu-18-04/
' -----------------------------------------------------------------------------------------------------------------------
Sub JuliaCreateSystemImage(Ask As Boolean, UnderLinux As Boolean)
          Dim OS As String
          Dim Prompt As String
          Dim SystemImage As String
            
1         On Error GoTo ErrHandler

2         OS = IIf(UnderLinux, "Linux", "Windows")
3         SystemImage = IIf(UnderLinux, gSysImageXVALinux, gSysImageXVAWindows)

4         If Ask Then
5             Prompt = "Create Julia system image file under " & OS & "?" & vbLf & vbLf & _
                  "This process takes about 10 minutes, but makes the Julia code run faster the first time it's used " & _
                  "in a given session by doing ahead-of-time instead of just-in-time compilation of the Julia code " & _
                  "in the XVA package."
              
6             If sFileExists(SystemImage) Then
7                 Prompt = Prompt & vbLf & vbLf & "The new system image will replace the existing one at:" & vbLf & _
                      SystemImage & vbLf & _
                      "that is dated " & Format(sFileInfo(SystemImage, "C"), "dd-mmm-yyyy hh:mm") & "."
8             Else
9                 Prompt = Prompt & vbLf & vbLf & "The system image will be at:" & vbLf & SystemImage
10            End If
11            If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, "Create System Image", "Create System Image") <> vbOK Then
12                Throw "User Cancelled", True
13            End If
14        End If

15        If JuliaIsRunning() Then
16            JuliaEval "exit()"
17        End If

          'PackageCompiler ignores the `--threads=auto` command-line option and it's necessary to set the `JULIA_NUM_THREADS` environment variable
18        ThrowIfError JuliaLaunch(UnderLinux, False, , , "export JULIA_NUM_THREADS=8", 60)

19        JuliaEval "using Cayley;Cayley.create_system_image()"

          'PGS 15 March 2022 - old way below - don't delete until happy with calling function Cayley.create_system_image()

          '20        IncludeFileContents = _
          '              "if Sys.iswindows()" & vbLf & _
          '              "    sysimage_path = """ & MorphSlashes(gSysImageXVAWindows, False) & """" & vbLf & _
          '              "elseif Sys.islinux()" & vbLf & _
          '              "    sysimage_path = """ & MorphSlashes(gSysImageXVALinux, True) & """" & vbLf & _
          '              "end" & vbLf & _
          '              vbLf & _
          '              "# Better to delete file, since if it's locked then the process fails, but only after quite" & vbLf & _
          '              "# some time.." & vbLf & _
          '              "isfile(sysimage_path) && rm(sysimage_path)" & vbLf & _
          '              vbLf & _
          '              "using Pkg" & vbLf & _
          '              "Pkg.activate()" & vbLf & _
          '              "Pkg.add(""PackageCompiler"")" & vbLf & _
          '              "using PackageCompiler" & vbLf & _
          '              "using XVA" & vbLf & _
          '              "packagefolder = pkgdir(XVA)" & vbLf & _
          '              "precompile_execution_file = joinpath(packagefolder,""src"",""precompile_execution_file.jl"")" & vbLf & _
          '              "PackageCompiler.create_sysimage([""XVA""];sysimage_path,precompile_execution_file)"
          '
          '21        IncludeFile = Environ("TEMP") & "\create_system_image.jl"
          '22        ThrowIfError sFileSave(IncludeFile, IncludeFileContents, , , , "Unix")
          '23        ThrowIfError JuliaInclude(MorphSlashes(IncludeFile, UnderLinux))

20        Exit Sub
ErrHandler:
21        Throw "#JuliaCreateSystemImage (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub


