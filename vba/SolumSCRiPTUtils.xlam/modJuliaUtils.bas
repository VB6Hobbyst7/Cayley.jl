Attribute VB_Name = "modJuliaUtils"
Option Explicit
Public Const gSysImageXVALinux = "c:\Users\Public\Solum\XVA_Linux.sox"
Public Const gSysImageXVAWindows = "c:\Users\Public\Solum\XVA_Windows.sox"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MorphSlashes
' Author     : Philip Swannell
' Date       : 07-Dec-2021
' Purpose    : Given a file's name, windows style, returns the address of the file as seen from Julia, which may be
'              running under Linux (if UseLinux = True).
' -----------------------------------------------------------------------------------------------------------------------
Function MorphSlashes(WindowsFileLocation As String, UnderLinux As Boolean)
          Dim Res As String
1         Res = Replace(WindowsFileLocation, "\", "/")
2         If UnderLinux Then
3             If Mid(Res, 2, 2) = ":/" Then
4                 Res = "/mnt/" & LCase(Left(Res, 1)) & Mid(Res, 3)
5             End If
6         End If
7         MorphSlashes = Res
End Function


'' -----------------------------------------------------------------------------------------------------------------------
'' Procedure  : JuliaCreateSystemImage
'' Author     : Philip Swannell
'' Purpose    : Create a system image including compiled package XVA.
'' N.B. Requires that a compiler has been installed for use from WSL.
'' See https://linuxize.com/post/how-to-install-gcc-compiler-on-ubuntu-18-04/
'' -----------------------------------------------------------------------------------------------------------------------
'Sub JuliaCreateSystemImage(Ask As Boolean, UnderLinux As Boolean)
'          Dim SystemImage As String
'          Dim IncludeFile As String
'          Dim IncludeFileContents
'          Dim Prompt As String
'          Dim OS As String
'
'1         On Error GoTo ErrHandler
'
'2         OS = IIf(UnderLinux, "Linux", "Windows")
'3         SystemImage = IIf(UnderLinux, gSysImageXVALinux, gSysImageXVAWindows)
'
'4         If Ask Then
'5             Prompt = "Create Julia system image file under " & OS & "?" & vbLf & vbLf & _
'                  "This process takes about 10 minutes, but makes the Julia code run faster the first time it's used " & _
'                  "in a given session by doing ahead-of-time instead of just-in-time compilation of the Julia code " & _
'                  "in the XVA package."
'
'6             If sFileExists(SystemImage) Then
'7                 Prompt = Prompt & vbLf & vbLf & "The new system image will replace the existing one at:" & vbLf & _
'                      SystemImage & vbLf & _
'                      "that is dated " + Format(sFileInfo(SystemImage, "C"), "dd-mmm-yyyy hh:mm") & "."
'8             Else
'9                 Prompt = Prompt & vbLf & vbLf & "The system image will be at:" & vbLf & SystemImage
'10            End If
'11            If MsgBoxPlus(Prompt, vbOKCancel + vbQuestion, "Create System Image", "Create System Image") <> vbOK Then
'12                Throw "User Cancelled", True
'13            End If
'14        End If
'
'15        If JuliaIsRunning() Then
'16            JuliaEval "exit()"
'17        End If
'
'
'          'PackageCompiler ignores the `--threads=auto` command-line option and it's necessary to set the `JULIA_NUM_THREADS` environment variable
'18        ThrowIfError julialaunch(UnderLinux, False, , , "export JULIA_NUM_THREADS=8", 60)
'
'
'19        JuliaEval "using Cayley;Cayley.create_system_image()"
'
'
'          'PGS 15 March 2022 - old way below - don't delete until happy with calling function Cayley.create_system_image()
'
'
''20        IncludeFileContents = _
''              "if Sys.iswindows()" & vbLf & _
''              "    sysimage_path = """ & MorphSlashes(gSysImageXVAWindows, False) & """" & vbLf & _
''              "elseif Sys.islinux()" & vbLf & _
''              "    sysimage_path = """ & MorphSlashes(gSysImageXVALinux, True) & """" & vbLf & _
''              "end" & vbLf & _
''              vbLf & _
''              "# Better to delete file, since if it's locked then the process fails, but only after quite" & vbLf & _
''              "# some time.." & vbLf & _
''              "isfile(sysimage_path) && rm(sysimage_path)" & vbLf & _
''              vbLf & _
''              "using Pkg" & vbLf & _
''              "Pkg.activate()" & vbLf & _
''              "Pkg.add(""PackageCompiler"")" & vbLf & _
''              "using PackageCompiler" & vbLf & _
''              "using XVA" & vbLf & _
''              "packagefolder = pkgdir(XVA)" & vbLf & _
''              "precompile_execution_file = joinpath(packagefolder,""src"",""precompile_execution_file.jl"")" & vbLf & _
''              "PackageCompiler.create_sysimage([""XVA""];sysimage_path,precompile_execution_file)"
''
''21        IncludeFile = Environ("TEMP") & "\create_system_image.jl"
''22        ThrowIfError sFileSave(IncludeFile, IncludeFileContents, , , , "Unix")
''23        ThrowIfError JuliaInclude(MorphSlashes(IncludeFile, UnderLinux))
'
'24        Exit Sub
'ErrHandler:
'25        Throw "#JuliaCreateSystemImage (line " & CStr(Erl) & "): " & Err.Description & "!"
'End Sub





