Attribute VB_Name = "modCreateSystemImage"
Option Explicit

'THIS CODE COPIED FROM CAYLEY2022.XLSM
'It has a dependency on JuliaExcel, so can't live in one of the addins...
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
                      "that is dated " + Format(sFileInfo(SystemImage, "C"), "dd-mmm-yyyy hh:mm") & "."
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
18        ThrowIfError julialaunch(UnderLinux, False, , , "export JULIA_NUM_THREADS=8", 60)

19        JuliaEval "using Cayley;Cayley.create_system_image()"

20        Exit Sub
ErrHandler:
21        Throw "#JuliaCreateSystemImage (line " & CStr(Erl) & "): " & Err.Description & "!"
End Sub


