Attribute VB_Name = "modEtc"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileFromConfig
' Author     : Philip Swannell
' Date       : 14-Mar-2022
' Purpose    : For better relocatability, files and folders on the Config sheet may be given as paths relative to
'              the folder of this workbook. This function returns a file's full path.
' -----------------------------------------------------------------------------------------------------------------------
Function FileFromConfig(NameOnConfig As String)

          Dim Res As String
1         On Error GoTo ErrHandler
2         Res = RangeFromSheet(shConfig, NameOnConfig, False, True, False, False, False).Value
3         FileFromConfig = sJoinPath(ThisWorkbook.Path, Res)

4         Exit Function
ErrHandler:
5         Throw "#FileFromConfig (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

