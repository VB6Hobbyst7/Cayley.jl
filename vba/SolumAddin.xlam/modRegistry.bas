Attribute VB_Name = "modRegistry"
' -----------------------------------------------------------------------------------------------------------------------
' Module    : Module1
' Author    : Philip Swannell
' Date      : 25-Apr-2016
' Purpose   : Manipulating the Windows Registry. Code adapted from  http://vba-corner.livejournal.com/3054.html
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRegistryRead
' Author    : Philip Swannell
' Date      : 25-Apr-2016
' Purpose   : Returns the value of a key from the Windows Registry.
' Arguments
' RegKey    : The key in the Registry. Must start with one of the following strings: HKEY_CURRENT_USER\,
'             HKEY_LOCAL_MACHINE\, HKEY_CLASSES_ROOT\, HKEY_USERS\, HKEY_CURRENT_CONFIG
' -----------------------------------------------------------------------------------------------------------------------
Function sRegistryRead(RegKey As String) As String
Attribute sRegistryRead.VB_Description = "Returns the value of a key from the Windows Registry."
Attribute sRegistryRead.VB_ProcData.VB_Invoke_Func = " \n28"
          Dim myWS As Object

          'access Windows scripting
1         On Error GoTo ErrHandler
2         ValidateRegKey RegKey

3         Set myWS = CreateObject("WScript.Shell")
          'read key from registry
4         sRegistryRead = myWS.RegRead(RegKey)        ' See https://msdn.microsoft.com/en-us/library/x05fawxd(v=vs.84).aspx
5         Exit Function
ErrHandler:
6         sRegistryRead = "#sRegistryRead (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ValidateRegKey
' Author    : Philip Swannell
' Date      : 25-Apr-2016
' Purpose   : Throws an error if the input RegKey does not start with one of the allowed values
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ValidateRegKey(RegKey As String)
          Const ErrorString = "RegKey must be a string starting with one of the following ""roots"": HKEY_CURRENT_USER\, HKEY_LOCAL_MACHINE\, HKEY_CLASSES_ROOT\, HKEY_USERS\, HKEY_CURRENT_CONFIG"
1         If InStr(RegKey, "\") = 0 Then Throw ErrorString
2         Select Case Left$(RegKey, InStr(RegKey, "\") - 1)
              Case "HKEY_CURRENT_USER", "HKCU", "HKEY_LOCAL_MACHINE", "HKLM", "HKEY_CLASSES_ROOT", "HKCR", "HKEY_USERS", "HKEY_CURRENT_CONFIG"
3             Case Else
4                 Throw ErrorString
5         End Select
End Sub
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRegistryKeyExists
' Author    : Philip Swannell
' Date      : 25-Apr-2016
' Purpose   : Determines whether a particular key exists in the Windows Registry. Returns True if it
'             exists, False if it does not exist.
' Arguments
' RegKey    : The key in the Registry. Must start with one of the following strings: HKEY_CURRENT_USER\,
'             HKEY_LOCAL_MACHINE\, HKEY_CLASSES_ROOT\, HKEY_USERS\, HKEY_CURRENT_CONFIG
' -----------------------------------------------------------------------------------------------------------------------
Function sRegistryKeyExists(RegKey As String) As Boolean
Attribute sRegistryKeyExists.VB_Description = "Determines whether a particular key exists in the Windows Registry. Returns True if it exists, False if it does not exist."
Attribute sRegistryKeyExists.VB_ProcData.VB_Invoke_Func = " \n28"
          Dim myWS As Object
1         On Error GoTo ErrorHandler
2         Set myWS = CreateObject("WScript.Shell")
3         myWS.RegRead RegKey
4         sRegistryKeyExists = True
5         Exit Function
ErrorHandler:
6         sRegistryKeyExists = False
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRegistryWrite
' Author    : Philip Swannell
' Date      : 25-Apr-2016
' Purpose   : USE THIS FUNCTION CAREFULLY, mis-use could make your PC malfunction!
'             Creates a new key in the Windows Registry, adds another value-name to an
'             existing key (and assigns it a value), or changes the value of an existing
'             value-name.
' Arguments
' RegKey    : The key in the Registry. Specify a key-name by ending RegKey with a final backslash. Do
'             not include a final backslash to specify a value name.
' NewValue  : The value to write
' TheType   : The type of the value to write. Should be REG_SZ if NewValue is a string or REG_DWORD if
'             NewValue is an integer.
'
' Notes     : This function "wraps" the RegWrite method of the Windows Script Host. Full documentation
'             at https://msdn.microsoft.com/en-us/library/yfdfhz1b(v=vs.84).aspx
' -----------------------------------------------------------------------------------------------------------------------
Function sRegistryWrite(RegKey As String, _
        NewValue As String, _
        Optional TheType As String = "REG_SZ")
Attribute sRegistryWrite.VB_Description = "USE THIS FUNCTION CAREFULLY, mis-use could make your PC malfunction!\nCreates a new key in the Windows Registry, adds another value-name to an existing key (and assigns it a value), or changes the value of an existing value-name."
Attribute sRegistryWrite.VB_ProcData.VB_Invoke_Func = " \n28"
          Dim myWS As Object

1         On Error GoTo ErrHandler
2         ValidateRegKey RegKey
3         Select Case TheType
              Case "REG_SZ", "REG_EXPAND_SZ", "REG_DWORD"
4             Case Else
5                 Throw "TheType must be one of REG_SZ, REG_EXPAND_SZ or REG_DWORD"
6         End Select

7         Set myWS = CreateObject("WScript.Shell")
8         myWS.RegWrite RegKey, NewValue, TheType        'See https://msdn.microsoft.com/en-us/library/yfdfhz1b(v=vs.84).aspx
9         Exit Function
ErrHandler:
10        sRegistryWrite = "#sRegistryWrite (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : sRegistryDelete
' Author    : Philip Swannell
' Date      : 25-Apr-2016
' Purpose   : Writes information to a key in the Windows Registry.
'             USE THIS FUNCTION CAREFULLY, mis-use could make your PC malfunction!
'             The function returns True if the deletion is successful, False if not
'             successful or an error if RegKey is mal-formed.
' Arguments
' RegKey    : The key in the Registry to be deleted. Must start with one of the following strings:
'             HKEY_CURRENT_USER\, HKEY_LOCAL_MACHINE\, HKEY_CLASSES_ROOT\, HKEY_USERS\,
'             HKEY_CURRENT_CONFIG
' -----------------------------------------------------------------------------------------------------------------------
Function sRegistryDelete(RegKey As String)
Attribute sRegistryDelete.VB_Description = "Writes information to a key in the Windows Registry.\nUSE THIS FUNCTION CAREFULLY, mis-use could make your PC malfunction!\nThe function returns True if the deletion is successful, False if not successful or an error if RegKey is mal-formed."
Attribute sRegistryDelete.VB_ProcData.VB_Invoke_Func = " \n28"
          Dim myWS As Object

1         On Error GoTo ErrHandler
2         ValidateRegKey RegKey
3         On Error GoTo ErrHandler2
4         Set myWS = CreateObject("WScript.Shell")
5         myWS.RegDelete RegKey
6         sRegistryDelete = True
7         Exit Function
ErrHandler2:
8         sRegistryDelete = False
9         Exit Function
ErrHandler:
10        sRegistryDelete = "#sRegistryDelete (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
