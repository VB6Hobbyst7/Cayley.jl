Attribute VB_Name = "modOptionCompareText"
Option Explicit
Option Compare Text
Option Private Module        ' stop the functions in this module being visible from Excel

' For Syntax for Pattern, see:
'See https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/like-operator
Function SafeLike(TheString As String, Pattern As String)
1         On Error GoTo ErrHandler
2         SafeLike = TheString Like Pattern
3         Exit Function
ErrHandler:
4         SafeLike = "#" + Err.Description + "!"
End Function
