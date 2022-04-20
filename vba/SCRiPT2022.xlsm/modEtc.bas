Attribute VB_Name = "modEtc"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : OtherBooksAreOpen
' Author     : Philip Swannell
' Date       : 07-Mar-2022
' Purpose    :
' -----------------------------------------------------------------------------------------------------------------------
Function OtherBooksAreOpen(ByRef MarketBookIsOpen As Boolean, ByRef LinesBookIsOpen As Boolean) As Boolean
          Dim LinesBookName As String
          Dim MarketBookName As String

1         On Error GoTo ErrHandler
2         LinesBookName = FileFromConfig("LinesWorkbook")
3         MarketBookName = FileFromConfig("MarketDataWorkbook")

4         LinesBookName = ThrowIfError(sSplitPath(LinesBookName))
5         MarketBookName = ThrowIfError(sSplitPath(MarketBookName))

6         If IsInCollection(Application.Workbooks, MarketBookName) Then
7             MarketBookIsOpen = True
8         End If

9         If IsInCollection(Application.Workbooks, LinesBookName) Then
10            LinesBookIsOpen = True
11        End If

12        OtherBooksAreOpen = MarketBookIsOpen And LinesBookIsOpen

13        Exit Function
ErrHandler:
14        Throw "#OtherBooksAreOpen (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
