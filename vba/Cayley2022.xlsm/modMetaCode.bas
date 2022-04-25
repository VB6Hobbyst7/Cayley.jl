Attribute VB_Name = "modMetaCode"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RefactorCode
' Author     : Philip Swannell
' Date       : 11-Mar-2022
' Purpose    : Code to amend code. Idea is to automate the task of doing multiple "search and replace"
'             operations within a chunk of code. Each search and replace is done on a case sensitive, "whole word"
'             basis, and the order in which the search and replace operations are performed is by decreasing length
'             of the SearchFors. This was useful when refactoring method GetTradesInJuliaFormat to cope with reforms
'             to column names in the csv files.

' Parameters :
'  ExistingCode : The contents of a method or module - cut and paste from VBA (or VSCode or whatever)
'  SearchFors   : Column array of words or phrases to search for.
'  ReplaceWiths : Column array of replacements.
' -----------------------------------------------------------------------------------------------------------------------
Function RefactorCode(ByVal ExistingCode, ByVal SearchFors, ByVal ReplaceWiths)

          Dim ArrayToSort
          Dim i As Long
          Dim k As Long
          Dim l As Long
          Dim NCReplaceWiths As Long
          Dim NCSearchFors As Long
          Dim NCText As Long
          Dim NRReplaceWiths As Long
          Dim NRSearchFors As Long
          Dim NRText As Long
          Dim WordLengths
          Const CaseSensitive = True

1         On Error GoTo ErrHandler
2         Force2DArrayR ExistingCode, NRText, NCText
3         Force2DArrayR SearchFors, NRSearchFors, NCSearchFors
4         Force2DArrayR ReplaceWiths, NRReplaceWiths, NCReplaceWiths

5         If NRSearchFors <> NRReplaceWiths Or NCSearchFors <> 1 Or NCReplaceWiths <> 1 Then
6             Throw "SearchFors and ReplaceWiths must be arrays of the same number of rows and one column"
7         End If

8         WordLengths = sReshape(0, NRSearchFors, 1)
9         For i = 1 To NRSearchFors
10            WordLengths(i, 1) = Len(SearchFors(i, 1))
11        Next i

12        ArrayToSort = sArrayRange(SearchFors, ReplaceWiths, WordLengths)
13        ArrayToSort = sSortedArray(ArrayToSort, 3, , , False)

14        SearchFors = sSubArray(ArrayToSort, 1, 1, , 1)
15        ReplaceWiths = sSubArray(ArrayToSort, 1, 2, , 1)

          'In case searchfor strings contain characters that have to be escaped
16        SearchFors = sRegExpFromLiteral(SearchFors)
          'Whole word only
17        SearchFors = sArrayConcatenate("\b", SearchFors, "\b")

18        RefactorCode = sArrayRange(SearchFors, ReplaceWiths)

19        For k = 1 To NRSearchFors
20            For l = 1 To NCSearchFors
21                ExistingCode = sRegExReplace(ExistingCode, CStr(SearchFors(k, l)), CStr(ReplaceWiths(k, l)), CaseSensitive)
22            Next
23        Next

24        RefactorCode = ExistingCode

25        Exit Function
ErrHandler:
26        RefactorCode = "#RefactorCode (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

