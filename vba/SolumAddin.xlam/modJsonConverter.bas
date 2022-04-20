Attribute VB_Name = "modJsonConverter"
''MODULE DOWNLOADED FROM GITHUB 22-Jan-2018
'Changes by PGS
'1) Added argument ArrayStyle to control how arrays are handled, with ArrayStyle = 2 (the default) arrays are more human-readable
'2) Simplified old argument WhiteSpace (which could be string or number) to new argument NumSpaces which must be just number
'3) Reduced use of On Error Resume Next by using my function NumDimensions
'4) Removed declarations for Mac and for Excel <= Office 2007
'5) Removed function json_BufferAppend. This was a way of avoiding the "Shlemiel The Painter" problem when building strings, but is slower and more complicated (uses Win API) than my clsAppendString
'6) throw errors when encountering objects of type other than dictionary or Collection, old code silently ignored
'7) Use Private function MyCStr instead of what was presumably intended to be use of CStr but in fact was not used
'8) Incorporated bug-fix suggested in pull request from "Sophist-UK" at https://github.com/VBA-tools/VBA-JSON/pull/44
'9) Re-wrote json_StringIsLargeNumber (original verison is below as json_StringIsLargeNumberORIGINAL)
'10) All functions made Private except ParseJson and ConvertToJson
'TODO contribute some of these changes to Tim Hall via his GitHub presence?

' VBA-JSON v2.2.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA
'
' Errors:
' 10001 - JSON parse error
'
' @class JsonConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' === VBA-UTC Headers

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

' === End VBA-UTC

Private Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `modJsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

Enum EnmArrayStyle
    as_Compact = 0    'All elements of arrays are written to the output with no line breaks or indentation
    AS_RowByRow = 1    '1D arrays written to a single line, 2D arrays written line-by-line
    as_ElementByElement = 2    ' every element written to its own line
End Enum

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseJsonFile
' Author     : Philip Swannell
' Date       : 18-Aug-2020
' Purpose    : Wrap to ParseJson but takes name of file whose contents are valid Json
' -----------------------------------------------------------------------------------------------------------------------
Function ParseJsonFile(FileName As String) As Object
          Dim FileContents As String
          Dim FSO As Scripting.FileSystemObject
          Dim t As TextStream
1         On Error GoTo ErrHandler
2         CheckFileNameIsAbsolute FileName
3         Set FSO = New Scripting.FileSystemObject
4         Set t = FSO.OpenTextFile(FileName, ForReading)
5         FileContents = t.ReadAll
6         t.Close: Set t = Nothing: Set FSO = Nothing
7         Set ParseJsonFile = ParseJson(FileContents)
8         Exit Function
ErrHandler:
9         Throw "#ParseJsonFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @method ParseJson
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
''
Public Function ParseJson(ByVal JsonString As String) As Object
          'PGS This is a change to Tim Hall's code. Can never imagine wanting some (i.e. >15 chars) strings that look like numbers to remain as strings while shorter strings get converted to numbers...
1         modJsonConverter.JsonOptions.UseDoubleForLargeNumbers = True

          Dim json_Index As Long
2         json_Index = 1

          ' Remove vbCr, vbLf, and vbTab from json_String
3         JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, vbNullString), VBA.vbLf, vbNullString), VBA.vbTab, vbNullString)

4         json_SkipSpaces JsonString, json_Index
5         Select Case VBA.Mid$(JsonString, json_Index, 1)
              Case "{"
6                 Set ParseJson = json_ParseObject(JsonString, json_Index)
7             Case "["
8                 Set ParseJson = json_ParseArray(JsonString, json_Index)
9             Case Else
                  ' Error: Invalid JSON string
10                Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
11        End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestCtJ
' Author     : Philip Swannell
' Date       : 26-Jan-2018
' Purpose    : Test harness for array handling in method ConvertToJson
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestCtJ()
          Dim Array1
          Dim Array2
          Dim ArrayStyle As Long
          Const NumSpaces = 1
          Const CurrentIndentation = 5

1         Array1 = VBA.Array("a", "b", "c")
2         Array2 = sReshape(sArrayStack("a", "b", "c", "d", "e", "f", "g", "h", "i"), 3, 3)

3         For ArrayStyle = 0 To 2
4             Debug.Print String(56, "-")
5             Debug.Print "1D, ArrayStyle = " & ArrayStyle & ", NumSpaces = " & NumSpaces & " CurrentIndentation = " & CurrentIndentation
6             Debug.Print "Previous lines..." & ConvertToJson(Array1, NumSpaces, ArrayStyle, CurrentIndentation)
7             Debug.Print String(56, "-")
8             Debug.Print "2D, ArrayStyle = " & ArrayStyle & ", NumSpaces = " & NumSpaces & " CurrentIndentation = " & CurrentIndentation
9             Debug.Print "Previous lines..." & ConvertToJson(Array2, NumSpaces, ArrayStyle, CurrentIndentation)
10            Debug.Print String(56, "-")
11        Next
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestRoundTripMulti
' Author     : Philip Swannell
' Date       : 19-Mar-2018
' Purpose    : Test all two character ascii strings for correct round tripping, ran after correcting handling of
'              line feed characters as suggested at https://github.com/VBA-tools/VBA-JSON/pull/44/files
'              i.e try to find other "edge cases" not correctly handled, but found none
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestRoundTripMulti()
          Dim i As Long
          Dim j As Long
1         For i = 1 To 255
2             For j = 1 To 255
3                 TestRoundTrip Chr$(i) & CStr(j)
4             Next
5         Next
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TestRoundTrip
' Author     : Philip Swannell
' Date       : 19-Mar-2018
' Purpose    : Test round-tripping of ConvertToJSON and ParseJSON for strings
' Parameters :
'  InputString:
' -----------------------------------------------------------------------------------------------------------------------
Private Sub TestRoundTrip(InputString)
          Dim DCT1 As Dictionary
          Dim DCT2 As Dictionary
          Dim JSON1 As String
          Dim JSON2 As String

1         On Error GoTo ErrHandler
2         Set DCT1 = New Dictionary
3         Set DCT2 = New Dictionary

4         DCT1.Add "AString", InputString
5         JSON1 = ConvertToJson(DCT1)

6         Set DCT2 = ParseJson(JSON1)
7         JSON2 = ConvertToJson(DCT2)

8         If JSON1 <> JSON2 Then
9             Debug.Print "InputString = " & InputString
10            Debug.Print "JSON1 = " & JSON1
11            Debug.Print "JSON2 = " & JSON2
12            Stop
13        End If

14        Exit Sub
ErrHandler:
15        Throw "#TestRoundTrip (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @method ConvertToJson
' @param {Variant} JsonValue (Dictionary, Collection, or Array)
' @param {Integer|String} NumSpaces "Pretty" print json with given number of spaces per indentation (Integer)
' @param {Boolean} ArrayStyle (PGS modification)
' @return {String}
''
Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal NumSpaces As Integer, Optional ByVal ArrayStyle As EnmArrayStyle, Optional ByVal json_CurrentIndentation As Long = 0) As String
          Dim cSA As clsStringAppend
          Dim json_Converted As String
          Dim json_DateStr As String
          Dim json_Indentation As String
          Dim json_Index As Long
          Dim json_Index2D As Long
          Dim json_InnerIndentation As String
          Dim json_IsFirstItem As Boolean
          Dim json_IsFirstItem2D As Boolean
          Dim json_Key As Variant
          Dim json_LBound As Long
          Dim json_LBound2D As Long
          Dim json_PrettyPrint As Boolean
          Dim json_SkipItem As Boolean
          Dim json_UBound As Long
          Dim json_UBound2D As Long
          Dim json_Value As Variant

1         Set cSA = New clsStringAppend
2         json_LBound = -1
3         json_UBound = -1
4         json_IsFirstItem = True
5         json_LBound2D = -1
6         json_UBound2D = -1
7         json_IsFirstItem2D = True
8         If NumSpaces < 0 Then Throw "NumSpaces must be zero or positive"
9         json_PrettyPrint = NumSpaces > 0
10        If NumSpaces = 0 Then ArrayStyle = as_Compact
11        If ArrayStyle < 0 Or ArrayStyle > 2 Then Throw "ArrayStyle must be 0, 1 or 2"

12        Select Case VBA.VarType(JsonValue)
              Case VBA.vbNull
13                ConvertToJson = "null"
14            Case VBA.vbDate
                  ' Date
15                json_DateStr = ConvertToIso(VBA.CDate(JsonValue))
16                ConvertToJson = """" & json_DateStr & """"
17            Case VBA.vbString
                  ' String (or large number encoded as string)
18                If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
19                    ConvertToJson = JsonValue
20                Else
21                    ConvertToJson = """" & json_Encode(JsonValue) & """"
22                End If
23            Case VBA.vbBoolean
24                If JsonValue Then
25                    ConvertToJson = "true"
26                Else
27                    ConvertToJson = "false"
28                End If
29            Case VBA.vbArray To VBA.vbArray + VBA.vbByte
30                If ArrayStyle >= AS_RowByRow Then
31                    json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * NumSpaces)
32                    json_InnerIndentation = VBA.Space$((json_CurrentIndentation + 2) * NumSpaces)
33                End If

                  ' Array
34                cSA.Append "["

                  Dim ND As Long
35                ND = NumDimensions(JsonValue)
36                If ND = 0 Then Throw "Unexpected error. Array has zero dimensions"    'Is that even possible?
37                If ND > 2 Then Throw "Cannot handle arrays with more than two dimensions"

38                json_LBound = LBound(JsonValue, 1)
39                json_UBound = UBound(JsonValue, 1)
40                If ND > 1 Then
41                    json_LBound2D = LBound(JsonValue, 2)
42                    json_UBound2D = UBound(JsonValue, 2)
43                End If

44                For json_Index = json_LBound To json_UBound
45                    If json_IsFirstItem Then
46                        json_IsFirstItem = False
47                    Else
                          ' Append comma to previous line
48                        cSA.Append ","
49                    End If

50                    If ND = 2 Then
                          ' 2D Array
51                        If json_PrettyPrint And (ArrayStyle = AS_RowByRow Or ArrayStyle = as_ElementByElement) Then
52                            cSA.Append vbNewLine + json_Indentation
53                        End If
54                        cSA.Append "["

55                        For json_Index2D = json_LBound2D To json_UBound2D
56                            If json_IsFirstItem2D Then
57                                json_IsFirstItem2D = False
58                            Else
59                                cSA.Append ","
60                            End If

61                            json_Converted = ConvertToJson(JsonValue(json_Index, json_Index2D), NumSpaces, ArrayStyle, json_CurrentIndentation + 2)

                              ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
62                            If json_Converted = vbNullString Then
                                  ' (nest to only check if converted = vbNullString)
63                                If json_IsUndefined(JsonValue(json_Index, json_Index2D)) Then
64                                    json_Converted = "null"
65                                End If
66                            End If

67                            If json_PrettyPrint And ArrayStyle = as_ElementByElement Then
68                                json_Converted = vbNewLine & json_InnerIndentation & json_Converted
69                            End If

70                            cSA.Append json_Converted
71                        Next json_Index2D

72                        If json_PrettyPrint And ArrayStyle = as_ElementByElement Then
73                            cSA.Append vbNewLine & json_Indentation
74                        End If

75                        cSA.Append "]"
76                        json_IsFirstItem2D = True
77                    Else
                          ' 1D Array
78                        json_Converted = ConvertToJson(JsonValue(json_Index), NumSpaces, ArrayStyle, json_CurrentIndentation + 1)

                          ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
79                        If json_Converted = vbNullString Then
                              ' (nest to only check if converted = vbNullString)
80                            If json_IsUndefined(JsonValue(json_Index)) Then
81                                json_Converted = "null"
82                            End If
83                        End If

84                        If ArrayStyle = as_ElementByElement Then
85                            json_Converted = vbNewLine & json_Indentation & json_Converted
86                        End If

87                        cSA.Append json_Converted
88                    End If
89                Next json_Index

90                If json_PrettyPrint Then
91                    json_Indentation = VBA.Space$(json_CurrentIndentation * NumSpaces)
92                    If ArrayStyle = as_ElementByElement Then
93                        cSA.Append vbNewLine & json_Indentation
94                    End If
95                End If

96                cSA.Append "]"

97                ConvertToJson = cSA.Report

                  ' Dictionary or Collection
98            Case VBA.vbObject
99                If json_PrettyPrint Then

100                   json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * NumSpaces)
101               End If

                  ' Dictionary
102               If VBA.TypeName(JsonValue) = "Dictionary" Then
103                   cSA.Append "{"
104                   For Each json_Key In JsonValue.Keys
                          ' For Objects, undefined (Empty/Nothing) is not added to object
105                       json_Converted = ConvertToJson(JsonValue(json_Key), NumSpaces, ArrayStyle, json_CurrentIndentation + 1)
106                       If json_Converted = vbNullString Then
107                           json_SkipItem = json_IsUndefined(JsonValue(json_Key))
108                       Else
109                           json_SkipItem = False
110                       End If

111                       If Not json_SkipItem Then
112                           If json_IsFirstItem Then
113                               json_IsFirstItem = False
114                           Else
115                               cSA.Append ","
116                           End If

117                           If json_PrettyPrint Then
118                               json_Converted = vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
119                           Else
120                               json_Converted = """" & json_Key & """:" & json_Converted
121                           End If

122                           cSA.Append json_Converted
123                       End If
124                   Next json_Key

125                   If json_PrettyPrint Then
126                       cSA.Append vbNewLine
127                       json_Indentation = VBA.Space$(json_CurrentIndentation * NumSpaces)
128                   End If

129                   cSA.Append json_Indentation & "}"

                      ' Collection
130               ElseIf VBA.TypeName(JsonValue) = "Collection" Then
131                   cSA.Append "["
132                   For Each json_Value In JsonValue
133                       If json_IsFirstItem Then
134                           json_IsFirstItem = False
135                       Else
136                           cSA.Append ","
137                       End If

138                       json_Converted = ConvertToJson(json_Value, NumSpaces, ArrayStyle, json_CurrentIndentation + 1)

                          ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
139                       If json_Converted = vbNullString Then
                              ' (nest to only check if converted = vbNullString)
140                           If json_IsUndefined(json_Value) Then
141                               json_Converted = "null"
142                           End If
143                       End If

144                       If json_PrettyPrint Then
145                           json_Converted = vbNewLine & json_Indentation & json_Converted
146                       End If

147                       cSA.Append json_Converted
148                   Next json_Value

149                   If json_PrettyPrint Then
150                       cSA.Append vbNewLine
151                       json_Indentation = VBA.Space$(json_CurrentIndentation * NumSpaces)
152                   End If

153                   cSA.Append json_Indentation & "]"
154               Else
155                   Throw "Cannot convert object of type " + TypeName(JsonValue) + " to string"
156               End If

157               ConvertToJson = cSA.Report
158           Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
                  ' Number (use decimals for numbers)
159               ConvertToJson = VBA.Replace(JsonValue, ",", ".")
160           Case Else
                  ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
                  ' Use VBA's built-in to-string
                  'On Error Resume Next
                  'ConvertToJson = JsonValue <-- PGS corrected this bug, 29/1/18
161               ConvertToJson = MyCStr(JsonValue)
                  'On Error GoTo 0
162       End Select
End Function

Private Function MyCStr(x As Variant)
1         On Error GoTo ErrHandler
2         MyCStr = CStr(x)
3         Exit Function
ErrHandler:
4         Throw "#MyCStr (line " & CStr(Erl) + "): " & "Cannot convert object of type " + TypeName(x) + " to a string"
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Dictionary
          Dim json_Key As String
          Dim json_NextChar As String

1         Set json_ParseObject = New Dictionary
2         json_SkipSpaces json_String, json_Index
3         If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
4             Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
5         Else
6             json_Index = json_Index + 1

7             Do
8                 json_SkipSpaces json_String, json_Index
9                 If VBA.Mid$(json_String, json_Index, 1) = "}" Then
10                    json_Index = json_Index + 1
11                    Exit Function
12                ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
13                    json_Index = json_Index + 1
14                    json_SkipSpaces json_String, json_Index
15                End If

16                json_Key = json_ParseKey(json_String, json_Index)
17                json_NextChar = json_Peek(json_String, json_Index)
18                If json_NextChar = "[" Or json_NextChar = "{" Then
19                    Set json_ParseObject.item(json_Key) = json_ParseValue(json_String, json_Index)
20                Else
21                    json_ParseObject.item(json_Key) = json_ParseValue(json_String, json_Index)
22                End If
23            Loop
24        End If
End Function

Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
1         Set json_ParseArray = New Collection

2         json_SkipSpaces json_String, json_Index
3         If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
4             Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
5         Else
6             json_Index = json_Index + 1

7             Do
8                 json_SkipSpaces json_String, json_Index
9                 If VBA.Mid$(json_String, json_Index, 1) = "]" Then
10                    json_Index = json_Index + 1
11                    Exit Function
12                ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
13                    json_Index = json_Index + 1
14                    json_SkipSpaces json_String, json_Index
15                End If

16                json_ParseArray.Add json_ParseValue(json_String, json_Index)
17            Loop
18        End If
End Function

Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
1         json_SkipSpaces json_String, json_Index
2         Select Case VBA.Mid$(json_String, json_Index, 1)
              Case "{"
3                 Set json_ParseValue = json_ParseObject(json_String, json_Index)
4             Case "["
5                 Set json_ParseValue = json_ParseArray(json_String, json_Index)
6             Case """", "'"
7                 json_ParseValue = json_ParseString(json_String, json_Index)
8             Case Else
9                 If VBA.Mid$(json_String, json_Index, 4) = "true" Then
10                    json_ParseValue = True
11                    json_Index = json_Index + 4
12                ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
13                    json_ParseValue = False
14                    json_Index = json_Index + 5
15                ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
16                    json_ParseValue = Null
17                    json_Index = json_Index + 4
18                ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
19                    json_ParseValue = json_ParseNumber(json_String, json_Index)
20                Else
21                    Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
22                End If
23        End Select
End Function

Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
          Dim cSA As clsStringAppend
          Dim json_Char As String
          Dim json_Code As String
          Dim json_Quote As String

1         Set cSA = New clsStringAppend

2         json_SkipSpaces json_String, json_Index

          ' Store opening quote to look for matching closing quote
3         json_Quote = VBA.Mid$(json_String, json_Index, 1)
4         json_Index = json_Index + 1

5         Do While json_Index > 0 And json_Index <= Len(json_String)
6             json_Char = VBA.Mid$(json_String, json_Index, 1)

7             Select Case json_Char
                  Case "\"
                      ' Escaped string, \\, or \/
8                     json_Index = json_Index + 1
9                     json_Char = VBA.Mid$(json_String, json_Index, 1)

10                    Select Case json_Char
                          Case """", "\", "/", "'"
11                            cSA.Append json_Char
12                            json_Index = json_Index + 1
13                        Case "b"
14                            cSA.Append vbBack
15                            json_Index = json_Index + 1
16                        Case "f"
17                            cSA.Append vbFormFeed
18                            json_Index = json_Index + 1
19                        Case "n"
20                            If VBA.Mid$(json_String, json_Index + 1, 2) = "\r" Then
21                                cSA.Append vbCrLf
22                                json_Index = json_Index + 3
23                            Else
24                                cSA.Append vbLf
25                                json_Index = json_Index + 1
26                            End If
27                        Case "r"
28                            If VBA.Mid$(json_String, json_Index + 1, 2) = "\n" Then
29                                cSA.Append vbCrLf
30                                json_Index = json_Index + 3
31                            Else
32                                cSA.Append vbCr
33                                json_Index = json_Index + 1
34                            End If
35                        Case "t"
36                            cSA.Append vbTab
37                            json_Index = json_Index + 1
38                        Case "u"
                              ' Unicode character escape (e.g. \u00a9 = Copyright)
39                            json_Index = json_Index + 1
40                            json_Code = VBA.Mid$(json_String, json_Index, 4)
41                            cSA.Append VBA.ChrW$(VBA.Val("&h" + json_Code))
42                            json_Index = json_Index + 4
43                    End Select
44                Case json_Quote
45                    json_ParseString = cSA.Report()
46                    json_Index = json_Index + 1
47                    Exit Function
48                Case Else
49                    cSA.Append json_Char
50                    json_Index = json_Index + 1
51            End Select
52        Loop
End Function

Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
          Dim json_Char As String
          Dim json_IsLargeNumber As Boolean
          Dim json_Value As String

1         json_SkipSpaces json_String, json_Index

2         Do While json_Index > 0 And json_Index <= Len(json_String)
3             json_Char = VBA.Mid$(json_String, json_Index, 1)

4             If VBA.InStr("+-0123456789.eE", json_Char) Then
                  ' Unlikely to have massive number, so use simple append rather than buffer here
5                 json_Value = json_Value & json_Char
6                 json_Index = json_Index + 1
7             Else
                  ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
                  ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
                  ' See: http://support.microsoft.com/kb/269370
                  '
                  ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
                  ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
8                 json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
9                 If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
10                    json_ParseNumber = json_Value
11                Else
                      ' VBA.Val does not use regional settings, so guard for comma is not needed
12                    json_ParseNumber = VBA.Val(json_Value)
13                End If
14                Exit Function
15            End If
16        Loop
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
          ' Parse key with single or double quotes
1         If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
2             json_ParseKey = json_ParseString(json_String, json_Index)
3         ElseIf JsonOptions.AllowUnquotedKeys Then
              Dim json_Char As String
4             Do While json_Index > 0 And json_Index <= Len(json_String)
5                 json_Char = VBA.Mid$(json_String, json_Index, 1)
6                 If (json_Char <> " ") And (json_Char <> ":") Then
7                     json_ParseKey = json_ParseKey & json_Char
8                     json_Index = json_Index + 1
9                 Else
10                    Exit Do
11                End If
12            Loop
13        Else
14            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
15        End If

          ' Check for colon and skip if present or throw if not present
16        json_SkipSpaces json_String, json_Index
17        If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
18            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
19        Else
20            json_Index = json_Index + 1
21        End If
End Function

Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
          ' Empty / Nothing -> undefined
1         Select Case VBA.VarType(json_Value)
              Case VBA.vbEmpty
2                 json_IsUndefined = True
3             Case VBA.vbObject
4                 Select Case VBA.TypeName(json_Value)
                      Case "Empty", "Nothing"
5                         json_IsUndefined = True
6                 End Select
7         End Select
End Function

Private Function json_Encode(ByVal json_Text As Variant) As String
          ' Reference: http://www.ietf.org/rfc/rfc4627.txt
          ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
          Dim cSA As clsStringAppend
          Dim json_AscCode As Long
          Dim json_Char As String
          Dim json_Index As Long
          
1         Set cSA = New clsStringAppend

2         For json_Index = 1 To VBA.Len(json_Text)
3             json_Char = VBA.Mid$(json_Text, json_Index, 1)
4             json_AscCode = VBA.AscW(json_Char)

              ' When AscW returns a negative number, it returns the twos complement form of that number.
              ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
              ' https://support.microsoft.com/en-us/kb/272138
5             If json_AscCode < 0 Then
6                 json_AscCode = json_AscCode + 65536
7             End If

              ' From spec, ", \, and control characters must be escaped (solidus is optional)

8             Select Case json_AscCode
                  Case 34
                      ' " -> 34 -> \"
9                     json_Char = "\"""
10                Case 92
                      ' \ -> 92 -> \\
11                    json_Char = "\\"
12                Case 47
                      ' / -> 47 -> \/ (optional)
13                    If JsonOptions.EscapeSolidus Then
14                        json_Char = "\/"
15                    End If
16                Case 8
                      ' backspace -> 8 -> \b
17                    json_Char = "\b"
18                Case 12
                      ' form feed -> 12 -> \f
19                    json_Char = "\f"
20                Case 10
                      ' line feed -> 10 -> \n
21                    json_Char = "\n"
22                Case 13
                      ' carriage return -> 13 -> \r
23                    json_Char = "\r"
24                Case 9
                      ' tab -> 9 -> \t
25                    json_Char = "\t"
26                Case 0 To 31, 127 To 65535
                      ' Non-ascii characters -> convert to 4-digit hex
27                    json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
28            End Select

29            cSA.Append json_Char
30        Next json_Index

31        json_Encode = cSA.Report
End Function

Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
          ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
1         json_SkipSpaces json_String, json_Index
2         json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
          ' Increment index to skip over spaces
1         Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
2             json_Index = json_Index + 1
3         Loop
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : json_StringIsLargeNumber
' Author     : Philip Swannell
' Date       : 19-Mar-2018
' Purpose    : Significant re-write of code downloaded from GitHub, a copy of which is below as json_StringIsLargeNumberORIGINAL
'              Code written with reference to the last "number" diagram at https://www.json.org/
' -----------------------------------------------------------------------------------------------------------------------
Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
          ' Check if the given string is considered a "large number"
          ' (See json_ParseNumber)

          Dim json_CharIndex As Long
          Dim json_DotCount As Long
          Dim json_ECount As Long
          Dim json_Length As Long

1         json_Length = VBA.Len(json_String)

          ' Length with be at least 16 characters and assume will be less than 100 characters
2         If json_Length >= 16 And json_Length <= 100 Then
              Dim json_CharCode As String

3             json_StringIsLargeNumber = True

4             For json_CharIndex = 1 To json_Length
5                 json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
6                 Select Case json_CharCode
                      Case 45    ' Minus sign. Allowed as first character or immediately after E or e, cannot be the last character.
7                         If json_CharIndex = json_Length Then
8                             json_StringIsLargeNumber = False
9                             Exit Function
10                        End If
11                        If json_CharIndex > 1 Then
12                            If VBA.UCase$(VBA.Mid$(json_String, json_CharIndex - 1, 1)) <> "E" Then
13                                json_StringIsLargeNumber = False
14                                Exit Function
15                            End If
16                        End If
17                    Case 43    ' Plus sign. Only allowed immediately after E or e, and cannot be the first or last character.
18                        If json_CharIndex = 1 Or json_CharIndex = json_Length Then
19                            json_StringIsLargeNumber = False
20                            Exit Function
21                        ElseIf VBA.UCase$(VBA.Mid$(json_String, json_CharIndex - 1, 1)) <> "E" Then
22                            json_StringIsLargeNumber = False
23                            Exit Function
24                        End If
25                    Case 69, 101    'E (or e). Cannot be the first or last character and can be only one occurrence.
26                        If json_CharIndex = 1 Or json_CharIndex = json_Length Then
27                            json_StringIsLargeNumber = False
28                            Exit Function
29                        End If
30                        json_ECount = json_ECount + 1
31                        If json_ECount > 1 Then
32                            json_StringIsLargeNumber = False
33                            Exit Function
34                        End If
35                    Case 46    ' Decimal point. Can only be one occurrence and cannot appear after the E (or e), cannot be the first character or the last character.
36                        If json_ECount > 0 Then
37                            json_StringIsLargeNumber = False
38                            Exit Function
39                        End If
40                        If json_CharIndex = 1 Or json_CharIndex = json_Length Then
41                            json_StringIsLargeNumber = False
42                            Exit Function
43                        End If
44                        json_DotCount = json_DotCount + 1
45                        If json_DotCount > 1 Then
46                            json_StringIsLargeNumber = False
47                            Exit Function
48                        End If
49                    Case 48 To 57    '0 thru 9 are valid
                          ' Continue through characters
50                    Case Else
51                        json_StringIsLargeNumber = False
52                        Exit Function
53                End Select
54            Next json_CharIndex
55        End If
End Function

'Version without changes by PGS
Private Function json_StringIsLargeNumberORIGINAL(json_String As Variant) As Boolean
          ' Check if the given string is considered a "large number"
          ' (See json_ParseNumber)

          Dim json_CharIndex As Long
          Dim json_Length As Long
1         json_Length = VBA.Len(json_String)

          ' Length with be at least 16 characters and assume will be less than 100 characters
2         If json_Length >= 16 And json_Length <= 100 Then
              Dim json_CharCode As String

3             json_StringIsLargeNumberORIGINAL = True

4             For json_CharIndex = 1 To json_Length
5                 json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
6                 Select Case json_CharCode
                      ' Look for .|0-9|E|e
                      Case 46, 48 To 57, 69, 101
                          ' Continue through characters
7                     Case Else
8                         json_StringIsLargeNumberORIGINAL = False
9                         Exit Function
10                End Select
11            Next json_CharIndex
12        End If
End Function

Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
          ' Provide detailed parse error message, including details of where and what occurred
          '
          ' Example:
          ' Error parsing JSON:
          ' {"abcde":True}
          '          ^
          ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

          Dim json_StartIndex As Long
          Dim json_StopIndex As Long

          ' Include 10 characters before and after error (if possible)
1         json_StartIndex = json_Index - 10
2         json_StopIndex = json_Index + 10
3         If json_StartIndex <= 0 Then
4             json_StartIndex = 1
5         End If
6         If json_StopIndex > VBA.Len(json_String) Then
7             json_StopIndex = VBA.Len(json_String)
8         End If

9         json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
              VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
              VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
              ErrorMessage
End Function

Private Function json_UnsignedAdd(json_Start As LongPtr, json_Increment As Long) As LongPtr
1         If json_Start And &H80000000 Then
2             json_UnsignedAdd = json_Start + json_Increment
3         ElseIf (json_Start Or &H80000000) < -json_Increment Then
4             json_UnsignedAdd = json_Start + json_Increment
5         Else
6             json_UnsignedAdd = (json_Start + &H80000000) + (json_Increment + &H80000000)
7         End If
End Function

''
' VBA-UTC v1.0.3
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error
''
Private Function ParseUtc(utc_UtcDate As Date) As Date
1         On Error GoTo utc_ErrorHandling

          Dim utc_LocalDate As utc_SYSTEMTIME
          Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION

2         utc_GetTimeZoneInformation utc_TimeZoneInfo
3         utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate

4         ParseUtc = utc_SystemTimeToDate(utc_LocalDate)

5         Exit Function

utc_ErrorHandling:
6         Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
Private Function ConvertToUtc(utc_LocalDate As Date) As Date
1         On Error GoTo utc_ErrorHandling

          Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
          Dim utc_UtcDate As utc_SYSTEMTIME

2         utc_GetTimeZoneInformation utc_TimeZoneInfo
3         utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate

4         ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)

5         Exit Function

utc_ErrorHandling:
6         Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

''
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
''
Private Function ParseIso(utc_IsoString As String) As Date
1         On Error GoTo utc_ErrorHandling

          Dim utc_DateParts() As String
          Dim utc_HasOffset As Boolean
          Dim utc_NegativeOffset As Boolean
          Dim utc_Offset As Date
          Dim utc_OffsetIndex As Long
          Dim utc_OffsetParts() As String
          Dim utc_Parts() As String
          Dim utc_TimeParts() As String

2         utc_Parts = VBA.Split(utc_IsoString, "T")
3         utc_DateParts = VBA.Split(utc_Parts(0), "-")
4         ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))

5         If UBound(utc_Parts) > 0 Then
6             If VBA.InStr(utc_Parts(1), "Z") Then
7                 utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", vbNullString), ":")
8             Else
9                 utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
10                If utc_OffsetIndex = 0 Then
11                    utc_NegativeOffset = True
12                    utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
13                End If

14                If utc_OffsetIndex > 0 Then
15                    utc_HasOffset = True
16                    utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
17                    utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")

18                    Select Case UBound(utc_OffsetParts)
                          Case 0
19                            utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
20                        Case 1
21                            utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
22                        Case 2
                              ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
23                            utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
24                    End Select

25                    If utc_NegativeOffset Then: utc_Offset = -utc_Offset
26                Else
27                    utc_TimeParts = VBA.Split(utc_Parts(1), ":")
28                End If
29            End If

30            Select Case UBound(utc_TimeParts)
                  Case 0
31                    ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
32                Case 1
33                    ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
34                Case 2
                      ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
35                    ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
36            End Select

37            ParseIso = ParseUtc(ParseIso)

38            If utc_HasOffset Then
39                ParseIso = ParseIso + utc_Offset
40            End If
41        End If

42        Exit Function

utc_ErrorHandling:
43        Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Private Function ConvertToIso(utc_LocalDate As Date) As String
1         On Error GoTo utc_ErrorHandling

2         ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")

3         Exit Function

utc_ErrorHandling:
4         Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
1         utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
2         utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
3         utc_DateToSystemTime.utc_wDay = VBA.day(utc_Value)
4         utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
5         utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
6         utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
7         utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
1         utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
              TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

