Attribute VB_Name = "modBloomberg"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : IsBloombergInstalled
' Author    : Philip
' Date      : 05-Oct-2017
' Purpose   : Test if Bloomberg Addin is installed. Would be better to also test if data
'             is available (e.g. what if user is not logged in to Bloomberg?)
' -----------------------------------------------------------------------------------------------------------------------
Function IsBloombergInstalled() As Boolean
          Dim Res
1         On Error GoTo ErrHandler
2         Res = Application.Evaluate("=BToday(TRUE)")
3         IsBloombergInstalled = Not (IsError(Res))
4         Exit Function
ErrHandler:
5         Throw "#IsBloombergInstalled (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergTickerBasisSwap
' Author    : Philip Swannell
' Date      : 06-Jul-2016
' Purpose   : Encapsulate the generation of the Bloomberg code (argument to BDH and BDP)
'             for basis swaps
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergTickerBasisSwap(Ccy As String, Tenor As String)
1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler
          Dim Contributor As String
          Dim Part1 As String
          Dim Part2 As String
          Dim Part4 As String
          Const Part3 = " Curncy"

3         Part2 = "EBS"
4         Select Case Ccy
              Case "AED"
5                 Part1 = "UD": Contributor = "CMPN"
6             Case "AUD"
7                 Part1 = "AD": Contributor = "CMPN"
8             Case "CAD"
9                 Part1 = "CD": Contributor = "CMPN"
10            Case "CHF"
11                Part1 = "SF": Contributor = "CMPN"
12            Case "CNH"
13                Part1 = "CG": Contributor = "CMPN"
14            Case "CZK"
15                Part1 = "CK": Contributor = "CMPN"
16                Part2 = "EUBS"
17            Case "DKK"
18                Part1 = "DK": Contributor = "CMPN"
19            Case "EUR"
20                Part1 = "EU": Contributor = "CMPN": Part2 = "SWE"
21            Case "GBP"
22                Part1 = "BP": Contributor = "CMPN"
23            Case "HKD"
24                Part1 = "HD": Contributor = "CMPN"
25            Case "HUF"
26                Part1 = "HF": Contributor = "CMPN"
27            Case "JPY"
28                Part1 = "JY": Contributor = "CMPN"
29            Case "KRW"
30                Part1 = "KR": Contributor = "CMPN"
31            Case "MXN"
32                Part1 = "MP": Contributor = "CMPN"
33            Case "MYR"
34                Part1 = "MR": Contributor = "CMPN"
35            Case "NOK"
36                Part1 = "NK": Contributor = "CMPN"
37            Case "NZD"
38                Part1 = "ND": Contributor = "CMPN"
39            Case "QAR"
40                Part1 = "QR": Contributor = "CMPN"
41            Case "RON"
42                Part1 = "RN": Contributor = "CMPN"
43                Part2 = "EUBS"
44            Case "SAR"
45                Part1 = "SR": Contributor = "CMPN"
46            Case "SEK"
47                Part1 = "SK": Contributor = "CMPN"
48            Case "SGD"
49                Part1 = "SD": Contributor = "CMPN"
50            Case "THB"
51                Part1 = "TB": Contributor = "CMPN"
52            Case "TWD"
53                Part1 = "TR": Contributor = "CMPN"
54            Case "USD"
55                Part1 = "EUBS": Contributor = "CMPN": Part2 = ""
56            Case "ZAR"
57                Part1 = "SA": Contributor = "CMPN"
58            Case Else: Throw "Unrecognised Ccy"
59        End Select

60        If Right(Tenor, 1) = "M" Then
61            Part4 = Left(Tenor, Len(Tenor) - 1)
62            Part4 = Chr(64 + CDbl(Part4))    '1 > "A". 2 > "B" etc.
63            BloombergTickerBasisSwap = Part1 & Part2 & Part4 & " " & Contributor & Part3
64        ElseIf Right(Tenor, 1) = "Y" Then
65            BloombergTickerBasisSwap = Part1 & Part2 & Left(Tenor, Len(Tenor) - 1) & " " & Contributor & Part3
66        Else
67            BloombergTickerBasisSwap = "#Unrecognised Tenor!"
68        End If
69        Exit Function
ErrHandler:
70        BloombergTickerBasisSwap = "#BloombergTickerBasisSwap (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergTickerInterestRateSwap
' Author    : Philip Swannell
' Date      : 14-Jun-2016
' Purpose   : Encapsulate the generation of a Bloomberg code for each point on the (normal) swaption vol surface.
' -----------------------------------------------------------------------------------------------------------------------
'TODO bring logic here up to same standard as in SCRiPT by moving validator to SolumAddin

Function BloombergTickerInterestRateSwap(Ccy As String, Tenor As String, FixedFrequency As String, FloatingFrequency As String)
1         On Error GoTo ErrHandler
          Dim Contributor As String
          Dim MonthCodePrefix As String
          Dim MonthsUseLetters As Boolean
          Dim Part1 As String
          Dim Part2 As String
          Dim Part3 As String
          Dim Part4 As String
          Dim Part5 As String
          Const Part6 = " Curncy"
2         Part3 = Left(Tenor, Len(Tenor) - 1)

3         Select Case Ccy
              Case "AED"
4                 Part1 = "UD"
5                 Contributor = "CMPN"
6                 MonthCodePrefix = "EIBO" & Tenor
7                 If Right(Tenor, 1) = "M" Then
8                     BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
9                 ElseIf FixedFrequency = "Ann" Then
10                    Part2 = "SW"
11                Else
12                    Throw "FixedFrequency not recognised"
13                End If
14            Case "AUD"
15                Part1 = "AD"
16                Contributor = "CMPN"
17                MonthCodePrefix = Part1 & "BB" & Tenor
18                If Right(Tenor, 1) = "M" Then
19                    BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
20                ElseIf FixedFrequency = "Quart" Then
21                    Part3 = Part3 + "Q"
22                    Part2 = "SWAP"
23                ElseIf FixedFrequency = "Semi" Then
24                    Part2 = "SWAP"
25                Else
26                    Throw "FixedFrequency not recognised"
27                End If
28            Case "CAD"
29                Part1 = "CD"
30                Contributor = "BGN"
31                MonthCodePrefix = Part1 & String(3 - Len(Tenor) + 2, "0") & Tenor
32                If Right(Tenor, 1) = "M" Then
33                    BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
34                ElseIf FixedFrequency = "Semi" Then
35                    Part2 = "SW"
36                Else
37                    Throw "FixedFrequency not recognised"
38                End If
39                If FloatingFrequency = "Quart" Then
40                    Part2 = "SW"
41                    Part4 = ""
42                Else
43                    Throw "FloatingFrequency must be Quart"
44                End If
45            Case "CHF"
46                Part1 = "SF"
47                Contributor = "BGN"
48                MonthCodePrefix = Part1 & String(3 - Len(Tenor) + 2, "0") & Tenor
49                If Right(Tenor, 1) = "M" Then
50                    BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
51                ElseIf FixedFrequency = "Ann" Then
52                    Part2 = "SW"
53                Else
54                    Throw "FixedFrequency not recognised"
55                End If
56                If FloatingFrequency = "Semi" Then
57                    Part2 = "SA"
58                    Part4 = ""
59                ElseIf FloatingFrequency = "Quart" Then
60                    Part2 = "SW"
61                    Part4 = "V3"
62                Else
63                    Throw "FloatingFrequency not recognised"
64                End If
65            Case "CNH"
66                Part1 = "CG"
67                Contributor = "CMPN"
68                MonthCodePrefix = "HI" & Ccy & Tenor
69                If Right(Tenor, 1) = "M" Then
70                    BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
71                ElseIf FixedFrequency = "Quart" Then
72                    Part2 = "SWH"
73                Else
74                    Throw "FixedFrequency not recognised"
75                End If
76            Case "CZK"
77                Part1 = "CK"
78                Contributor = "CMPN"
79                MonthCodePrefix = "PRIB0" & Tenor
80                If Right(Tenor, 1) = "M" Then
81                    BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
82                ElseIf FixedFrequency = "Ann" Then
83                    Part2 = "SW"
84                Else
85                    Throw "FixedFrequency not recognised"
86                End If
87            Case "DKK"
88                Part1 = "DK"
89                Contributor = "CMPN"
90                MonthCodePrefix = "CIBO" & Format(CLng(Left(Tenor, Len(Tenor) - 1)), "00") & "M"
91                If Right(Tenor, 1) = "M" Then
92                    BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
93                ElseIf FixedFrequency = "Ann" Then
94                    Part2 = "SW"
95                Else
96                    Throw "FixedFrequency not recognised"
97                End If
98                If FloatingFrequency = "Semi" Then
99                    Part2 = "SW"
100                   Part4 = ""
101               ElseIf FloatingFrequency = "Quart" Then
102                   Part2 = "SW"
103                   Part4 = "V3"
104               Else
105                   Throw "FloatingFrequency not recognised"
106               End If
107           Case "EUR"
108               Part1 = "EU"
109               Contributor = "BGN"
110               MonthCodePrefix = Ccy & String(3 - Len(Tenor) + 1, "0") & Tenor
111               If Right(Tenor, 1) = "M" Then
112                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
113               ElseIf FixedFrequency = "Ann" Then
114                   Part2 = "SA"
115               Else
116                   Throw "FixedFrequency not recognised"
117               End If
118               If FloatingFrequency = "Semi" Then
119                   Part2 = "SA"
120                   Part4 = ""
121               ElseIf FloatingFrequency = "Quart" Then
122                   Part2 = "SW"
123                   Part4 = "V3"
124               Else
125                   Throw "FloatingFrequency not recognised"
126               End If

127           Case "GBP"
128               Part1 = "BP"
129               Contributor = "BGN"
130               MonthCodePrefix = Part1 & String(3 - Len(Tenor) + 2, "0") & Tenor
131               If Right(Tenor, 1) = "M" Then
132                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
133               ElseIf FixedFrequency = "Semi" Then
134                   Part2 = "SW"
135               Else
136                   Throw "FixedFrequency not recognised"
137               End If
138               If FloatingFrequency = "Semi" Then
139                   Part2 = "SA"
140                   Part4 = ""
141               ElseIf FloatingFrequency = "Quart" Then
142                   Part2 = "SW"
143                   Part4 = "V3"
144               Else
145                   Throw "FloatingFrequency not recognised"
146               End If
147           Case "HKD"
148               Part1 = "HD"
149               Contributor = "BGN"
150               MonthsUseLetters = True
151               MonthCodePrefix = Part1 & "DR" & Chr(CLng(Left(Tenor, Len(Tenor) - 1)) + 64)
152               If Right(Tenor, 1) = "M" Then
153                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
154               ElseIf FixedFrequency = "Quart" Then
155                   Part2 = "SW"
156               Else
157                   Throw "FixedFrequency not recognised"
158               End If
159           Case "HUF"
160               Part1 = "HF"
161               Contributor = "CMPN"
162               MonthCodePrefix = "BUBOR0" & Tenor
163               If Right(Tenor, 1) = "M" Then
164                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
165               ElseIf FixedFrequency = "Ann" Then
166                   Part2 = "SW"
167               Else
168                   Throw "FixedFrequency not recognised"
169               End If
170               If FloatingFrequency = "Semi" Then
171                   Part2 = "SW"
172                   Part4 = ""
173               ElseIf FloatingFrequency = "Quart" Then
174                   Part2 = "SW"
175                   Part4 = "V3"
176               Else
177                   Throw "FloatingFrequency not recognised"
178               End If
179           Case "JPY"
180               Part1 = "JY"
181               Contributor = "CMPL"
182               MonthCodePrefix = Part1 & String(3 - Len(Tenor) + 2, "0") & Tenor
183               If Right(Tenor, 1) = "M" Then
184                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
185               ElseIf FixedFrequency = "Semi" Then
186                   Part2 = "SWAP"
187               Else
188                   Throw "FixedFrequency not recognised"
189               End If
190               If FloatingFrequency = "Semi" Then
191                   Part2 = "SWAP"
192                   Part4 = ""
193               ElseIf FloatingFrequency = "Quart" Then
194                   Part2 = "SW"
195                   Part4 = "V3"
196               Else
197                   Throw "FloatingFrequency not recognised"
198               End If
199           Case "KRW"
200               Part1 = "KW"
201               Contributor = "CMPN"
202               MonthCodePrefix = "KWSWOF"
203               If Right(Tenor, 1) = "M" Then
204                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
205               ElseIf FixedFrequency = "Quart" Then
206                   Part2 = "SWO"
207               Else
208                   Throw "FixedFrequency not recognised"
209               End If
210           Case "MXN"
211               Part1 = "MP"
212               Contributor = "CMPN"
213               BloombergTickerInterestRateSwap = "TBD"
214           Case "MYR"
215               Part1 = "MR"
216               Contributor = "BGN"
217               MonthCodePrefix = "KLIB" & Tenor
218               If Right(Tenor, 1) = "M" Then
219                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
220               ElseIf FixedFrequency = "Quart" Then
221                   Part2 = "SWQO"
222               Else
223                   Throw "FixedFrequency not recognised"
224               End If
225           Case "NOK"
226               Part1 = "NK"
227               Contributor = "BGN"
228               MonthCodePrefix = "NIBOR" & Tenor
229               If Right(Tenor, 1) = "M" Then
230                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
231               ElseIf FixedFrequency = "Ann" Then
232                   Part2 = "SW"
233               Else
234                   Throw "FixedFrequency not recognised"
235               End If
236               If FloatingFrequency = "Semi" Then
237                   Part2 = "SW"
238                   Part4 = ""
239               ElseIf FloatingFrequency = "Quart" Then
240                   Part2 = "SW"
241                   Part4 = "V3"
242               Else
243                   Throw "FloatingFrequency not recognised"
244               End If
245           Case "NZD"
246               Part1 = "ND"
247               Contributor = "BGN"
248               MonthCodePrefix = Part1 & "BB" & Tenor
249               If Right(Tenor, 1) = "M" Then
250                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
251               ElseIf FixedFrequency = "Semi" Then
252                   Part2 = "SWAP"
253               Else
254                   Throw "FixedFrequency not recognised"
255               End If
256           Case "QAR"
257               Part1 = "QA"
258               Contributor = "CMPN"
259               If Right(Tenor, 1) = "M" Then
260                   BloombergTickerInterestRateSwap = "QRIFR" & Tenor & " Index"
261               ElseIf FixedFrequency = "Ann" Then
262                   Part2 = "SW"
263               Else
264                   Throw "FixedFrequency not recognised"
265               End If
266           Case "RON"
267               Part1 = "RN"
268               Contributor = "CMPN"
269               MonthCodePrefix = "BUBR" & Tenor
270               If Right(Tenor, 1) = "M" Then
271                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
272               ElseIf FixedFrequency = "Ann" Then
273                   Part2 = "SW"
274               Else
275                   Throw "FixedFrequency not recognised"
276               End If
277           Case "SAR"
278               Part1 = "SR"
279               Contributor = "CMPN"
280               MonthCodePrefix = "SAIB" & Tenor
281               If Right(Tenor, 1) = "M" Then
282                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
283               ElseIf FixedFrequency = "Ann" Then
284                   Part2 = "SW"
285               Else
286                   Throw "FixedFrequency not recognised"
287               End If
288           Case "SEK"
289               Part1 = "SK"
290               Contributor = "BGN"
291               MonthCodePrefix = "STIB" & Tenor
292               If Right(Tenor, 1) = "M" Then
293                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
294               ElseIf FixedFrequency = "Ann" Then
295                   Part2 = "SW"
296               Else
297                   Throw "FixedFrequency not recognised"
298               End If
299           Case "SGD"
300               Part1 = "SD"
301               Contributor = "BGN"
302               MonthCodePrefix = "SORF" & Tenor
303               If Right(Tenor, 1) = "M" Then
304                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
305               ElseIf FixedFrequency = "Quart" Then
306                   Part2 = "SW"
307               Else
308                   Throw "FixedFrequency not recognised"
309               End If
310           Case "THB"
311               Part1 = "TB"
312               Contributor = "CMPN"
313               MonthCodePrefix = "BOFX" & Tenor
314               If Right(Tenor, 1) = "M" Then
315                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
316               ElseIf FixedFrequency = "Quart" Then
317                   Part2 = "SWBC"
318               Else
319                   Throw "FixedFrequency not recognised"
320               End If
321           Case "TWD"
322               Part1 = "TD"
                  'In other codes Part1 = TR
323               Contributor = "BGN"
324               MonthCodePrefix = "TAIBOR" & Tenor
325               If Right(Tenor, 1) = "M" Then
326                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
327               ElseIf FixedFrequency = "Quart" Then
328                   Part2 = "SWO"
329               Else
330                   Throw "FixedFrequency not recognised"
331               End If
                  'Unless want to use TAIBIR:'MonthCodePrefix = "TDSF" & (Left(Tenor, 1) * 30) & "D"
332           Case "USD"
333               Part1 = "US"
334               Contributor = "BGN"
335               MonthCodePrefix = Part1 & String(3 - Len(Tenor) + 2, "0") & Tenor
336               If Right(Tenor, 1) = "M" Then
337                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
338               ElseIf FixedFrequency = "Ann" Then
339                   Part2 = "SW"
340               ElseIf FixedFrequency = "Quart" Then
341                   Part2 = "SWO"
342               ElseIf FixedFrequency = "Semi" Then
343                   Part2 = "SWAP"
344               Else
345                   Throw "FixedFrequency not recognised"
346               End If
347           Case "ZAR"
348               Part1 = "SA"
349               Contributor = "CMPL"
350               MonthCodePrefix = "JIBA" & Tenor
351               If Right(Tenor, 1) = "M" Then
352                   BloombergTickerInterestRateSwap = MonthCodePrefix & " Index"
353               ElseIf FixedFrequency = "Quart" Then
354                   Part2 = "SW"
355               Else
356                   Throw "FixedFrequency not recognised"
357               End If

358           Case Else: Throw "Unrecognised Ccy"
359       End Select

360       If (Right(Tenor, 1) = "Y") Then
361           If Part4 = "" Then
362               Select Case FloatingFrequency
                      Case "Semi"
363                       Part4 = ""
364                   Case "Quart"
365                       Part4 = ""
366                   Case Else
367                       Throw "FloatingFrequency must be Semi or Quart"
368               End Select
369           End If
370           Part5 = " " & Contributor

371           BloombergTickerInterestRateSwap = Part1 & Part2 & Part3 & Part4 & Part5 & Part6

372       ElseIf (Right(Tenor, 1) <> "Y") And (Right(Tenor, 1) <> "M") Then
373           Throw "Tenor must be a string ending with either Y or M"

374       End If

375       Exit Function
ErrHandler:
376       BloombergTickerInterestRateSwap = "#BloombergTickerInterestRateSwap (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergTickerSwaptionVol
' Author    : Philip Swannell
' Date      : 08-Jun-2016
' Purpose   : For use with the Bloomberg functions BDH and BDP, returns the string required
'             as the first argument to those functions when we wish to access a point within
'             the ATM swaption vol matrix. We'll need to take care for each currency as
'             to what contributor and what QuoteType to use.
'             Bloomberg appears to be quite inconsistent with its choice of codes for swap rates
'             so this function is a mess as I encounter those inconsistencies. Hence rather than
'             being called from code, this function is called from the cells of each currency sheet
'             (in the column headed "Bloomberg Code" within the Swaps range of data). Thus the call to
'             this function can simply be overridden with text if necessary.
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergTickerSwaptionVol(Ccy As String, Exercise As String, Tenor As String, Optional QuoteType As String = "Normal", Optional Contributor = "CMPN") As String
1         On Error GoTo ErrHandler
          Dim Part1 As String
          Dim Part2 As String
          Dim Part3 As String
          Dim Part4 As String
          Dim Part5 As String
          Const Part6 = " Curncy"
          Dim NumberFormatExercise As String
          Dim NumberFormatTenor As String
          Dim StrangeFor18M As Boolean

2         NumberFormatExercise = "00"
3         NumberFormatTenor = "0"

4         Select Case Ccy
              Case "AED": Part1 = "UD"
5             Case "AUD": Part1 = "AD"
6             Case "CAD": Part1 = "CD"
7             Case "CHF": Part1 = "SF"
                  'Changed to match tickers used by Airbus March 2022
                  '  NumberFormatTenor = "00"
8             Case "CZK": Part1 = "CK"
9             Case "DKK": Part1 = "DK"
10                NumberFormatTenor = "00"
11            Case "EUR": Part1 = "EU"
                  'Changed to match tickers used by Airbus March 2022
                  ' NumberFormatExercise = "0"
12            Case "GBP": Part1 = "BP"
                  'Changed to match tickers used by Airbus March 2022
                  'NumberFormatExercise = "0"
13            Case "HKD": Part1 = "HD"
14            Case "HUF": Part1 = "HF"
                  'Changed to match tickers used by Airbus March 2022
15            Case "JPY": Part1 = "JY"
                  '    Case "JPY": Part1 = "JP"
                  ' NumberFormatExercise = "0"
16            Case "KRW": Part1 = "KW"
17            Case "MXN": Part1 = "MP"
18            Case "MYR": Part1 = "MR"
19            Case "NOK": Part1 = "NK"
20            Case "NZD": Part1 = "ND"
21            Case "RON": Part1 = "RN"
22            Case "SAR": Part1 = "SR"
23            Case "SEK": Part1 = "SK"
24                NumberFormatTenor = "00"
25            Case "SGD": Part1 = "SD"
26            Case "THB": Part1 = "TB"
27            Case "TWD": Part1 = "TR"
28            Case "USD": Part1 = "US"
29                StrangeFor18M = True
30            Case "ZAR": Part1 = "SA"
31            Case "CNH", "QAR", "RON": Throw "Swaption Vol Code not found"
32            Case Else: Throw "Unrecognised Ccy"
33        End Select

34        Select Case QuoteType
              Case "Normal": Part2 = "SN"
35            Case "OIS Normal"
36                If Ccy = "USD" Then
37                    Part2 = "SN"
38                ElseIf Ccy = "CHF" Then
39                    Part2 = "NO"
40                ElseIf Ccy = "DKK" Or Ccy = "SEK" Then
41                    Part2 = "NV"
42                Else
43                    Part2 = "NE"
44                End If
45            Case Else: Throw "Unrecognised QuoteType. Allowed values are: Normal, OIS Normal"
46        End Select

47        Select Case UCase(Exercise)
              Case "1M", "2M", "3M", "4M", "5M", "6M", "7M", "8M", "9M", "10M", "11M"
48                Part3 = "0" & Chr(64 + CLng(Left(Exercise, Len(Exercise) - 1)))  '1M = 0A, 2M = 0B etc
49            Case "13M", "14M", "15M", "16M", "17M", "18M"
50                If StrangeFor18M Then
51                    Part3 = "0" & Chr(64 + CLng(Left(Exercise, Len(Exercise) - 1)))  '1M = 0A, 2M = 0B, 18M = 0R etc
52                Else
53                    Part3 = "1" & Chr(64 - 12 + CLng(Left(Exercise, Len(Exercise) - 1)))    '13M = 1A, 14M = 1B etc
54                End If
55            Case "1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y", "25Y", "30Y"
56                Part3 = Format(CLng(Left(Exercise, Len(Exercise) - 1)), NumberFormatExercise)
57            Case Else
58                Throw "Unrecognised Exercise. Allowed values are: 1M, 2M, 3M,...18M, 1Y, 2Y, 3Y, 4Y, 5Y, 6Y, 7Y, 8Y, 9Y, 10Y, 15Y, 20Y, 25Y, 30Y"
59        End Select

60        Select Case UCase(Tenor)
              Case "1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y", "25Y", "30Y"
61                Part4 = Format(CLng(Left(Tenor, Len(Tenor) - 1)), NumberFormatTenor)
62            Case Else
63                Throw "Unrecognised Tenor. Allowed values are: 1Y, 2Y, 3Y, 4Y, 5Y, 6Y, 7Y, 8Y, 9Y, 10Y, 15Y, 20Y, 25Y, 30Y"
64        End Select

65        Select Case Contributor
              Case "CMPN", "LAST", "BBIR", "CFIR", "CNTR", "ICPL", "SMKO", "TRPU", "GFIS", "CMPL"
66                Part5 = " " & Contributor
67            Case Else
68                Throw "Unrecognised Contributor. Allowed values are: CMPN, LAST, BBIR, CFIR, CNTR, ICPL, SMKO, TRPU, GFIS, CMPL"
69        End Select

70        BloombergTickerSwaptionVol = Part1 & Part2 & Part3 & Part4 & Part5 & Part6

71        Exit Function
ErrHandler:
72        BloombergTickerSwaptionVol = "#BloombergTickerSwaptionVol (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : CurrencyOrder
' Author    : Philip Swannell
' Date      : 11-Jul-2016
' Purpose   : An attempt to code the rules for the Fx market's preferred quote order for
'             currency pairs. By construction matches Bloomberg for those currency pairs
'             for which Bloomberg has FxVol data.
' -----------------------------------------------------------------------------------------------------------------------
Function CurrencyOrder(Ccy1 As String, Ccy2 As String)
          Static BBPairs As Variant
          Static Majors As Variant
          Dim MatchRes1 As Variant
          Dim MatchRes2 As Variant

1         On Error GoTo ErrHandler
2         If IsEmpty(BBPairs) Then
              'Complete list of currency pairs for which 1 year vols are available from Bloomberg as of 14 June 2016
              'See https://d.docs.live.net/4251b448d4115355/Excel Sheets/CurrencyOrderInvestigation.xlsx
3             BBPairs = sArrayStack( _
                  sTokeniseString("AUDCHF,AUDCNH,AUDJPY,AUDKRW,AUDSEK,AUDUSD,BRLJPY,BTNJPY,CADCHF,CADDKK,CADMXN,CADNOK,CADSEK,CADSGD,CHFHKD,CHFHUF,CHFMXN,CHFNOK,CHFSGD,CHFZAR"), _
                  sTokeniseString("CNYJPY,EURAUD,EURBRL,EURCHF,EURCNH,EURCZK,EURFKP,EURGBP,EURHKD,EURHRK,EURILS,EURJPY,EURNOK,EURNZD,EURPLN,EURRON,EURRUB,EURSGD,EURSHP,EURTRY"), _
                  sTokeniseString("EURTWD,EURUSD,EURZAR,FKPAUD,FKPBTN,FKPCAD,FKPCHF,FKPCNH,FKPCNY,FKPIDR,FKPINR,FKPKRW,FKPSGD,FKPTWD,FKPZAR,GBPCAD,GBPCNH,GBPDKK,GBPHKD,GBPIDR"), _
                  sTokeniseString("GBPJPY,GBPNZD,GBPSEK,GBPSGD,GBPTWD,GBPUSD,GBPZAR,GIPAUD,GIPBTN,GIPCHF,GIPCNH,GIPCNY,GIPDKK,GIPINR,GIPJPY,GIPKRW,GIPNOK,GIPNZD,GIPSEK,GIPSGD"), _
                  sTokeniseString("GIPTWD,HKDJPY,JPYCNH,JPYKRW,JPYTWD,MXNJPY,NZDCAD,NZDDKK,NZDHKD,NZDSEK,NZDSGD,NZDUSD,PLNCZK,PLNDKK,PLNHUF,PLNSEK,RUBDKK,RUBJPY,SHPCAD,SHPCHF"), _
                  sTokeniseString("SHPCNY,SHPIDR,SHPINR,SHPJPY,SHPNOK,SHPNZD,SHPPLN,SHPTWD,SHPZAR,USDBRL,USDCAD,USDCHF,USDCLP,USDCNY,USDCOP,USDCZK,USDDKK,USDHKD,USDHUF,USDIDR"), _
                  sTokeniseString("USDILS,USDINR,USDJPY,USDKRW,USDMXN,USDNOK,USDPEN,USDPHP,USDPLN,USDQAR,USDRUB,USDSAR,USDSEK,USDSGD,USDTRY,USDTWD,USDZAR,XAGCHF,XAGUSD,XAUFKP,XAUUSD,XPDUSD,XPTUSD,ZARDKK,ZARJPY"))
4             BBPairs = sSortedArray(BBPairs)
5         End If

6         If IsNumber(sMatch(Ccy1 & Ccy2, BBPairs, True)) Then
7             CurrencyOrder = Ccy1 & Ccy2
8         ElseIf IsNumber(sMatch(Ccy2 & Ccy1, BBPairs, True)) Then
9             CurrencyOrder = Ccy2 & Ccy1
10        Else
11            If IsEmpty(Majors) Then
12                Majors = sTokeniseString("XAG,XAU,XPD,XPT,EUR,GBP,GIP,FKP,SHP,AUD,NZD,USD,CAD,CHF,CNY,HKD,BTN,BRL,ZAR,MXN,RUB,JPY,PLN")
13            End If
14            MatchRes1 = sMatch(Ccy1, Majors)
15            If Not IsNumber(MatchRes1) Then MatchRes1 = sNRows(Majors) + 1
16            MatchRes2 = sMatch(Ccy2, Majors)
17            If Not IsNumber(MatchRes2) Then MatchRes2 = sNRows(Majors) + 1
18            If MatchRes1 < MatchRes2 Then
19                CurrencyOrder = Ccy1 & Ccy2
20            Else
21                CurrencyOrder = Ccy2 & Ccy1
22            End If
23        End If

24        Exit Function
ErrHandler:
25        Throw "#CurrencyOrder (line " & CStr(Erl) + "): " & Err.Description & "!"

End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergFormulaSwapRate
' Author    : Hermione Glyn
' Date      : 11-Jul-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergFormulaSwapRate(Ccy As String, Tenor As String, FixFreq As String, FloatFreq As String, _
          Live As Boolean, AsOfDate As Long)

1         On Error GoTo ErrHandler

          Dim DateString As String
          Dim Formula As String
          Dim Ticker As String

2         Ticker = BloombergTickerInterestRateSwap(Ccy, Tenor, FixFreq, FloatFreq)

3         If Not Live Then DateString = "DATE(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"

4         If Live Then
5             If Right(Tenor, 1) = "M" Then
6                 Formula = "BDP(""" + Ticker + """,""PX_LAST"")"
7             Else
8                 Formula = "Avg(BDP(""" + Ticker + """,""Bid""),BDP(""" + Ticker + """,""Ask""))"
9             End If
10        Else
11            Formula = "BDH(""" + Ticker + """,""PX_LAST""," & DateString + ")"
12        End If

13        If gApplyRandomAdjustments Then Formula = "RandomAdjust(" & Formula & ")"

14        BloombergFormulaSwapRate = Formula

15        Exit Function
ErrHandler:
16        BloombergFormulaSwapRate = "#BloombergFormulaSwapRate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergFormulaBasisSwapRate
' Author    : Hermione Glyn
' Date      : 11-Jul-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergFormulaBasisSwapRate(Ccy As String, Tenor As String, Live As Boolean, AsOfDate As Long)

1         On Error GoTo ErrHandler

          Dim DateString As String
          Dim Formula As String
          Dim Ticker As String

          'If TypeName(Application.Caller) <> "Range" Then
          'Record "BloombergFormulaBasisSwapRate", Ccy, Tenor, Live, AsOfDate
          'End If

2         Ticker = BloombergTickerBasisSwap(Ccy, Tenor)

3         If Not Live Then DateString = "DATE(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"

4         If Live Then
5             If Right(Tenor, 1) = "M" Then
6                 Formula = "BDP(""" + Ticker + """,""PX_LAST"")"
7             Else
8                 Formula = "Avg(BDP(""" + Ticker + """,""Bid""),BDP(""" + Ticker + """,""Ask""))"
9             End If
10        Else
11            Formula = "BDH(""" + Ticker + """,""PX_LAST""," & DateString + ")"
12        End If

13        If gApplyRandomAdjustments Then Formula = "RandomAdjust(" & Formula & ")"

14        BloombergFormulaBasisSwapRate = Formula

15        Exit Function
ErrHandler:
16        BloombergFormulaBasisSwapRate = "#BloombergFormulaBasisSwapRate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergFormulaSwaptionVol
' Author    : Hermione Glyn
' Date      : 11-Jul-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergFormulaSwaptionVol(Ccy As String, Exercise As String, Tenor As String, QuoteType As String, Contributor As String, Live As Boolean, AsOfDate As Long)

1         On Error GoTo ErrHandler

          Dim DateString As String
          Dim Formula As String
          Dim Ticker As String

          'If TypeName(Application.Caller) <> "Range" Then
          'Record "BloombergFormulaSwaptionVol", Ccy, Exercise, Tenor, QuoteType, Contributor, Live, AsOfDate
          'End If

2         Ticker = BloombergTickerSwaptionVol(Ccy, Exercise, Tenor, QuoteType, Contributor)

3         If Not Live Then DateString = "DATE(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"

4         If Live Then
5             Formula = "BDP(""" + Ticker + """,""PX_LAST"")"
6         Else
7             Formula = "BDH(""" + Ticker + """,""PX_LAST""," & DateString + ")"
8         End If

9         If gApplyRandomAdjustments Then Formula = "RandomAdjust(" & Formula & ")"

10        BloombergFormulaSwaptionVol = Formula

11        Exit Function
ErrHandler:
12        BloombergFormulaSwaptionVol = "#BloombergFormulaSwaptionVol (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : RandomAdjust
' Author    : Philip Swannell
' Date      : 16-Mar-2018
' Purpose   : Implements method applying random adjustment to data fed from Bloomberg so
'             that we are not sending Bloomberg numbers off the one PC we have the rights
'             to use Bloomberg numbers.
'             Number is first multiplied by a call to Rnd, then rounded to 5 significant figures
' -----------------------------------------------------------------------------------------------------------------------
Function RandomAdjust(ByVal x As Variant)
          Const Nsf = 5    'Number of sgnificant figures in the return
          Dim NDigits As Long
          Dim Pow As Double
1         If IsNumber(x) Then
2             x = x * 1 + (Rnd() - 0.5) * 0.05
3             If x <> 0 Then
4                 NDigits = Log(x) / Log(10)
5                 Pow = 10 ^ (NDigits + 1 - Nsf)
6                 x = Pow * CLng(x / Pow)
7             End If
8         End If
9         RandomAdjust = x
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergFormulaFxSpot
' Author    : Hermione Glyn
' Date      : 11-Jul-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergFormulaFxSpot(CcyPair As String, Live As Boolean, AsOfDate As Long)
1         On Error GoTo ErrHandler

          Dim DateString As String
          Dim Formula As String
          Dim Ticker As String

2         If Not Live Then DateString = "DATE(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"

3         Ticker = BloombergTickerFxSpot(CcyPair)

4         If Live Then
5             Formula = "BDP(""" + Ticker + """,""PX_LAST"")"
6         Else
7             Formula = "BDH(""" + Ticker + """,""PX_LAST""," & DateString + "," + DateString + ")"
8         End If

9         If gApplyRandomAdjustments Then Formula = "RandomAdjust(" & Formula & ")"
10        BloombergFormulaFxSpot = Formula

11        Exit Function
ErrHandler:
12        BloombergFormulaFxSpot = "#BloombergFormulaFxSpot (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergTickerFxSpot
' Author    : Hermione Glyn
' Date      : 30-Sep-2016
' Purpose   : To be called by BloombergFormula equivalent and to get data
'             from airbus text file of rates.
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergTickerFxSpot(CcyPair As String)
1         On Error GoTo ErrHandler
          Dim Ticker As String
2         Ticker = CcyPair + " Curncy"
3         BloombergTickerFxSpot = Ticker
4         Exit Function
ErrHandler:
5         Throw "#BloombergTickerFxSpot (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergFormulaFxVol
' Author    : Hermione Glyn
' Date      : 11-Jul-2016
' Purpose   :
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergFormulaFxVol(CcyPair As String, Tenor As String, Live As Boolean, AsOfDate As Long)
1         On Error GoTo ErrHandler

          Dim DateString As String
          Dim Formula As String
          Dim Ticker As String

2         Ticker = BloombergTickerFxVol(CcyPair, Tenor)
3         If Not Live Then DateString = "DATE(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"

4         If Live Then
5             Formula = "BDP(""" + Ticker + """,""PX_LAST"")"
6         Else
7             Formula = "BDH(""" + Ticker + """,""PX_LAST""," & DateString + "," + DateString + ")"
8         End If
9         If gApplyRandomAdjustments Then Formula = "RandomAdjust(" & Formula & ")"
10        BloombergFormulaFxVol = Formula

11        Exit Function
ErrHandler:
12        BloombergFormulaFxVol = "#BloombergFormulaFxVol (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergTickerFxVol
' Author    : Hermione Glyn
' Date      : 30-Sep-2016
' Purpose   : To be called by BloombergFormula equivalent and to get data
'             from airbus text file of rates.
'             Need to check if they have data for AED, QAR, SAR or if treated as USD
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergTickerFxVol(CcyPair As String, Tenor As String)
1         On Error GoTo ErrHandler
          Dim Ticker As String
2         Ticker = CurrencyOrder(Left(CcyPair, 3), Right(CcyPair, 3)) + "V" + UCase(Tenor) + " Curncy"
3         Select Case Left(Ticker, 6)
              Case "EURAED", "EURQAR", "EURSAR"
4                 Ticker = "EURUSD" + Mid(Ticker, 7)
5         End Select
6         BloombergTickerFxVol = Ticker
7         Exit Function
ErrHandler:
8         Throw "#BloombergTickerFxVol (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ParseBBGFile
' Author    : Hermione Glyn
' Date      : 30-Sep-2016
' Purpose   : Reads a text file into an array, where tickers are the first column and
'             other bloomberg fields follow (eg: PX_LAST). Used by FeedRatesFromTextFile.
' -----------------------------------------------------------------------------------------------------------------------
Function ParseBBGFile(FileName As String, ByRef AnchorDate As Long)

1         On Error GoTo ErrHandler
          Dim DataEnd As Long
          Dim DataStart As Long
          Dim FileIsComplete As Boolean
          Dim Headers As Variant
          Dim HeadersEnd As Long
          Dim HeadersStart As Long
          Dim i As Long
          Dim NC As Long
          Dim NR As Long
          Dim OutputData As Variant
          Dim RunDate As String
          Dim TheInputData As Variant

2         If Not sFileExists(FileName) Then Throw "Cannot find file: " + FileName + vbLf + _
              "Is the 'MarketDataFile' setting correct on the Config sheet of the market data workbook?", True

3         TheInputData = ThrowIfError(sFileShow(FileName, "|", True, False, False, False))
4         NR = sNRows(TheInputData)
5         NC = sNCols(TheInputData)

6         For i = 1 To NR
7             If Left(UCase(CStr(TheInputData(i, 1))), 7) = "RUNDATE" Then
8                 RunDate = TheInputData(i, 1)
9                 AnchorDate = AnchorDateFromRunDate(RunDate)
10            End If
11            Select Case UCase(CStr(TheInputData(i, 1)))
                  Case "START-OF-FIELDS"
12                    HeadersStart = i + 1
13                Case "END-OF-FIELDS"
14                    HeadersEnd = i
15                Case "START-OF-DATA"
16                    DataStart = i + 1
17                Case "END-OF-DATA"
18                    DataEnd = i
19                Case "END-OF-FILE"
20                    FileIsComplete = True
21            End Select
22        Next i

23        If HeadersStart = 0 Then Throw "#Cannot find line: START-OF-FIELDS in " & FileName
24        If HeadersEnd = 0 Then Throw "#Cannot find line: END-OF-FIELDS in " & FileName
25        If DataStart = 0 Then Throw "#Cannot find line: START-OF-DATA in " & FileName
26        If DataEnd = 0 Then Throw "#Cannot find line: END-OF-DATA in " & FileName
27        If Len(RunDate) = 0 Then Throw "#Cannot find RUNDATE in " & FileName
28        If Not FileIsComplete Then Throw "#Cannot find line: END-OF-FILE in " & FileName

29        Headers = sArrayRange("Ticker", "?", "?", sArrayTranspose(sSubArray(TheInputData, HeadersStart, 1, HeadersEnd - HeadersStart, 1)))
30        OutputData = sSubArray(TheInputData, DataStart, 1, DataEnd - DataStart, NC - 1)
31        OutputData = sArrayStack(Headers, OutputData)

32        ParseBBGFile = OutputData
33        Exit Function

ErrHandler:
34        Throw "#ParseBBGFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : AnchorDateFromRunDate
' Author    : Philip Swannell
' Date      : 15-Dec-2016
' Purpose   : Convert as string of the form RUNDATE=YYYYMMDD into the indicated date as a long
' -----------------------------------------------------------------------------------------------------------------------
Function AnchorDateFromRunDate(RunDateText As String)
          Dim TheDay As String
          Dim TheMonth As String
          Dim TheYear As String
          Dim yyyymmdd As String
          Const ErrString = "RUNDATE not found in file. Should appear as a line in the file of the form 'RUNDATE=YYYYMMDD'"

1         On Error GoTo ErrHandler
2         yyyymmdd = Trim(sStringBetweenStrings(RunDateText, "="))
3         If Len(yyyymmdd) <> 8 Then Throw ErrString
4         If Not sIsRegMatch("^[0-9]*$", yyyymmdd) Then Throw ErrString

5         TheYear = Left(yyyymmdd, 4)
6         TheMonth = Mid(yyyymmdd, 5, 2)
7         TheDay = Right(yyyymmdd, 2)
8         If CLng(TheMonth) > 12 Or CLng(TheMonth) < 1 Then Throw ErrString
9         If CLng(TheDay) > 31 Or CLng(TheDay) < 1 Then Throw ErrString
10        AnchorDateFromRunDate = DateSerial(CInt(TheYear), CInt(TheMonth), CInt(TheDay))

11        Exit Function
ErrHandler:
12        Throw "#AnchorDateFromRunDate (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergFormulaZCInflationSwap
' Author    : Philip Swannell
' Date      : 24-Apr-2017
' Purpose   : Bloomberg formula for inflation zero coupon swaps...
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergFormulaZCInflationSwap(Index As String, Tenor As String, Live As Boolean, AsOfDate As Long)
          Dim DateString As String
          Dim Formula As String
          Dim Ticker As String

1         On Error GoTo ErrHandler

2         If Not Live Then DateString = "DATE(" + Format(AsOfDate, "yyyy") + "," + Format(AsOfDate, "m") + "," + Format(AsOfDate, "d") + ")"

3         Ticker = BloombergTickerZCInflationSwap(Index, Tenor)

4         If Live Then
5             Formula = "BDP(""" + Ticker + """,""PX_LAST"")"
6         Else
7             Formula = "BDH(""" + Ticker + """,""PX_LAST""," & DateString + ")"
8         End If

9         If gApplyRandomAdjustments Then Formula = "RandomAdjust(" & Formula & ")"
10        BloombergFormulaZCInflationSwap = Formula

11        Exit Function
ErrHandler:
12        BloombergFormulaZCInflationSwap = "#BloombergFormulaZCInflationSwap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergTickerZCInflationSwap
' Author    : Philip Swannell
' Date      : 18-May-2017
' Purpose   : Returns the Bloomberg ticker for a given inflation index and maturity
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergTickerZCInflationSwap(Index As String, Tenor As String)
          Static IndexError As String
          Dim Ticker As String
          Const TenorError = "Invalid tenor. Must be of the form nY for a positive integer n"
          Dim nY As Double

1         On Error GoTo ErrHandler
2         If IndexError = "" Then IndexError = "Invalid Index. Valid are: " + sConcatenateStrings(shSAIStaticData.Range("InflationIndices").Columns(1), ", ")

3         If UCase(Right(Tenor, 1)) <> "Y" Then Throw TenorError
4         If Not IsNumeric(Left(Tenor, Len(Tenor) - 1)) Then Throw TenorError

5         nY = CDbl(Left(Tenor, Len(Tenor) - 1))
6         If nY <= 0 Or CLng(nY) <> nY Then Throw TenorError

7         Select Case LCase(Index)
              Case "ukrpi"
8                 Ticker = "BPSWIT" + CStr(nY) + " Curncy"
9             Case "uscpi"
10                Ticker = "USSWIT" + CStr(nY) + " Curncy"
11            Case "eurohicpxt"
12                Ticker = "EUSWI" + CStr(nY) + " Curncy"
13            Case "frenchcpixt"
14                Ticker = "FRSWI" + CStr(nY) + " Curncy"
15            Case Else
16                If IsNumber(sMatch(Index, shSAIStaticData.Range("InflationIndices").Columns(1))) Then
17                    Throw "Not yet implemented for index " + Index
18                Else
19                    Throw IndexError
20                End If
21        End Select

22        BloombergTickerZCInflationSwap = Ticker
23        Exit Function
ErrHandler:
24        Throw "#BloombergTickerZCInflationSwap (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergTickerInflationIndex
' Author    : Philip
' Date      : 28-Sep-2017
' Purpose   : Our language to Bloomberg ticker language for inflation indices...
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergTickerInflationIndex(Index As String)
          Static IndexError As String
          Dim Ticker As String

1         On Error GoTo ErrHandler
2         If IndexError = "" Then IndexError = "Invalid Index. Valid are: " + sConcatenateStrings(shSAIStaticData.Range("InflationIndices").Columns(1), ", ")

3         Select Case LCase(Index)
              Case "ukrpi"
4                 Ticker = "UKRPI Index"
5             Case "uscpi"
6                 Ticker = "CPURNSA Index"
7             Case "eurohicpxt"
8                 Ticker = "CPTFEMU Index"
9             Case "frenchcpixt"
10                Ticker = "FRCPXTOB Index"
11            Case Else
12                If IsNumber(sMatch(Index, shSAIStaticData.Range("InflationIndices").Columns(1))) Then
13                    Throw "Not yet implemented for index " + Index
14                Else
15                    Throw IndexError
16                End If
17        End Select

18        BloombergTickerInflationIndex = Ticker
19        Exit Function
ErrHandler:
20        Throw "#BloombergTickerInflationIndex (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : BloombergFormulaInflationIndex
' Author    : Philip Swannell
' Date      : 28-Sep-2017
' Purpose   : Returns the formula to get historical index for a single month
' -----------------------------------------------------------------------------------------------------------------------
Function BloombergFormulaInflationIndex(Index As String, Year As Long, Month As Long)
          Dim DatePart As String
          Dim DaysInMonth As Long
          Dim Ticker As String

1         On Error GoTo ErrHandler
2         Ticker = BloombergTickerInflationIndex(Index)
3         DaysInMonth = Day(DateSerial(Year, Month + 1, 1) - 1)
4         DatePart = "DATE(" + CStr(Year) + "," + CStr(Month) + "," + CStr(DaysInMonth) + ")"
5         BloombergFormulaInflationIndex = "BDH(""" + Ticker + """,""PX_LAST""," & DatePart & "," + DatePart + ")"
6         Exit Function
ErrHandler:
7         Throw "#BloombergFormulaInflationIndex (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

