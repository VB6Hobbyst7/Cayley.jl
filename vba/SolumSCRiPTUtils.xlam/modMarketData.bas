Attribute VB_Name = "modMarketData"
Option Explicit
Private Const gModel = "SCRiPTModel" 'Don't understand why code here can't see this constant declared in SolumAddin
Public Const gRInDev = False

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : MyAnchorDate
' Author    : Philip Swannell
' Date      : 09-Sep-2016
' Purpose   : Return the AnchorDate by interogating the named model in the R environment.
' -----------------------------------------------------------------------------------------------------------------------
Function MyAnchorDate(ModelName As String)
          Dim Expression1 As String
          Dim Expression2 As String
          Dim Res As Variant
1         On Error GoTo ErrHandler
2         Expression1 = "exists(""" & ModelName & """)"
3         If Not ThrowIfError(sExecuteRCode(Expression1)) Then
4             Throw "Cannot find model: '" + ModelName + "' Please ensure the Hull-White model is built", True
5         End If
6         Expression2 = "toexceldate(" + ModelName + "$AnchorDate)"
7         Res = ThrowIfError(sExecuteRCode(Expression2))
8         MyAnchorDate = Res
9         Exit Function
ErrHandler:
10        Throw "#MyAnchorDate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: SCRiPT_SurvProb
' Purpose:  Returns Survival Probabilities for a given credit for a market built by the R code, SCRiPT workbook etc.
' Parameter Credit (String): name of credit
' Parameter Dates (Variant): Column array of dates or omitted, in which case return is two column array of grid-dates and probabilities
' Parameter MultiplicativeShockToSpreads (Double):
' Parameter ModelName (): Text name of the HW model in R
' Author: Philip Swannell
' Date: 23-Nov-2017
' -----------------------------------------------------------------------------------------------------------------------
Function SCRiPT_SurvProb(Credit As String, Optional ByVal Dates As Variant, Optional MultiplicativeShockToSpreads As Double = 1, Optional ModelName As String = gModel)
Attribute SCRiPT_SurvProb.VB_Description = "Returns (from the local R environment) survival probabilities used by Solum's SCRiPT software."
Attribute SCRiPT_SurvProb.VB_ProcData.VB_Invoke_Func = " \n32"
1         On Error GoTo ErrHandler
          Static HaveChecked As Boolean
2         If gRInDev Or Not HaveChecked Then CheckR "SCRiPT_SurvProb", gPackages, gRSourcePath + "UtilsPGS.R", "SourceAllFiles()": HaveChecked = True
3         SCRiPT_SurvProb = ThrowIfError(Application.Run("BERT.Call", "SurvivalProbabilityXL", ModelName, Credit, Dates, MultiplicativeShockToSpreads))
4         Exit Function
ErrHandler:
5         SCRiPT_SurvProb = "#SCRiPT_SurvProb (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: SCRiPT_DF
' Purpose: Returns discount factors as calculated by the R code used by SCRiPT
' Parameter Ccy (String): Three letter ISO currency code
' Parameter Dates (Variant): Column array of dates. Non numbers cause #N/A! to appear in the return
' Parameter Funding (Boolean): TRUE = funding discount factors, FALSE = discount factors for Libor projection
' Parameter FundingSpread (Double): Funding spread. Enter 1% funding spread as 0.01
' Parameter ModelName (String): The name of the model in R. Defaults to the string "SCRiPTModel" which is the name of the model created by the SCRiPT workbook
' Author: Philip Swannell
' Date: 24-Nov-2017
' -----------------------------------------------------------------------------------------------------------------------
Function SCRiPT_DF(Ccy As String, Optional Dates As Variant, Optional Funding As Boolean, Optional FundingSpread As Double = 0, Optional ModelName As String = gModel)
Attribute SCRiPT_DF.VB_Description = "Returns (from the local R environment) discount factors used by Solum's SCRiPT software."
Attribute SCRiPT_DF.VB_ProcData.VB_Invoke_Func = " \n32"
1         On Error GoTo ErrHandler
          Static HaveChecked As Boolean
          If gRInDev Or Not HaveChecked Then CheckR "SCRiPT_DF", gPackages, gRSourcePath + "UtilsPGS.R", "SourceAllFiles()": HaveChecked = True
2         SCRiPT_DF = ThrowIfError(Application.Run("BERT.Call", "DiscountFactorXL", Ccy, Dates, Funding, FundingSpread, ModelName))
3         Exit Function
ErrHandler:
4         SCRiPT_DF = "#SCRiPT_DF (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Function SCRiPT_LIBOR(Ccy As String, Start As Variant, Maturity As Variant, DCT As String, Optional ModelName As String = gModel)
Attribute SCRiPT_LIBOR.VB_Description = "Returns (from the local R environment) forward libor rates used by Solum's SCRiPT software."
Attribute SCRiPT_LIBOR.VB_ProcData.VB_Invoke_Func = " \n32"
1         On Error GoTo ErrHandler
          Static HaveChecked As Boolean
2         If gRInDev Or Not HaveChecked Then CheckR "SCRiPT_LIBOR", gPackages, gRSourcePath + "UtilsPGS.R", "SourceAllFiles()": HaveChecked = True
3         SCRiPT_LIBOR = ThrowIfError(Application.Run("BERT.Call", "LIBORXL", Ccy, Start, Maturity, DCT, ModelName))
4         Exit Function
ErrHandler:
5         SCRiPT_LIBOR = "#SCRiPT_LIBOR (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

Function SCRiPT_SwapRate(Ccy As String, Start As Variant, Maturity As Variant, _
          FixedFrequency As String, FixedDCT As String, FixedBDC As String, _
          FloatingFrequency As String, FloatingDCT As String, FloatingBDC As String, Optional ModelName As String = gModel, Optional ControlString As String = "SwapRate")
Attribute SCRiPT_SwapRate.VB_Description = "Returns (from the local R environment) forward swap rates used by Solum's SCRiPT software. Can also return the ""annuity factor"", i.e. the value of the fixed leg of the swap swap with the coupon set to 1% and the notional to 100."
Attribute SCRiPT_SwapRate.VB_ProcData.VB_Invoke_Func = " \n32"
1         On Error GoTo ErrHandler
          Static HaveChecked As Boolean
          Dim StringArgs
2         If gRInDev Or Not HaveChecked Then CheckR "SCRiPT_SwapRate", gPackages, gRSourcePath + "UtilsPGS.R", "SourceAllFiles()": HaveChecked = True
3         StringArgs = Array(Ccy, FixedDCT, FixedBDC, FloatingDCT, FloatingBDC, ModelName, LCase(Replace(ControlString, " ", "")))
4         SCRiPT_SwapRate = Application.Run("BERT.Call", "SwapRateXL", StringArgs, Start, Maturity, sParseFrequencyString(FixedFrequency, True, True), sParseFrequencyString(FloatingFrequency, True, True))
5         Exit Function
ErrHandler:
6         SCRiPT_SwapRate = "#SCRiPT_SwapRate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure Name: SCRiPT_Results
' Purpose:        Straight wrap of R function of the same name.
' Parameter ID (String):The name of a "netting set" i.e. Counterparty or else a TradeID.
' Parameter ControlString (String): A comma-delimited string to set the output columns Allowed tokens: 'Time', 'Date' or ABC
'                                   where A is 'Our' or 'Their'; B is 'EE', 'EPE', 'ENE', 'PFE' or 'Paths'; and C is absent or
'                                   'WhatIf'. E.g. TheirEPEWhatIf = EPE from bank PoV, including WHATIF trades.
' Parameter WithHeaders (Boolean):  If TRUE, a row of header text appears at the top of the return. If FALSE no headers are returned.
'                                   This argument is optional, defaulting to FALSE.
' Parameter IDIsNetset (Variant):   If TRUE, ID is taken to be a netting set, if FALSE it is taken to be a Trade ID. This argument
'                                   is optional and if omitted the function makes an intelligent guess (producing an error in the
'                                   unlikely event that ID matches both a TradeID and a netting set).
' Author: Philip Swannell
' Date: 06-Dec-2017
' -----------------------------------------------------------------------------------------------------------------------
Function SCRiPT_Results(ID As String, ControlString As String, Optional WithHeaders As Boolean, Optional IDIsNetset As Variant)
Attribute SCRiPT_Results.VB_Description = "For use with the SCRiPT workbook. Returns (from the local R environment) data generated for ""netting sets"" and trades. For example expected exposures, potential future exposures, expected positive exposure."
Attribute SCRiPT_Results.VB_ProcData.VB_Invoke_Func = " \n32"
          Static HaveChecked As Boolean
1         On Error GoTo ErrHandler
2         If gRInDev Or Not HaveChecked Then CheckR "SCRiPT_Results", gPackages, gRSourcePath + "UtilsPGS.R", "SourceAllFiles()": HaveChecked = True
3         SCRiPT_Results = Application.Run("BERT.Call", "SCRiPT_Results", ID, ControlString, WithHeaders, IDIsNetset)
4         Exit Function
ErrHandler:
5         SCRiPT_Results = "#SCRiPT_Results (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : SCRiPT_DateSchedule
' Author    : Philip Swannell
' Date      : 20-Sep-2016
' Purpose   : Wrap private function DateSchedule, though argument Frequency is expressed differently

'Prior to 18 Dec 2021 this function was a wrapper to to R function SCRiPT_DateSchedule
' -----------------------------------------------------------------------------------------------------------------------
Function SCRiPT_DateSchedule(StartDate As Long, EndDate As Long, ByVal Frequency As String, Optional BDC As String = "Mod Foll", Optional ByVal WhatToReturn As String = "Dates")
Attribute SCRiPT_DateSchedule.VB_Description = "Returns the dates of (one leg of) an interest rate swap."
Attribute SCRiPT_DateSchedule.VB_ProcData.VB_Invoke_Func = " \n32"
          Dim FrequencyNum As Long
1         On Error GoTo ErrHandler

3         Select Case UCase(Frequency)
              Case "1Y", "12M", "ANNUAL", "ANN"
4                 FrequencyNum = 1
5             Case "6M", "SEMI", "SEMI-ANNUAL", "SEMI ANNUAL"
6                 FrequencyNum = 2
7             Case "3M", "QUARTERLY"
8                 FrequencyNum = 4
9             Case "1M", "MONTHLY"
10                FrequencyNum = 12
11            Case Else
12                Throw "Frequency must be '1Y', '6M', '3M', '1M' or equivalently 'Annual', 'Semi Annual', 'Quarterly', 'Monthly'"
13        End Select

27        SCRiPT_DateSchedule = DateSchedule(StartDate, EndDate, FrequencyNum, BDC, WhatToReturn)
28        Exit Function
ErrHandler:
29        SCRiPT_DateSchedule = "#SCRiPT_DateSchedule (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SCRiPT_AdjustDate
' Author     : Philip Swannell
' Date       : 10-Jan-2018
' Purpose    : Applies date adjustment by calling R function AdjustDate, but Strings, errors and Booleans passed are left unchanged
'              numbers have to be whole numbers and are adjusted
' Parameters :
'  Dates: An array of arbitrary values. Strings, errors and Booleans passed are left unchanged, numbers have to be whole numbers and are adjusted
'  BDC  : Mod Foll etc
' -----------------------------------------------------------------------------------------------------------------------
Function SCRiPT_AdjustDate(ByVal Dates, Optional BDC As String = "Mod Foll")
Attribute SCRiPT_AdjustDate.VB_Description = "Adjusts a date (or array of dates) to be a weekday according to the supplied business day convention."
Attribute SCRiPT_AdjustDate.VB_ProcData.VB_Invoke_Func = " \n32"
          Static HaveChecked As Boolean, AnyErrors As Boolean
          Dim AmendedDates
          Dim i As Long
          Dim Indicators() As Boolean
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result
1         On Error GoTo ErrHandler
2         Force2DArrayR Dates
3         AmendedDates = Dates
4         NR = sNRows(Dates): NC = sNCols(Dates)
5         ReDim Indicators(1 To NR, 1 To NC)
6         For i = 1 To NR
7             For j = 1 To NC
8                 If Not IsNumberOrDate(Dates(i, j)) Then
9                     AnyErrors = True
10                    Indicators(i, j) = True
11                    AmendedDates(i, j) = 0
12                Else
13                    If Dates(i, j) <> CLng(Dates(i, j)) Then Throw "Dates must be an array of dates, but element " + CStr(i) + "," + CStr(j) + " is not a date (because it is not a whole number)"
14                End If
15            Next j
16        Next i
17        If gRInDev Or Not HaveChecked Then CheckR "SCRiPT_AdjustDate", gPackages, gRSourcePath + "UtilsPGS.R", "SourceAllFiles()": HaveChecked = True

18        Result = Application.Run("BERT.Call", "AdjustDateExcel", AmendedDates, BDC)
19        If AnyErrors Then
20            Result = sArrayIf(Indicators, Dates, Result)
21        End If
22        SCRiPT_AdjustDate = Result

23        Exit Function
ErrHandler:
24        SCRiPT_AdjustDate = "#SCRiPT_AdjustDate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : NextIMMDate
' Author     : Philip Swannell
' Date       : 10-Jan-2018
' Purpose    : The IMM date on or after the input array of dates
' Parameters :
'  Dates:       May be an array, numbers must be whole numbers, non numbers passed through unchanged
' -----------------------------------------------------------------------------------------------------------------------
Function SCRiPT_NextIMMDate(ByVal Dates, Optional WithSerial As Boolean = False)
Attribute SCRiPT_NextIMMDate.VB_Description = "Returns the IMM date (third Wednesday of Mar, Jun, Sep or Dec) that is on or after the input Dates, which may be an array."
Attribute SCRiPT_NextIMMDate.VB_ProcData.VB_Invoke_Func = " \n32"
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result
1         On Error GoTo ErrHandler
2         Force2DArrayR Dates
3         NR = sNRows(Dates): NC = sNCols(Dates)
          Result = Dates
4         For i = 1 To NR
5             For j = 1 To NC
8                 Result(i, j) = NIMM(Dates(i, j), WithSerial)
10            Next j
11        Next i
12        SCRiPT_NextIMMDate = Result
13        Exit Function
ErrHandler:
14        SCRiPT_NextIMMDate = "#SCRiPT_NextIMMDate (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : NIMM
' Author     : Philip Swannell
' Date       : 10-Jan-2018
' Purpose    : The IMM date on or after the input date x
' -----------------------------------------------------------------------------------------------------------------------
Private Function NIMM(x As Variant, WithSerial As Boolean)
          Dim d As Long
          Dim m As Long
          Dim MPlus As Long
          Dim TW As Long
          Dim y As Long

1         On Error GoTo ErrHandler
2         If Not IsNumber(x) Then
3             NIMM = x
4             Exit Function
5         End If

6         y = Year(x): m = Month(x): d = Day(x)
7         If m Mod 3 = 0 Or WithSerial Then
8             TW = ThirdWed(y, m)
9             If d <= TW Then
10                NIMM = DateSerial(y, m, TW)
11            Else
12                MPlus = IIf(WithSerial, m + 1, m + 3)
13                NIMM = DateSerial(y, MPlus, ThirdWed(y, MPlus))
14            End If
15        Else
16            NIMM = DateSerial(y, m + 3 - (m Mod 3), ThirdWed(y, m + 3 - (m Mod 3)))
17        End If

18        Exit Function
ErrHandler:
19        Throw "#NIMM (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ThirdWed
' Author     : Philip Swannell
' Date       : 10-Jan-2018
' Purpose    : The third wednesday in the input Y/M
' -----------------------------------------------------------------------------------------------------------------------
Private Function ThirdWed(y As Long, m As Long)
1         On Error GoTo ErrHandler
2         ThirdWed = 21 - (DateSerial(y, m, 3) Mod 7)
3         Exit Function
ErrHandler:
4         Throw "#ThirdWed (line " & CStr(Erl) & "): " & Err.Description & "!"
End Function

