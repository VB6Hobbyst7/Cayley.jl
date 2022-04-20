Attribute VB_Name = "modWorldBank"
Option Explicit
'Utility functions for WorlBank project

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WorldBankParseFile
' Author     : Philip Swannell
' Date       : 27-Apr-2021
' Purpose    : Translated from the function csvnoheaderstodataframe in Julia project IPVALC, but for convenience can guess
'             the csvtype and lender from file contents.
' -----------------------------------------------------------------------------------------------------------------------
Function WorldBankParseFile(FileName As String, Optional ByVal csvtype As String = "guess", Optional ByVal lender As String = "guess")
          Dim Data, Headers

1         On Error GoTo ErrHandler
2         Data = ThrowIfError(sFileShow(FileName, vbTab, True, , , True))
3         If csvtype = "guess" Or lender = "guess" Then
4             guesscvstypeandlender Data, csvtype, lender
5         End If

6         Headers = ThrowIfError(WorldBankFileHeaders(csvtype, lender, False))
7         If sNCols(Headers) <> sNCols(Data) Then Throw "inferred headers have different number of columns to contents of file"

8         WorldBankParseFile = sArrayStack(Headers, Data)

9         Exit Function
ErrHandler:
10        WorldBankParseFile = "#WorldBankParseFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WorldBankFileHeaders
' Author     : Philip Swannell
' Date       : 27-Apr-2021
' Purpose    : Translated from the function WorldBankFileHeaders in Julia project IPVALC
' -----------------------------------------------------------------------------------------------------------------------
Private Function WorldBankFileHeaders(csvtype As String, Optional lender As String = "notspecified", Optional longform As Boolean)
          Dim Headers As Variant

1         On Error GoTo ErrHandler
2         If csvtype = "calendar" Then

3             Headers = sArrayRange("Year", "Month", "Day")
4         ElseIf csvtype = "capatmvolatility" Then

5             Headers = sArrayRange("Volatility Type", "Frequency", "Maturity", "Implied Volatility")
6         ElseIf csvtype = "cdsspreads" Then
7             Headers = sArrayRange("Maturity", "CDS Spreads", "Market Recovery")
8         ElseIf csvtype = "fixings" Then
              ' There are no Fixings files for IDA

9             Headers = sArrayRange("Currency", "Year", "Month", "Day", "Fixing Value")
10        ElseIf csvtype = "loancashflowsrepayment" Then
              '= In this case IDA files have two extra columns:
              'Interest Relief Amount in USD", "Principal Relief Amount in USD" '
11            If lender = "ibrd" Then

12                Headers = sArrayRange("Period Start in Years (ACT/365)", "Period End in Years (ACT/365)", _
                      "Outstanding Notional in Local Currency", "Yearfrac", "Redeployment Spread")
13            ElseIf lender = "ida" Then

14                Headers = sArrayRange("Period Start in Years (ACT/365)", "Period End in Years (ACT/365)", _
                      "Outstanding Notional in Local Currency", "Yearfrac", "Principal Relief Amount in USD", _
                      "Interest Relief Amount in USD", "Redeployment Spread")
15            Else
16                Throw ("lender must be specified as either 'ibrd' or 'ida' when csvtype is 'loancashflowsrepayment'")
17            End If
18        ElseIf csvtype = "loancashflowsalldata" Then
              'IBRD files have two columns that IDA files don't have:
              '"Fixing Calendar"
              'IDA files have two columns that IBRD files don't have:
              '"Credit_Type", "Non Concessional Credit (1 = Yes; 0 = No)" =#
19            If lender = "ibrd" Then
20                Headers = sArrayRange("LoanID", "Currency", "Coupon Projection Curve", _
                      "Loan Cashflows Discounting Curve", "Borrower ID", "FX Spot", _
                      "Fixity (0 = Float; 1 = Fix)", "Loan Spread", "Fixed Rate", "WB Recovery", _
                      "Effective Notional", "Accrued Interest Yearfrac", "Daycount Convention", _
                      "Cap Vol ID", "Fixing Calendar", "Cashflow Calendar", "CDS Calendar", _
                      "Hedgeswap Discount Curve", "CDS Discounting Curve", _
                      "Redeployment Cost Discount Curve")
21            ElseIf lender = "ida" Then

22                Headers = sArrayRange("LoanID", "Currency", "Coupon Projection Curve", _
                      "Loan Cashflows Discounting Curve", "Borrower ID", "FX Spot", _
                      "Fixity (0 = Float; 1 = Fix)", "Loan Spread", "Fixed Rate", "WB Recovery", _
                      "Effective Notional", "Accrued Interest Yearfrac", "Daycount Convention", _
                      "Credit_Type", "Non Concessional Credit (1 = Yes; 0 = No)", "Cap Vol ID", _
                      "Cashflow Calendar", "CDS Calendar", "Hedgeswap Discount Curve", _
                      "CDS Discounting Curve", "Redeployment Cost Discount Curve")
23            Else
24                Throw ("lender must be specified as either :ibrd or :ida when csvtype is 'loancashflowsalldata'")
25            End If
26        ElseIf csvtype = "yieldcurve" Then
27            Headers = sArrayRange("Tenor in Years (ACT/365)", "Discount Factors")
28        Else
29            Throw ("csvtype " & csvtype & " not recognised, should be one of 'calendar', 'capatmvolatility', " & _
                  "'cdsspreads', 'fixings', 'loancashflowsalldata', 'loancashflowsrepayment', 'yieldcurve'")
30        End If

31        If Not longform Then
32            Headers = abbreviateheaders(Headers)
33        End If

34        WorldBankFileHeaders = Headers

35        Exit Function
ErrHandler:
36        WorldBankFileHeaders = "#WorldBankFileHeaders (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function abbreviateheader_core(header As String)
          Dim Res As String, InstrRes As Long

1         On Error GoTo ErrHandler

2         Select Case header
              Case "Principal Relief Amount in USD"
3                 abbreviateheader_core = ""
4                 Exit Function
5             Case "Interest Relief Amount in USD"
6                 abbreviateheader_core = ""
7                 Exit Function
8         End Select

9         Res = LCase(Replace(Replace(header, " ", ""), "_", ""))
10        InstrRes = InStr(Res, "(")
11        If InstrRes > 0 Then
12            Res = Left(Res, InstrRes - 1)
13        End If

14        abbreviateheader_core = Res

15        Exit Function
ErrHandler:
16        Throw "#abbreviateheader_core (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function abbreviateheaders(ByVal Headers As Variant)
          Dim NR As Long, NC As Long, i As Long, j As Long
1         On Error GoTo ErrHandler
2         Force2DArrayR Headers, NR, NC

3         For i = 1 To NR
4             For j = 1 To NC
5                 Headers(i, j) = abbreviateheader_core(CStr(Headers(i, j)))
6             Next j
7         Next i
8         abbreviateheaders = Headers
9         Exit Function
ErrHandler:
10        Throw "#abbreviateheaders (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function guesscvstypeandlender(FileContents As Variant, ByRef csvtype As String, lender As String)

          Const ErrMessage = "Cannot guess csvtype and lender from file contents"

1         On Error GoTo ErrHandler
2         Select Case sNCols(FileContents)
              Case 3
3                 If VarType(FileContents(1, 1)) = vbString Then
4                     csvtype = "cdsspreads"
5                 Else
6                     csvtype = "calendar"
7                 End If
8             Case 5
9                 If VarType(FileContents(1, 1)) = vbString Then
10                    csvtype = "fixings"
11                Else
12                    csvtype = "loancashflowsrepayment"
13                    lender = "ibrd"
14                End If
15            Case 2
16                csvtype = "yieldcurve"
17            Case 4
18                csvtype = "capatmvolatility"
19            Case 7
20                csvtype = "loancashflowsrepayment"
21                lender = "ida"
22            Case 20
23                csvtype = "loancashflowsalldata"
24                lender = "ibrd"
25            Case 21
26                csvtype = "loancashflowsalldata"
27                lender = "ida"
28            Case Else
29                Throw ErrMessage
30        End Select
31        Exit Function
ErrHandler:
32        Throw "#guesscvstypeandlender (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
