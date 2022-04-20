Attribute VB_Name = "modISDASIMMPrivate"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ISDASIMMMakeID_Core
' Author     : Philip Swannell
' Date       : 07-Apr-2020
' Purpose    : For use from the workbook "ISDA SIMM 2020 Workbook Summary.xlsm" Idea is to have a unique identifier for
'              each number calculated for isda simm calibration. Can use the identifier to compare results year to year and (potentially)
'              for communication of results Solum to ISDA.
' Parameters :
'  AssetClass:
'  Parameter :
'  Label1    :
'  Label2    :
'  Label3    :
'  Lag       : either 10 or 1
' -----------------------------------------------------------------------------------------------------------------------
Function ISDASIMMMakeID_Core(ByVal AssetClass As String, ByVal Parameter As String, Optional ByVal Label1 As String, _
          Optional ByVal Label2 As String, Optional ByVal Label3 As String, Optional Lag As Long = 10)

          Dim Labels As String
          Dim LagString

1         On Error GoTo ErrHandler
2         Select Case Lag
              Case 10, 1
3                 LagString = "-" & CStr(Lag) & "d"
4             Case 0
5                 LagString = ""
6             Case Else
7                 Throw "Lag must be 10 or 1 or 0"
8         End Select

9         AssetClass = LCase(ISDASIMMStandardiseAssetClass(AssetClass))

10        If LCase(Parameter) = LCase("Cross-Currency – IR correlation") Then
11            Parameter = "corr"
12        End If

13        Parameter = Replace(LCase(Parameter), "-", " ")
14        Parameter = Replace(LCase(Parameter), "intra bucket correlation", "intra corr")
15        Parameter = Replace(LCase(Parameter), "inter bucket correlation", "inter corr")
16        Parameter = Replace(LCase(Parameter), "correlations", "corr")
17        Parameter = Replace(LCase(Parameter), "correlation", "corr")
18        Parameter = Replace(Parameter, " (1d)", "") 'For convenience on Workbook Summary sheet

19        If InStr(Parameter, "corr") > 0 Then
20            LagString = ""
21        End If

22        Select Case UCase(Replace(Parameter, " ", ""))
              Case "STRESSPERIOD"
23                Parameter = "stress period"
24                LagString = ""
25            Case "DELTARISKWEIGHT", "DRW", "DELTARISKWEIGHTS"
26                Parameter = "drw"
27            Case "VEGARISKWEIGHT", "VRW", "IRVEGARISKWEIGHTS"
28                Parameter = "vrw"
29        End Select

30        Select Case LCase(Label1)
              Case "1 ig"
31                Label1 = "1"
32            Case "2 hy / nr"
33                Label1 = "2"
34            Case "buckets 1- 11", "buckets 1 - 11"
35                Label1 = "1 to 11"
36        End Select

37        Label1 = Replace(LCase(Label1), "bucket ", "")
38        Label1 = Replace(LCase(Label1), ", ", ",")
39        Label1 = Replace(LCase(Label1), "regular", "reg")
40        Label1 = Replace(LCase(Label1), "-", ",")

41        Labels = Label1

42        If Label2 <> "" Then
43            Labels = Labels & "," & Label2
44        End If
45        If Label3 <> "" Then
46            Labels = Labels & "," & Label3
47        End If

          'Now split the labels and re-order in the case of correlation
48        If InStr(Parameter, "corr") > 0 Then
49            If AssetClass = "cm" Or AssetClass = "crq" Or AssetClass = "crnq" Or AssetClass = "eq" Then
50                If Len(Labels) - Len(Replace(Labels, ",", "")) = 1 Then
51                    Label1 = VBA.Split(Labels, ",")(0)
52                    Label2 = VBA.Split(Labels, ",")(1)
53                    If IsNumeric(Label1) Then
54                        If IsNumeric(Label2) Then
55                            If CDbl(Label1) > CDbl(Label2) Then
56                                Labels = Label2 + "," + Label1
57                            End If
58                        End If
59                    End If
60                End If
61            End If
62        End If
63        If Parameter <> "" Then
64            Parameter = "-" & Parameter
65        End If
66        If Labels <> "" Then
67            Labels = "-" & Labels
68        End If

69        ISDASIMMMakeID_Core = LCase(AssetClass & Parameter & Labels & LagString)

70        Exit Function
ErrHandler:
71        ISDASIMMMakeID_Core = "#ISDASIMMMakeID_Core (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function



