Attribute VB_Name = "modGlobals"
Option Explicit

Public Const gCayleyTradesWorkbookName = "CayleyTrades.xlsm"
Public Const g_Col_GreyText = 8421504

Public Const MN_CM As String = "cayleymodel"
Public Const MN_CMS As String = "cayleymodelshocked"
Public Const MN_CMH As String = "cayleymodelhistoric"
Public Const MN_CMHS As String = "cayleymodelhistoricshocked"

Public gModel_CM As Dictionary
Public gModel_CMS As Dictionary
Public gModel_CMH As Dictionary
Public gModel_CMHS As Dictionary
Public gMarketData As Dictionary

Public Const gRegKey_Defn = "CayleyScenarioDefinitionFiles"
Public Const gRegKey_Res = "CayleyScenarioResultsFiles"
Public Const gOldestSupportedScenarioVersion = 632 'i.e. Cayley 2022

Public Const gUseThreads = False 'In tests 6/4/2022 using threads was slower than not using threads (sigh), so removed UseThreads from Config sheet, and have this setting instead, set to False

'Release script checks that these two are up-to-date.
Public Const gMinimumSolumAddinVersion = 2336
Public Const gMinimumSolumSCRiPTUtilsVersion = 262

Public Const gMinimumMarketDataWorkbookVersion = 252
Public Const gGitHubRepo = "https://github.com/SolumCayley/Cayley.jl"
Public Const gGitHubAccount = "treasury.support@airbus.com"

Sub PleaseOpenOtherBooks()
          Dim SheetName As String
1         If ActiveSheet Is shScenarioDefinition Then
2             SheetName = shScenarioDefinition.Name
3         Else
4             SheetName = shCreditUsage.Name
5         End If

6         Throw "The Trades files, the Market Data workbook and the Lines workbook " & _
              "must all be open before you can use the this worksheet." & vbLf & vbLf & _
              "Use the Menu button on the " & SheetName & " worksheet to open them.", True
End Sub

