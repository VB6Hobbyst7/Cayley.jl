Attribute VB_Name = "modGlobals"
Option Explicit

Public Const gCayleyTradesWorkbookName = "CayleyTrades.xlsm"
Public Const g_Col_GreyText = 8421504
Public Const gDebugMode = False  'If True then routines call Debug.Print and sExcelWorkingSetSize

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
Public Const gMinimumSolumAddinVersion = 2324
Public Const gMinimumSolumSCRiPTUtilsVersion = 260

Public Const gMinimumMarketDataWorkbookVersion = 245
Public Const gGitHubRepo = "https://github.com/SolumCayley/Cayley.jl"
Public Const gGitHubAccount = "treasury.support@airbus.com"
