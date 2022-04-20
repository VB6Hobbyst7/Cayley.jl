REM Author Philip Swannell 24 Jan 2022
REM copy addin files
copy C:\ProgramData\Solum\Addins\SolumAddin.xlam C:\Projects\Cayley\addins\SolumAddin.xlam
copy C:\ProgramData\Solum\Addins\SolumSCRiPTUtils.xlam C:\Projects\Cayley\addins\SolumSCRiPTUtils.xlam
copy C:\ProgramData\Solum\ExcelDNA\ExcelDna.IntelliSense.xll C:\Projects\Cayley\addins\ExcelDna.IntelliSense.xll
copy C:\ProgramData\Solum\ExcelDNA\ExcelDna.IntelliSense64.xll C:\Projects\Cayley\addins\ExcelDna.IntelliSense64.xll
REM copy the vba code, from where the current release scripts for the addins put them
del /F /Q C:\Projects\Cayley\vba\SolumAddin.xlam\*.*
del /F /Q C:\Projects\Cayley\vba\SolumAddinSCRiPTUtils.xlam\*.*
copy C:\Projects\ExcelVBA\SolumAddin\*.* C:\Projects\Cayley\vba\SolumAddin.xlam\
copy C:\Projects\ExcelVBA\SolumSCRiPTUtils\*.* C:\Projects\Cayley\vba\SCRiPTUtils.xlam\
