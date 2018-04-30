@echo off

RD /S /Q %~dp0\OpenCsvByExcel\bin
"%ProgramFiles(x86)%\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe" %~dp0\OpenCsvByExcel.sln /p:Configuration=Release /p:Platform="Any CPU" /p:CodeAnalysisRuleSet= /t:Clean;Build

powershell -ExecutionPolicy RemoteSigned -Command %~dp0\build.ps1

pause