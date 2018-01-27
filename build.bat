@echo off

RD /S /Q %~dp0\OpenCsvByExcel\bin
"%ProgramFiles(x86)%\MSBuild\14.0\Bin\msbuild.exe" %~dp0\OpenCsvByExcel.sln /p:Configuration=Release /p:Platform=x86 /p:CodeAnalysisRuleSet= /t:Clean;Build
"%ProgramFiles(x86)%\MSBuild\14.0\Bin\msbuild.exe" %~dp0\OpenCsvByExcel.sln /p:Configuration=Release /p:Platform=x64 /p:CodeAnalysisRuleSet= /t:Clean;Build

powershell -ExecutionPolicy RemoteSigned -Command %~dp0\binary.ps1

pause