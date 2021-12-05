@echo off
rem Get Current Path
echo %path%
set defaultPath=%CD%
set defaultPath=%defaultPath: =%
rem Print Current Path
echo %defaultPath%

rem Iterate over directory
for /D %%w in ("*Excel*","*Powerpoint*", "*Word*","*Outlook*") do (
SETLOCAL ENABLEDELAYEDEXPANSION
set workloadPath=%defaultPath%\%%~nxw
rem Print Workload path
echo !workloadPath!
cd !workloadPath!
nuget install !workloadPath!\%%~nxw\packages.config -o !workloadPath!\packages\
echo !workloadPath!\%%~nxw.sln
msbuild !workloadPath!\%%~nxw.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU"
cd %defaultPath%
ENDLOCAL
)