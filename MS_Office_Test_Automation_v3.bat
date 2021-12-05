@echo off
echo MSOffice Test Suite

SETLOCAL ENABLEDELAYEDEXPANSION
set workloadPath=bin\\Release\\Executable
set key=#        rem keyword to identify comments
Set inputtxtfile=..\\input\\
set CmdLine=
set wprbinary="C:\\Program Files (x86)\\Windows Kits\\10\\Windows Performance Toolkit\\wpr.exe"
SET "timestamp=%date:~10,4%%date:~4,2%%date:~7,2%-%time:~0,2%%time:~3,2%%time:~6,2%"
SET "savedir=OPTS-%timestamp%"
SET savedir=%savedir: =%

if exist TestSuite_OverAll_CSVResult.csv Call del TestSuite_OverAll_CSVResult.csv 
if exist TestSuite_CSVResults if exist ".\\TestSuite_CSVResults\\*.csv" Call del ".\\TestSuite_CSVResults\\*.csv"
if not exist TestSuite_CSVResults Call !mkdir TestSuite_CSVResults

if exist TestSuite_OverAll_Logs.log Call del TestSuite_OverAll_Logs.log 
if exist TestSuite_Logs if exist ".\\TestSuite_Logs\\*.log" Call del ".\\TestSuite_CSVResults\\*.log"
if not exist TestSuite_Logs Call !mkdir TestSuite_Logs

rem iterate over the workloads folder

Set /p "option=Choose Input type 1.Default 2.Custom"

for /D %%W in ("*Excel_*","*Word_*","*Powerpoint_*","*Outlook_*") DO (  
   echo Workload %%W 
   pushd %CD%   
   cd %%W\\%%W\\%workloadPath%
   set binary=%%W
   
   rem exit /b 1
   
   set benchmarkname=null   
   for /F "tokens=1*  delims=_" %%a in ("!binary!") do (
       set benchmarkname=%%b
   )
   
   if !option!==1 ( 
     
     Call !wprbinary! -start %~dp0%%W\OfficeSuiteWprp.wprp
     Call !binary! default > !binary!_default.log
     Call !wprbinary! -stop !binary!.etl

     
   ) else (
     set CmdLine=!binary!
     echo !CmdLine!
     set /a iterationCount = 1
    
     
     for /f "tokens=*  usebackq delims= " %%a in (`"findstr /n ^^ %inputtxtfile%%%W.txt "`) do ( 
        rem check for new line
        set "var=%%a"
        set "var=!var:*:=!"     
        if not defined var  (  
          rem new line is encountered - Run the workload      
          Call !wprbinary! -start %~dp0%%W\OfficeSuiteWprp.wprp          
          call !CmdLine!   > !binary!_Custom_!iterationCount!.log
          Call !wprbinary! -stop !binary!_!iterationCount!.etl
          set /a iterationCount+=1
          echo.
          rem Initialize the CmdLine variable to the binary for next Input Config
          echo Next TestCase -----       
          set CmdLine=!binary! 
        ) else (
             rem check for comments
             set temp=!var:~0,1!     rem check the first letter of the line
             if !temp!==!key! (     
             rem do nothing - move to next line
             ) else (
          call set "CmdLine=%%CmdLine%% !var!" ) )
      )
     rem run the final set of argument
     Call !wprbinary! -start %~dp0%%W\OfficeSuiteWprp.wprp
     call !CmdLine!   > !binary!_Custom_!iterationCount!.log
     Call !wprbinary! -stop !binary!_!iterationCount!.etl

     set /a iterationCount+=1
     echo.
     echo. )
     
     rem Copying CSV to target results folder in root
     rem first it fetches all csv's in a directory and then checks if there is a number after "-" if its so it will copy to target folder
     rem we are using "-" because usually the format is {Workload}_date-time.csv
     set timestamp=null
     for %%Q in (*.csv) do (        
          for /F "tokens=1*  delims=-" %%a in ("%%~nQ") do (
             SET "csv_var="&for /f "delims=0123456789" %%i in ("%%b") do set csv_var=%%i
             if NOT "%%b"=="" (
                if NOT defined csv_var (
                    copy %%Q  %~dp0\\TestSuite_CSVResults\\
                )
             )
          
         )
     )

     copy !binary!_*.log %~dp0\\TestSuite_Logs\\
     timeout /t 2 >nul
     rem Cleaning up files

     ECHO Cleaning Up
     mkdir %savedir%     
     move !binary!*.json %savedir%
     move !binary!_*.log %savedir%
     move !binary!*.etl %savedir%
     move *!benchmarkname!*.x* %savedir%  > nul 2>&1
     move *!benchmarkname!*.pdf* %savedir% > nul 2>&1
     move *!benchmarkname!*.csv* %savedir% > nul 2>&1
     move *!benchmarkname!*.doc* %savedir% > nul 2>&1
     move *!benchmarkname!*.pp* %savedir% > nul 2>&1
	 timeout /t 2 >nul
   popd
  )


copy .\\TestSuite_CSVResults\\*.csv TestSuite_OverAll_CSVResult.csv
timeout /t 4
del /Q .\\TestSuite_CSVResults 
rmdir /Q TestSuite_CSVResults   

copy .\\TestSuite_Logs\\*.log TestSuite_OverAll_Logs.log
timeout /t 4
del /Q .\\TestSuite_Logs 
rmdir /Q TestSuite_Logs   


ENDLOCAL

