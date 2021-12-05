@echo off
echo MSOffice Test Suite

SETLOCAL ENABLEDELAYEDEXPANSION
set workloadPath=.\
set inputtxtfile=..\\input\\Excel_Sort.txt
set wprpprofilepath=..\\..\\..\\..\\OfficeSuiteWprp.wprp
set binary=Excel_Sort
set workloadname=Sort
rem set wprprecord = 1 for recording events and generating etl file, set  wprprecord = 0 for not recording events 
set /a wprprecord=1
set wprbinary="C:\\Program Files (x86)\\Windows Kits\\10\\Windows Performance Toolkit\\wpr.exe"
SET "timestamp=%date:~10,4%%date:~4,2%%date:~7,2%-%time:~0,2%%time:~3,2%%time:~6,2%"
SET "savedir=OPTS-%timestamp%"
SET savedir=%savedir: =%

pushd %CD%
cd %workloadPath%

Set /p "option=Choose Input type 1.Default 2.Custom"   
   
   if !option!==1 ( 
     
     if %wprprecord% == 1 Call !wprbinary! -start %~dp0!wprpprofilepath!        
     Call !binary! default > !binary!_default.log
     if %wprprecord% == 1 Call !wprbinary! -stop !binary!.etl

     ECHO Cleaning Up
     mkdir %savedir%
     move !binary!_*.csv %savedir%
	 timeout /t 1 > nul 2>&1
     move !binary!_*.json %savedir%
	 timeout /t 1 > nul 2>&1
     move !binary!_*.log %savedir%
	 timeout /t 1 > nul 2>&1
     move !binary!*.etl %savedir%
	 timeout /t 1 > nul 2>&1
     move *!workloadname!*.xl* %savedir%

   ) else (
     set CmdLine=!binary!
     echo !CmdLine!
     set /a iterationCount = 1
     
     for /f "tokens=*  usebackq delims= " %%a in (`"findstr /n ^^ %inputtxtfile% "`) do ( 
        rem check for new line
        set "var=%%a"
        set "var=!var:*:=!"     
        if not defined var  (  
          rem new line is encountered - Run the workload
          if %wprprecord% == 1 Call !wprbinary! -start %~dp0!wprpprofilepath!       
          call !CmdLine! > !binary!_Custom_!iterationCount!.log
          if %wprprecord% == 1 Call !wprbinary! -stop !binary!_!iterationCount!.etl            
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
     if %wprprecord% == 1 Call !wprbinary!  -start %~dp0!wprpprofilepath!        
     call !CmdLine! > !binary!_Custom_!iterationCount!.log
     if %wprprecord% == 1 Call !wprbinary! -stop  !binary!_!iterationCount!.etl        
     set /a iterationCount+=1
     ECHO Cleaning Up
     mkdir %savedir%
     move !binary!_*.csv %savedir%
	 timeout /t 1 > nul 2>&1
     move !binary!_*.json %savedir%
	 timeout /t 1 > nul 2>&1
     move !binary!_*.log %savedir%
	 timeout /t 1 > nul 2>&1
     move !binary!*.etl %savedir%
	 timeout /t 1 > nul 2>&1
     move *!workloadname!*.xl* %savedir%
     
     echo.
     echo. 
     
     
     )

ECHO Done.
popd


