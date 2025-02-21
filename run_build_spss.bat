@ECHO OFF



SET PROJECT_NUM=2400814M



REM set delivery folder name, automatically sets to MM-DD-YY
REM works automatically
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YY=%dt:~2,2%" & set "YYYY=20%dt:~2,2%" & set "MM=%dt:~4,2%" & set "DD=%dt:~6,2%" & set "HHMMSS=%dt:~8,6%"
set "DELIVERYNAME_TEAM=%MM%.%DD%.%YYYY%"
REM set "DELIVERYNAME_LOCAL=%YYYY%%MM%%DD%_%HHMMSS%_script"


@REM :: you can set a delivery location - files will be copied automatically - unsuppress this line and several lines at the bottom
@REM :: date in folder name is generated autocatically
@REM SET "DELIVERY_LOCATION_TEAM=P:\150_Jagou\150_BES Fiscal Year 25 Global\IPS\SPSS\%DELIVERYNAME_TEAM%"







@REM ECHO please make sure AA definitions are applied
@REM pause



DEL "Outputs\R%PROJECT_NUM%.spss"
DEL "Outputs\R%PROJECT_NUM%.sav"
DEL "Outputs\R%PROJECT_NUM%.mdd"
DEL "Dara\S%PROJECT_NUM%.mdd"


ECHO - Apply AA definitions
CALL run_aarevival_apply_write_mrs
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && exit /b %errorlevel% )


ECHO - 601_SavPrepRevival
mrscriptcl 601_SavPrepRevival.mrs
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && exit /b %errorlevel% )


ECHO - 602_SavCreateUnicode
dmsrun 602_SavCreateUnicode.dms
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && exit /b %errorlevel% )



@REM :: copy to project team location - unsuppress if needed
@REM IF NOT EXIST "%DELIVERY_LOCATION_TEAM%\" (
@REM     MKDIR "%DELIVERY_LOCATION_TEAM%\"
@REM )
@REM COPY "Outputs\R%PROJECT_NUM%.sav" "%DELIVERY_LOCATION_TEAM%\"

