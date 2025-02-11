@ECHO OFF
SETLOCAL enabledelayedexpansion


@REM :: insert your files here
SET "MDD_FILE=..\tests\working\current\R2400814M.mdd"


@REM :: insert your files here
SET "MAP_FILE="








@REM :: a config option
@REM :: 3 possible choices here
@REM :: you can set it to:
@REM :: 1. "from_mdd_or_blank"
@REM ::    - it means, analysis values are filled with what was stored in MDD, and what is missing is left blank
@REM ::    this is useful if you have AA applied and you need to load the definitions - you need to keep using analysis values that were saved from AA
@REM :: 2. "from_mdd_and_autofill"
@REM ::    - it means, analysis values are filled with what is stored in MDD, and the rest is auto-filled
@REM ::    this is an autofill option but is based on analysis values that were stored in MDD
@REM :: 3. "autofill_override"
@REM ::    - it means, all analysis values are filled automatically starting from 1
@REM ::    this is useful when we have a fresh MDD that never had AA map before, and if it has analysis values possibly coming from WA MDD - they should be ignored
@REM :: 4. "blank"
@REM ::    - leave all blank, no auto-population, so that the team will have to review and fill
SET "CONFIG_FILL_ANALYSISVALUES=from_mdd_or_blank"





ECHO -
ECHO 1. read MDD and update the map
python dist/mdmtools_aarevival.py --program update_map --mdd "%MDD_FILE%" --map "%MAP_FILE%" --config-fill-analysisvalues "%CONFIG_FILL_ANALYSISVALUES%"
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && goto CLEANUP && exit /b %errorlevel% )

@REM pause
