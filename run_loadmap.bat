@ECHO OFF
SETLOCAL enabledelayedexpansion


SET PROJECT_NUM=2400814M



@REM :: insert your files here
SET "MDD_FILE=.\Data\R%PROJECT_NUM%.mdd"


@REM :: insert your files here
SET "MAP_FILE=.\AnalysisAuthorRevival.xlsx"








@REM @REM :: a config option
@REM @REM :: 3 possible choices here
@REM @REM :: you can set it to:
@REM @REM :: 1. "from_mdd_or_blank"
@REM @REM ::    - it means, analysis values are filled with what was stored in MDD, and what is missing is left blank
@REM @REM ::    this is useful if you have AA applied and you need to load the definitions - you need to keep using analysis values that were saved from AA
@REM @REM :: 2. "from_mdd_and_autofill"
@REM @REM ::    - it means, analysis values are filled with what is stored in MDD, and the rest is auto-filled
@REM @REM ::    this is an autofill option but is based on analysis values that were stored in MDD
@REM @REM :: 3. "autofill_override"
@REM @REM ::    - it means, all analysis values are filled automatically starting from 1
@REM @REM ::    this is useful when we have a fresh MDD that never had AA map before, and if it has analysis values possibly coming from WA MDD - they should be ignored
@REM @REM :: 4. "blank"
@REM @REM ::    - leave all blank, no auto-population, so that the team will have to review and fill
@REM SET "CONFIG_FILL_ANALYSISVALUES=from_mdd_or_blank"





ECHO -
ECHO 1. read MDD and update the map
@REM python dist/mdmtools_aarevival.py --program update_map --mdd "%MDD_FILE%" --map "%MAP_FILE%" --config-fill-analysisvalues "%CONFIG_FILL_ANALYSISVALUES%"
python dist/mdmtools_aarevival.py --program update_map --mdd "%MDD_FILE%" --map "%MAP_FILE%" --create-map-if-not-exists yes
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && goto CLEANUP && exit /b %errorlevel% )

@REM pause
