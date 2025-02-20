@ECHO OFF
SETLOCAL enabledelayedexpansion


SET PROJECT_NUM=2400814M



@REM :: insert your files here
SET "MDD_FILE=.\Data\R%PROJECT_NUM%.mdd"


@REM :: insert your files here
SET "MAP_FILE=.\AnalysisAuthorRevival.xlsx"




@REM :: the path where outout files are saved
@REM :: "." means the same directory as this script
@REM :: empty path ("") means the same directory as your MDD
@REM :: temporary files are still created at the location of your MDD (it all is configured within this BAT file - adjust if you need)
@REM :: but temp files are deleted at the end of script (if you have CONFIG_CLEAN_TEMP_MIDDLE_FILES==1==1)

REM SET "OUT_PATH="
SET "OUT_PATH=./601_SavPrepRevival.mrs"









ECHO -
ECHO 1. read MDD and update the map
python dist/mdmtools_aarevival.py --program write_mrs --mdd "%MDD_FILE%" --map "%MAP_FILE%" --out "%OUT_PATH%"
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && goto CLEANUP && exit /b %errorlevel% )

@REM pause
