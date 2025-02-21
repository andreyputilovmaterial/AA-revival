@ECHO OFF

ECHO Clear up dist\...
IF EXIST dist (
    REM -
) ELSE (
    MKDIR dist
)
DEL /F /Q dist\*

ECHO Calling pinliner...
REM REM :: comment: please delete .pyc files before every call of the mdmtools_aarevival - this is implemented in my fork of the pinliner
@REM python src-make\lib\pinliner\pinliner\pinliner.py src -o dist/mdmtools_aarevival.py --verbose
python src-make\lib\pinliner\pinliner\pinliner.py src -o dist/mdmtools_aarevival.py
if %ERRORLEVEL% NEQ 0 ( echo ERROR: Failure && pause && exit /b %errorlevel% )
ECHO Done

ECHO Patching mdmtools_aarevival.py...
ECHO # ... >> dist/mdmtools_aarevival.py
ECHO # print('within mdmtools_aarevival') >> dist/mdmtools_aarevival.py
REM REM :: no need for this, the root package is loaded automatically
@REM ECHO # import mdmtools_aarevival >> dist/mdmtools_aarevival.py
ECHO from src import program >> dist/mdmtools_aarevival.py
ECHO program.cli() >> dist/mdmtools_aarevival.py
ECHO # print('out of mdmtools_aarevival') >> dist/mdmtools_aarevival.py

PUSHD dist
COPY ..\run_loadmap.bat .\run_aarevival_reload_map.bat
COPY ..\run_applymap.bat .\run_aarevival_apply_write_mrs.bat
powershell -Command "(gc 'run_aarevival_reload_map.bat' -encoding 'Default') -replace '(dist[/\\])?mdmtools_aarevival.py', 'mdmtools_aarevival.py' | Out-File -encoding 'Default' 'run_aarevival_reload_map.bat'"
powershell -Command "(gc 'run_aarevival_apply_write_mrs.bat' -encoding 'Default') -replace '(dist[/\\])?mdmtools_aarevival.py', 'mdmtools_aarevival.py' | Out-File -encoding 'Default' 'run_aarevival_apply_write_mrs.bat'"
COPY ..\run_build_spss.bat .\run_aarevival_build_spss.bat
@REM COPY ..\601_SavPrepRevival.mrs .\601_SavPrepRevival.mrs
python ..\build_copy_with_utf8_bom.py  ..\SavPrepTemplate.mrs .\601_SavPrepRevival.mrs
POPD


ECHO End

