::  ===================================================== ::
::  ================== ProtectPages.com ================= ::
::  ===================================================== ::
::  ================== SFX-Packager v1.1 ================ ::
::  https://protectpages.com/blog/creating-portable-exe-or-self/
::  ===================================================== ::
::  ===================================================== ::

@echo off

if "%~1"=="" (
	GOTO :ERROR  
) else (
	GOTO :SUCCESS 
)

:ERROR
mshta javascript:alert("You should drag folder onto this file .\n\n\n(Please read the instruction you will see now)");close(); 
start "" "https://www.protectpages.com/blog/creating-portable-exe-or-self/"
exit

:SUCCESS
:: set path to the dragged folder path
set target_path=%~1

:: current dir path of this bat
set compiler_path=%~dp0

:: ======== removed ===========
:: get folder name 
::		for /f "delims=" %%i in ("%target_path%") do SET original_basename=%%~ni

:: open folder
cd "%target_path%\"

:: find .exe file in target folder
FOR /F "delims=" %%i IN ('dir  /s /b *.exe') DO set exe_full_path=%%i
:: find .msi file in target folder
FOR /F "delims=" %%i IN ('dir  /s /b *.msi') DO set msi_full_path=%%i

:: open parent folder
cd "%target_path%\..\"

:: if file
for /f "delims=" %%i in ("%exe_full_path%") do SET exe_file_name=%%~nxi
for /f "delims=" %%i in ("%msi_full_path%") do SET msi_file_name=%%~ni

:: write config.txt 
set target_config=config.txt
@echo ;!@Install@!UTF-8!> %target_config%

:: if same drives, then allow user to choose HARD method. Otherwise, only SOFT can be used
set question=""
set /P question= If you want to a confirmation question to come up during installation, type it here (Otherwise, press ENTER for no-question installation):

IF %question%=="" ( 
 GOTO :NO_QUESTION
)
@echo Title="Unpacking">> %target_config%
@echo BeginPrompt="%question%">> %target_config%

:NO_QUESTION
@echo RunProgram="%exe_file_name%">> %target_config%
@echo ;!@InstallEnd@!>> %target_config%

:: create archive
set temp_archive_name=temp_.7z
:: no compression (change 0 to other level for compression)
"%PROGRAMFILES%/7-zip/7z" a -mx1 %temp_archive_name% "%target_path%\*"

:: compile final exe
copy /b "%compiler_path%/7zS.sfx" + config.txt + %temp_archive_name% %msi_file_name%.exe

:: delete temp_files
del %temp_archive_name%
del config.txt

mshta javascript:alert("Complete! See package aside the source folder.");close(); 