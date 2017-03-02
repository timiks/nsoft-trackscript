@echo off
:start
if not exist %CERT_FILE% goto certificate

:: Package .air
echo.
echo Packaging %AIR_FILE% using certificate %CERT_FILE%...
call adt -package -tsa none %SIGNING_OPTIONS% %AIR_FILE% %APP_XML% %FILE_OR_DIR%
if errorlevel 1 goto failed

:: Bundle
echo.
echo Creating bundle...
call adt -package -target bundle %BUNDLE_DIR% %AIR_FILE%
if errorlevel 1 goto bundle_failed
goto end

:certificate
echo.
echo Certificate not found: %CERT_FILE%
echo Creating certificate: %CERT_FILE%...
call bat\CreateCertificate.bat
goto start

:failed
echo AIR setup creation FAILED.
echo.
echo Troubleshooting:
echo - did you build your project in FlashDevelop?
echo - verify AIR SDK target version in %APP_XML%
echo.
if %PAUSE_ERRORS%==1 pause
exit

:bundle_failed
echo Bundle creation failed.
if %PAUSE_ERRORS%==1 pause
exit

:end
echo.
echo Done.