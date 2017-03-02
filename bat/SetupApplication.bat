:: NOTICE: All paths are relative to project root

:: General AIR Application Settings
:: 1. Application Name
:: 2. Application ID (must match <id> of Application Descriptor)
:: 3. Application Descriptor (.xml file)
set AIR_NAME=TrackScript
set APP_ID=TrackScript
set APP_XML=application.xml

:: Certificate Settings
set CERT_NAME="%AIR_NAME%"
set CERT_PASS=diesel
set CERT_FILE="bat\Certificate.p12"
set SIGNING_OPTIONS=-storetype pkcs12 -keystore %CERT_FILE% -storepass %CERT_PASS%

:: Packaging Settings
set APP_DIR=bin
set AIR_FILE=%AIR_NAME%.air
set FILE_OR_DIR=-C %APP_DIR% .
set BUNDLE_DIR=bundle

:validation
%SystemRoot%\System32\find /C "<id>%APP_ID%</id>" "%APP_XML%" > NUL
if errorlevel 1 goto badid
goto end

:badid
echo.
echo ERROR: 
echo   Application ID in 'bat\SetupApplication.bat' (APP_ID) 
echo   does NOT match Application descriptor '%APP_XML%' (id)
echo.
if %PAUSE_ERRORS%==1 pause
exit

:end