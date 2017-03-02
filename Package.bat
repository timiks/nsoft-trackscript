@echo off
set PAUSE_ERRORS=1
call bat\SetupSDK.bat
call bat\SetupApplication.bat
call bat\Packager.bat

echo.
echo Deleting unneeded files from bundle...

cd bundle
del mimetype
rd icons /s /q
ren META-INF meta-inf

cd Adobe AIR\Versions\1.0\Resources\
del WebKit.dll
rd  WebKit /s /q

echo.
echo Done.
pause