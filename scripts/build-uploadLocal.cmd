
rem all paths are relative to the git scripts folder

set appName=menucrm0210

rem prompt user if appName is correct
@echo Build project and install on site: %appName%
pause

call build.cmd

rem upload to contensive application
c:
cd "%deploymentFolderRoot%%versionNumber%"
cc -a %appName% --installFile "%collectionName%.zip"
cd ..\..\scripts
