
rem echo off

rem Must be run from the projects git\project\scripts folder - everything is relative
rem run >build [versionNumber]
rem versionNumber is YY.MM.DD.build-number, like 20.5.8.1
rem

c:
cd \Git\aoContactManager\scripts

set collectionName=aoContactManager
set collectionPath=..\collections\aoContactManager\
set solutionName=ContactManager.sln
set binPath=..\source\ContactManager\bin\debug\
set msbuildLocation=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\
set deploymentFolderRoot=C:\deployments\aoContactManager\Dev\

set year=%date:~12,4%
set month=%date:~4,2%
if %month% GEQ 10 goto monthOk
set month=%date:~5,1%
:monthOk
set day=%date:~7,2%
if %day% GEQ 10 goto dayOk
set day=%date:~8,1%
:dayOk
set versionMajor=%year%
set versionMinor=%month%
set versionBuild=%day%
set versionRevision=1
rem
rem if deployment folder exists, delete it and make directory
rem
:tryagain
set versionNumber=%versionMajor%.%versionMinor%.%versionBuild%.%versionRevision%
if not exist "%deploymentFolderRoot%%versionNumber%" goto :makefolder
set /a versionRevision=%versionRevision%+1
goto tryagain
:makefolder
md "%deploymentFolderRoot%%versionNumber%"

rem ==============================================================
rem
echo build 
rem
cd ..\source
"%msbuildLocation%msbuild.exe" %solutionName%
if errorlevel 1 (
   echo failure building
   pause
   exit /b %errorlevel%
)
cd ..\scripts

rem ==============================================================
rem
echo Build addon collection
rem

rem clean collection folder (leave html)
del "%collectionPath%%collectionName%.zip" /Q
del "%collectionPath%*.dll" /Q
del "%collectionPath%*.config" /Q
del "%collectionPath%*.lic" /Q
del "%collectionPath%*.dep" /Q

rem copy bin folder assemblies to collection folder
copy "%binPath%*.dll" "%collectionPath%"
copy "%binPath%*.lic" "%collectionPath%"
copy "%binPath%*.config" "%collectionPath%"
copy "%binPath%*.dep" "%collectionPath%"

cd %collectionPath%
"c:\program files\7-zip\7z.exe" a "%collectionName%.zip"
xcopy "%collectionName%.zip" "%deploymentFolderRoot%%versionNumber%" /Y
xcopy "%collectionName%.zip" "c:\deployments\_current_sprint" /Y
cd ..\..\scripts

rem clean collection folder (leave html and collection xml)
del "%collectionPath%*.dll" /Q
del "%collectionPath%*.config" /Q
del "%collectionPath%*.lic" /Q
del "%collectionPath%*.dep" /Q
del "%collectionPath%%collectionName%.zip" /Q
