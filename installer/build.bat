SET ZIPCMD=winrar.exe
SET PLUGIN_FOLDER=NavRun500ImporterPlugin
SET INSTALL_PACKAGE_NAME=NavRun500ImporterPlugin.st3plugin
SET PLUGIN_GUID=ceb3610c-6caf-46ec-b6bb-7c857ca075af
SET INSTALL_FOLDER=C:\ProgramData\ZoneFiveSoftware\SportTracks\3\Plugins\Installed\%PLUGIN_GUID%\%PLUGIN_FOLDER%

RMDIR /S /Q %INSTALL_FOLDER%
MD %INSTALL_FOLDER%
COPY ..\plugin.xml %INSTALL_FOLDER%
COPY ..\bin\Debug\NavRun500ImporterPlugin.dll %INSTALL_FOLDER%

RMDIR /S /Q %PLUGIN_FOLDER%
MD %PLUGIN_FOLDER%
DEL %INSTALL_PACKAGE_NAME%
COPY ..\plugin.xml %PLUGIN_FOLDER%
COPY ..\bin\Release\NavRun500ImporterPlugin.dll %PLUGIN_FOLDER%
%ZIPCMD% a -afzip %INSTALL_PACKAGE_NAME% %PLUGIN_FOLDER%

