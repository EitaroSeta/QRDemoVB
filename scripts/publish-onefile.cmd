@echo off
set OUT=%~dp0..\pub-onefile
dotnet publish %~dp0..\src\QrDemoVB -c Release -r win-x64 --self-contained true ^
  -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true ^
  -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=None -p:DebugSymbols=false ^
  -o "%OUT%"
echo.
echo === DONE ===
echo %OUT%\QrDemoVB.exe
