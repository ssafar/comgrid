@echo off
rem Builds bin\cgr.exe with the Roslyn compiler from Visual Studio / Build Tools,
rem against the .NET Framework 4.8 that ships with Windows. No SDK or NuGet needed.
setlocal
set VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe
set CSC=
for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -find MSBuild\**\Bin\Roslyn\csc.exe`) do set CSC=%%i
if not defined CSC (
    echo No Roslyn csc.exe found. Install Visual Studio Build Tools with the MSBuild component.
    exit /b 1
)
cd /d "%~dp0"
if not exist bin mkdir bin
"%CSC%" -nologo -langversion:latest -optimize -warn:4 -out:bin\cgr.exe ^
    -r:Microsoft.CSharp.dll -r:System.Web.Extensions.dll src\*.cs
