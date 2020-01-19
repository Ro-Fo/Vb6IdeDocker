# Recipe for compiling Legacy VB6 Apps in Docker

This recipe explains how to build a Legacy VB6 App in a windows docker environment.

Please read the [windows docker documentation](https://docs.docker.com/docker-for-windows/)  before you start reading this article.

## Overview

- Step 1: requirements
- Step 2: pre-existing image
- Step 3: Install VB6 IDE and Visual Basic 6 SP6
- Step 4: Register third-party components
- FAQ

## Step 1: requirements

- [Docker for Windows](https://docs.docker.com/docker-for-windows/)
- VB6 IDE Installation (CD)
- [Visual Basic 6 SP6](https://www.microsoft.com/de-de/download/details.aspx?id=5721) (is needed for bigger size projects)
- your Legacy third-party components (ActiveX, *.oca, *.ocx, *.dll, ...)

Keep in mind if you install third-party components in docker, install it  ```silent```.

For obvious legal reasons, I'm not providing any prebuilt images of this on docker hub nor anywhere else.

## Step 2: pre-existing image

Use an official image from Microsoft [Docker Hub](https://hub.docker.com/u/microsoft/).

You have some options:

1. Windows 10:

``` dockerfile
	FROM mcr.microsoft.com/windows:1903
```

2. Windows Server Core:

``` dockerfile
    FROM mcr.microsoft.com/windows/servercore:1909
```

3. Windows with .NET Framework:

``` dockerfile
    FROM mcr.microsoft.com/dotnet/framework/sdk:4.8
```

My recommendation is to start with the full windows version and if you get it working you can change to a smaller image. Keep in mind that their can be problems with third-party components on the Windows-Server-Core platform.

## Step 3: Install VB6 IDE and Visual Basic 6 SP6

These both installations are very old so you need to tweak the installations.

```VB6 IDE```:

``` bat
    REM Set Installation Source Folder
    SET INSTDIR=%cd%\Visual_Basic_6_Professional
    REM registry patch which skip the first GUI section of the install (32bit,64bit)
    REGEDIT /S "%INSTDIR%\KEY.DAT"
    REM Run the VB Installer itself with the serial number
    START /WAIT %INSTDIR%\setup\acmsetup.exe /K "12345678" /T "%INSTDIR%\setup\VB98PRO.STF" /S "%INSTDIR%\" /n "User Name" /o "Company name" /b 1 /gc %cd%\vb6_install_log.txt /qtn
```

KEY.DAT:

``` dat
    REGEDIT4

    [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic\SetupWizard]
    "aspo"=dword:00000000

    [HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic\SetupWizard]
    "aspo"=dword:00000000
```

Without this registry-key the Installer will get stucked in a modal window.

```VB6 SP6```:

The Service Pack is needed to make the VB6 IDE stable. When your compiling step is getting an ```Access Violation``` Exception, you will need the Service Pack since your project is getting too big. Otherwise you can leave it.

To install vbsp6, extract `VS6sp61.cab` (VS6sp62.cab, VS6sp63.cab, VS6sp64.cab must be beside the file). Then you copy the `VB98` Dir over the existing VB6 installation.

## Step 4: Install old third-party components

If you use third party components in the compile process you will need to install them in the Docker-Container. In the most cases the installers are too old so you need register the DLLs in the container by yourself. It helps when you have a VM configured with the installed components. So you can have a look how they got installed and what are the local files you need.

Copy Files in the Container and then register them like this:

``` bat
    @ECHO OFF

    start /wait C:\Windows\System32\regsvr32.exe /s .\AnyNative.dll
    echo Exit Code is %errorlevel%

    start /wait C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe .\AnyComInteropNetAssemblies.dll /codebase /tlb /nologo
    echo Exit Code is %errorlevel%
```

If u having problems with the Native-DLL Registration check the Exitcode if <> 0.

[regsvr32 Exitcodes](https://devblogs.microsoft.com/oldnewthing/20180920-00/?p=99785):
- 0 Success
- 1	Error parsing command line
- 2	OleInitialize failed
- 3	LoadLibrary failed
- 4	GetProcAddress failed
- 5	Registration function failed

### third-party licences

If you installed the third party component on the normal way, you can skip this step. If you have registered the component, then you need to activate the third party component.

There are two license types. A License which is needed for production and one for development. In the most cases its for development so you need them activated when you compile your VB6 project.

1. license key as a file:
If their are files like '*.lic' besides the third party components, then just copy these files to the registered components and you are fine.

2. license key in windows registry:
Try to find the license in the registry. Try to search for the company name or for for the license. If u having problems, you can try the windows tool 'Procmon.exe' (see FAQ).

### compiling

Use the following command to compile your VB6-project.

```
    ..\Microsoft Visual Studio\VB98\VB6.EXE /MAKE /OUT .\vbpFile.log .\project.vbp /outdir .\someoutputdir
```

##### trouble shooting

to analyse compile erros:

* always try to reproduce the error in small projects
* use `procmon` on third party problems

##### language pack for Windows 

The default Windows Container language pack is English. You can get compile errors if you use symbols in the method or variable signature. The Default File encoding of the vb6 classes should be ANSI.

For example: In Germany dont use "ü, ö, ä, ß, ..."


# FAQ

## Why you really want to do this? 

We have a huge VB6 Legacy Code-Base (1mil LOC) and use the [COM Interop](https://en.wikipedia.org/wiki/COM_Interop) technology to use the .NET Framework.

Of couse the biggest advantage is that the Integration of Build Servers are much easier with Docker.

Also we had repeating problems with our CI-VM Windows-Registry, since references between compiled VB6 DLLs requires to be registered. After many many builds the Windows-Registry dies.

## Process Monitor - procmon

[Process Monitor](https://docs.microsoft.com/en-us/sysinternals/downloads/procmon) is an advanced monitoring tool for Windows that shows real-time file system, Registry and process/thread activity.

In our case we can use it to analyse errors in the compile process. If u having problems and can not find a solution try to use procmon to find missing third party components or license keys. "`name not found`" is the category where you have to look.

## Install Msi-Installer Packages in Docker

If you are trying to install TortoiseSVN for example, you will need to install it like this:

``` dockerfile
    COPY Docker/TortoiseSVN.msi .
    RUN msiexec /package TortoiseSVN.msi /quiet ADDLOCAL=ALL INSTALLDIR="C:\Program Files\TortoiseSVN"
```

for git:

``` dockerfile
    COPY Docker/git.inf .
    COPY Docker/Git-2.24.0.2-64-bit.exe .
    RUN Git-2.24.0.2-64-bit.exe /SILENT /NORESTART /LOADINF=git.inf
```

### Windows Server Core

The Windows Server Core image has problems with MSI-Installer-Packages. Copy the ```oledlg.dll``` DLL to the Windows System Directory if its missing.

``` dockerfile
    COPY /oledlg.dll "C:/Windows/System32"
```
