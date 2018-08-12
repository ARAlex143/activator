@echo off&color a&mode con: cols=90 lines=29
cls
fltmc >nul 2>&1 || (
	echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
	echo UAC.ShellExecute "%~fs0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
	"%temp%\getadmin.vbs"
	del /f /q "%temp%\getadmin.vbs"
	exit /b
)
REG QUERY "HKU\S-1-5-19" >NUL 2>&1 && (
GOTO clock
) || (
echo Right click and run as administrator...
echo.
GOTO exit
)

:clock
w32tm /config /manualpeerlist:"time.windows.com" /syncfromflags:manual /update
w32tm /resync /nowait
goto DetectWindows

:DetectWindows
cls
set SYSTEMID=
set WINVER=
set OFFICECHECK=No Office installation detected
set officever= No compatible version detected, try converting to VL
set PF=
set office=
set custom=
set port=1688
goto productname

:productname
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | find "Windows Server 2012 R2" >nul
if "%ERRORLEVEL%" == "0" set SYSTEMID=Windows Server 2012 R2&goto editionid1
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | find "Windows Server 2016" >nul
if "%ERRORLEVEL%" == "0" set SYSTEMID=Windows Server 2016&goto editionid1
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | find "Windows 8.1" >nul
if "%ERRORLEVEL%" == "0" set SYSTEMID=Windows 8.1&goto editionid2
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | find "Windows 8" >nul
if "%ERRORLEVEL%" == "0" set SYSTEMID=Windows 8&goto editionid2
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | find "Windows 7" >nul
if "%ERRORLEVEL%" == "0" set SYSTEMID=Windows 7&goto editionid4
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName | find "Windows 10" >nul
if "%ERRORLEVEL%" == "0" set SYSTEMID=Windows 10&goto editionid3
goto wversion

:editionid1
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Server Standard" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Server Standard
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Datacenter" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Datacenter
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Essentials" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Essentials
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Cloud Storage" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Cloud Storage
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Azure Core" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Azure Core
goto wversion

:editionid2
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Professional" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Professional
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "ProfessionalWMC" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=ProfessionalWMC
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core Single Language" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core Single Language
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core Country Specific" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core Country Specific
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Professional N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Professional N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Embedded Industry Automotive" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Embedded Industry Automotive
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Embedded Industry Enterprise" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Embedded Industry Enterprise
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Embedded Industry Professional" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Embedded Industry Professional
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core ARM" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core ARM
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core Connected" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core Connected
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core Connected N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core Connected N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core Connected Country Specific" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core Country Specific
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core Connected Single Language" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core Connected Single Language
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Professional Student" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Professional Student
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Professional Student N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Professional Student N
goto wversion

:editionid3
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Education" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Education
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Education N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Education N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise G" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise G
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise G N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise G N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Professional" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Professional
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Professional N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Professional N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise 2015 LTSB" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise 2015 LTSB
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise 2015 LTSB N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise 2015 LTSB N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise 2016 LTSB" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise 2016 LTSB
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise 2016 LTSB N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise 2016 LTSB N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Home" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Home
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Home N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Home N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Home Single Language" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Home Single Language
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Home Country Specific" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Home Country Specific
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core Single Language" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core Single Language
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Core Country Specific" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Core Country Specific
goto wversion

:editionid4
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Professional" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Professional
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Professional N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Professional N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Professional E" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Professional E
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise N
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Enterprise E" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Enterprise E
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Embedded POSReady" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Embedded POSReady
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Embedded ThinPC" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Embedded ThinPC
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Embedded Standard" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Embedded Standard
goto wversion
)

:wversion
cls
if "%SYSTEMID%" == "Windows Server 2016" (
if "%WINVER%" == "Server Standard" set KEY=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
if "%WINVER%" == "Datacenter" set KEY=CB7KF-BWN84-R7R2Y-793K2-8XDDG
if "%WINVER%" == "Essentials" set KEY=JCKRF-N37P4-C2D82-9YXRT-4M63B
if "%WINVER%" == "Cloud Storage" set KEY=QN4C6-GBJD2-FB422-GHWJK-GJG2R
if "%WINVER%" == "Azure Core" set KEY=VP34G-4NPPG-79JTQ-864T4-R3MQX
goto wtapadapter
)

if "%SYSTEMID%" == "Windows Server 2012 R2" (
if "%WINVER%" == "Server Standard" set KEY=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
if "%WINVER%" == "Datacenter" set KEY=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
if "%WINVER%" == "Essentials" set KEY=KNC87-3J2TX-XB4WP-VCPJV-M4FWM
if "%WINVER%" == "Cloud Storage" set KEY=3NPTF-33KPT-GGBPR-YX76B-39KDD
goto wkmsserverold
)

if "%SYSTEMID%" == "Windows 7" (
if "%WINVER%" == "Professional" set KEY=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
if "%WINVER%" == "Professional N" set KEY=MRPKT-YTG23-K7D7T-X2JMM-QY7MG
if "%WINVER%" == "Professional E" set KEY=W82YF-2Q76Y-63HXB-FGJG9-GF7QX
if "%WINVER%" == "Enterprise" set KEY=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
if "%WINVER%" == "Enterprise N" set KEY=YDRBP-3D83W-TY26F-D46B2-XCKRJ
if "%WINVER%" == "Enterprise E" set KEY=C29WB-22CC8-VJ326-GHFJW-H9DH4
if "%WINVER%" == "Embedded POSReady" set KEY=YBYF6-BHCR3-JPKRB-CDW7B-F9BK4
if "%WINVER%" == "Embedded ThinPC" set KEY=73KQT-CD9G6-K7TQG-66MRP-CQ22C
if "%WINVER%" == "Embedded Standard" set KEY=XGY72-BRBBT-FF8MH-2GG8H-W7KCW
goto wkmsserverold
)

if "%SYSTEMID%" == "Windows 8.1" (
if "%WINVER%" == "Professional" set KEY=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
if "%WINVER%" == "ProfessionalWMC" set KEY=789NJ-TQK6T-6XTH8-J39CJ-J8D3P
if "%WINVER%" == "Core" set KEY=M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK
if "%WINVER%" == "Core Single Language" set KEY=BB6NG-PQ82V-VRDPW-8XVD2-V8P66
if "%WINVER%" == "Core N" set KEY=7B9N3-D94CG-YTVHR-QBPX3-RJP64
if "%WINVER%" == "Enterprise" set KEY=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
if "%WINVER%" == "Professional N" set KEY=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
if "%WINVER%" == "Enterprise N" set KEY=TT4HM-HN7YT-62K67-RGRQJ-JFFXW
if "%WINVER%" == "Embedded Industry Automotive" set KEY=VHXM3-NR6FT-RY6RT-CK882-KW2CJ
if "%WINVER%" == "Embedded Industry Enterprise" set KEY=FNFKF-PWTVT-9RC8H-32HB2-JB34X
if "%WINVER%" == "Embedded Industry Professional" set KEY=NMMPB-38DD4-R2823-62W8D-VXKJB
if "%WINVER%" == "Core Country Specific" set KEY=NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3
if "%WINVER%" == "Core ARM" set KEY=XYTND-K6QKT-K2MRH-66RTM-43JKP
if "%WINVER%" == "Core Connected" set KEY=3PY8R-QHNP9-W7XQD-G6DPH-3J2C9
if "%WINVER%" == "Core Connected N" set KEY=Q6HTR-N24GM-PMJFP-69CD8-2GXKR
if "%WINVER%" == "Core Connected Country Specific" set KEY=R962J-37N87-9VVK2-WJ74P-XTMHR
if "%WINVER%" == "Core Connected Single Language" set KEY=KF37N-VDV38-GRRTV-XH8X6-6F3BB
if "%WINVER%" == "Professional Student" set KEY=MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B
if "%WINVER%" == "Professional Student N" set KEY=TNFGH-2R6PB-8XM3K-QYHX2-J4296
goto wtapadapter
)

if "%SYSTEMID%" == "Windows 8" (
if "%WINVER%" == "Professional" set KEY=NG4HW-VH26C-733KW-K6F98-J8CK4
if "%WINVER%" == "ProfessionalWMC" set KEY=GNBB8-YVD74-QJHX6-27H4K-8QHDG
if "%WINVER%" == "Core" set KEY=BN3D2-R7TKB-3YPBD-8DRP2-27GG4
if "%WINVER%" == "Core Single Language" set KEY=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
if "%WINVER%" == "Core N" set KEY=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
if "%WINVER%" == "Enterprise" set KEY=32JNW-9KQ84-P47T8-D8GGY-CWCK7
if "%WINVER%" == "Professional N" set KEY=XCVCF-2NXM9-723PB-MHCB7-2RYQQ
if "%WINVER%" == "Enterprise N" set KEY=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
if "%WINVER%" == "Embedded Industry Professional" set KEY=JVPDN-TBWJW-PD94V-QYKJ2-KWYQM
if "%WINVER%" == "Embedded Industry Enterprise" set KEY=NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2
if "%WINVER%" == "Core Country Specific " set KEY=4K36P-JN4VD-GDC6V-KDT89-DYFKP
if "%WINVER%" == "Core ARM" set KEY=DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV
goto wkmsserverold
)

if "%SYSTEMID%" == "Windows 10" (
if "%WINVER%" == "Enterprise" set KEY=NPPR9-FWDCX-D2C8J-H872K-2YT43
if "%WINVER%" == "Professional" set KEY=W269N-WFGWX-YVC9B-4J6C9-T83GX
if "%WINVER%" == "Education" set KEY=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
if "%WINVER%" == "Enterprise 2015 LTSB" set KEY=WNMTR-4C88C-JK8YV-HQ7T2-76DF9
if "%WINVER%" == "Professional N" set KEY=MH37W-N47XK-V7XM9-C7227-GCQG9
if "%WINVER%" == "Enterprise N" set KEY=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
if "%WINVER%" == "Education N" set KEY=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
if "%WINVER%" == "Enterprise 2015 LTSB N" set KEY=2F77B-TNFGY-69QQF-B8YKP-D69TJ
if "%WINVER%" == "Home" set KEY=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
if "%WINVER%" == "Home N" set KEY=3KHY7-WNT83-DGQKR-F7HPR-844BM
if "%WINVER%" == "Home Single Language" set KEY=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
if "%WINVER%" == "Home Country Specific" set KEY=PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
goto wtapadapter
)

goto wtapadapter

:DelReg
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55C92734-D682-4D71-983E-D6EC3F16059F" /f >nul 2>&1
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0FF1CE15-A989-479D-AF46-F275C6370663" /f >nul 2>&1
reg delete "HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55C92734-D682-4D71-983E-D6EC3F16059F" /f >nul 2>&1
rem reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\59A52881-A989-479D-AF46-F275C6370663" /f >nul 2>&1
rem reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\0FF1CE15-A989-479D-AF46-F275C6370663" /f >nul 2>&1
exit /b

:RemoveIFEOEntry
if '%~1' NEQ 'osppsvc.exe' (
	reg delete "%_regKey%\%~1" /f >nul 2>&1
)
if '%~1' EQU 'osppsvc.exe' (
    reg delete "%_regKey%\osppsvc.exe" /f /v "Debugger" >nul 2>&1
    reg delete "%_regKey%\osppsvc.exe" /f /v "KMS_Emulation" >nul 2>&1
    reg delete "%_regKey%\osppsvc.exe" /f /v "KMS_ActivationInterval" >nul 2>&1
    reg delete "%_regKey%\osppsvc.exe" /f /v "KMS_RenewalInterval" >nul 2>&1
    reg delete "%_regKey%\osppsvc.exe" /f /v "Office2010" >nul 2>&1
    reg delete "%_regKey%\osppsvc.exe" /f /v "Office2013" >nul 2>&1
)
exit /b

:StartService
sc query "%1" | findstr /i "RUNNING" >nul 2>&1
if %errorlevel% NEQ 0 net start "%1" /y >nul 2>&1
sc query "%1" | findstr /i "RUNNING" >nul 2>&1
if %errorlevel% NEQ 0 sc start "%1" >nul 2>&1
exit /b

:StopService
sc query "%1" | findstr /i "STOPPED" >nul 2>&1
if %errorlevel% NEQ 0 net stop "%1" /y >nul 2>&1
sc query "%1" | findstr /i "STOPPED" >nul 2>&1
if %errorlevel% NEQ 0 sc stop "%1" >nul 2>&1
exit /b

:wtapadapter
cls
title Windows KMS Emulator By AR_Alex
setlocal EnableExtensions
setlocal EnableDelayedExpansion
for /f "tokens=2 delims==" %%a in ('wmic os get osarchitecture /value') do (
	set _OSarch=%%a
)
call :StopService
Call :DelReg
regedit /s %WINDIR%\KMS_Service\settings.reg
xcopy %WINDIR%\KMS_Service\Kms_Files\"%_OSarch%\SppExtComObjPatcher.exe" "%SystemRoot%\system32\" /y
xcopy %WINDIR%\KMS_Service\Kms_Files\"%_OSarch%\SppExtComObjHook.dll" "%SystemRoot%\system32\" /y
Call :DelReg
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /upk 
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ckms 
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /skms 10.0.0.20 
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ipk %KEY%
echo. Activating %SYSTEMID% %WINVER%
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ato
call :StopService "sppsvc"
del /f /q "%SystemRoot%\system32\SppExtComObjPatcher.exe" 
del /f /q "%SystemRoot%\system32\SppExtComObjHook.dll" 
call :RemoveIFEOEntry "SppExtComObj.exe"
sc start sppsvc trigger=timer;sessionid=0 >nul 2>&1
exit

:kmsactivatewindows
cls
title Windows 8.1 KMS Emulator By AR_Alex
echo. Preparing Files...

sc stop sppsvc >nul 2>&1
set port=1688
echo. Detecting processor type...
reg.exe query "hklm\software\microsoft\Windows NT\currentversion" /v buildlabex | find /i "amd64" >nul 2>&1
if %errorlevel% equ 0 set xOS=x64
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set xOS=x86

echo. Preparing the Firewall...
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
%WINDIR%\System32\Netsh Advfirewall Firewall add rule name="0pen Port KMS" dir=in action=allow protocol=TCP localport=%port% >nul 2>&1

echo. Installing registry settings...
regedit /s %WINDIR%\KMS_Service\Kms_Files\settings.reg

Call :DelReg

echo. Setting up KMS Address...
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /upk >nul 2>&1
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ckms >nul 2>&1
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /skms 8.7.8.7 >nul 2>&1
echo. Installing KMS Client key...
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ipk %KEY%

echo. Starting the KMS Emulator...
cd %WINDIR%\KMS_Service\Kms_Files
SECOInjector_%xOS%.exe /s 8.7.8.7 /f /l >nul 2>&1
title Windows 8.1 KMS Emulator By AR_Alex

Call :DelReg

echo. Activating...
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ato
echo. Generating Expiration Date
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /xpr

echo. Closing the KMS Emulator and cleaning up...
SECOInjector_%xOS%.exe /f /u >nul 2>&1
title Windows 8.1 KMS Emulator By AR_Alex

echo. Processing the Firewall...
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1

echo. DONE!
exit

:wkmsserverold
cls
title Windows KMS Emulator By AR_Alex
echo. Please wait while we work...
setlocal EnableExtensions
setlocal EnableDelayedExpansion
set port=1688
set kmsip=10.3.0.111
set ActivationInterval=43200
set RenewalInterval=43200
set i=0
set keyid=
:genkeyid
set /a rnd=%random% %% 10
set keyid=%keyid%%rnd%
set /a i+=1
if %i% LEQ 5 goto :genkeyid
set epid=55041-00206-152-%keyid%-03-1033-9600.0000-3002013
set xOS=x64
if "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set xOS=x86
%WINDIR%\KMS_Service\Kms_Files\%xOS%\devcon.exe remove tap0901 >nul 2>&1
echo. Installing network tap adapter...
%WINDIR%\KMS_Service\Kms_Files\%xOS%\devcon.exe install %WINDIR%\KMS_Service\Kms_Files\%xOS%\OemWin2k.inf tap0901 >nul 2>&1
echo. Setting Adapter IP...
for /f "tokens=1 delims=. " %%i in ('route print ^| find /i "TAP-Windows Adapter V9"') do (netsh interface ip set address %%i static 10.3.0.1 255.255.255.0) >nul 2>&1
echo. Preparing the Firewall...
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
%WINDIR%\System32\Netsh Advfirewall Firewall add rule name="0pen Port KMS" dir=in action=allow protocol=TCP localport=%port% >nul 2>&1
echo.&echo Starting TunMirror...
IF EXIST "%systemroot%\Microsoft.NET\Framework\v4.0.30319" goto continuetun
echo. Downloading .NET Framework 4...
certutil.exe -urlcache -split -f https://download.microsoft.com/download/5/6/2/562A10F9-C9F4-4313-A044-9C94E0A8FAC8/dotNetFx40_Client_x86_x64.exe %temp%\dotNetFx40_Client_x86_x64.exe
echo. Installing .NET Framework 4
start /wait %temp%\dotNetFx40_Client_x86_x64.exe /norestart /passive
del %temp%\dotNetFx40_Client_x86_x64.exe >nul 2>&1
:continuetun
start /b cmd /c %WINDIR%\KMS_Service\Kms_Files\TunMirror.exe >nul 2>&1
for /f "tokens=5 delims=, " %%i in ('netstat -ano ^|find ":%port% "') do taskkill /pid %%i /f >nul 2>&1
timeout /t 5 /nobreak >nul 2>&1
echo. Starting KMS Server Emulator...
start /b cmd /c %WINDIR%\KMS_Service\Kms_Files\KMSServer.exe %port% %epid% %ActivationInterval% %RenewalInterval% >nul 2>&1
timeout /t 5 /nobreak >nul 2>&1
echo. Setting up KMS Address...
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /upk
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ckms
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /skms %kmsip%
echo. Installing KMS Client key...
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ipk %KEY%
sc stop sppsvc >nul 2>&1
echo. Activating...
echo.
Call :DelReg
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ato
echo. Generating Expiration Date
echo.
sc stop sppsvc >nul 2>&1
goto closeserver
:closeserver
echo. Closing Server...
taskkill /t /f /IM kmsserver.exe >nul 2>&1
taskkill /t /f /IM tunmirror.exe >nul 2>&1
echo. Processing the firewall...
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
exit