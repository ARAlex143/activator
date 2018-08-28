@echo off&color a&mode con: cols=90 lines=29
cls
pushd "%~dp0"
title Windows 10 KMS Activator By AR_Alex
fltmc >nul 2>&1 || (
	echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
	echo UAC.ShellExecute "%~fs0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
	"%temp%\getadmin.vbs"
	del /f /q "%temp%\getadmin.vbs"
	exit /b
)
REG QUERY "HKU\S-1-5-19" >NUL 2>&1 && (
GOTO update
) || (
echo Right click and run as administrator...
echo.
pause
GOTO exit
)
:update
cls
if /i not exist Kms_Files\ goto error1
if /i not exist Office_Files\ goto error1
Ping www.google.de -n 1 -w 1000 >nul 2>&1
if errorlevel 1 (
  cls
  set online=no
  goto DetectWindows
)
set online=yes
echo. Checking for updates...
bitsadmin /Transfer myDownloadJob https://raw.githubusercontent.com/ARAlex143/activator/master/kmsactivator.txt %temp%\kmsactivator.txt >nul 2>&1
if not exist %temp%\kmsactivator.txt goto certu
goto continueupd
:certu
echo Using certutil to check for updates
certutil.exe -urlcache -split -f https://raw.githubusercontent.com/ARAlex143/activator/master/kmsactivator.txt %temp%\kmsactivator.txt >nul 2>&1
:continueupd
if not exist %temp%\kmsactivator.txt goto DetectWindows
set /p update= < %temp%\kmsactivator.txt
del %temp%\kmsactivator.txt

if "%update%" == "3.3" goto DetectWindows
goto downloadupdate

:DetectWindows
cls
echo Checking windows installation
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
if "%SYSTEMID%" == "" set SYSTEMID=No Compatible OS Found&goto next

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
goto next

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
goto next

:editionid3
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Cloud" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Cloud
reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionID | find "Cloud N" >nul
if "%ERRORLEVEL%" == "0" SET WINVER=Cloud N
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
goto next

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
goto next

:next
echo Checking office installation
If Exist "%installdir%\Program Files (x86)\Microsoft Office\Office14\OSPP.vbs" (
set OFFICECHECK=Activate Office 2010
set office=Office 2010
set PF=alternative
set OSPP=cscript.exe /nologo "%installdir%\Program Files (x86)\Microsoft Office\Office14\ospp.vbs"
)

If Exist "%installdir%\Program Files\Microsoft Office\Office14\OSPP.vbs" ( 
set OFFICECHECK=Activate Office 2010
set office=Office 2010
set PF=normal
set OSPP=cscript.exe /nologo "%installdir%\Program Files\Microsoft Office\Office14\ospp.vbs"
)

If Exist "%installdir%\Program Files (x86)\Microsoft Office\Office15\OSPP.vbs" (
set OFFICECHECK=Activate Office 2013
set office=Office 2013
set PF=alternative
set OSPP=cscript.exe /nologo "%installdir%\Program Files (x86)\Microsoft Office\Office15\ospp.vbs"
)

If Exist "%installdir%\Program Files\Microsoft Office\Office15\OSPP.vbs" ( 
set OFFICECHECK=Activate Office 2013
set office=Office 2013
set PF=normal
set OSPP=cscript.exe /nologo "%installdir%\Program Files\Microsoft Office\Office15\ospp.vbs"
)

If Exist "%installdir%\Program Files (x86)\Microsoft Office\Office16\OSPP.vbs" (
set OFFICECHECK=Activate Office 2016
set office=Office 2016
set PF=alternative6
set path6="%installdir%\Program Files (x86)\Microsoft Office\"
set OSPP=cscript.exe /nologo "%installdir%\Program Files (x86)\Microsoft Office\Office16\ospp.vbs"
)

If Exist "%installdir%\Program Files\Microsoft Office\Office16\OSPP.vbs" ( 
set OFFICECHECK=Activate Office 2016
set office=Office 2016
set PF=normal6
set path6="%installdir%\Program Files\Microsoft Office\"
set OSPP=cscript.exe /nologo "%installdir%\Program Files\Microsoft Office\Office16\ospp.vbs"
)

If Exist "%installdir%\Program Files (x86)\Microsoft Office\Office17\OSPP.vbs" (
set OFFICECHECK=Activate Office 2019
set office=Office 2019
set PF=alternative
set OSPP=cscript.exe /nologo "%installdir%\Program Files (x86)\Microsoft Office\Office17\ospp.vbs"
)

If Exist "%installdir%\Program Files\Microsoft Office\Office17\OSPP.vbs" ( 
set OFFICECHECK=Activate Office 2019
set office=Office 2019
set PF=normal
set OSPP=cscript.exe /nologo "%installdir%\Program Files\Microsoft Office\Office17\ospp.vbs"
)
for /f "tokens=*" %%A in (
  '%OSPP% /dstatus ^| findstr /i sku'
) do set number=%%A
if "%office%" == "Office 2010" goto office2010
if "%office%" == "Office 2013" goto office2013
if "%office%" == "Office 2016" goto office2016
if "%office%" == "Office 2019" goto office2019
goto mainmenu
:office2019
rem Professional Plus C2R-P
if "%number%" == "SKU ID: 0bc88885-718c-491d-921f-6f214349e79c" set key=VQ9DP-NVHPH-T9HJC-J9PDT-KTQRG & set officever=Professional Plus C2R-P
rem Project Professional C2R-P
if "%number%" == "SKU ID: fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9" set key=XM2V9-DN9HH-QB449-XDGKC-W2RMW & set officever=Project Professional C2R-P
rem Visio Professional C2R-P
if "%number%" == "SKU ID: 500f6619-ef93-4b75-bcb4-82819998a3ca" set key=N2CG9-YD3YK-936X4-3WR82-Q3X4H & set officever=Visio Professional C2R-P
rem Professional Plus
if "%number%" == "SKU ID: 85dd8b5f-eaa4-4af3-a628-cce9e77c9a03" set key=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP & set officever=Professional Plus
rem Standard
if "%number%" == "SKU ID: 6912a74b-a5fb-401a-bfdb-2e3ab46f4b02" set key=6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK & set officever=Standard
rem Project Professional
if "%number%" == "SKU ID: 2ca2bf3f-949e-446a-82c7-e25a15ec78c4" set key=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B & set officever=Project Professional
rem Project Standard
if "%number%" == "SKU ID: 1777f0e3-7392-4198-97ea-8ae4de6f6381" set key=C4F7P-NCP8C-6CQPT-MQHV9-JXD2M & set officever=Project Standard
rem Visio Professional
if "%number%" == "SKU ID: 5b5cf08f-b81a-431d-b080-3450d8620565" set key=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB & set officever=Visio Professional
rem Visio Standard
if "%number%" == "SKU ID: e06d7df3-aad0-419d-8dfb-0ac37e2bdf39" set key=7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2 & set officever=Visio Standard
rem Access
if "%number%" == "SKU ID: 9e9bceeb-e736-4f26-88de-763f87dcc485" set key=9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT & set officever=Access
rem Excel
if "%number%" == "SKU ID: 237854e9-79fc-4497-a0c1-a70969691c6b" set key=TMJWT-YYNMB-3BKTF-644FC-RVXBD & set officever=Excel
rem Outlook
if "%number%" == "SKU ID: c8f8a301-19f5-4132-96ce-2de9d4adbd33" set key=7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK & set officever=Outlook
rem PowerPoint
if "%number%" == "SKU ID: 3131fd61-5e4f-4308-8d6d-62be1987c92c" set key=RRNCX-C64HY-W2MM7-MCH9G-TJHMQ & set officever=PowerPoint
rem Publisher
if "%number%" == "SKU ID: 9d3e4cca-e172-46f1-a2f4-1d2107051444" set key=G2KWX-3NW6P-PY93R-JXK2T-C9Y9V & set officever=Publisher
rem Skype for Business
if "%number%" == "SKU ID: 734c6c6e-b0ba-4298-a891-671772b2bd1b" set key=NCJ33-JHBBY-HTK98-MYCV8-HMKHJ & set officever=Skype for Business
rem Word
if "%number%" == "SKU ID: 059834fe-a8ea-4bff-b67b-4d006b5447d3" set key=PBX3G-NWMT6-Q7XBW-PYJGG-WXD33 & set officever=Word
goto mainmenu
:office2016
rem Project Professional C2R-P
if "%number%" == "SKU ID: 829b8110-0e6f-4349-bca4-42803577788d" set key=WGT24-HCNMF-FQ7XH-6M8K7-DRTW9 & set officever= Project Professional C2R-P
rem Project Standard C2R-P
if "%number%" == "SKU ID: cbbaca45-556a-4416-ad03-bda598eaa7c8" set key=D8NRQ-JTYM3-7J2DX-646CT-6836M & set officever=Project Standard C2R-P
rem Visio Professional C2R-P
if "%number%" == "SKU ID: b234abe3-0857-4f9c-b05a-4dc314f85557" set key=69WXN-MBYV6-22PQG-3WGHK-RM6XC & set officever=Visio Professional C2R-P
rem Visio Standard C2R-P
if "%number%" == "SKU ID: 361fe620-64f4-41b5-ba77-84f8e079b1f7" set key=NY48V-PPYYH-3F4PX-XJRKJ-W4423 & set officever=Visio Standard C2R-P
rem MondoR
if "%number%" == "SKU ID: e914ea6e-a5fa-4439-a394-a9bb3293ca09" set key=DMTCJ-KNRKX-26982-JYCKT-P7KB6 & set officever=MondoR
rem Mondo
if "%number%" == "SKU ID: 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce" set key=HFTND-W9MK4-8B7MJ-B6C4G-XQBR2 & set officever=Mondo
rem Professional Plus
if "%number%" == "SKU ID: d450596f-894d-49e0-966a-fd39ed4c4c64" set key=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 & set officever=Professional Plus
rem Standard
if "%number%" == "SKU ID: dedfa23d-6ed1-45a6-85dc-63cae0546de6" set key=JNRGM-WHDWX-FJJG3-K47QV-DRTFM & set officever=Standard
rem Project Professional
if "%number%" == "SKU ID: 4f414197-0fc2-4c01-b68a-86cbb9ac254c" set key=YG9NW-3K39V-2T3HJ-93F3Q-G83KT & set officever=Project Professional
rem Project Standard
if "%number%" == "SKU ID: da7ddabc-3fbe-4447-9e01-6ab7440b4cd4" set key=GNFHQ-F6YQM-KQDGJ-327XX-KQBVC & set officever=Project Standard
rem Visio Professional
if "%number%" == "SKU ID: 6bf301c1-b94a-43e9-ba31-d494598c47fb" set key=PD3PC-RHNGV-FXJ29-8JK7D-RJRJK & set officever=Visio Professional
rem Visio Standard
if "%number%" == "SKU ID: aa2a7821-1827-4c2c-8f1d-4513a34dda97" set key=7WHWN-4T7MP-G96JF-G33KR-W8GF4 & set officever=Visio Standard
rem Access
if "%number%" == "SKU ID: 67c0fc0c-deba-401b-bf8b-9c8ad8395804" set key=GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW & set officever=Access
rem Excel
if "%number%" == "SKU ID: c3e65d36-141f-4d2f-a303-a842ee756a29" set key=9C2PK-NWTVB-JMPW8-BFT28-7FTBF & set officever=Excel
rem OneNote
if "%number%" == "SKU ID: d8cace59-33d2-4ac7-9b1b-9b72339c51c8" set key=DR92N-9HTF2-97XKM-XW2WJ-XW3J6 & set officever=OneNote
rem Outlook
if "%number%" == "SKU ID: ec9d9265-9d1e-4ed0-838a-cdc20f2551a1" set key=R69KK-NTPKF-7M3Q4-QYBHW-6MT9B & set officever=Outlook
rem Powerpoint
if "%number%" == "SKU ID: d70b1bba-b893-4544-96e2-b7a318091c33" set key=J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6 & set officever=Powerpoint
rem Publisher
if "%number%" == "SKU ID: 041a06cb-c5b8-4772-809f-416d03d16654" set key=F47MM-N3XJP-TQXJ9-BP99D-8K837 & set officever=Publisher
rem Skype for Business
if "%number%" == "SKU ID: 83e04ee1-fa8d-436d-8994-d31a862cab77" set key=869NQ-FJ69K-466HW-QYCP2-DDBV6 & set officever=Skype for Business
rem Word
if "%number%" == "SKU ID: bb11badf-d8aa-470e-9311-20eaf80fe5cc" set key=WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6 & set officever=Word
goto mainmenu
:office2013
rem Mondo
if "%number%" == "SKU ID: dc981c6b-fc8e-420f-aa43-f8f33e5c0923" set key=42QTK-RN8M7-J3C4G-BBGYM-88CYV & set officever=Mondo
rem Professional Plus
if "%number%" == "SKU ID: b322da9c-a2e2-4058-9e4e-f59a6970bd69" set key=YC7DK-G2NP3-2QQC3-J6H88-GVGXT & set officever=Professional Plus
rem Standard
if "%number%" == "SKU ID: b13afb38-cd79-4ae5-9f7f-eed058d750ca" set key=KBKQT-2NMXY-JJWGP-M62JB-92CD4 & set officever=Standard
rem Project Professional
if "%number%" == "SKU ID: 4a5d124a-e620-44ba-b6ff-658961b33b9a" set key=FN8TT-7WMH6-2D4X9-M337T-2342K & set officever=Project Professional
rem Project Standard
if "%number%" == "SKU ID: 427a28d1-d17c-4abf-b717-32c780ba6f07" set key=6NTH3-CW976-3G3Y2-JK3TX-8QHTT & set officever=Project Standard
rem Visio Professional
if "%number%" == "SKU ID: e13ac10e-75d0-4aff-a0cd-764982cf541c" set key=C2FG9-N6J68-H8BTJ-BW3QX-RM3B3 & set officever=Visio Professional
rem Visio Standard
if "%number%" == "SKU ID: ac4efaf0-f81f-4f61-bdf7-ea32b02ab117" set key=J484Y-4NKBF-W2HMG-DBMJC-PGWR7 & set officever=Visio Standard
rem Access
if "%number%" == "SKU ID: 6ee7622c-18d8-4005-9fb7-92db644a279b" set key=NG2JY-H4JBT-HQXYP-78QH9-4JM2D & set officever=Access
rem Excel
if "%number%" == "SKU ID: f7461d52-7c2b-43b2-8744-ea958e0bd09a" set key=VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB & set officever=Excel
rem Groove
if "%number%" == "SKU ID: fb4875ec-0c6b-450f-b82b-ab57d8d1677f" set key=H7R7V-WPNXQ-WCYYC-76BGV-VT7GH & set officever=Groove
rem InfoPath
if "%number%" == "SKU ID: a30b8040-d68a-423f-b0b5-9ce292ea5a8f" set key=DKT8B-N7VXH-D963P-Q4PHY-F8894 & set officever=InfoPath
rem Lync
if "%number%" == "SKU ID: 1b9f11e3-c85c-4e1b-bb29-879ad2c909e3" set key=2MG3G-3BNTT-3MFW9-KDQW3-TCK7R & set officever=Lync
rem OneNote
if "%number%" == "SKU ID: efe1f3e6-aea2-4144-a208-32aa872b6545" set key=TGN6P-8MMBC-37P2F-XHXXK-P34VW & set officever=OneNote
rem Outlook
if "%number%" == "SKU ID: 771c3afa-50c5-443f-b151-ff2546d863a0" set key=QPN8Q-BJBTJ-334K3-93TGY-2PMBT & set officever=Outlook
rem Powerpoint
if "%number%" == "SKU ID: 8c762649-97d1-4953-ad27-b7e2c25b972e" set key=4NT99-8RJFH-Q2VDH-KYG2C-4RD4F & set officever=Powerpoint
rem Publisher
if "%number%" == "SKU ID: 00c79ff1-6850-443d-bf61-71cde0de305f" set key=PN2WF-29XG2-T9HJ7-JQPJR-FCXK4 & set officever=Publisher
rem Word
if "%number%" == "SKU ID: d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3" set key=6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7 & set officever=Word
goto mainmenu
:office2010
rem Mondo
if "%number%" == "SKU ID: 09ed9640-f020-400a-acd8-d7d867dfd9c2" set key=YBJTT-JG6MD-V9Q7P-DBKXJ-38W9R & set officever=Mondo
rem Mondo 2
if "%number%" == "SKU ID: ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c" set key=7TC2V-WXF6P-TD7RT-BQRXR-B8K32 & set officever=Mondo 2
rem Professional Plus
if "%number%" == "SKU ID: 6f327760-8c5c-417c-9b61-836a98287e0c" set key=VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB & set officever=Professional Plus
rem Standard
if "%number%" == "SKU ID: 9da2a678-fb6b-4e67-ab84-60dd6a9c819a" set key=V7QKV-4XVVR-XYV4D-F7DFM-8R6BM & set officever=Standard
rem Project Professional
if "%number%" == "SKU ID: df133ff7-bf14-4f95-afe3-7b48e7e331ef" set key=YGX6F-PGV49-PGW3J-9BTGG-VHKC6 & set officever=Project Professional
rem Project Standard
if "%number%" == "SKU ID: 5dc7bf61-5ec9-4996-9ccb-df806a2d0efe" set key=4HP3K-88W3F-W2K3D-6677X-F9PGB & set officever=Project Standard
rem Visio Premium
if "%number%" == "SKU ID: 92236105-bb67-494f-94c7-7f7a607929bd" set key=D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ & set officever=Visio Premium
rem Visio Professional
if "%number%" == "SKU ID: e558389c-83c3-4b29-adfe-5e4d7f46c358" set key=7MCW8-VRQVK-G677T-PDJCM-Q8TCP & set officever=Visio Professional
rem Visio Standard
if "%number%" == "SKU ID: 9ed833ff-4f92-4f36-b370-8683a4f13275" set key=767HD-QGMWX-8QTDB-9G3R2-KHFGJ & set officever=Visio Standard
rem Access
if "%number%" == "SKU ID: 8ce7e872-188c-4b98-9d90-f8f90b7aad02" set key=V7Y44-9T38C-R2VJK-666HK-T7DDX & set officever=Access
rem Excel
if "%number%" == "SKU ID: cee5d470-6e3b-4fcc-8c2b-d17428568a9f" set key=H62QG-HXVKF-PP4HP-66KMR-CW9BM & set officever=Excel
rem Groove SharePoint Workspace
if "%number%" == "SKU ID: 8947d0b8-c33b-43e1-8c56-9b674c052832" set key=QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4 & set officever=Groove SharePoint Workspace
rem InfoPath
if "%number%" == "SKU ID: ca6b6639-4ad6-40ae-a575-14dee07f6430" set key=K96W8-67RPQ-62T9Y-J8FQJ-BT37T & set officever=InfoPath
rem OneNote
if "%number%" == "SKU ID: ab586f5c-5256-4632-962f-fefd8b49e6f4" set key=Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX & set officever=OneNote
rem Outlook
if "%number%" == "SKU ID: ecb7c192-73ab-4ded-acf4-2399b095d0cc" set key=7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ & set officever=Outlook
rem Powerpoint
if "%number%" == "SKU ID: 45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a" set key=RC8FX-88JRY-3PF7C-X8P67-P4VTT & set officever=Powerpoint
rem Publisher
if "%number%" == "SKU ID: b50c4f75-599b-43e8-8dcd-1081a7967241" set key=BFK7F-9MYHM-V68C7-DRQ66-83YTP & set officever=Publisher
rem Word
if "%number%" == "SKU ID: 2d0882e7-a4e7-423b-8ccc-70d91e0158b1" set key=HVHB3-C6FV7-KQX9W-YQG79-CRY7T & set officever=Word
rem Home and Business
if "%number%" == "SKU ID: ea509e87-07a1-4a45-9edc-eba5a39f36af" set key=D6QFG-VBYP2-XQHM7-J97RH-VVRCK & set officever=Home and Business
goto mainmenu

:mainmenu
cls
title Windows KMS Activator Menu By AR_Alex
if "%OFFICECHECK%"=="No Office installation detected" set officever=
echo. Windows KMS Activator Menu By AR_Alex
echo.
echo. VERSION: 3.3
echo. 
echo. Your Windows edition is compatible
echo. (%SYSTEMID% %WINVER%)
echo.
if "%online%"=="yes" echo. Online Mode
if "%online%"=="no" echo. Offline Mode
if "%update%" == "3.3" echo. You are using the latest version
echo.
echo. Choose Your Option...
echo.
echo. (1) Activate %SYSTEMID% %WINVER% with KMS
echo. (2) Permanently Activate with a Digital License
echo. (3) %OFFICECHECK% %officever%
echo. (4) Create a Scheduled Task for the Activator
echo. (5) Convert Retail %office% to Volume
echo. (6) Fix Windows Activation files
echo. (7) Check Activation for Windows
echo. (8) Check Activation for Office
echo. (9) Activate %SYSTEMID% With Your Own key Using KMS
echo. (u) Uninstall KMS
echo. (r) Re-Download the Program from Server
echo. (0) Close the Program

set /p userinp=    ^   Make your selection: 
set userinp=%userinp:~0,1%
if /i "%userinp%"=="1" goto wversion
if /i "%userinp%"=="2" goto digital
if /i "%userinp%"=="3" goto officeinfo
if /i "%userinp%"=="4" goto chooseoption
if /i "%userinp%"=="5" goto retailtovolume1
if /i "%userinp%"=="6" goto fix
if /i "%userinp%"=="7" goto chwindows
if /i "%userinp%"=="8" goto choffice
if /i "%userinp%"=="9" goto w8custommenu
if /i "%userinp%"=="u" goto uninstallkmsm
if /i "%userinp%"=="0" exit
if /i "%userinp%"=="r" goto dupdate
GOTO mainmenu

:digital
cls
title Digital License
if "%systemid%" == "Windows 7" goto digitaluns
if "%systemid%" == "Windows 8" goto digitaluns
if "%systemid%" == "Windows 8.1" goto digitaluns
if "%systemid%" == "" goto digitaluns
if "%winver%" == "Cloud" set key=V3WVW-N2PV2-CGWC3-34QGF-VMJ2C & set sku=178
if "%winver%" == "Cloud N" set key=NH9J3-68WK7-6FB93-4K3DF-DJ4F6 & set sku=179
if "%winver%" == "Core" set key=YTMG3-N6DKC-DKB77-7M9GH-8HVX7 & set sku=101
if "%winver%" == "Core Country Specific" set key=N2434-X9D7W-8PF6X-8DV9T-8TYMD & set sku=99
if "%winver%" == "Core N" set key=4CPRK-NM3K3-X6XXQ-RXX86-WXCHW & set sku=98
if "%winver%" == "Core Single Language" set key=BT79Q-G7N6G-PGBYW-4YWX6-6F4BT & set sku=100
if "%winver%" == "Education" set key=YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY & set sku=121
if "%winver%" == "Education N" set key=84NGF-MHBT6-FXBX8-QWJK7-DRR8H & set sku=122
if "%winver%" == "Enterprise" set key=XGVPP-NMH47-7TTHJ-W3FW7-8HV2C & set sku=4
if "%winver%" == "Enterprise N" set key=3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT & set sku=27
if "%winver%" == "Enterprise S" set key=NK96Y-D9CD8-W44CQ-R8YTK-DYJWX & set sku=125
if "%winver%" == "Enterprise S N" set key=2DBW3-N2PJG-MVHW3-G7TDK-9HKR4 & set sku=126
if "%winver%" == "Professional" set key=VK7JG-NPHTM-C97JM-9MPGT-3V66T & set sku=48
if "%winver%" == "Professional Education" set key=8PTT6-RNW4C-6V7J2-C2D3X-MHBPB & set sku=164
if "%winver%" == "Professional Education N" set key=GJTYN-HDMQY-FRR76-HVGC7-QPF8P & set sku=165
if "%winver%" == "Professional N" set key=2B87N-8KFHP-DKV6R-Y2C8J-PKCK & set sku=49
if "%winver%" == "Professional Workstation" set key=DXG7C-N36C4-C4HTG-X4T3X-2YV77 & set sku=161
if "%winver%" == "Professional Workstation N" set key=WYPNQ-8C467-V2W6J-TX4WX-WT2RQ & set sku=162
if "%sku%" == "" goto digitaluns
if "%key%" == "" goto digitaluns
echo. This option will permanently and legitimately activate
echo. %systemid% %winver% using a digital license directly from microsoft.
echo.
echo. (1) Continue
echo. (2) Go back to the Main Menu
set /p userinp=    ^   Make your selection: 
set userinp=%userinp:~0,1%
if /i "%userinp%"=="1" goto digitalcont
if /i "%userinp%"=="2" goto mainmenu
goto mainmenu

:digitaluns
cls
echo.
echo. %systemid% %winver% is Unsupported at this time
echo. If you believe this is a mistake contact me asap.
echo. 
pause
goto mainmenu

:digitalcont
cls
title Digital License
echo. Installing product key for %systemid% %winver%...
echo.
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ipk %key%
cd /d "%~dp0"
set "gatherosstate=Digital\gatherosstate.exe"
echo. Setting up registry...
reg add "HKLM\SYSTEM\Tokens" /v "Channel" /t REG_SZ /d "Retail" /f
reg add "HKLM\SYSTEM\Tokens\Kernel" /v "Kernel-ProductInfo" /t REG_DWORD /d %sku% /f
reg add "HKLM\SYSTEM\Tokens\Kernel" /v "Security-SPP-GenuineLocalStatus" /t REG_DWORD /d 1 /f
echo.
echo Generating GenuineTicket.xml...
start /wait "" "%gatherosstate%"
timeout /t 3 >nul 2>&1
echo.
echo. Migrating Windows Genuine Authorization blob...
echo.
clipup -v -o -altto Digital\
echo. Activating...
echo.
cscript /nologo %windir%\system32\slmgr.vbs -ato
cscript /nologo %windir%\system32\slmgr.vbs -dli
cscript /nologo %windir%\system32\slmgr.vbs -xpr
echo. Deleting registry entries...
reg delete "HKLM\SYSTEM\Tokens" /f
pause
goto done


:uninstallkmsm
cls
title Uninstall KMS by AR_Alex
echo. Uninstall any traces of KMS. This includes Windows
echo. Office KMS activations, scheduled activations, and
echo. any files related to this program copied into your
echo. computer.
echo.
echo. (1) Continue
echo. (2) Go back to the Main Menu
set /p userinp=    ^   Make your selection: 
set userinp=%userinp:~0,1%
if /i "%userinp%"=="1" goto uninstallkms
if /i "%userinp%"=="2" goto mainmenu
goto mainmenu

:uninstallkms
cls
title Uninstalling KMS.... AR_Alex
echo. Starting...
echo. Uninstalling Windows KMS activations
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /upk 
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ckms
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /cpky
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /rilc
sc stop sppsvc
sc start sppsvc
echo. Uninstalling Office KMS activations
if "%PF%" == "alternative" set OSPP=cscript.exe /nologo "%installdir%\Program Files (x86)\Microsoft Office\Office15\ospp.vbs"
if "%PF%" == "normal" set OSPP=cscript.exe /nologo "%installdir%\Program Files\Microsoft Office\Office15\ospp.vbs"
if "%PF%" == "alternative6" set OSPP=cscript.exe /nologo "%installdir%\Program Files (x86)\Microsoft Office\Office16\ospp.vbs"
if "%PF%" == "normal6" set OSPP=cscript.exe /nologo "%installdir%\Program Files\Microsoft Office\Office16\ospp.vbs"
%OSPP% /unpkey:GVGXT >nul 2>&1
%OSPP% /unpkey:WFG99 >nul 2>&1
%OSPP% /unpkey:92CD4 >nul 2>&1
%OSPP% /unpkey:2342K >nul 2>&1
%OSPP% /unpkey:8QHTT >nul 2>&1
%OSPP% /unpkey:RM3B3 >nul 2>&1
%OSPP% /unpkey:PGWR7 >nul 2>&1
echo. Uninstalling Scheduled tasks created by this software
schtasks /delete /f /tn "KMS Service"
echo. Deleting files
rd %WINDIR%\KMS_Service\ /s /q
echo. Done!
pause
goto done
exit

:downloadupdate
cls
title. Install new update - By AR_Alex
echo. There is a new update available
echo. Version: %update%
echo. 
echo. (1) Download the update now
echo. (2) Continue without updating
echo.

set /p userinp=    ^   Make your selection: 
set userinp=%userinp:~0,1%
if /i "%userinp%"=="1" goto dupdate
if /i "%userinp%"=="2" goto DetectWindows
goto downloadupdate

:dupdate
cls
title. Downloading update from server...
echo. Downloading update from server...
bitsadmin /Transfer myDownloadJob https://github.com/ARAlex143/activator/archive/master.zip %USERPROFILE%\Downloads\KMS_Activator_By_AR_Alex.zip  >nul 2>&1
if not exist %USERPROFILE%\Downloads\KMS_Activator_By_AR_Alex.zip goto dupdateretry
echo. 
echo. Success! The downloaded file is located in: %USERPROFILE%\Downloads\KMS_Activator_By_AR_Alex.zip
echo. Unzip the downloaded file then, enjoy.
pause.
exit

:dupdateretry
cls
title. Downloading update from server...
echo. Downloading update from server using certutil...
certutil.exe -urlcache -split -f https://github.com/ARAlex143/activator/archive/master.zip %USERPROFILE%\Downloads\KMS_Activator_By_AR_Alex.zip
if not exist %USERPROFILE%\Downloads\KMS_Activator_By_AR_Alex.zip echo. Error the update could not be downloaded please contact AR_Alex & pause & goto mainmenu
echo.
echo. Success! The downloaded file is located in: %USERPROFILE%\Downloads\KMS_Activator_By_AR_Alex.zip
echo. Unzip the downloaded file then, enjoy.
pause.
exit

:w8custommenu
cls
title By AR_Alex - Install your own key
echo. Custom Key Installation
echo.
echo. Install your own key:
echo.
echo. Use this format XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
echo.
set /p KEY= ^> Enter the full product key to install:

set custom=yes
goto wkmsserverold

:choffice
cls
title Office Activation Checker by AR_Alex
echo. We will now check the status of your activation.
echo.
%OSPP% /dstatus 
echo.
pause.
goto mainmenu

:chwindows
cls
title Windows Activation Checker by AR_Alex
echo. We will now check the status of your activation.
echo.
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /dli
echo.
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /xpr
echo.
pause.
goto mainmenu

:chooseoption
cls
echo. Choose the type of task you would like to create
echo.
echo. (1) Windows only
echo. (2) Office only
echo. (3) Windows and Office at the same time
echo.
echo.

set /p userinp=    ^   Make your selection: 
set userinp=%userinp:~0,1%
if /i "%userinp%"=="1" set taskactivator=Windows_Task

if /i "%userinp%"=="2" (
set taskactivator=Office_Task
)

if /i "%userinp%"=="3" (
set taskactivator=Windows_Office_Task
)

:createtaskmenu
cls
title Task Creator by AR_Alex
echo. Scheduled Task Creator Menu
echo.
echo. Select an Option...
echo.
echo. (1) Activate every 24 hours
echo. (2) Activate every 1 week
echo. (3) Activate every 1 month
echo. (4) Go back to the Main Menu

set /p userinp=    ^   Make your selection: 
set userinp=%userinp:~0,1%
if /i "%userinp%"=="1" goto 24hours
if /i "%userinp%"=="2" goto 1week
if /i "%userinp%"=="3" goto 1month
if /i "%userinp%"=="4" goto mainmenu
GOTO createtaskmenu

:comingsoon
cls
echo. Sorry these options are not yet available
echo.
pause.
goto chooseoption

:24hours
cls
if /i not exist Kms_Files\ goto error1
robocopy Kms_Files\ %WINDIR%\KMS_Service\Kms_Files\ /e
schtasks /delete /f /tn "KMS Service" >nul 2>&1
schtasks /create /XML %WINDIR%\KMS_Service\Kms_Files\Task_Files\1day_%taskactivator%.xml /tn "KMS Service"
echo.
if %errorlevel% neq 0 (echo Error while creating kms service!&echo.&echo Press any key to exit...&pause>nul&exit)
pause.
goto done

:1week
cls
if /i not exist Kms_Files\ goto error1
robocopy Kms_Files\ %WINDIR%\KMS_Service\Kms_Files\ /e
schtasks /delete /f /tn "KMS Service" >nul 2>&1
schtasks /create /XML %WINDIR%\KMS_Service\Kms_Files\Task_Files\1week_%taskactivator%.xml /tn "KMS Service"
echo.
if %errorlevel% neq 0 (echo Error while creating kms service!&echo.&echo Press any key to exit...&pause>nul&exit)
pause.
goto done

:1month
cls
if /i not exist Kms_Files\ goto error1
robocopy Kms_Files\ %WINDIR%\KMS_Service\Kms_Files\ /e
schtasks /delete /f /tn "KMS Service" >nul 2>&1
schtasks /create /XML %WINDIR%\KMS_Service\Kms_Files\Task_Files\1month_%taskactivator%.xml /tn "KMS Service"
echo.
if %errorlevel% neq 0 (echo Error while creating kms service!&echo.&echo Press any key to exit...&pause>nul&exit)
pause.
goto done

:fix
cls
title Windows Repair Menu
echo. Windows Repair Menu
echo.
echo. (1) Scan and fix any broken windows files
echo. (2) Go back to the Main Menu
echo.
echo.
echo.

:CHOOSEACTION
set /p userinp=    ^   Make your selection: 
set userinp=%userinp:~0,1%
if /i "%userinp%"=="1" goto w8repair
if /i "%userinp%"=="2" goto mainmenu
goto fix

:w8repair
cls
sfc /scannow
echo.
pause.
GOTO done

:retailtovolume1
if "%office%" == "Office 2016" (
goto retailtovolume16
)
pause
exit

:retailtovolume16
cls
title Convert Retail Office 2016 to Volume AR_Alex
echo. Convert Retail Office 2016 to Volume AR_Alex
echo.
echo. Choose your edition currently installed...
echo.
echo. (1) %office% Professional Plus
echo. (2) %office% Standard
echo. (3) %office% Project Standard
echo. (4) %office% Project Pro
echo. (5) %office% Visio Standard
echo. (6) %office% Visio Pro
echo. (7) Go back to the Main Menu

set /p userinp=    ^   Make your selection: 
set userinp=%userinp:~0,1%
if /i "%userinp%"=="1" set convertpth=Proplus16& goto convertproplus6
if /i "%userinp%"=="2" set convertpth=Standard16& goto convertproplus6
if /i "%userinp%"=="3" set convertpth=Projectstd16& goto convertproplus6
if /i "%userinp%"=="4" set convertpth=Projectpro16& goto convertproplus6
if /i "%userinp%"=="5" set convertpth=Visiostd16& goto convertproplus6
if /i "%userinp%"=="6" set convertpth=Visiopro16& goto convertproplus6
if /i "%userinp%"=="7" goto mainmenu

:convertproplus6
cls
title Office 2K16 VL AR_Alex 
if /i not exist Office_Files\ goto error1
echo. Installing Volume License.
pushd "%~dp0"
for /f %%x in ('dir /b Office_Files\%convertpth%\*.xrm-ms') do %ospp% /inslic:Office_Files\%convertpth%\%%x
echo.
echo. Success you should now have %office% %convertpth% Volume Edition
pause.
goto mainmenu

:retailtovolume13
cls
title Convert Retail Office 2013 to Volume By AR_Alex
echo. Convert Retail Office 2013 to Volume By AR_Alex
echo. 
echo. Choose your edition currently installed...
echo.
echo. (1) Office 2013 Professional Plus
echo. (2) Project 2013 Professional
echo. (3) Visio 2013 Professional
echo. (4) Go back to the Main Menu

set /p userinp=    ^   Make your selection: 
set userinp=%userinp:~0,1%
if /i "%userinp%"=="1" set convertpth=Proplus13& goto convertproplus15
if /i "%userinp%"=="2" set convertpth=Project13& goto convertproplus15
if /i "%userinp%"=="3" set convertpth=Visio13& goto convertproplus15
if /i "%userinp%"=="4" goto mainmenu
GOTO retailtovolume13

:convertproplus15
if /i not exist Office_Files\ goto error1
for /f %%x in ('dir /b Office_Files\%convertpth%\*.xrm-ms') do %ospp% /inslic:Office_Files\%convertpth%\%%x
regedit /s Office_Files\%convertpth%\%convertpth%.reg
if /i not %errorlevel%==0 goto error2
echo.
echo. Success you should now have %office% %convertpth% Volume Edition
pause.
goto done

:error1
echo. 
echo. ERROR 0x000MF21
echo.
echo. You are missing files, please redownload...
echo. Your antivirus might be removing them.
echo. Or extract the whole zip into the desktop.
echo. No changes have been made...
pause. 
exit

:error2
echo. 
echo. ERROR 0x000PGMF35
echo.
echo. Please contact AR_Alex of this error...
pause. 
exit

:error3
echo. 
echo. ERROR 0x000IP56
echo.
echo. Please contact AR_Alex of this error...
pause. 
exit

:error4
echo. 
echo. ERROR 0x000WVW78
echo.
echo. We're sorry but your version of windows is not supported...
pause. 
exit

:error5
echo. 
echo. ERROR 0x000SPP97
echo.
echo. Error while working with SppExtComObj.Exe
echo. Contact AR_Alex to solve this problem...
pause. 
exit

:error6
echo. 
echo. ERROR 0x000WVE46
echo.
echo. Error while detecting the version of windows you are using
echo. Either we have still not added support for it or it is not compatible at all...
pause. 
exit

:error7
cls
echo. 
echo. ERROR 0x000RGS41
echo.
echo. An error occurred while activating
echo. Please install the registry settings...
echo. The registry settings installer is located
echo. in the main folder which is called "Install Me First.reg"
echo.
echo. When you have finished installing it close this program
echo. Then re-open it again.
echo.
echo. If this problem persists please contact AR_Alex
echo. Contact info of AR_Alex is in the Read Me.txt
pause.
exit

:officeinfo
cls
set ac=
if "%systemid%" == "Windows 7" set ac="1" & goto wkmsserverold
if "%systemid%" == "Windows 8" set ac="1" & goto wkmsserverold
goto cinjectoro

:cinjectoro
cls
title Windows KMS Emulator By AR_Alex

setlocal EnableExtensions
setlocal EnableDelayedExpansion
set port=1688
for /f "tokens=2 delims==" %%a in ('wmic os get osarchitecture /value') do (
	set _OSarch=%%a
)
call :StopService
Call :DelReg
regedit /s settings.reg
xcopy Kms_Files\"%_OSarch%\SppExtComObjPatcher.exe" "%SystemRoot%\system32\" /y
xcopy Kms_Files\"%_OSarch%\SppExtComObjHook.dll" "%SystemRoot%\system32\" /y
Call :DelReg
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
%WINDIR%\System32\Netsh Advfirewall Firewall add rule name="0pen Port KMS" dir=in action=allow protocol=TCP localport=1688 >nul 2>&1
echo. Setting up KMS Address...
echo.
echo. Installing KMS Client key...
%OSPP% /inpkey:%key%
%OSPP% /actype:2
%OSPP% /remhst
%OSPP% /setprt:1688
%OSPP% /sethst:10.0.0.20
echo. Activating
Call :DelReg
%OSPP% /act
call :StopService "sppsvc"
del /f /q "%SystemRoot%\system32\SppExtComObjPatcher.exe" 
del /f /q "%SystemRoot%\system32\SppExtComObjHook.dll" 
call :RemoveIFEOEntry "SppExtComObj.exe"
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=1688 >nul 2>&1
sc start sppsvc trigger=timer;sessionid=0 >nul 2>&1
pause
goto done

:kmsactivateoffice
cls
title Windows KMS Emulator By AR_Alex
echo. Preparing Files...
echo.
sc stop sppsvc >nul 2>&1
if /i not exist Kms_Files\ goto error1
sc stop sppsvc >nul 2>&1
set port=1688
echo. Detecting processor type...
echo.
reg.exe query "hklm\software\microsoft\Windows NT\currentversion" /v buildlabex | find /i "amd64" >nul 2>&1
if %errorlevel% equ 0 set xOS=x64
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set xOS=x86
echo. Preparing the Firewall...
echo.
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
%WINDIR%\System32\Netsh Advfirewall Firewall add rule name="0pen Port KMS" dir=in action=allow protocol=TCP localport=%port% >nul 2>&1
Call :DelReg
echo. Setting up KMS Address...
echo.
%OSPP% /actype:2
%OSPP% /remhst
%OSPP% /setprt:%port%
%OSPP% /sethst:8.7.8.7
echo. Installing KMS Client key...
echo.
%OSPP% /inpkey:%key%
echo. Starting the KMS Emulator...
echo.
cd Kms_Files
SECOInjector_%xOS%.exe /s 8.7.8.7 /f /l >nul 2>&1
title Windows 8.1 KMS Emulator By AR_Alex
Call :DelReg
echo. Activating...
echo.
%OSPP% /act
echo.
echo. Verifying that office activated successfully...
echo.
%OSPP% /act | findstr /i /c:"Product activation successful"> NUL 2>&1
if [%errorlevel%]==[0] (echo Office Activated Successfully! ) else (goto error7)
echo.
sc stop sppsvc >nul 2>&1
echo. Closing the KMS Emulator and cleaning up...
echo.
SECOInjector_%xOS%.exe /f /u >nul 2>&1
title Windows 8.1 KMS Emulator By AR_Alex
echo. Processing the Firewall...
echo.
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
sc start sppsvc >nul 2>&1
echo. DONE!
echo.
pause.
goto done

:DelReg
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55C92734-D682-4D71-983E-D6EC3F16059F" /f >nul 2>&1
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0FF1CE15-A989-479D-AF46-F275C6370663" /f >nul 2>&1
reg delete "HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55C92734-D682-4D71-983E-D6EC3F16059F" /f >nul 2>&1
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\59A52881-A989-479D-AF46-F275C6370663" /f >nul 2>&1
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\0FF1CE15-A989-479D-AF46-F275C6370663" /f >nul 2>&1
exit /b

:wversion
cls
if "%SYSTEMID%" == "Windows Server 2016" (
if "%WINVER%" == "Server Standard" set KEY=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
if "%WINVER%" == "Datacenter" set KEY=CB7KF-BWN84-R7R2Y-793K2-8XDDG
if "%WINVER%" == "Essentials" set KEY=JCKRF-N37P4-C2D82-9YXRT-4M63B
if "%WINVER%" == "Cloud Storage" set KEY=QN4C6-GBJD2-FB422-GHWJK-GJG2R
if "%WINVER%" == "Azure Core" set KEY=VP34G-4NPPG-79JTQ-864T4-R3MQX
goto cinjector
)

if "%SYSTEMID%" == "Windows Server 2012 R2" (
if "%WINVER%" == "Server Standard" set KEY=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
if "%WINVER%" == "Datacenter" set KEY=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
if "%WINVER%" == "Essentials" set KEY=KNC87-3J2TX-XB4WP-VCPJV-M4FWM
if "%WINVER%" == "Cloud Storage" set KEY=3NPTF-33KPT-GGBPR-YX76B-39KDD
set ac=" " & goto wkmsserverold
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
set ac=" " & goto wkmsserverold
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
goto cinjector
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
set ac=" " & goto wkmsserverold
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
goto cinjector
)

goto cinjector

:cinjector
cls
title Windows KMS Emulator By AR_Alex
setlocal EnableExtensions
setlocal EnableDelayedExpansion
for /f "tokens=2 delims==" %%a in ('wmic os get osarchitecture /value') do (
	set _OSarch=%%a
)
call :StopService
Call :DelReg
regedit /s settings.reg
xcopy Kms_Files\"%_OSarch%\SppExtComObjPatcher.exe" "%SystemRoot%\system32\" /y
xcopy Kms_Files\"%_OSarch%\SppExtComObjHook.dll" "%SystemRoot%\system32\" /y
Call :DelReg
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
%WINDIR%\System32\Netsh Advfirewall Firewall add rule name="0pen Port KMS" dir=in action=allow protocol=TCP localport=%port% >nul 2>&1
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /upk 
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ckms 
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /skms 10.0.0.20 
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ipk %KEY%
echo. Activating %SYSTEMID% %WINVER%
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ato
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /xpr
cscript "c:\windows\system32\slmgr.vbs" //nologo -dli
pause
call :StopService "sppsvc"
del /f /q "%SystemRoot%\system32\SppExtComObjPatcher.exe" 
del /f /q "%SystemRoot%\system32\SppExtComObjHook.dll" 
call :RemoveIFEOEntry "SppExtComObj.exe"
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
sc start sppsvc trigger=timer;sessionid=0 >nul 2>&1
goto done

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

:wkmsinjector
cls
sc stop sppsvc >nul 2>&1
if /i not exist Kms_Files\ goto error1

sc stop sppsvc >nul 2>&1
set port=1688
echo. Detecting processor type...
echo.
reg.exe query "hklm\software\microsoft\Windows NT\currentversion" /v buildlabex | find /i "amd64" >nul 2>&1
if %errorlevel% equ 0 set xOS=x64
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set xOS=x86

echo. Preparing the Firewall...
echo.
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
%WINDIR%\System32\Netsh Advfirewall Firewall add rule name="0pen Port KMS" dir=in action=allow protocol=TCP localport=%port% >nul 2>&1

Call :DelReg

echo. Setting up KMS Address...
echo.
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /upk >nul 2>&1
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ckms >nul 2>&1
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /skms 8.7.8.7 >nul 2>&1
echo. Installing KMS Client key...
echo.
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ipk %KEY%

sc stop sppsvc >nul 2>&1

echo. Starting the KMS Emulator...
echo.
cd Kms_Files
SECOInjector_%xOS%.exe /s 8.7.8.7 /f /l >nul 2>&1
title Windows 8.1 KMS Emulator By AR_Alex

Call :DelReg

echo. Activating...
echo.
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /ato
echo. Generating Expiration Date
echo.
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /xpr

sc stop sppsvc >nul 2>&1

echo. Closing the KMS Emulator and cleaning up...
echo.
SECOInjector_%xOS%.exe /f /u >nul 2>&1
title Windows 8.1 KMS Emulator By AR_Alex

echo. Processing the Firewall...
echo.
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
sc start sppsvc >nul 2>&1
echo. DONE!
echo.
pause.
goto done

:wkmsserverold
cls
title Windows KMS Emulator By AR_Alex
echo. Please wait while we work...
if /i not exist Kms_Files\ goto error1
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


pushd "%~dp0"
set xOS=x64
if "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set xOS=x86
Kms_Files\%xOS%\devcon.exe remove tap0901 >nul 2>&1
echo. Installing network tap adapter...
Kms_Files\%xOS%\devcon.exe install Kms_Files\%xOS%\OemWin2k.inf tap0901 >nul 2>&1
if %errorlevel% neq 0 (echo Error installing TAP adapter!&echo.&echo Press any key to exit...&pause>nul&exit)
echo. Setting Adapter IP...
for /f "tokens=1 delims=. " %%i in ('route print ^| find /i "TAP-Windows Adapter V9"') do (netsh interface ip set address %%i static 10.3.0.1 255.255.255.0) >nul 2>&1
if %errorlevel% neq 0 (echo Error! Could not setup ip...&echo.&echo Press any key to exit...&pause>nul&exit)
echo. Preparing the Firewall...
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
%WINDIR%\System32\Netsh Advfirewall Firewall add rule name="0pen Port KMS" dir=in action=allow protocol=TCP localport=%port% >nul 2>&1
if %errorlevel% neq 0 (echo Error while shutting down the firewall!&echo.&echo Press any key to exit...&pause>nul&exit)
echo.&echo Starting TunMirror...
IF EXIST "%systemroot%\Microsoft.NET\Framework\v4.0.30319" goto continuetun
echo. Downloading .NET Framework 4...
certutil.exe -urlcache -split -f https://download.microsoft.com/download/5/6/2/562A10F9-C9F4-4313-A044-9C94E0A8FAC8/dotNetFx40_Client_x86_x64.exe %temp%\dotNetFx40_Client_x86_x64.exe
echo. Installing .NET Framework 4
start /wait %temp%\dotNetFx40_Client_x86_x64.exe /norestart /passive
del %temp%\dotNetFx40_Client_x86_x64.exe >nul 2>&1
:continuetun
start /b cmd /c Kms_Files\TunMirror.exe >nul 2>&1
for /f "tokens=5 delims=, " %%i in ('netstat -ano ^|find ":%port% "') do taskkill /pid %%i /f >nul 2>&1
timeout /t 5 /nobreak >nul 2>&1
echo. Starting KMS Server Emulator...
start /b cmd /c Kms_Files\KMSServer.exe %port% %epid% %ActivationInterval% %RenewalInterval% >nul 2>&1
if %errorlevel% neq 0 (echo Error! KMS server not running...&exit /b)
timeout /t 5 /nobreak >nul 2>&1
if %ac%=="1" goto officeact
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
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /xpr
sc stop sppsvc >nul 2>&1
cscript "c:\windows\system32\slmgr.vbs" //nologo -dli
goto closeserver
:officeact
echo. Installing KMS Client key...
%OSPP% /inpkey:%KEY%
echo. Setting up KMS Address...
%OSPP% /remhst
%OSPP% /setprt:%port%
%OSPP% /sethst:%kmsip%
echo. Activating...
echo.
%OSPP% /act
echo.
:closeserver
echo. Closing Server...
taskkill /t /f /IM kmsserver.exe >nul 2>&1
taskkill /t /f /IM tunmirror.exe >nul 2>&1
echo. Processing the firewall...
%WINDIR%\System32\Netsh Advfirewall Firewall delete rule name="0pen Port KMS" protocol=TCP localport=%port% >nul 2>&1
pause
goto done

:done
cls
title Windows KMS Activator By AR_Alex
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo. 
echo.
echo. Done, if something went wrong talk to AR_Alex with screenshots and system info
echo. To make that watermark disappear if you did activate just open my program with admin and close it.
echo.
echo. Thank AR_Alex for taking his time to make this program for you :)
echo. A special thanks for the members in nsaneforums and a big thanks to MDL
echo. 
echo. Credits... Nsaneforums.com, forums.mydigitallife.info, and others...
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
echo.
pause.
goto mainmenu
exit /b
exit
