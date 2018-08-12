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
cls
if "%SYSTEMID%" == "Windows Server 2016" (
if "%WINVER%" == "Server Standard" set KEY=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
if "%WINVER%" == "Datacenter" set KEY=CB7KF-BWN84-R7R2Y-793K2-8XDDG
if "%WINVER%" == "Essentials" set KEY=JCKRF-N37P4-C2D82-9YXRT-4M63B
if "%WINVER%" == "Cloud Storage" set KEY=QN4C6-GBJD2-FB422-GHWJK-GJG2R
if "%WINVER%" == "Azure Core" set KEY=VP34G-4NPPG-79JTQ-864T4-R3MQX
goto next2
)

if "%SYSTEMID%" == "Windows Server 2012 R2" (
if "%WINVER%" == "Server Standard" set KEY=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
if "%WINVER%" == "Datacenter" set KEY=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
if "%WINVER%" == "Essentials" set KEY=KNC87-3J2TX-XB4WP-VCPJV-M4FWM
if "%WINVER%" == "Cloud Storage" set KEY=3NPTF-33KPT-GGBPR-YX76B-39KDD
goto next2
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
goto next2
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
goto next2
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
goto next2
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
goto next2
)
:next2
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
set OSPP=cscript.exe /nologo "%installdir%\Program Files\Microsoft Office\Office15\ospp.vbs"
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
echo Checking office installation
for /f "tokens=*" %%A in (
  '%OSPP% /dstatus ^| findstr /i sku'
) do set number=%%A
if "%office%" == "Office 2010" goto office2010
if "%office%" == "Office 2013" goto office2013
if "%office%" == "Office 2016" goto office2016
if "%office%" == "Office 2019" goto office2019
goto officeinfo
:office2019
rem Professional Plus C2R-P
if "%number%" == "SKU ID: 0bc88885-718c-491d-921f-6f214349e79c" set key2=VQ9DP-NVHPH-T9HJC-J9PDT-KTQRG & set officever=Professional Plus C2R-P
rem Project Professional C2R-P
if "%number%" == "SKU ID: fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9" set key2=XM2V9-DN9HH-QB449-XDGKC-W2RMW & set officever=Project Professional C2R-P
rem Visio Professional C2R-P
if "%number%" == "SKU ID: 500f6619-ef93-4b75-bcb4-82819998a3ca" set key2=N2CG9-YD3YK-936X4-3WR82-Q3X4H & set officever=Visio Professional C2R-P
rem Professional Plus
if "%number%" == "SKU ID: 85dd8b5f-eaa4-4af3-a628-cce9e77c9a03" set key2=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP & set officever=Professional Plus
rem Standard
if "%number%" == "SKU ID: 6912a74b-a5fb-401a-bfdb-2e3ab46f4b02" set key2=6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK & set officever=Standard
rem Project Professional
if "%number%" == "SKU ID: 2ca2bf3f-949e-446a-82c7-e25a15ec78c4" set key2=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B & set officever=Project Professional
rem Project Standard
if "%number%" == "SKU ID: 1777f0e3-7392-4198-97ea-8ae4de6f6381" set key2=C4F7P-NCP8C-6CQPT-MQHV9-JXD2M & set officever=Project Standard
rem Visio Professional
if "%number%" == "SKU ID: 5b5cf08f-b81a-431d-b080-3450d8620565" set key2=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB & set officever=Visio Professional
rem Visio Standard
if "%number%" == "SKU ID: e06d7df3-aad0-419d-8dfb-0ac37e2bdf39" set key2=7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2 & set officever=Visio Standard
rem Access
if "%number%" == "SKU ID: 9e9bceeb-e736-4f26-88de-763f87dcc485" set key2=9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT & set officever=Access
rem Excel
if "%number%" == "SKU ID: 237854e9-79fc-4497-a0c1-a70969691c6b" set key2=TMJWT-YYNMB-3BKTF-644FC-RVXBD & set officever=Excel
rem Outlook
if "%number%" == "SKU ID: c8f8a301-19f5-4132-96ce-2de9d4adbd33" set key2=7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK & set officever=Outlook
rem PowerPoint
if "%number%" == "SKU ID: 3131fd61-5e4f-4308-8d6d-62be1987c92c" set key2=RRNCX-C64HY-W2MM7-MCH9G-TJHMQ & set officever=PowerPoint
rem Publisher
if "%number%" == "SKU ID: 9d3e4cca-e172-46f1-a2f4-1d2107051444" set key2=G2KWX-3NW6P-PY93R-JXK2T-C9Y9V & set officever=Publisher
rem Skype for Business
if "%number%" == "SKU ID: 734c6c6e-b0ba-4298-a891-671772b2bd1b" set key2=NCJ33-JHBBY-HTK98-MYCV8-HMKHJ & set officever=Skype for Business
rem Word
if "%number%" == "SKU ID: 059834fe-a8ea-4bff-b67b-4d006b5447d3" set key2=PBX3G-NWMT6-Q7XBW-PYJGG-WXD33 & set officever=Word
goto officeinfo
:office2016
rem Project Professional C2R-P
if "%number%" == "SKU ID: 829b8110-0e6f-4349-bca4-42803577788d" set key2=WGT24-HCNMF-FQ7XH-6M8K7-DRTW9 & set officever= Project Professional C2R-P
rem Project Standard C2R-P
if "%number%" == "SKU ID: cbbaca45-556a-4416-ad03-bda598eaa7c8" set key2=D8NRQ-JTYM3-7J2DX-646CT-6836M & set officever=Project Standard C2R-P
rem Visio Professional C2R-P
if "%number%" == "SKU ID: b234abe3-0857-4f9c-b05a-4dc314f85557" set key2=69WXN-MBYV6-22PQG-3WGHK-RM6XC & set officever=Visio Professional C2R-P
rem Visio Standard C2R-P
if "%number%" == "SKU ID: 361fe620-64f4-41b5-ba77-84f8e079b1f7" set key2=NY48V-PPYYH-3F4PX-XJRKJ-W4423 & set officever=Visio Standard C2R-P
rem MondoR
if "%number%" == "SKU ID: e914ea6e-a5fa-4439-a394-a9bb3293ca09" set key2=DMTCJ-KNRKX-26982-JYCKT-P7KB6 & set officever=MondoR
rem Mondo
if "%number%" == "SKU ID: 9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce" set key2=HFTND-W9MK4-8B7MJ-B6C4G-XQBR2 & set officever=Mondo
rem Professional Plus
if "%number%" == "SKU ID: d450596f-894d-49e0-966a-fd39ed4c4c64" set key2=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 & set officever=Professional Plus
rem Standard
if "%number%" == "SKU ID: dedfa23d-6ed1-45a6-85dc-63cae0546de6" set key2=JNRGM-WHDWX-FJJG3-K47QV-DRTFM & set officever=Standard
rem Project Professional
if "%number%" == "SKU ID: 4f414197-0fc2-4c01-b68a-86cbb9ac254c" set key2=YG9NW-3K39V-2T3HJ-93F3Q-G83KT & set officever=Project Professional
rem Project Standard
if "%number%" == "SKU ID: da7ddabc-3fbe-4447-9e01-6ab7440b4cd4" set key2=GNFHQ-F6YQM-KQDGJ-327XX-KQBVC & set officever=Project Standard
rem Visio Professional
if "%number%" == "SKU ID: 6bf301c1-b94a-43e9-ba31-d494598c47fb" set key2=PD3PC-RHNGV-FXJ29-8JK7D-RJRJK & set officever=Visio Professional
rem Visio Standard
if "%number%" == "SKU ID: aa2a7821-1827-4c2c-8f1d-4513a34dda97" set key2=7WHWN-4T7MP-G96JF-G33KR-W8GF4 & set officever=Visio Standard
rem Access
if "%number%" == "SKU ID: 67c0fc0c-deba-401b-bf8b-9c8ad8395804" set key2=GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW & set officever=Access
rem Excel
if "%number%" == "SKU ID: c3e65d36-141f-4d2f-a303-a842ee756a29" set key2=9C2PK-NWTVB-JMPW8-BFT28-7FTBF & set officever=Excel
rem OneNote
if "%number%" == "SKU ID: d8cace59-33d2-4ac7-9b1b-9b72339c51c8" set key2=DR92N-9HTF2-97XKM-XW2WJ-XW3J6 & set officever=OneNote
rem Outlook
if "%number%" == "SKU ID: ec9d9265-9d1e-4ed0-838a-cdc20f2551a1" set key2=R69KK-NTPKF-7M3Q4-QYBHW-6MT9B & set officever=Outlook
rem Powerpoint
if "%number%" == "SKU ID: d70b1bba-b893-4544-96e2-b7a318091c33" set key2=J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6 & set officever=Powerpoint
rem Publisher
if "%number%" == "SKU ID: 041a06cb-c5b8-4772-809f-416d03d16654" set key2=F47MM-N3XJP-TQXJ9-BP99D-8K837 & set officever=Publisher
rem Skype for Business
if "%number%" == "SKU ID: 83e04ee1-fa8d-436d-8994-d31a862cab77" set key2=869NQ-FJ69K-466HW-QYCP2-DDBV6 & set officever=Skype for Business
rem Word
if "%number%" == "SKU ID: bb11badf-d8aa-470e-9311-20eaf80fe5cc" set key2=WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6 & set officever=Word
goto officeinfo
:office2013
rem Mondo
if "%number%" == "SKU ID: dc981c6b-fc8e-420f-aa43-f8f33e5c0923" set key2=42QTK-RN8M7-J3C4G-BBGYM-88CYV & set officever=Mondo
rem Professional Plus
if "%number%" == "SKU ID: b322da9c-a2e2-4058-9e4e-f59a6970bd69" set key2=YC7DK-G2NP3-2QQC3-J6H88-GVGXT & set officever=Professional Plus
rem Standard
if "%number%" == "SKU ID: b13afb38-cd79-4ae5-9f7f-eed058d750ca" set key2=KBKQT-2NMXY-JJWGP-M62JB-92CD4 & set officever=Standard
rem Project Professional
if "%number%" == "SKU ID: 4a5d124a-e620-44ba-b6ff-658961b33b9a" set key2=FN8TT-7WMH6-2D4X9-M337T-2342K & set officever=Project Professional
rem Project Standard
if "%number%" == "SKU ID: 427a28d1-d17c-4abf-b717-32c780ba6f07" set key2=6NTH3-CW976-3G3Y2-JK3TX-8QHTT & set officever=Project Standard
rem Visio Professional
if "%number%" == "SKU ID: e13ac10e-75d0-4aff-a0cd-764982cf541c" set key2=C2FG9-N6J68-H8BTJ-BW3QX-RM3B3 & set officever=Visio Professional
rem Visio Standard
if "%number%" == "SKU ID: ac4efaf0-f81f-4f61-bdf7-ea32b02ab117" set key2=J484Y-4NKBF-W2HMG-DBMJC-PGWR7 & set officever=Visio Standard
rem Access
if "%number%" == "SKU ID: 6ee7622c-18d8-4005-9fb7-92db644a279b" set key2=NG2JY-H4JBT-HQXYP-78QH9-4JM2D & set officever=Access
rem Excel
if "%number%" == "SKU ID: f7461d52-7c2b-43b2-8744-ea958e0bd09a" set key2=VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB & set officever=Excel
rem Groove
if "%number%" == "SKU ID: fb4875ec-0c6b-450f-b82b-ab57d8d1677f" set key2=H7R7V-WPNXQ-WCYYC-76BGV-VT7GH & set officever=Groove
rem InfoPath
if "%number%" == "SKU ID: a30b8040-d68a-423f-b0b5-9ce292ea5a8f" set key2=DKT8B-N7VXH-D963P-Q4PHY-F8894 & set officever=InfoPath
rem Lync
if "%number%" == "SKU ID: 1b9f11e3-c85c-4e1b-bb29-879ad2c909e3" set key2=2MG3G-3BNTT-3MFW9-KDQW3-TCK7R & set officever=Lync
rem OneNote
if "%number%" == "SKU ID: efe1f3e6-aea2-4144-a208-32aa872b6545" set key2=TGN6P-8MMBC-37P2F-XHXXK-P34VW & set officever=OneNote
rem Outlook
if "%number%" == "SKU ID: 771c3afa-50c5-443f-b151-ff2546d863a0" set key2=QPN8Q-BJBTJ-334K3-93TGY-2PMBT & set officever=Outlook
rem Powerpoint
if "%number%" == "SKU ID: 8c762649-97d1-4953-ad27-b7e2c25b972e" set key2=4NT99-8RJFH-Q2VDH-KYG2C-4RD4F & set officever=Powerpoint
rem Publisher
if "%number%" == "SKU ID: 00c79ff1-6850-443d-bf61-71cde0de305f" set key2=PN2WF-29XG2-T9HJ7-JQPJR-FCXK4 & set officever=Publisher
rem Word
if "%number%" == "SKU ID: d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3" set key2=6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7 & set officever=Word
goto officeinfo
:office2010
rem Mondo
if "%number%" == "SKU ID: 09ed9640-f020-400a-acd8-d7d867dfd9c2" set key2=YBJTT-JG6MD-V9Q7P-DBKXJ-38W9R & set officever=Mondo
rem Mondo 2
if "%number%" == "SKU ID: ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c" set key2=7TC2V-WXF6P-TD7RT-BQRXR-B8K32 & set officever=Mondo 2
rem Professional Plus
if "%number%" == "SKU ID: 6f327760-8c5c-417c-9b61-836a98287e0c" set key2=VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB & set officever=Professional Plus
rem Standard
if "%number%" == "SKU ID: 9da2a678-fb6b-4e67-ab84-60dd6a9c819a" set key2=V7QKV-4XVVR-XYV4D-F7DFM-8R6BM & set officever=Standard
rem Project Professional
if "%number%" == "SKU ID: df133ff7-bf14-4f95-afe3-7b48e7e331ef" set key2=YGX6F-PGV49-PGW3J-9BTGG-VHKC6 & set officever=Project Professional
rem Project Standard
if "%number%" == "SKU ID: 5dc7bf61-5ec9-4996-9ccb-df806a2d0efe" set key2=4HP3K-88W3F-W2K3D-6677X-F9PGB & set officever=Project Standard
rem Visio Premium
if "%number%" == "SKU ID: 92236105-bb67-494f-94c7-7f7a607929bd" set key2=D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ & set officever=Visio Premium
rem Visio Professional
if "%number%" == "SKU ID: e558389c-83c3-4b29-adfe-5e4d7f46c358" set key2=7MCW8-VRQVK-G677T-PDJCM-Q8TCP & set officever=Visio Professional
rem Visio Standard
if "%number%" == "SKU ID: 9ed833ff-4f92-4f36-b370-8683a4f13275" set key2=767HD-QGMWX-8QTDB-9G3R2-KHFGJ & set officever=Visio Standard
rem Access
if "%number%" == "SKU ID: 8ce7e872-188c-4b98-9d90-f8f90b7aad02" set key2=V7Y44-9T38C-R2VJK-666HK-T7DDX & set officever=Access
rem Excel
if "%number%" == "SKU ID: cee5d470-6e3b-4fcc-8c2b-d17428568a9f" set key2=H62QG-HXVKF-PP4HP-66KMR-CW9BM & set officever=Excel
rem Groove SharePoint Workspace
if "%number%" == "SKU ID: 8947d0b8-c33b-43e1-8c56-9b674c052832" set key2=QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4 & set officever=Groove SharePoint Workspace
rem InfoPath
if "%number%" == "SKU ID: ca6b6639-4ad6-40ae-a575-14dee07f6430" set key2=K96W8-67RPQ-62T9Y-J8FQJ-BT37T & set officever=InfoPath
rem OneNote
if "%number%" == "SKU ID: ab586f5c-5256-4632-962f-fefd8b49e6f4" set key2=Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX & set officever=OneNote
rem Outlook
if "%number%" == "SKU ID: ecb7c192-73ab-4ded-acf4-2399b095d0cc" set key2=7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ & set officever=Outlook
rem Powerpoint
if "%number%" == "SKU ID: 45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a" set key2=RC8FX-88JRY-3PF7C-X8P67-P4VTT & set officever=Powerpoint
rem Publisher
if "%number%" == "SKU ID: b50c4f75-599b-43e8-8dcd-1081a7967241" set key2=BFK7F-9MYHM-V68C7-DRQ66-83YTP & set officever=Publisher
rem Word
if "%number%" == "SKU ID: 2d0882e7-a4e7-423b-8ccc-70d91e0158b1" set key2=HVHB3-C6FV7-KQX9W-YQG79-CRY7T & set officever=Word
rem Home and Business
if "%number%" == "SKU ID: ea509e87-07a1-4a45-9edc-eba5a39f36af" set key2=D6QFG-VBYP2-XQHM7-J97RH-VVRCK & set officever=Home and Business
goto officeinfo

goto officeinfo

:DelReg
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55C92734-D682-4D71-983E-D6EC3F16059F" /f >nul 2>&1
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0FF1CE15-A989-479D-AF46-F275C6370663" /f >nul 2>&1
reg delete "HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55C92734-D682-4D71-983E-D6EC3F16059F" /f >nul 2>&1
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\59A52881-A989-479D-AF46-F275C6370663" /f >nul 2>&1
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\0FF1CE15-A989-479D-AF46-F275C6370663" /f >nul 2>&1
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
:officeinfo
cls
set ac=
if "%systemid%" == "Windows 7" goto wkmsserverold
if "%systemid%" == "Windows 8" goto wkmsserverold
goto wtapadapter
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
cscript.exe /nologo %WINDIR%\System32\slmgr.vbs /xpr
cscript "c:\windows\system32\slmgr.vbs" //nologo -dli
%OSPP% /actype:2
%OSPP% /remhst
%OSPP% /setprt:1688
%OSPP% /sethst:10.0.0.20
echo. Installing KMS Client key...
echo.
%OSPP% /inpkey:%KEY2%
echo. Activating
Call :DelReg
%OSPP% /act
call :StopService "sppsvc"
del /f /q "%SystemRoot%\system32\SppExtComObjPatcher.exe" 
del /f /q "%SystemRoot%\system32\SppExtComObjHook.dll" 
call :RemoveIFEOEntry "SppExtComObj.exe"
sc start sppsvc trigger=timer;sessionid=0 >nul 2>&1
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
echo.
:officeact
echo. Installing KMS Client key...
%OSPP% /inpkey:%KEY2%
echo. Setting up KMS Address...
%OSPP% /remhst
%OSPP% /setprt:%port%
%OSPP% /sethst:%kmsip%
echo. Activating...
echo.
%OSPP% /act
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