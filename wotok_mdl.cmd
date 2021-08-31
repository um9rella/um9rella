:: Created by Krakatoa@MDL
:: Support for Windows 10 and Office 2019

@echo off
color 1F
echo Please wait...
(NET FILE||(powershell start-process -FilePath '%0' -verb runas)&&(exit /B)) >NUL 2>&1
(NET FILE||(exit)) >NUL 2>&1
setlocal EnableExtensions
setlocal EnableDelayedExpansion

set "ctrConfig=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
set "CTRF=%CommonProgramFiles%\Microsoft Shared\ClickToRun"
set "cdnms=http://officecdn.microsoft.com/pr"
set "cdn=%cdnms%"
for /f "tokens=2 delims==" %%G in ('wmic path Win32_Processor get AddressWidth /value') do (set "bitos=%%G")

:MainMenu
cls
title WOTOK
pushd %~dp0
set "choice=-"
echo.
echo ============================================================
echo  WOTOK v2018.10.07.02
echo ============================================================
echo.
echo  [ 0 ] Windows/Office - Check status
echo  [ 1 ] Office - CDN version
echo  [ 2 ] Office - Download
echo  [ 3 ] Office - Install Offline
echo  [ 4 ] Office - Install Online
echo  [ 5 ] Office - Uninstall All
echo  [ 6 ] Office - Check update
echo  [ 7 ] Office - Set CDNBaseUrl
echo  [ 8 ] Office - Install VL
echo  [ 9 ] Office - Uninstall License
echo  [ * ] Windows - Install key
echo  [ + ] Activation VL
echo  [ / ] Windows - Digital License
echo  [ - ] EXIT
echo.
set /p choice= Choice: 
if %choice%==0 goto:status
if %choice%==1 goto:cdn
if %choice%==2 goto:Download
if %choice%==3 goto:InstallOffline
if %choice%==4 goto:InstallOnline
if %choice%==5 goto:uninstall
if %choice%==6 goto:checkupdate
if %choice%==7 goto:setCDNBaseUrl
if %choice%==8 goto:installvl
if %choice%==9 goto:upk
if %choice%==* goto:installwinkey
if %choice%==+ goto:activation
if %choice%==/ goto:DigitalLicense
if %choice%==- goto:EXIT
goto:MainMenu


:upk
cls
echo.
echo Office - Uninstall License
echo.
set "c=0"
for /f "TOKENS=2 DELIMS==" %%D in ('WMIC PATH SoftwareLicensingProduct WHERE "Name LIKE 'Office%%' AND LicenseStatus>'1'" GET ID /VALUE') do (
  set /a c+=1
  set "i_upkid!c!=%%D"
  for /f "TOKENS=2 DELIMS==" %%M in ('WMIC PATH SoftwareLicensingProduct WHERE "ID='%%D'" GET LicenseFamily /VALUE') do (set "i_upklf!c!=%%M")
)
if %c%==0 (echo Licence for uninstall not exist&&pause&&goto:MainMenu)
FOR /L %%k in (1,1,%c%) do (echo [%%k] !i_upklf%%k!)
set userinp=-
set /p userinp=Choice: 
if %userinp%==0 goto:MainMenu
if %userinp%==- goto:MainMenu
if %userinp% GTR %c% goto:MainMenu
set "upkid=!i_upkid%userinp%!"
set "upklf=!i_upklf%userinp%!"
cscript //nologo "%windir%\system32\slmgr.vbs" /upk %upkid%
echo.
pause
goto:MainMenu



:activation
cls
echo.
echo [1] Install KMS seco
echo [2] Uninstall KMS seco
echo [3] Activation VL
echo [-] MainMenu
echo.
set activation=-
set /p activation= Choice: 
if %activation%==1 goto:installkms
if %activation%==2 goto:uninstallkms
if %activation%==3 goto:activationvl
if %activation%==- goto:MainMenu
goto:MainMenu


:installwinkey
cls
echo.
echo Install Windows generic key
echo [1] Core VL
echo [2] Professional VL
echo [3] ProfessionalWorkstation VL
echo [4] Enterprise VL
echo [5] Education VL
echo [6] Core
echo [7] Professional
echo [8] ProfessionalWorkstation
echo [9] ProfessionalWorkstation Insider
echo [0] Enterprise
echo [/] Education
echo [*] Bonus: Migrate data from C:\GenuineTicket.xml
echo [+] Bonus: Activation slmgr /ato
echo [-] MainMenu
echo.
set installwinkey=-
set "ipk=cscript //nologo %systemroot%\system32\slmgr.vbs /ipk"
set /p installwinkey= Choice: 
echo.
if %installwinkey%==1 (%ipk% TX9XD-98N7V-6WMQ6-BX7FG-H8Q99)
if %installwinkey%==2 (%ipk% W269N-WFGWX-YVC9B-4J6C9-T83GX)
if %installwinkey%==3 (%ipk% NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J)
if %installwinkey%==4 (%ipk% NPPR9-FWDCX-D2C8J-H872K-2YT43)
if %installwinkey%==5 (%ipk% NW6C2-QMPVW-D7KKK-3GKT6-VCFB2)
if %installwinkey%==6 (%ipk% YTMG3-N6DKC-DKB77-7M9GH-8HVX7)
if %installwinkey%==7 (%ipk% VK7JG-NPHTM-C97JM-9MPGT-3V66T)
if %installwinkey%==8 (%ipk% DXG7C-N36C4-C4HTG-X4T3X-2YV77)
if %installwinkey%==9 (%ipk% 3VKCP-7QN23-D3GQV-MD2VH-XD863)
if %installwinkey%==0 (%ipk% XGVPP-NMH47-7TTHJ-W3FW7-8HV2C)
if %installwinkey%==/ (%ipk% YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY)
if %installwinkey%==* (clipup -v -o -altto c:)
if %installwinkey%==+ (cscript //nologo %systemroot%\system32\slmgr.vbs /ato)
echo.
pause
goto:MainMenu


:installvl
cls
call :product vl
echo %product% %pidkeys%
echo.
echo MainMenu (Enter) or Install VL (1)
set installvl=-
set /p installvl= Choice: 
if %installvl%==1 goto:installvlstart
if %installvl%==- goto:MainMenu
goto:MainMenu
:installvlstart
echo Installing VL, please wait...
"%ProgramFiles%\Microsoft Office\root\integration\integrator.exe" /I /License PRIDName=%product%.16 PidKey=%pidkeys% >nul
for /f "tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" 2^>nul') do set "ProductReleaseIds=%%b"
for %%a in (%ProductReleaseIds%) do (if %%a==%product% (goto:installvlcontinue))
reg add %ctrConfig% /v ProductReleaseIds /t REG_SZ /d "%ProductReleaseIds%,%product%" /f >nul
:installvlcontinue
reg add %ctrConfig% /v %product%.OSPPReady /t REG_SZ /d 1 /f >nul
echo Installed VL
echo.
pause
goto:MainMenu


:setCDNBaseUrl
cls
call :id
echo.
echo MainMenu (Enter) or set CDNBaseUrl %id% (1)
set setCDNBaseUrl=-
set /p setCDNBaseUrl= Choice: 
if %setCDNBaseUrl%==1 goto:setCDNBaseUrlstart
if %setCDNBaseUrl%==- goto:MainMenu
goto:MainMenu
:setCDNBaseUrlstart
reg add %ctrConfig% /v CDNBaseUrl /t REG_SZ /d "%cdnms%/%id%" /f >nul
echo Set CDNBaseUrl
echo.
pause
goto:MainMenu

:checkupdate
cls
echo.
echo MainMenu (Enter) or Check update Office (1)
set checkupdate=-
set /p checkupdate= Choice: 
if %checkupdate%==1 goto:checkupdatestart
if %checkupdate%==- goto:MainMenu
goto:MainMenu
:checkupdatestart
echo Checking update Office, please wait...
"%CTRF%\OfficeC2RClient.exe" /update user updatepromptuser=True displaylevel=True
echo Checked update Office
echo.
pause
goto:MainMenu


:uninstall
cls
echo.
echo MainMenu (Enter) or Uninstall All Office (1)
set uninstall=-
set /p uninstall= Choice: 
if %uninstall%==1 goto:uninstallstart
if %uninstall%==- goto:MainMenu
goto:MainMenu
:uninstallstart
echo Uninstaling All Office, please wait...
"%CTRF%\OfficeClickToRun.exe" productstoremove=AllProducts
echo All Office uninstalled
echo.
pause
goto:MainMenu


:status
cls
set "choice=-"
echo.
echo  Win/Office - Check status
echo.
echo  [ 1 ] Windows slmgr dli
echo  [ 2 ] Windows slmgr dlv
echo  [ 3 ] Office slmgr dli
echo  [ 4 ] Office slmgr dlv
echo  [ 5 ] slmgr dli all
echo  [ 6 ] slmgr dlv all
echo  [ 7 ] Office ospp dstatus
echo  [ 8 ] Office ospp dstatusall
echo  [ - ] MainMenu
echo.
set /p choice= Choice: 
if %choice%==1 call :status_slmgr_w dli
if %choice%==2 call :status_slmgr_w dlv
if %choice%==3 call :status_slmgr_o dli
if %choice%==4 call :status_slmgr_o dlv
if %choice%==5 call :slmgr_all dli
if %choice%==6 call :slmgr_all dlv
if %choice%==7 call :ospp dstatus
if %choice%==8 call :ospp dstatusall
if %choice%==- goto:MainMenu
goto:status


:InstallOnline
cls
echo.
echo Install Office Online
call :Lang
echo Lang: %lang%
call :Bit
echo Bit: %bit%
call :id
if "%version%"=="" (call :check %id%)
echo ID: %id%  %channel%  %version%
set "cdnbaseurl=%cdnms%/%id%"
set "baseurl=%cdnbaseurl%"
set "files=temp"
if not exist temp (md temp >nul 2>&1)
goto:InstallFinal


:InstallOffline
cls
echo Install Office Offline
echo.
echo Version 
set "c=0"
for /f %%D in ('dir /a:d /b Office\Data\') do (set /a c+=1 && set "i_ver!c!=%%D")
if %c%==0 (echo Version not exist&&pause&&goto:Download)
FOR /L %%k in (1,1,%c%) do (echo [   %%k	] !i_ver%%k!)
set userinp=-
set /p userinp=Choice: 
if %userinp%==0 goto:MainMenu
if %userinp%==- goto:MainMenu
if %userinp% GTR %c% goto:MainMenu
set "ver=!i_ver%userinp%!"
echo Version: %ver%
echo.
echo Bit
set "c=0"
for /f "tokens=2 delims=.x" %%D in ('dir /a /b Office\Data\%ver%\stream.x*.x-none.dat') do (set /a c+=1 && If %%D==86 (set "i_bit!c!=32") else (set "i_bit!c!=%%D"))
FOR /L %%k in (1,1,%c%) do (echo [   %%k	] !i_bit%%k!)
set userinp=-
set /p userinp=Choice: 
if %userinp%==0 goto:MainMenu
if %userinp%==- goto:MainMenu
if %userinp% GTR %c% goto:MainMenu
set "bit=!i_bit%userinp%!"
if %bit%==32 (set arch=86)
if %bit%==64 (set arch=64)
echo Bit: %bit%
echo.
echo Lang
set "c=0"
for /f "tokens=3 delims=." %%D in ('dir /a /b Office\Data\%ver%\stream.x%arch%.*.dat') do (If not %%D==x-none (set /a c+=1 && set "i_lang!c!=%%D"))
FOR /L %%k in (1,1,%c%) do (echo [   %%k	] !i_lang%%k!)
set userinp=-
set /p userinp=Choice: 
if %userinp%==0 goto:MainMenu
if %userinp%==- goto:MainMenu
if %userinp% GTR %c% goto:MainMenu
set "lang=!i_lang%userinp%!"
call :lcid %lang%
echo Lang: %lang%
call :id
echo ID: %id%
if %userinp% GTR 0 set version=%ver%
set "cdnbaseurl=%cdnms%/%id%"
set "baseurl=%cd%"
set "files=Office\Data\%version%"
goto:InstallFinal


:InstallFinal
call :product all
echo %product%
call :excludedapps
echo excludedapps: %excludedapps%
echo.
echo Install Office %id% %version% %bit%bit %lang%
echo "%CTRF%\OfficeClickToRun.exe"^
 cdnbaseurl=%cdnbaseurl%^
 baseurl="%baseurl%"^
 version=%version%^
 platform=x%arch%^
 productstoadd=%product%_%lang%_x-none^
 deliverymechanism=%id%%excludedapps%
echo.
echo MainMenu (Enter) or Install (1)
set installstart=-
set /p installstart= Choice: 
if %installstart%==1 goto:InstallStart
if %installstart%==- goto:MainMenu
goto:MainMenu
:InstallStart
echo.
if exist "%CTRF%\OfficeClickToRun.exe" (
  echo OfficeClickToRun.exe exist
) else (
  echo Copying installer, please wait...
  md "%CTRF%"
  if %files%==temp (
    powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/i%bitos%%lcid%.cab '%~dp0temp\i%bitos%%lcid%.cab' }"
    powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/i%bitos%0.cab '%~dp0temp\i%bitos%0.cab' }"
  )
  expand %files%\i%bitos%0.cab -F:* "%CTRF%" >nul 2>&1
  expand %files%\i%bitos%%lcid%.cab -F:* "%CTRF%" >nul 2>&1
  if exist temp (rd /S /Q temp >nul 2>&1)
)
echo Installing Office, please wait...
"%CTRF%\OfficeClickToRun.exe"^
 cdnbaseurl=%cdnbaseurl%^
 baseurl="%baseurl%"^
 version=%version%^
 platform=x%arch%^
 productstoadd=%product%_%lang%_x-none^
 deliverymechanism=%id%%excludedapps%
echo Office installed
echo.
pause
goto:MainMenu


:cdn
cls
echo.
echo AudienceId                            Version           AudienceData        Channel
set "aid=EA4A4090-DE26-49D7-93C1-91BFF9E53FC3"
call :check %aid%
echo %aid%  %version%  Dogfood::DevMain
set "ver_ddm=%version%"
set "aid=F3260CF1-A92C-4C75-B02E-D64C0A86A968"
call :check %aid%
echo %aid%  %version%  Dogfood::CC
set "ver_dcc=%version%"
set "aid=B61285DD-D9F7-41F2-9757-8F61CBA4E9C8"
call :check %aid%
echo %aid%  %version%  Microsoft::DevMain
set "ver_mdm=%version%"
set "aid=5462EEE5-1E97-495B-9370-853CD873BB07"
call :check %aid%
echo %aid%  %version%  Microsoft::CC
set "ver_mcc=%version%"
set "aid=1D2D2EA6-1680-4C56-AC58-A441C8C24FF9"
call :check %aid%
echo %aid%  %version%  Microsoft::LTSC
set "ver_mltsc=%version%"
set "aid=5440FD1F-7ECB-4221-8110-145EFAA6372F"
call :check %aid%
echo %aid%  %version%  Insiders::DevMain   InsiderFast
set "ver_idm=%version%"
set "aid=64256AFE-F5D9-4F86-8936-8840A6A4F5BE"
call :check %aid%
echo %aid%  %version%  Insiders::CC        Insiders
set "ver_icc=%version%"
set "aid=492350F6-3A01-4F97-B9C0-C7C6DDF67D60"
call :check %aid%
echo %aid%  %version%  Production::CC      Monthly
set "ver_pcc=%version%"
set "aid=F2E724C1-748F-4B47-8FB8-8E0D210E9208"
call :check %aid%
echo %aid%  %version%  Production::LTSC    Perpetual2019
set "ver_pltsc=%version%"
set version=
pause
goto:MainMenu


:Download
cls
echo.
echo Download Office
call :Lang
echo Lang: %lang%
call :Bit
echo Bit: %bit%
call :id
if "%version%"=="" (call :check %id%)
echo ID: %id%  %channel%  %version%
echo.
echo Download files for:
echo %id%  %channel%  %version%  %bit%bit  %lang%
echo Download (Enter) or MainMenu (-)
set "downmenu=1"
set /p downmenu= Choice: 
if %downmenu%==1 goto:DownloadStart
if %downmenu%==- goto:MainMenu
goto:MainMenu
:DownloadStart
echo Downloading files...
if not exist "Office\Data\%version%" (md "Office\Data\%version%" >nul 2>&1)
if not exist "Office\Data\%version%\i32%lcid%.cab" (powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/i32%lcid%.cab '%~dp0Office\Data\%version%\i32%lcid%.cab' }")
if not exist "Office\Data\%version%\i320.cab" (powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/i320.cab '%~dp0Office\Data\%version%\i320.cab' }")
if not exist "Office\Data\%version%\i64%lcid%.cab" (powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/i64%lcid%.cab '%~dp0Office\Data\%version%\i64%lcid%.cab' }")
if not exist "Office\Data\%version%\i640.cab" (powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/i640.cab '%~dp0Office\Data\%version%\i640.cab' }")
if not exist "Office\Data\%version%\s%bit%%lcid%.cab" (powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/s%bit%%lcid%.cab '%~dp0Office\Data\%version%\s%bit%%lcid%.cab' }")
if not exist "Office\Data\%version%\s%bit%0.cab" (powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/s%bit%0.cab '%~dp0Office\Data\%version%\s%bit%0.cab' }")
if not exist "Office\Data\%version%\stream.x%arch%.%lang%.dat" (powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/stream.x%arch%.%lang%.dat '%~dp0Office\Data\%version%\stream.x%arch%.%lang%.dat' }")
if not exist "Office\Data\%version%\stream.x%arch%.x-none.dat" (powershell -command "& { Start-BitsTransfer %cdn%/%id%/Office/Data/%version%/stream.x%arch%.x-none.dat '%~dp0Office\Data\%version%\stream.x%arch%.x-none.dat' }")
echo.
echo MainMenu (Enter) or Install (1)
set downinstall=-
set /p downinstall= Choice: 
if %downinstall%==1 (set "cdnbaseurl=%cdnms%/%id%"&&set "baseurl=%cd%"&&set "files=Office\Data\%version%"&&goto:InstallFinal)
if %downinstall%==- goto:MainMenu
goto:MainMenu


:excludedapps
echo.
set "excludedapps=0"
echo excludedapps
echo access^^,excel^^,groove^^,lync^^,onedrive^^,onenote^^,outlook^^,powerpoint^^,publisher^^,word^^,visio^^,project
echo Example: access^^,publisher^^,groove^^,lync^^,onenote^^,onedrive
echo Default (Enter):
set /p excludedapps=Choice: 
if not %excludedapps%==0 (set "excludedapps= %product%.excludedapps^=%excludedapps%") else (set "excludedapps=")
goto:eof


:product
echo.
set "product=ProPlus2019Volume"
echo Product
if %1 NEQ vl (
echo ProPlus2019Retail
echo ProjectPro2019Retail 
echo VisioPro2019Retail
)
echo ProPlus2019Volume
echo ProjectPro2019Volume
echo VisioPro2019Volume
echo Example: ProPlus2019Volume
echo Default (Enter): %product%
set /p product=Choice: 
if %1 NEQ vl (
if %product%==ProPlus2019Retail goto:eof
if %product%==ProjectPro2019Retail goto:eof
if %product%==VisioPro2019Retail goto:eof
)
if %product%==ProPlus2019Volume (set "pidkeys=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"&&goto:eof)
if %product%==ProjectPro2019Volume (set "pidkeys=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B"&&goto:eof)
if %product%==VisioPro2019Volume (set "pidkeys=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB"&&goto:eof)
goto:product


:check
powershell -command "& { (New-Object Net.WebClient).DownloadFile('%cdn%/%1/Office/Data/MRO.cab', '%~dp0MRO.cab') }"
expand MRO.cab -F:MRO.xml "%cd%" >nul 2>&1
for /f "tokens=2 delims=<> " %%G in ('findstr /i Available MRO.xml') do set "version=%%G"
del /f /q MRO.xml MRO.cab >nul 2>&1
goto:eof


:Lang
if "%lang%"=="" (for /f "tokens=2 delims=={}," %%G in ('wmic os get MUILanguages /Value') do (set "lang=%%~G"))
echo.
echo Lang:
echo ar-sa bg-bg cs-cz da-dk de-de el-gr en-us es-es et-ee fi-fi
echo fr-fr he-il hi-in hr-hr hu-hu it-it ja-jp ko-kr lt-lt lv-lv
echo ms-my nb-no nl-nl pl-pl pt-br pt-pt ro-ro ru-ru sk-sk sl-si
echo sv-se th-th tr-tr uk-ua vi-vn zh-cn zh-tw sr-latn-rs
echo Example: en-us
echo Default (Enter): %lang%
set /p lang= Choice: 
call :lcid %lang%
if not %lcid%==0 goto:eof
echo Bad choice.
pause
goto:Lang


:Bit
set "bit=%bitos%"
echo.
echo Bit:
echo 32 64
echo Example: 64
echo Default (Enter): %bit%
set /p bit= Choice: 
if %bit%==32 (set arch=86&&goto:eof)
if %bit%==64 (set arch=64&&goto:eof)
echo Bad choice.
pause
goto:Bit


:id
if "%id%"=="" set "id=492350F6-3A01-4F97-B9C0-C7C6DDF67D60"
set version=
echo.
echo ID:
echo EA4A4090-DE26-49D7-93C1-91BFF9E53FC3  Dogfood::DevMain    %ver_ddm%
echo F3260CF1-A92C-4C75-B02E-D64C0A86A968  Dogfood::CC         %ver_dcc%
echo B61285DD-D9F7-41F2-9757-8F61CBA4E9C8  Microsoft::DevMain  %ver_mdm%
echo 5462EEE5-1E97-495B-9370-853CD873BB07  Microsoft::CC       %ver_mcc%
echo 1D2D2EA6-1680-4C56-AC58-A441C8C24FF9  Microsoft::LTSC     %ver_mltsc%
echo 5440FD1F-7ECB-4221-8110-145EFAA6372F  Insiders::DevMain   %ver_idm%
echo 64256AFE-F5D9-4F86-8936-8840A6A4F5BE  Insiders::CC        %ver_icc%
echo 492350F6-3A01-4F97-B9C0-C7C6DDF67D60  Production::CC      %ver_pcc%
echo F2E724C1-748F-4B47-8FB8-8E0D210E9208  Production::LTSC    %ver_pltsc%
echo Example: 492350F6-3A01-4F97-B9C0-C7C6DDF67D60
echo Default (Enter): %id%
set /p id= Choice: 
if /i %id%==EA4A4090-DE26-49D7-93C1-91BFF9E53FC3 (set "channel=Dogfood::DevMain"&&(set "version=%ver_ddm%")&&goto:eof)
if /i %id%==F3260CF1-A92C-4C75-B02E-D64C0A86A968 (set "channel=Dogfood::CC"&&(set "version=%ver_dcc%")&&goto:eof)
if /i %id%==B61285DD-D9F7-41F2-9757-8F61CBA4E9C8 (set "channel=Microsoft::DevMain"&&(set "version=%ver_mdm%")&&goto :eof)
if /i %id%==5462EEE5-1E97-495B-9370-853CD873BB07 (set "channel=Microsoft::CC"&&(set "version=%ver_mcc%")&&goto:eof)
if /i %id%==1D2D2EA6-1680-4C56-AC58-A441C8C24FF9 (set "channel=Microsoft::LTSC"&&(set "version=%ver_mltsc%")&&goto:eof)
if /i %id%==5440FD1F-7ECB-4221-8110-145EFAA6372F (set "channel=Insiders::DevMain"&&(set "version=%ver_idm%")&&goto:eof)
if /i %id%==64256AFE-F5D9-4F86-8936-8840A6A4F5BE (set "channel=Insiders::CC"&&(set "version=%ver_icc%")&&goto:eof)
if /i %id%==492350F6-3A01-4F97-B9C0-C7C6DDF67D60 (set "channel=Production::CC"&&(set "version=%ver_pcc%")&&goto:eof)
if /i %id%==F2E724C1-748F-4B47-8FB8-8E0D210E9208 (set "channel=Production::LTSC"&&(set "version=%ver_pltsc%")&&goto:eof)
echo Bad choice.
pause
goto:id


:lcid
set lcid=0
if /i %1==ar-sa (set lcid=1025)
if /i %1==bg-bg (set lcid=1026)
if /i %1==cs-cz (set lcid=1029)
if /i %1==da-dk (set lcid=1030)
if /i %1==de-de (set lcid=1031)
if /i %1==el-gr (set lcid=1032)
if /i %1==en-us (set lcid=1033)
if /i %1==es-es (set lcid=3082)
if /i %1==et-ee (set lcid=1061)
if /i %1==fi-fi (set lcid=1035)
if /i %1==fr-fr (set lcid=1036)
if /i %1==he-il (set lcid=1037)
if /i %1==hi-in (set lcid=1081)
if /i %1==hr-hr (set lcid=1050)
if /i %1==hu-hu (set lcid=1038)
if /i %1==it-it (set lcid=1040)
if /i %1==ja-jp (set lcid=1041)
if /i %1==ko-kr (set lcid=1042)
if /i %1==lt-lt (set lcid=1063)
if /i %1==lv-lv (set lcid=1062)
if /i %1==ms-my (set lcid=1086)
if /i %1==nb-no (set lcid=1044)
if /i %1==nl-nl (set lcid=1043)
if /i %1==pl-pl (set lcid=1045)
if /i %1==pt-br (set lcid=1046)
if /i %1==pt-pt (set lcid=2070)
if /i %1==ro-ro (set lcid=1048)
if /i %1==ru-ru (set lcid=1049)
if /i %1==sk-sk (set lcid=1051)
if /i %1==sl-si (set lcid=1060)
if /i %1==sv-se (set lcid=1053)
if /i %1==th-th (set lcid=1054)
if /i %1==tr-tr (set lcid=1055)
if /i %1==uk-ua (set lcid=1058)
if /i %1==vi-vn (set lcid=1066)
if /i %1==zh-cn (set lcid=2052)
if /i %1==zh-tw (set lcid=1028)
if /i %1==sr-latn-rs (set lcid=9242)
goto:eof


:status_slmgr_w
cls
for /f "tokens=2 delims==" %%G in ('"wmic path SoftwareLicensingProduct where (PartialProductKey is not null) get ID /value"') do (call :slmgr_w %1 %%G)
pause
goto:eof

:slmgr_w
wmic path SoftwareLicensingProduct where ID='%2' get Name /value | findstr /i "Windows" 1>nul && (cscript //nologo %systemroot%\System32\slmgr.vbs /%1)
exit /b

:status_slmgr_o
cls
for /f "tokens=2 delims==" %%G in ('"wmic path SoftwareLicensingProduct where (PartialProductKey is not null) get ID /value"') do (call :slmgr_o %1 %%G)
pause
goto:eof

:slmgr_o
wmic path SoftwareLicensingProduct where ID='%2' get Name /value | findstr /i "Windows" 1>nul || (cscript //nologo %systemroot%\system32\slmgr.vbs /%1 %2)
exit /b

:slmgr_all
cls
cscript //nologo %systemroot%\system32\slmgr.vbs /%1 all
pause
goto:eof

:ospp
cls
cscript //nologo "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /%1
pause
goto:eof


:installkms
if not exist %bitos% (md %bitos% >nul 2>&1)
if not exist %bitos%\SppExtComObjHook.dll (cls && echo The SppExtComObjHook.dll is requested. Please download the SppExtComObjHook.dll and insert it into folder %bitos%. && pause && goto:installkms)
cls
echo Please wait
set kmsip=172.16.16.16
set port=1688
set hSpp="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
reg delete %hSpp%\55c92734-d682-4d71-983e-d6ec3f16059f /f >nul 2>&1
reg delete %hSpp%\0ff1ce15-a989-479d-af46-f275c6370663 /f >nul 2>&1
set sps=SoftwareLicensingService
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /t REG_DWORD /d 1 /f >nul 2>&1
If Not Exist "%SystemRoot%\System32\SppExtComObjHook.dll" copy /y "%bitos%\SppExtComObjHook.dll" "%SystemRoot%\System32" >nul 2>&1
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\SppExtComObj.exe" /f /v "Debugger" /t REG_SZ /d "rundll32.exe SppExtComObjHook.dll,PatcherMain" >nul 2>&1
for /f "tokens=2 delims==" %%A in ('"wmic path %sps% get version /format:list"') do set ver=%%A
wmic path %sps% where version='%ver%' call SetKeyManagementServiceMachine "%kmsip%" >nul 2>&1
wmic path %sps% where version='%ver%' call SetKeyManagementServicePort %port% >nul 2>&1
sc query sppsvc | find /i "STOPPED" >nul 2>&1 || net stop sppsvc /y >nul 2>&1
sc query sppsvc | find /i "STOPPED" >nul 2>&1 || sc stop sppsvc >nul 2>&1
echo.
echo KMS seco installed
echo.
pause
goto:MainMenu


:uninstallkms
set sps=SoftwareLicensingService
sc query sppsvc | find /i "STOPPED" >nul 2>&1 || net stop sppsvc /y >nul 2>&1
sc query sppsvc | find /i "STOPPED" >nul 2>&1 || sc stop sppsvc >nul 2>&1
if exist "%SystemRoot%\system32\SppExtComObjHook.dll" (del /f /q "%SystemRoot%\system32\SppExtComObjHook.dll") >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\SppExtComObj.exe" /f >nul 2>&1
wmic path %sps% where version='%ver%' call ClearKeyManagementServiceMachine >nul 2>&1
wmic path %sps% where version='%ver%' call ClearKeyManagementServicePort >nul 2>&1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceDnsPublishing 1 >nul 2>&1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceHostCaching 1 >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f" /f >nul 2>&1
reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies\55c92734-d682-4d71-983e-d6ec3f16059f" /f >nul 2>&1
reg delete "HKEY_USERS\S-1-5-20\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies\0ff1ce15-a989-479d-af46-f275c6370663" /f >nul 2>&1
reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /f >nul 2>&1
sc start sppsvc trigger=timer;sessionid=0 >nul 2>&1
echo.
echo KMS seco uninstalled
echo.
pause
goto:MainMenu


:activationvl
set spp=SoftwareLicensingProduct
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where ProductKeyChannel='Volume:GVLK' get ID /format:list"') do (set ID=%%G&echo.&call :Activate %%G)
echo.
pause
goto:MainMenu

:Activate
wmic path %spp% where ID='%1' call ClearKeyManagementServiceMachine >nul 2>&1
wmic path %spp% where ID='%1' call ClearKeyManagementServicePort >nul 2>&1
for /f "tokens=2 delims==" %%x in ('"wmic path %spp% where ID='%1' get Name /format:list"') do echo Activate %%x
wmic path %spp% where ID='%1' call Activate >nul 2>&1
if %errorlevel% neq 0 echo Activation failed&exit /b
for /f "tokens=2 delims==" %%x in ('"wmic path %spp% where ID='%1' get GracePeriodRemaining /format:list"') do (set gpr=%%x&set /a gpr2=%%x/60/24)
echo Activated for %gpr2% days (%gpr% minutes)&exit /b



:DigitalLicense
cls
echo.
echo Note from mod: We don't support the HWID exploit in MDL.
echo Source code for generating digital license in the post is deleted.
echo.
pause
goto:MainMenu

:Exit