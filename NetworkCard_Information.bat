@echo off
Title NetworkCard Informations
color 0A
setlocal EnableDelayedExpansion
set /a nbrcard=0
set /a nbrcurrent=0
call :UP_BAR 0

REM Récupère l'index des cartes réseaux qui ont un nom/ID
for /F %%a in ('wmic nic where "NetConnectionID is not null" get index ^|findstr /r "^[0-9]"') do (set /a nbrcard+=1 && set /a NICIndex[!nbrcard!]=%%a)

REM Pour chaque carte réseaux
for /L %%b in (1,1,%nbrcard%) do (
	
	REM Nom de la connexion
	for /F "skip=1 delims=" %%c in ('wmic nic where "Index=!NICIndex[%%b]!" get NetConnectionID ^|findstr /r "[0-9-A-Z-a-z]"') do (set aResult[%%b][1]=%%c)

	REM Fabriquant
	for /F "skip=1 delims=" %%d in ('wmic nic where "Index=!NICIndex[%%b]!" get Manufacturer ^|findstr /r "[0-9-A-Z-a-z]"') do (set aResult[%%b][2]=%%d)
	
	REM Modèle
	for /F "skip=1 delims=" %%e in ('wmic nic where "Index=!NICIndex[%%b]!" get Name ^|findstr /r "[0-9-A-Z-a-z]"') do (set aResult[%%b][3]=%%e)

	REM Etat de la connexion
	for /F "delims=" %%f in ('wmic nic where "Index=!NICIndex[%%b]!" get NetConnectionStatus ^|findstr /r "^[0-9]"') do (call :TradStatus %%b %%f)
	call :UP_BAR ^(^(4+^(16*(%%b-1^)^)^)*100^)/^(16*%nbrcard%^)

	REM Adresse MAC
	for /F "skip=1 delims=" %%g in ('wmic nic where "Index=!NICIndex[%%b]!" get MACAddress ^|findstr /r "[0-9-A-Z-a-z]"') do (set aResult[%%b][5]=%%g)

	REM Type de connexion (càble, wireless)
	for /F "skip=1 delims=" %%h in ('wmic nic where "Index=!NICIndex[%%b]!" get AdapterType ^|findstr /r "[0-9-A-Z-a-z]"') do (set aResult[%%b][6]=%%h)

	REM DHCP activé
	for /F "skip=1 delims=" %%i in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get DHCPEnabled ^|findstr /r "[0-9-A-Z-a-z]"') do (call :TradENtoFR %%b %%i)

	REM Serveur DHCP
	for /F "delims=" %%j in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get DHCPServer ^|findstr /r "[0-9]"') do (set aResult[%%b][8]=%%j)
	call :UP_BAR ^(^(8+^(16*(%%b-1^)^)^)*100^)/^(16*%nbrcard%^)

	REM Bail DHCP - Obtenu
	for /F "delims=" %%k in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get DHCPLeaseObtained ^|findstr /r "[0-9]"') do (call :DateFR %%b O %%k)

	REM Bail DHCP - Expiration
	for /F "delims=" %%l in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get DHCPLeaseExpires ^|findstr /r "[0-9]"') do (call :DateFR %%b E %%l)

	REM Adresse IP
	for /F "delims=" %%m in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get IPAddress ^|findstr /r "[0-9]"') do (call :Ipv4Only %%b IP %%m)

	REM Masque
	for /F "delims=" %%n in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get IPSubnet ^|findstr /r "[0-9]"') do (call :Ipv4Only %%b Masque %%n)
	call :UP_BAR ^(^(12+^(16*(%%b-1^)^)^)*100^)/^(16*%nbrcard%^)

	REM Passerelle
	for /F "delims=" %%o in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get DefaultIPGateway ^|findstr /r "[0-9]"') do (call :Ipv4Only %%b Pass %%o)

	REM Serveur DNS
	for /F "delims=" %%p in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get DNSServerSearchOrder ^|findstr /r "[0-9]"') do (call :Ipv4Only %%b DNS %%p)

	REM Suffixe DNS recherché
	for /F "skip=1 delims=" %%q in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get DNSDomainSuffixSearchOrder ^|findstr /r "[0-9-A-Z-a-z]"') do (call :SuffixeDNSFormat %%b %%q)

	REM Nom DNS du PC
	for /F "skip=1 delims=" %%r in ('wmic nicconfig where "Index=!NICIndex[%%b]!" get DNSHostName ^|findstr /r "[0-9-A-Z-a-z]"') do (set aResult[%%b][16]=%%r)
	call :UP_BAR ^(^(16+^(16*(%%b-1^)^)^)*100^)/^(16*%nbrcard%^)
)

REM On affiche le résultat
timeout /t 1 >nul
cls
color 0F
for /L %%s in (1,1,%nbrcard%) do (
	echo.
	If defined aResult[%%s][1] (echo Nom de la connexion	: !aResult[%%s][1]!)
	If defined aResult[%%s][2] (echo Fabriquant		: !aResult[%%s][2]!)
	If defined aResult[%%s][3] (echo ModŠle			: !aResult[%%s][3]!)
	If defined aResult[%%s][4] (echo Etat de la connexion	: !aResult[%%s][4]!)
	If defined aResult[%%s][5] (echo Adresse MAC		: !aResult[%%s][5]!)
	If defined aResult[%%s][6] (echo Type de connexion	: !aResult[%%s][6]!)
	If defined aResult[%%s][7] (echo DHCP activ‚		: !aResult[%%s][7]!)
	If defined aResult[%%s][8] (echo Serveur DHCP		: !aResult[%%s][8]!)
	If defined aResult[%%s][9] (echo Bail DHCP - Obtenu	: !aResult[%%s][9]!)
	If defined aResult[%%s][10] (echo Bail DHCP - Expiration	: !aResult[%%s][10]!)
	If defined aResult[%%s][11] (echo Adresse IP		: !aResult[%%s][11]!)
	If defined aResult[%%s][12] (echo Masque			: !aResult[%%s][12]!)
	If defined aResult[%%s][13] (echo Passerelle		: !aResult[%%s][13]!)
	If defined aResult[%%s][14] (echo Serveur DNS		: !aResult[%%s][14]!)
	If defined aResult[%%s][15] (echo Suffixe DNS recherch‚	: !aResult[%%s][15]!)
	If defined aResult[%%s][16] (echo Nom DNS du PC		: !aResult[%%s][16]!)
	echo.
	echo.
)

pause>nul|echo Appuyez sur une touche pour quitter le script...
exit

:TradStatus
set /a ConnectStatusN=%2
If %ConnectStatusN% EQU 0 set ConnectStatus=D‚connect‚
If %ConnectStatusN% EQU 1 set ConnectStatus=Connexion en cours
If %ConnectStatusN% EQU 2 set ConnectStatus=Connect‚
If %ConnectStatusN% EQU 3 set ConnectStatus=D‚connexion en cours
If %ConnectStatusN% EQU 4 set ConnectStatus=Carte d‚sactiv‚
If %ConnectStatusN% EQU 5 set ConnectStatus=Mat‚riel désactiv‚
If %ConnectStatusN% EQU 6 set ConnectStatus=Dysfonctionnement mat‚riel
If %ConnectStatusN% EQU 7 set ConnectStatus=M‚dia d‚connect‚
If %ConnectStatusN% EQU 8 set ConnectStatus=Authentification
If %ConnectStatusN% EQU 9 set ConnectStatus=Authentification r‚ussie
If %ConnectStatusN% EQU 10 set ConnectStatus=Authentification ‚chou‚e
If %ConnectStatusN% EQU 11 set ConnectStatus=Adresse invalide
If %ConnectStatusN% EQU 12 set ConnectStatus=Informations requises
If not defined ConnectStatus set ConnectStatus=N/A
set aResult[%1][4]=%ConnectStatus%
GOTO:EOF

:TradENtoFR
set ENtoFR=%2
set ENtoFR=%ENtoFR:True=Oui%
set ENtoFR=%ENtoFR:False=Non%
set aResult[%1][7]=%ENtoFR%
GOTO:EOF

:DateFR
set DateFR=%3
set DateFR=%DateFR:~6,2%/%DateFR:~4,2%/%DateFR:~0,4% %DateFR:~8,2%:%DateFR:~10,2%:%DateFR:~12,2%
If "%2"=="O" (set aResult[%1][9]=%DateFR%) else (set aResult[%1][10]=%DateFR%)
GOTO:EOF

:Ipv4Only
If "%2"=="DNS" (set IPv4=%3 %4 & set tokens=2^^,4) else (set IPv4=%3 & set tokens=2)
for /F delims^=^"^ tokens^=%tokens% %%w in ('echo %IPv4%') do (If "%2"=="DNS" (set IPv4=%%w, %%x) else (set IPv4=%%w))
If "%2"=="IP" (set aResult[%1][11]=%IPv4%)
If "%2"=="Masque" (set aResult[%1][12]=%IPv4%)
If "%2"=="Pass" (set aResult[%1][13]=%IPv4%)
If "%2"=="DNS" (set aResult[%1][14]=%IPv4%)
GOTO:EOF

:SuffixeDNSFormat
set /a NumberCard=%1
set SuffixeDNS=%*
If %NumberCard% LEQ 9 (set SuffixeDNS=%SuffixeDNS:~2%) else (set SuffixeDNS=%SuffixeDNS:~3%)
set SuffixeDNS=%SuffixeDNS:{=%
set SuffixeDNS=%SuffixeDNS:}=%
set SuffixeDNS=%SuffixeDNS:"=%
set aResult[%1][15]=%SuffixeDNS%
GOTO:EOF

:UP_BAR
If defined Percent (set /a PercentOld=%Percent%+2) else (set /a PercentOld=2)
set /a CheckPercentOld=%PercentOld%%%2
If %CheckPercentOld% EQU 1 (set /a PercentOld=%PercentOld%-1)
set /a Percent=%1
If %Percent% LSS 3 (set /a Percent=3)
If %Percent% GEQ 100 (set /a Percent=100)
for /L %%i in (%PercentOld%,2,%Percent%) do (
	set BAR=!BAR!Û
	set complete=%%i
)
cls
echo.
echo.
echo        Chargement ... %complete%%%
echo      ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
echo       %BAR%
echo       %BAR%
echo      ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
echo.
GOTO:EOF
