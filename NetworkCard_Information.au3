#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Data\Image\Icon.ico
#AutoIt3Wrapper_Outfile=Release\NetworkCard_Information.Exe
#AutoIt3Wrapper_Outfile_x64=Release\NetworkCard_Information_x64.Exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Res_Description=NetworkCard Information
#AutoIt3Wrapper_Res_Fileversion=1.0
#AutoIt3Wrapper_Res_LegalCopyright=Copyright (C) 2020-2025 Zaibai Software Production
#AutoIt3Wrapper_Res_Language=1036
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Include
#include "Data\Include\GuiConstructor\GuiConstructor.au3"
#include "Data\Include\GIFAnimation.au3"
#include "Data\Include\Network.au3"
#include <WindowsConstants.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>

;Option
#NoTrayIcon
OnAutoItExitRegister("_Quit")
Opt("WinTitleMatchMode", 3) ;Exact title match (pour If ProcessExists($PIDini) And WinExists($TITLE))
Opt("GUIResizeMode", 802)
Opt("GUIOnEventMode", 1)

;Variable Constante
Dim $TITLE = "NetworkCard Information"
Dim $VERSION = "1.0"

;Variable Constante - Fichier
Dim $DIR_INSTALL = @AppDataDir & "\" & $TITLE & "\"
Dim $DIR_DATA = $DIR_INSTALL & "Data\"
Dim $DIR_IMAGE = $DIR_DATA & "Image\"
Dim $INI_SETTING = $DIR_INSTALL & "paramètres.ini"

#cs
;Check si une instance du programme est déjà en cours
Dim $bPIDExit = False
Local $PIDini = IniRead($INI_SETTING, "Paramètres", "PID", "null")
If ProcessExists($PIDini) And WinExists($Title) Then
	If MsgBox(4 + 32 + 256 + 262144, "Programme déjà lancé", 'Nous avons détecté que ' & $Title & ' est déjà lancé.' & @CRLF & $Title & ' est limité à une instance. Voulez-vous arrêter la première instance ?') = 7 Then
		$bPIDExit = True
		Exit
	EndIf
	ProcessClose($PIDini)
EndIf
IniWrite($INI_SETTING, "Paramètres", "PID", @AutoItPID)
#ce

;Check Version/Installation
Local $CheckVersion = IniRead($INI_SETTING, "Paramètres", "Version", "null")
If $CheckVersion <> $VERSION Or Not FileExists($DIR_INSTALL) Then _Install()

;Chargement du fichier $INI_SETTING
Dim $Ini_XPos = IniRead($INI_SETTING, "Paramètres", "XPos", -1)
Dim $Ini_YPos = IniRead($INI_SETTING, "Paramètres", "YPos", -1)

;Variable utilisé dans diverses fonctions
Dim $aInfoNetworkAdapter = _GetNetworkAdapterInfosModif() ;Récupération des informations réseaux (besoin pour la construction de la GUI)
Dim $iNbrNetworkAdapter = UBound($aInfoNetworkAdapter)
Dim $aNetworkButtonImg[$iNbrNetworkAdapter][2]
Dim $aGuiInfo[$iNbrNetworkAdapter][2]
Dim $aGuiInfoLabel[16] = ["Nom de la connexion", "Fabriquant", "Modèle", "Etat de la connexion", "Adresse MAC", "Type de connexion", "DHCP activé", _
		"Serveur DHCP", "Bail DHCP - Obtenu", "Bail DHCP - Expiration", "Adresse IP", "Masque", "Passerelle", "Serveur DNS", "Suffixe DNS recherché", "Nom DNS du PC"]
Dim $CtrlButtonNetworkClick = 0, $bUpdateNetworkInfoFunc = False

;Création de la GUI
_NoFocusLines_Global_Set() ;Avant création de la GUI pour tout les contrôles OU _NoFocusLines_Set($ControlID) après création de la GUI
Global $GUIP = GUICreate($TITLE, 380, 45 + (85 * $iNbrNetworkAdapter), $Ini_XPos, $Ini_YPos, $WS_POPUP)
GUISetFont(8.5 * _RatioFont()[0])
GUISetIcon($DIR_IMAGE & "Icon.ico")
GUISetBkColor("0xb6e1f4")
_WinAPI_SetWindowRgn($GUIP, _WinAPI_CreateRoundRectRgn(0, 0, 380, 45 + (85 * $iNbrNetworkAdapter), 15, 15)) ;Arrondis les bords de la GUI

;GUI - Background menu
GUICtrlCreateLabel("", 0, 0, 380, 25)
GUICtrlSetBkColor(-1, "0x5a7edc")
GUICtrlSetState(-1, $GUI_DISABLE)

;GUI - Icon menu
GUICtrlCreatePic('', 3, 1, 23, 23, Default, $GUI_WS_EX_PARENTDRAG)
_SetImage(-1, $DIR_IMAGE & "Icon.ico")

;GUI - Label Title
GUICtrlCreateLabel("  " & $TITLE, 25, 2, 292, 25, Default, $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetFont(-1, 14 * _RatioFont()[0])
GUICtrlSetColor(-1, "0xFFFFFF")

;GUI - Gif refresh
$GifRefresh = _GUICtrlCreateGIF($DIR_IMAGE & "Refresh.GIF", "", 315, 2)
GUICtrlSetTip($GifRefresh, "Actualiser les informations réseaux")
_GIF_PauseAnimation($GifRefresh)
GUICtrlSetOnEvent(-1, "_UpdateNetworkInfo")

;GUI - Button Minimize
$Minimize = GUICtrlCreateLabel("-", 344, -2, 11, 21, BitOR($GUI_SS_DEFAULT_LABEL, $SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetColor(-1, "0xFFFFFF")
GUICtrlSetFont(-1, 24 * _RatioFont()[0])
_GUICtrl_OnHoverRegister($Minimize, "_HoverButtonMinimize", "_HoverButtonMinimize")
GUICtrlSetOnEvent(-1, "_GuiMinimizeFunc")

;GUI - Button Close
$Close = GUICtrlCreateLabel("x", 363, 0, 13, 21, BitOR($GUI_SS_DEFAULT_LABEL, $SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
GUICtrlSetColor(-1, "0xFFFFFF")
GUICtrlSetFont(-1, 18 * _RatioFont()[0])
_GUICtrl_OnHoverRegister($Close, "_HoverButtonClose", "_HoverButtonClose")
GUICtrlSetOnEvent(-1, "_GuiClose")

;GUI - Barre menu
GUICtrlCreateLabel("", -2, 25, 383, 1, Default)
GUICtrlSetBkColor(-1, "0x000000")

;GUI - Button Network + label
_GuiButtonNetwork(True)

;Lancement de la GUI
GUISetOnEvent($GUI_EVENT_CLOSE, "_GuiClose")
GUISetState(@SW_SHOW)
GUIRegisterMsg($WM_MOVE, "_WMGuiMoveSave")
AdlibRegister("_UpdateNetworkInfo", 60000)

;Boucle Principal
While 1
	If $CtrlButtonNetworkClick <> 0 Then _GuiInfos($CtrlButtonNetworkClick)
	Sleep(20)
Wend

Func _GuiButtonNetwork($bCreate = False)
	Local $sColor
	For $i = 0 To $iNbrNetworkAdapter - 1
		Switch $aInfoNetworkAdapter[$i][3]
			Case "Connecté"
				$sColor = "Green"

			Case "Connexion en cours", "Déconnexion en cours", "Média déconnecté", "Authentification", "Authentification réussie", "Authentification échouée", "Adresse invalide", "Informations requises"
				$sColor = "Orange"

			Case "Déconnecté", "Carte désactivé", "Matériel désactivé", "Dysfonctionnement matériel", "N/A"
				$sColor = "Black"
		EndSwitch
		If $bCreate Then
			$aNetworkButtonImg[$i][0] = _3StatPic_Create($DIR_IMAGE & $sColor & "1.png", 5, 40 + (85 * $i), "", $DIR_IMAGE & $sColor & "3.png", "", "", "", Default, Default, $aInfoNetworkAdapter[$i][0], 14, 9, "", 300)
			$aNetworkButtonImg[$i][1] = $sColor
		ElseIf $aNetworkButtonImg[$i][1] <> $sColor Then
			_3StatPic_ChangeImg($aNetworkButtonImg[$i][0], $DIR_IMAGE & $sColor & "1.png", $DIR_IMAGE & $sColor & "1.png", $DIR_IMAGE & $sColor & "3.png")
			$aNetworkButtonImg[$i][1] = $sColor
		EndIf
	Next
EndFunc

Func _ButtonNetworkClick($Ctrl, $iClickMode) ;Fonction "appelé" via l'include GuiConstructor
	;Modifie l'image affiché quand on clique sur un des boutons réseaux
	Local $iCtrlIndex = __3StatPic_hCtrlToIndex($Ctrl)
	If $iCtrlIndex = -1 Then Return 0
	Switch $iClickMode
		Case 1 ; pressed
			_SetImage($Ctrl, $__3SP_Array[$iCtrlIndex][3])
		Case 2 ; released
			_SetImage($Ctrl, $__3SP_Array[$iCtrlIndex][2])
			$CtrlButtonNetworkClick = $Ctrl ;Déclenche la fonction _GuiInfos dans la boucle principal
	EndSwitch
EndFunc ;Il faut quitter cette fonction au plus vite pour éviter les clignotements

Func _GuiInfos($Ctrl)
	$CtrlButtonNetworkClick = 0

	;Identifie la "ligne" du bouton/tableau
	For $i = 0 To $iNbrNetworkAdapter - 1
		If $Ctrl = $aNetworkButtonImg[$i][0] Then
			Local $iLine = $i
			ExitLoop
		EndIf
	Next

	If Not WinExists($aInfoNetworkAdapter[$iLine][0] & " - " & $aInfoNetworkAdapter[$iLine][2]) Then

		;Récupère l'emplacement de la $GUIP si ça position actuelle = -1 (centrer)
		If $Ini_XPos = -1 Or $Ini_YPos = -1 Then
			Local $aWPos = WinGetPos($GUIP)
			If IsArray($aWPos) Then
				IniWrite($INI_SETTING, "Paramètres", "XPos", $aWPos[0])
				IniWrite($INI_SETTING, "Paramètres", "YPos", $aWPos[1])
				$Ini_XPos = $aWPos[0]
				$Ini_YPos = $aWPos[1]
			EndIf
		EndIf

		;Création de la GUI Info
		$aGuiInfo[$iLine][0] = GUICreate($aInfoNetworkAdapter[$iLine][0] & " - " & $aInfoNetworkAdapter[$iLine][2], 530, 440, 384 + (40*$iLine), 3 + (26*$iLine), $WS_POPUP, $WS_EX_MDICHILD, $GUIP)
		GUISetFont(12 * _RatioFont()[0])
		GUISetIcon($DIR_IMAGE & "Icon.ico")
		GUISetBkColor("0xb6e1f4")
		_WinAPI_SetWindowRgn($aGuiInfo[$iLine][0], _WinAPI_CreateRoundRectRgn(0, 0, 530, 440, 15, 15)) ;Arrondis les bords de la GUI

		;Background menu
		GUICtrlCreateLabel("", 0, 0, 530, 25)
		GUICtrlSetBkColor(-1, "0x5a7edc")
		GUICtrlSetState(-1, $GUI_DISABLE)

		;Icon menu
		GUICtrlCreatePic('', 3, 1, 23, 23, Default, $GUI_WS_EX_PARENTDRAG)
		_SetImage(-1, $DIR_IMAGE & "Icon.ico")

		;Label Title
		GUICtrlCreateLabel("  " & $aInfoNetworkAdapter[$iLine][0] & " - " & $aInfoNetworkAdapter[$iLine][2], 25, 2, 470, 25, Default, $GUI_WS_EX_PARENTDRAG)
		GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
		GUICtrlSetFont(-1, 14 * _RatioFont()[0])
		GUICtrlSetColor(-1, "0xFFFFFF")

		;Button Close
		GUICtrlCreateLabel("x", 513, 0, 13, 21, BitOR($GUI_SS_DEFAULT_LABEL, $SS_CENTER, $SS_CENTERIMAGE))
		GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
		GUICtrlSetColor(-1, "0xFFFFFF")
		GUICtrlSetFont(-1, 18 * _RatioFont()[0])
		_GUICtrl_OnHoverRegister(-1, "_HoverButtonClose", "_HoverButtonClose")
		GUICtrlSetOnEvent(-1, "_GuiInfoClose")

		;Barre menu
		GUICtrlCreateLabel("", -2, 25, 533, 1, Default)
		GUICtrlSetBkColor(-1, "0x000000")

		;Création des label + input "déguisé" (permet le copier/coller)
		For $i = 0 To UBound($aGuiInfoLabel) - 1
			GUICtrlCreateLabel($aGuiInfoLabel[$i], 5, 35 + (25 * $i))
			GUICtrlCreateLabel(":", 180, 35 + (25 * $i))
			$ControlIDInput = GUICtrlCreateInput($aInfoNetworkAdapter[$iLine][$i], 195, 35 + (25 * $i), 310, 25, BitOR($GUI_SS_DEFAULT_INPUT, $ES_READONLY), $WS_EX_TOOLWINDOW)
			GUICtrlSetBkColor(-1, "0xb6e1f4")
			$aGuiInfo[$iLine][1] &= $ControlIDInput & "|"
		Next
		$aGuiInfo[$iLine][1] = StringTrimRight($aGuiInfo[$iLine][1], 1)

		;Lancement de la GUI
		GUISetState(@SW_SHOW, $aGuiInfo[$iLine][0])
		GUISetOnEvent($GUI_EVENT_CLOSE, "_GuiInfoClose", $aGuiInfo[$iLine][0])
	Else
		GUISetState(@SW_SHOW, $aGuiInfo[$iLine][0])
		WinActivate($aGuiInfo[$iLine][0])
	EndIf
EndFunc

Func _UpdateNetworkInfo()
	;Signal que la fonction est active et lance le GIF
	$bUpdateNetworkInfoFunc = True
	$hTimer = TimerInit()
	_GIF_ResumeAnimation($GifRefresh)
	Local $iNbrNetworkAdapterNow, $CtrlDataInput

	;Récupération des informations réseaux
	$aInfoNetworkAdapter = _GetNetworkAdapterInfosModif()
	$iNbrNetworkAdapterNow = UBound($aInfoNetworkAdapter)

	;Check si le nombre de cartes réseau n'a pas changé depuis le lancement du programme
	If $iNbrNetworkAdapterNow <> $iNbrNetworkAdapter Then
		MsgBox(48, "Connexions réseau - Changement détecté", "Nous avons détecté des ajouts ou supressions de carte réseaux, " & $TITLE & " va redémarrer afin de mettre à jour les données.", 60, $GUIP)
		_GIF_PauseAnimation($GifRefresh)
		ShellExecute('"' & @ScriptFullPath & '"')
		_GuiClose()
	EndIf

	;Met à jour les boutons en fonction des statuts si besoin est
	_GuiButtonNetwork(False)

	;Met à jour les données dans les GUI Infos existantes
	For $i = 0 To $iNbrNetworkAdapter - 1
		If IsHWnd($aGuiInfo[$i][0]) And WinExists($aGuiInfo[$i][0]) Then
			$CtrlDataInput = StringSplit($aGuiInfo[$i][1], "|", 2)
			For $j = 0 To UBound($CtrlDataInput) - 1
				ControlSetText($aGuiInfo[$i][0], "", "[ID:" & $CtrlDataInput[$j] & "]", $aInfoNetworkAdapter[$i][$j])
			Next
		EndIf
	Next

	;Boucle si inférieur à deux secondes (pour l'animation)
	Do
		Sleep(100)
	Until TimerDiff($hTimer) >= 2000

	;Signal que la fonction est fini
	_GIF_PauseAnimation($GifRefresh)
	$bUpdateNetworkInfoFunc = False
EndFunc

Func _HoverButtonMinimize($Ctrl, $iHoverMode, $hWin)
	Switch $iHoverMode
		Case 1 ; hover
			GUICtrlSetColor($Ctrl, "0xd3d3d3")
		Case 2 ; leave
			GUICtrlSetColor($Ctrl, "0xFFFFFF")
	EndSwitch
EndFunc

Func _HoverButtonClose($Ctrl, $iHoverMode, $hWin)
	Switch $iHoverMode
		Case 1 ; hover
			GUICtrlSetColor($Ctrl, "0xFF0000")
		Case 2 ; leave
			GUICtrlSetColor($Ctrl, "0xFFFFFF")
	EndSwitch
EndFunc

Func _Install()
	;Récupération de valeur de $INI_SETTING (utile en cas de MAJ)
	Local $IniR_XPos = IniRead($INI_SETTING, "Paramètres", "XPos", -1)
	Local $IniR_YPos = IniRead($INI_SETTING, "Paramètres", "YPos", -1)

	;Suppression de donnée
	DirRemove($DIR_INSTALL, 1)
	Sleep(500)

	;Création de dossier
	DirCreate($DIR_INSTALL)
	DirCreate($DIR_DATA)
	DirCreate($DIR_IMAGE)

	;Installation de fichier
	FileInstall(".\Data\Image\Icon.ico", $DIR_IMAGE & "Icon.ico")
	FileInstall(".\Data\Image\Refresh.gif", $DIR_IMAGE & "Refresh.gif")
	FileInstall(".\Data\Image\Black1.png", $DIR_IMAGE & "Black1.png")
	FileInstall(".\Data\Image\Black3.png", $DIR_IMAGE & "Black3.png")
	FileInstall(".\Data\Image\Green1.png", $DIR_IMAGE & "Green1.png")
	FileInstall(".\Data\Image\Green3.png", $DIR_IMAGE & "Green3.png")
	FileInstall(".\Data\Image\Orange1.png", $DIR_IMAGE & "Orange1.png")
	FileInstall(".\Data\Image\Orange3.png", $DIR_IMAGE & "Orange3.png")

	;Restauration de valeur de $INI_SETTING (utile en cas de MAJ)
	IniWrite($INI_SETTING, "Paramètres", "Version", $VERSION)
	IniWrite($INI_SETTING, "Paramètres", "XPos", $IniR_XPos)
	IniWrite($INI_SETTING, "Paramètres", "YPos", $IniR_YPos)
EndFunc

Func _WMGuiMoveSave($hWnd, $nMsg, $wParam, $lParam)
	#forceref $hWnd, $nMsg, $wParam, $lParam

	;Enregistre l'emplacement de la $GUIP
	If $hWnd <> $GUIP Or StringRegExp($lParam, '(83008300)') Then Return $GUI_RUNDEFMSG ;83008300 correspond à "l'emplacement" quand on réduit la GUI (donc emplacement à ne pas enregistrer).
    Local $aWPos = WinGetPos($hWnd)
    If IsArray($aWPos) Then
		IniWrite($INI_SETTING, "Paramètres", "XPos", $aWPos[0])
		IniWrite($INI_SETTING, "Paramètres", "YPos", $aWPos[1])
		$Ini_XPos = $aWPos[0]
        $Ini_YPos = $aWPos[1]
    EndIf
	Return $GUI_RUNDEFMSG
EndFunc

Func _GuiMinimizeFunc()
	GUISetState(@SW_MINIMIZE, @GUI_WinHandle)
EndFunc

Func _GuiClose()
	If IsDeclared("GUIP") Then
		AdlibUnRegister("_UpdateNetworkInfo")
		GUISetState(@SW_HIDE, $GUIP)
		For $i = 0 To UBound($aGuiInfo) - 1
			If IsHWnd($aGuiInfo[$i][0]) And WinExists($aGuiInfo[$i][0]) Then GUISetState(@SW_HIDE, $aGuiInfo[$i][0])
		Next
	EndIf
	_NoFocusLines_Global_Exit()
	Exit
EndFunc

Func _GuiInfoClose()
	GUISetState(@SW_HIDE, @GUI_WinHandle)
EndFunc

Func _Quit()
	#cs
	If Not $bPIDExit Then IniWrite($INI_SETTING, "Paramètres", "PID", "null")
	#ce
EndFunc

;Amélioration possible:
;Modifier la gui "en live" si changement de nombre de carte/connexion réseau.
