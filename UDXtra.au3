#NoTrayIcon
#include <Array.au3>
#include <TrayConstants.au3> ; Required for the $TRAY_ICONSTATE_SHOW constant.

AutoItSetOption("MustDeclareVars", 1)
Global $settingSearchUD = 1
Global $settingChromeCase = 1

Func get_token($aToken)
   $aToken = StringReplace ($aToken, " ", "")
   $aToken = StringReplace ($aToken, "-", "")
   $aToken = StringReplace ($aToken, "(", "")
   $aToken = StringReplace ($aToken, ")", "")
   $aToken = StringReplace ($aToken, @CR, "")
   $aToken = StringReplace ($aToken, @LF, "")
   If (StringIsDigit ($aToken)) Then
	  Return $aToken
   Else
	  Return ""
   EndIf
EndFunc

Func search_UD()
   If Not $settingSearchUD Then
	  Return
   EndIF

   Local $aPos = MouseGetPos()
   Local $aType = ""
   Local $Location[1]
   _ArrayPop ($Location)
   Local $hUD = WinGetHandle ("Unified Desktop for IP phones")
   Local $aToken = get_token (ClipGet ())
   Local $allLocations[10] = ["Vancouver", "Calgary", "Kelowna", "Edmonton", "Victoria", "Winnipeg", "Manitoba", "Saskatoon", "Fort McMurray", "Shaw Direct"]
   Local $delay = 200

   If Not (StringLen ($aToken)) Then
	  WinActivate("Unified Desktop for IP phones")
	  WinWaitActive("Unified Desktop for IP phones")
	  Send("^c")
	  $aToken = get_token (ClipGet ())
		 If Not (StringLen ($aToken)) Then
		 Send("{HOME}+{END}^c")
		 $aToken = get_token (ClipGet ())
	  EndIf
   EndIf

   If Not (StringLen ($aToken)) Then
		 Return
	  Else
		 If (StringLen ($aToken) = 11) Then
			$aType = "Account Number"
		 ElseIf (StringLen ($aToken) = 10) Then
			$aType = "Customer Phone Number"
		 Else
			Return
		 EndIf
   EndIf
   MouseClick ("main", 655, 79, 1, 1)
   Sleep(100)
   If $aType = "Account Number" Then
	  Local $aPrefix = Int (StringLeft ($aToken, 3))
	  Switch $aPrefix
	  Case 003, 005
		 _ArrayAdd ($Location, "Victoria")
	  Case 009, 012
		 _ArrayAdd ($Location, "Victoria")
		 _ArrayAdd ($Location, "Vancouver")
	  Case 010, 011, 014
		 _ArrayAdd ($Location, "Vancouver")
	  Case 013
		 _ArrayAdd ($Location, "Vancouver")
		 _ArrayAdd ($Location, "Kelowna")
	  Case 016, 017, 018, 019
		 _ArrayAdd ($Location, "Kelowna")
	  Case 031
		 _ArrayAdd ($Location, "Kelowna")
		 _ArrayAdd ($Location, "Calgary")
	  Case 026, 028, 032
		 _ArrayAdd ($Location, "Calgary")
	  Case 065
		 _ArrayAdd ($Location, "Calgary")
		 _ArrayAdd ($Location, "Edmonton")
	  Case 027, 029, 030
		 _ArrayAdd ($Location, "Edmonton")
	  Case 033
		 _ArrayAdd ($Location, "Fort McMurray")
	  Case 055
		 _ArrayAdd ($Location, "Saskatchewan")
	  Case 039
		 _ArrayAdd ($Location, "Winnipeg")
	  Case 038
		 _ArrayAdd ($Location, "Manitoba")
	  EndSwitch
   ElseIf $aType = "Customer Phone Number" Then
	  Local $areaCode = Int (StringLeft ($aToken, 3))
	  Switch $areaCode
	  Case 403, 587, 780
		 _ArrayAdd ($Location, "Calgary")
		 _ArrayAdd ($Location, "Edmonton")
		 _ArrayAdd ($Location, "Fort McMurray")
	  Case 236, 250, 604
		 _ArrayAdd ($Location, "Vancouver")
		 _ArrayAdd ($Location, "Kelowna")
		 _ArrayAdd ($Location, "Victoria")
	  Case 204, 431
		 _ArrayAdd ($Location, "Winnipeg")
		 _ArrayAdd ($Location, "Manitoba")
	  Case 226, 249, 289, 343, 365, 416, 437, 519, 613, 647, 705, 807, 905
		 _ArrayAdd ($Location, "Manitoba")
	  Case 306, 639
		 _ArrayAdd ($Location, "Saskatoon")
	  EndSwitch
   EndIf

   For $loc In $allLocations
	  If (_ArraySearch ($Location, $loc) < 0) Then
		 _ArrayAdd ($Location, $loc)
	  EndIf
   Next

   For $loc In $Location
	  If StringLen($loc) = 0 Then
		 Break
	  EndIf
	  MouseClick ("main", 115, 115, 2, 1)
	  Sleep($delay)
	  Send ($loc)
	  Sleep($delay)
	  Send ("{TAB}")
	  Send ($aType)
	  Sleep($delay)
	  Send ("{TAB}")
	  Sleep($delay)
	  Send("^a")
	  Send($aToken)
	  Sleep($delay)
	  Send ("{ENTER}")
	  Do
		 Sleep(100)
	  Until PixelGetColor (525, 391, $hUD) <> 0xCCCCCC
	  If (PixelGetColor (33, 167, $hUD) <> 16714752) Or (WinExists ("Error Encountered in Unified Desktop")) Then
		 ExitLoop
	  EndIf
   Next
EndFunc

Func search_case()
   If Not $settingChromeCase Then
	  Return
   EndIF
   Local $proc="iexplore.exe"
   Local $strComputer="."
   Local $oWMI=ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & $strComputer & "\root\cimv2")
   Local $oProcessColl=$oWMI.ExecQuery("Select * from Win32_Process where Name= " & '"'& $Proc & '"')
   Local $Url = ""
   For $Process In $oProcessColl
	  Local $Cmd=$Process.Commandline
	  If StringLeft ($Cmd, 148) = '"C:\Program Files\Internet Explorer\iexplore.exe" -noframemerging https://shawprod.service-now.com/sn_customerservice_redirect.do?sysparm_accountid=' Then
		 $Url = StringSplit ($Cmd, "noframemerging ", 3)[1]
		 $Process.Terminate()
	  EndIf
   Next
   If StringLen ($Url) Then
	  ShellExecute ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", $Url)
   EndIf
EndFunc

HotKeySet ("^!s", "search_UD")

Opt("TrayMenuMode", 3)
Local $iSettings = TrayCreateMenu("Settings") ; Create a tray menu sub menu with two sub items.
Local $iSearchUD = TrayCreateItem("Search UD by 'Ctrl+Alt+S", $iSettings)
TrayItemSetState($iSearchUD, $TRAY_CHECKED)
Local $iChromeCase = TrayCreateItem("Open CaseMgmt in Chrome", $iSettings)
TrayItemSetState($iChromeCase, $TRAY_CHECKED)
TrayCreateItem("") ; Create a separator line.

Local $idAbout = TrayCreateItem("About")
TrayCreateItem("") ; Create a separator line.

Local $idExit = TrayCreateItem("Exit")

TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.

While 1
   If $settingChromeCase Then
	  search_case()
   EndIf
   ;Sleep (250)
   Switch TrayGetMsg()
      Case $idAbout ; Display a message box about UDXtra
         MsgBox($MB_SYSTEMMODAL, "", "UDXtra" & @CRLF & @CRLF & _
            "Version: 0.1" & @CRLF & _
			"Extra features for Unified Desktop" & @CRLF & _
			"(c) Denis Abakumov.")
	  Case $iSearchUD
		 $settingSearchUD = Not $settingSearchUD
		 If $settingSearchUD Then
			TrayItemSetState($iSearchUD, $TRAY_CHECKED)
		 Else
			TrayItemSetState($iSearchUD, $TRAY_UNCHECKED)
		 EndIf
	  Case $iChromeCase
		 $settingChromeCase = Not $settingChromeCase
		 If $settingChromeCase Then
			TrayItemSetState($iChromeCase, $TRAY_CHECKED)
		 Else
			TrayItemSetState($iChromeCase, $TRAY_UNCHECKED)
		 EndIf
      Case $idExit ; Exit
         Exit
   EndSwitch

WEnd
