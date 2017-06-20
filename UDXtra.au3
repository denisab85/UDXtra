#NoTrayIcon
#include <Array.au3>
#include <TrayConstants.au3> ; Required for the $TRAY_ICONSTATE_SHOW constant.
#include <Misc.au3>
#include <Date.au3>

#pragma compile(ProductVersion, 0.1)
;#pragma compile(Console, True)

AutoItSetOption("MustDeclareVars", 1)
Global $settingSearchUD = 1
Global $settingChromeCase = 1
Global $settingTruncateRf = 1


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
   If Not $settingSearchUD Then Return
   If Not $settingSearchUD Then Return
   If Not WinExists ("Unified Desktop") Then Return
   WinActivate ("Unified Desktop")
   Local $aType = ""
   Local $Location[1]
   _ArrayPop ($Location)
   Local $hUD = WinGetHandle ("Unified Desktop for IP phones")
   Local $aToken = get_token (ClipGet ())
   Local $allLocations[10] = ["Vancouver", "Calgary", "Kelowna", "Edmonton", "Victoria", "Winnipeg", "Manitoba", "Saskatoon", "Fort McMurray", "Shaw Direct"]
   Local $delay = 300

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
   Local $winPos = WinGetPos("Unified Desktop")
   Local $winLeft = $winPos[0]
   Local $winTop = $winPos[1]
   Local $winWidth = $winPos[2]

   Local $x = ($winWidth - 220) / 2 + $winLeft + 44
   Local $y = $winTop + 90
   MouseClick ("main", $x, $y, 1, 1)
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

   Local $n = 0
   ;_ArrayDisplay ($Location)
   For $loc In $Location
	  If _IsPressed("1B") Then Return
	  If StringLen($loc) = 0 Then Break
	  MouseClick ("main", $winLeft+120, $winTop+120, 2, 1)
	  Sleep($delay)
	  Send($loc)
	  Sleep($delay)
	  Send("{TAB}")
	  If ($n = 0) Then
		 Send ($aType)
		 Sleep($delay)
	  EndIf
	  Send ("{TAB}")
	  If ($n = 0) Then
		 Sleep($delay)
		 Send("^a")
		 Send($aToken)
		 Sleep($delay)
	  EndIf
	  Send ("{ENTER}")
	  While PixelGetColor ($winLeft+5, $winTop+399, $hUD) = 0xCCCCCC
		 ConsoleWrite ("Progress window detected. Waiting another " & String ($delay) & " ms.")
		 Sleep($delay)
	  WEnd
	  Sleep($delay)
	  Sleep($delay)
	  Local $pixColor = PixelGetColor ($winLeft+41, $winTop+175, $hUD) ; This pixel should be red-ish
	  If ($pixColor < 16711680) Or ($pixColor > 16721960) Or (WinExists ("Error Encountered in Unified Desktop")) Then
		 ConsoleWrite ("PixelGetColor (" & String($winLeft+41) & ", " & String($winTop+175) & ") = " & String(PixelGetColor ($winLeft+41, $winTop+175, $hUD)) & @CRLF)
		 ExitLoop
	  EndIf
	  $n += 1
   Next
   ClipPut("")
EndFunc


Func chrome_case()
   Local $result = False
   Local $strProc="iexplore.exe"
   If Not $settingChromeCase Or Not ProcessExists ($strProc) Then Return
   If _IsPressed (12) Then Return
   Local $strComputer="."
   Local $oWMI=ObjGet("winmgmts:{impersonationLevel=impersonate}!\\" & $strComputer & "\root\cimv2")
   Local $oProcessColl=$oWMI.ExecQuery("Select * from Win32_Process where Name= " & '"'& $strProc & '"')
   Local $Url = ""
   For $Process In $oProcessColl
	  Local $Cmd=$Process.Commandline
	  If StringLeft ($Cmd, 148) = '"C:\Program Files\Internet Explorer\iexplore.exe" -noframemerging https://shawprod.service-now.com/sn_customerservice_redirect.do?sysparm_accountid=' Then
		 If Not $result Then
			$result = True
			Local $winPos = WinGetPos("Unified Desktop")
			Local $winLeft = $winPos[0]
			Local $winTop = $winPos[1]
			Local $winWidth = $winPos[2]
			Local $aMousePos = MouseGetPos()
			Local $clip = ClipGet()
			ClipPut("")
			MouseClick ("main", $winLeft + $winWidth - 72, $winTop + 127, 4, 1)
			Sleep(100)
			Send("^c")
			Sleep(100)
			Local $calabrioID = ClipGet()
			If StringIsDigit($calabrioID) And StringLen($calabrioID) = 8 Then
			   MouseClick ("main", $winLeft + $winWidth - 72, $winTop + 595, 1, 1)
			   Sleep(100)
			   Send("^{END}" & @CRLF & _NowTime() & @CRLF & "Calabrio ID: ")
			   Send("^v")
			   MouseClick ("main", $winLeft + 63, $winTop + 275, 1, 1)
			EndIf
			ClipPut($clip)
			MouseMove ($aMousePos[0], $aMousePos[1], 1)
		 EndIf
		 $Url = StringSplit ($Cmd, "noframemerging ", 3)[1]
		 If StringLen ($Url) Then ShellExecute ("C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", $Url)
		 $Process.Terminate()
	  EndIf
   Next
   Return $result
EndFunc


Func self_update()
   SplashTextOn("HLC: Checking for HLC Update", "Please wait... This may take a few moments.", "400", "100", "-1", "-1", 50, "", "", "")
   If Not IsDeclared("$complete_hlcfile_master_data_path") Then
	  Global $complete_hlcfile_master_data_path = $site_master_data_path & "_HLC"
	  If FileGetVersion($complete_hlcfile_master_data_path & "\hlc.exe") <> FileGetVersion("C:\hes\locationchange\hlc.exe") Then FileCopy($complete_hlcfile_master_data_path & "\hlc.exe", @ScriptFullPath & ".new") ;changed the name of the new file.
	  Local $batchPath = @ScriptDir & '\hlc_update.bat'
	  Local $batchFile =  "@echo off"& @CRLF _
                            "ping localhost -n 2 > nul" & @CRLF _ ;not sure what you're doing here. Giving the script time to exit?
                            ":loop" & @CRLF _ ;specify the start of a zone
                            'del /Q "' & @ScriptFullPath & '"' & @CRLF _ ;the quotes are needed for long filepaths, and filepaths with spaces. The @SciptfullPath is for flexibility
                            'if exist "' & @ScriptFullPath & '" goto loop' & @CRLF _ ;if the delete failed, try again
                            'move "' & @ScriptFullPath & '.new" "' & @ScriptFullPath & '"' & @CRLF _ ;this is why I changed the new file's name.
                            'start "' & @ScriptFullPath & '"' & @CRLF _
                            'del /Q "' & $batchPath & '"' & @CRLF _
                            "exit"
	  FileWrite($batchPath,$batchFile)
	  Run($batchPath, "", @SW_HIDE)
	  SplashOff()
	  Exit
   Else
	  SplashOff()
   EndIf
EndFunc


Func get_min_max ($list)
   Local $split = StringSplit($list, ",", 2)
   Local $min = _ArrayMin($split, 1)
   Local $max = _ArrayMax($split, 1)
   If ($min = $max) Then
	  return $min
   Else
	  return ($min & "," & $max)
   EndIf
EndFunc

Func truncate_rf ($rf)
   Local $regExpFull = ""
   Local $regExpLong = '^(RF|INET|STB|ESTB|DP) RX((-?\d{1,2}(\.\d)?)(,-?\d{1,2}(\.\d)?)*)/TX((-?\d{1,2}(\.\d)?)(,-?\d{1,2}(\.\d)?)*) (RF|INET|STB|ESTB|DP) DSNR ((-?\d{1,2}(\.\d)?)(,-?\d{1,2}(\.\d)?)*) USNR ((-?\d{1,2}(\.\d)?)(,-?\d{1,2}(\.\d)?)*)'
   Local $aArray
   Local $result = ""
   Local $type = ""
   Local $rx = ""
   Local $tx = ""
   Local $dsnr = ""
   Local $usnr = ""

   If Not StringRegExp($rf, $regExpLong) Then
	  $regExpFull = "(Upstream Signal/Noise Ratio ((#\d{1,2} \(dB\):)|(\(dB\) #(\d{1,2})))\s+(-?\d{1,2}(\.\d)?))+"
	  $aArray = StringRegExp($rf, $regExpFull, 4)
	  If Not @error Then
		 For $match In $aArray
			;_ArrayDisplay($match)
			If (StringLen($usnr)) Then $usnr = $usnr & ","
			$usnr = $usnr & $match[6]
		 Next
	  EndIf
	  $regExpFull = "(Downstream Signal/Noise Ratio ((#\d{1,2} \(dB\):)|(\(dB\) #(\d{1,2})))\s+(-?\d{1,2}(\.\d)?))+"
	  $aArray = StringRegExp($rf, $regExpFull, 4)
	  If Not @error Then
		 For $match In $aArray
			;_ArrayDisplay($match)
			If (StringLen($dsnr)) Then $dsnr = $dsnr & ","
			$dsnr = $dsnr & $match[6]
		 Next
	  EndIf
	  $regExpFull = "(Receive Level(( #\d{1,2} \(dB\):)|(\(dB\) #(\d{1,2})))\s+(-?\d{1,2}(\.\d)?))+"
	  Local $aArray = StringRegExp($rf, $regExpFull, 4)
	  If Not @error Then
		 For $match In $aArray
			;_ArrayDisplay($match)
			If (StringLen($rx)) Then $rx = $rx & ","
			$rx = $rx & $match[6]
		 Next
	  EndIf
	  $regExpFull = "(Transmit Level(( #\d{1,2} \(dB\):)|(\(dB\) #(\d{1,2})))\s+(-?\d{1,2}(\.\d)?))+"
	  Local $aArray = StringRegExp($rf, $regExpFull, 4)
	  If Not @error Then
		 For $match In $aArray
			;_ArrayDisplay($match)
			If (StringLen($tx)) Then $tx = $tx & ","
			$tx = $tx & $match[6]
		 Next
	  EndIf

	  Select
		 Case StringInStr($rf, "Chelsea")
			$type = "STB"
		 Case StringInStr($rf, "CGNM_") Or StringInStr($rf, "D30GW-EAGLE") Or StringInStr($rf, "dpc3825")
            $type = "INET"
		 Case StringInStr($rf, "TS060151D")
			$type = "DP"
		 Case StringInStr($rf, "7.14.")
			$type = "ESTB"
		 Case Else ; If nothing matches
            $type = "RF"
	  EndSelect

	  If (StringLen($rx) And StringLen($tx) And StringLen($dsnr) And StringLen($usnr)) Then
		 $rf = $type & " " & "RX" & $rx & "/TX" & $tx & " " & $type & " DSNR " & $dsnr & " USNR " & $usnr
	  EndIf
   EndIf

   $aArray = StringRegExp($rf, $regExpLong, 1)
   If Not @error Then
	  ;_ArrayDisplay($aArray)
	  $type = $aArray[0]
	  $rx = $aArray[1]
	  $tx = $aArray[6]
	  $dsnr = $aArray[12]
	  $usnr = $aArray[17]
	  $rx = get_min_max ($rx)
	  $tx = get_min_max ($tx)
	  $dsnr = get_min_max ($dsnr)
	  $usnr = get_min_max ($usnr)
	  $result = $type & " " & $rx & "/" & $tx & " SNR " & $dsnr & "/" & $usnr
	  ClipPut ($result)
   EndIf
   Return $result
EndFunc


HotKeySet ("^!s", "search_UD")

Opt("TrayMenuMode", 1)
Local $iSettings = TrayCreateMenu("Settings") ; Create a tray menu sub menu with two sub items.

Local $iSearchUD = TrayCreateItem("Search UD by 'Ctrl+Alt+S", $iSettings)
TrayItemSetState($iSearchUD, $TRAY_CHECKED)
Local $iChromeCase = TrayCreateItem("Open CaseMgmt in Chrome", $iSettings)
TrayItemSetState($iChromeCase, $TRAY_CHECKED)
Local $iTruncateRf = TrayCreateItem("Truncate RF in clipboard", $iSettings)
TrayItemSetState($iTruncateRf, $TRAY_CHECKED)

TrayCreateItem("") ; Create a separator line.

Local $idAbout = TrayCreateItem("About")
TrayCreateItem("") ; Create a separator line.

Local $idExit = TrayCreateItem("Exit")

TraySetState($TRAY_ICONSTATE_SHOW) ; Show the tray menu.

Local $hTimer = TimerInit() ; Begin the timer and store the handle in a variable.
Local $calabrioID = ""
While 1
   If TimerDiff($hTimer) > 500 Then
	  $hTimer = TimerInit()
	  If $settingChromeCase Then
		 chrome_case()
	  EndIf
	  If $settingTruncateRf Then truncate_rf (ClipGet())
   EndIf

   Local $tMsg = TrayGetMsg()
   Switch $tMsg
      Case $idAbout ; Display a message box about UDXtra
         MsgBox($MB_SYSTEMMODAL, "About UDXtra", "UDXtra" & @CRLF & @CRLF & _
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
	  Case $iTruncateRf
		 $settingTruncateRf = Not $settingTruncateRf
		 If $settingTruncateRf Then
			TrayItemSetState($iTruncateRf, $TRAY_CHECKED)
		 Else
			TrayItemSetState($iTruncateRf, $TRAY_UNCHECKED)
		 EndIf
      Case $idExit ; Exit
         Exit
   EndSwitch
WEnd
