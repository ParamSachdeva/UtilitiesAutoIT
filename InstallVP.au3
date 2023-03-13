#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#include <IE.au3>


Global $ips
Global $build
Global $version
Global $oie
Global $ftp
Global $ftp = ""
Global $ftp_old = ""
Global $fileloc = @TempDir & "\InstallerFiles\"
Global $vpbuild
selectbuild()
copyinstallerfromftp()
updatevp($ips)

Func updatevp($ips)
	Opt("WinTitleMatchMode", 2)
	$wintitle = "Choose File to Upload"
	$ietitle = "Web App"
	MsgBox($mb_systemmodal, "Info", "Update VP on Hardware Intiating....", 2)
	$multipleips = StringSplit($ips, ";")
	For $i = 1 To $multipleips[0]
		$oie = _iecreate($multipleips[$i])
		$hwnd = _iepropertyget($oie, "hwnd")
		WinSetState($hwnd, "", @SW_MAXIMIZE)
		_ieaction($oie, "visible")
		_ieloadwait($oie)
		clickobjectbyname("Tools")
		clickobjectbyname("Manage Software")
		Sleep(2000)
		$fileupload = _iegetobjbyid($oie, "fileupload")
		_ieaction($fileupload, "focus")
		Sleep(3000)
		Send("{SPACE}")
		Sleep(3000)
		WinWaitActive($wintitle, "", 30)
		WinActivate($wintitle, "")
		Sleep(1000)
		ControlSend($wintitle, "", "[CLASS:Edit; INSTANCE:1]", $vpbuild)
		Sleep(2000)
		ControlClick($wintitle, "", "[CLASS:Button; INSTANCE:1]")
		Sleep(3000)
		WinActivate($ietitle, "")
		Sleep(2000)
		Local $apos = WinGetPos($ietitle)
		MouseClick("left", $apos[2] - 1185, $apos[3] - 160, 1)
		Sleep(10000)
		$hbmp = _screencapture_capture("")
		_screencapture_capture($fileloc & "/" & $multipleips[$i] & "_" & @MDAY & @MON & @YEAR & "_" & @HOUR & @MIN & @SEC & ".jpg", 0, 0, 3000, 1500)
		MsgBox($mb_systemmodal, "Info", "Refer File for status: " & $fileloc & "/" & $multipleips[$i] & "_" & @MDAY & @MON & @YEAR & "_" & @HOUR & @MIN & @SEC & ".jpg", 2)
		MsgBox($mb_systemmodal, "Info", "File uploaded , process will be done automatically in backgrund....", 2)
	Next
	While (WinExists($ietitle, ""))
		WinClose($ietitle)
	WEnd
	ShellExecute($fileloc)
EndFunc

Func clickobjectbyname($smystring)
	$smystring
	$olinks = _ielinkgetcollection($oie)
	For $olink In $olinks
		ConsoleWrite("InLoopName...####")
		Local $slinktext = _iepropertyget($olink, "innerText")
		If StringInStr($slinktext, $smystring) Then
			_ieaction($olink, "click")
			ExitLoop
		EndIf
	Next
EndFunc

Func clickobjectbyidtest($smystring)
	$smystring
	$olinks = _ielinkgetcollection($oie)
	For $olink In $olinks
		ConsoleWrite("InLoopID...####")
		Local $slinktext = _iepropertyget($olink, "id")
		If StringInStr($slinktext, $smystring) Then
			_ieaction($olink, "click")
			ExitLoop
		EndIf
	Next
EndFunc

Func clickobjectbyid(ByRef $oobject, $saction)
	If NOT IsObj($oobject) Then
		__ieconsolewriteerror("Error", "_IEAction(" & $saction & ")", "$_IESTATUS_InvalidDataType")
		Return SetError($_iestatus_invaliddatatype, 1, 0)
	EndIf
	$saction = StringLower($saction)
	Select
		Case $saction = "click"
			If __ieisobjtype($oobject, "documentContainer") Then
				__ieconsolewriteerror("Error", "_IEAction(click)", " $_IESTATUS_InvalidObjectType")
			EndIf
			Sleep(3000)
			$oobject.click()
			ConsoleWrite("Now")
			Exit
	EndSelect
EndFunc

Func selectbuild()
	$hgui = GUICreate("Update VP for Real Hardware", 520, 250)
	GUICtrlSetState(-1, $gui_checked)
	GUICtrlCreateLabel("Enter Real Hardware IP, Use semicolon(;) to provide multiple IP's ", 30, 10, 500, 18)
	GUICtrlSetFont(-1, 11, 2000)
	GUICtrlCreateLabel("For Trunk Latest build left everything blank and click OK button this will automatically install the latest trunk build", 20, 60, 500, 18)
	GUICtrlSetFont(-1, 9, 2000)
	GUICtrlCreateLabel("install the latest trunk build of VP", 20, 75, 500, 18)
	GUICtrlSetFont(-1, 9, 2000)
	GUICtrlCreateLabel("For Release mention release number here in X.Y format (Example:7.1) >>>", 20, 110, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("(if left blank then this will install trunk build)", 20, 125, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("Mention here which build number you want to install (Example: 56118) >>>", 20, 155, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("(if left blank then this will install latest build version)", 20, 170, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlSetState(-1, $gui_dropaccepted)
	$ips = GUICtrlCreateInput("", 150, 33, 250, 18)
	$build = GUICtrlCreateInput("", 440, 110, 30, 18)
	$version = GUICtrlCreateInput("", 440, 155, 45, 18)
	$hbutton = GUICtrlCreateButton("OK", 210, 210, 90, 30)
	GUICtrlSetFont(-1, 9, 700)
	GUISetState(@SW_SHOW)
	While 1
		Switch GUIGetMsg()
			Case $gui_event_close
				Exit
			Case $hbutton
				$ips = StringLower(GUICtrlRead($ips))
				$build = StringLower(GUICtrlRead($build))
				$version = StringLower(GUICtrlRead($version))
				If (StringIsSpace($build)) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\test\em_update_vp"
					WinClose("Select Build")
					ExitLoop
				ElseIf (StringIsSpace($version)) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\rc" & $build & "\em_update_vp"
					WinClose("Select Build")
					ExitLoop
				Else
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\rc" & $build & "\em_update_vp." & $build & "." & $version
					$ftp_old = "\\saceng-nas01\Share1\EventMaster\E2\release\rc" & $build & "\old\em_update_vp." & $build & "." & $version
				EndIf
				ExitLoop
		EndSwitch
	WEnd
	MsgBox(1, "Provided list of IPs: ", $ips, 2)
EndFunc

Func copyinstallerfromftp()
	ConsoleWrite($ftp)
	$fileloc = @TempDir & "\InstallerFiles\"
	ConsoleWrite($ftp & "*.tar.gz")
	ConsoleWrite($fileloc)
	FileDelete($fileloc)
	MsgBox($mb_systemmodal, "Info", "Copying Build...", 2)
	If (FileCopy($ftp & "*.tar.gz", $fileloc, $fc_nooverwrite + $fc_createpath) == 0) Then
		MsgBox($mb_systemmodal, "Info", "Copying failed due to slow network issue...", 2)
		If (FileCopy($ftp_old & "*.tar.gz", $fileloc, $fc_nooverwrite + $fc_createpath) == 0) Then
			MsgBox($mb_systemmodal, "Info", "Build Number Not available" & @CRLF & " in Pool folder location: " & $ftp_old, 7)
			Exit
		EndIf
	EndIf
	$apppath = _filelisttoarray(@TempDir & "\InstallerFiles\", Default, Default, True)
	$vpbuild = $apppath[1]
	MsgBox($mb_systemmodal, "Info", "Opening Build location" & @CRLF & $apppath[1], 2)
EndFunc
