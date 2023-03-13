#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>



Global $ftp
Global $ftp_old
Global $apppath
selectbuild()
copyinstallerfromftp()
installapplication()

Func selectbuild()
	$hgui = GUICreate("EM Installation: Enter build Version", 520, 190)
	GUICtrlSetState(-1, $gui_checked)
	GUICtrlCreateLabel("For Macchia type e3 in release box and click OK button and for E2 trunk Latest build", 20, 10, 500, 18)
	GUICtrlSetFont(-1, 9, 2000)
    GUICtrlCreateLabel("left everything blank and click OK button this will automatically install the latest trunk build", 20, 25, 500, 18)
	GUICtrlSetFont(-1, 9, 2000)
	GUICtrlCreateLabel("For Release mention release number including rc like 'rcx.y' format  >>>", 20, 60, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("(Example: rc7.1) if left blank then this will install trunk build", 20, 75, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("Mention here which build number you want to install (Example: 56118) >>>", 20, 115, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("(if left blank then this will install latest build release build version)", 20, 130, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlSetState(-1, $gui_dropaccepted)
	$build = GUICtrlCreateInput("", 440, 60, 45, 18)
	$version = GUICtrlCreateInput("", 440, 115, 45, 18)
	$hbutton = GUICtrlCreateButton("OK", 210, 150, 90, 30)
	GUISetState(@SW_SHOW)
	While 1
		Switch GUIGetMsg()
			Case $gui_event_close
				Exit
			Case $hbutton
				$ftp = ""
				$ftp_old = ""
				$build = StringLower(GUICtrlRead($build))
				$version = StringLower(GUICtrlRead($version))
				ConsoleWrite($build)
				ConsoleWrite($version)
				If (StringLen($build) == 0) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\test\Event Master Toolset Rev"
					WinClose("Select Build")
					ExitLoop
			   ElseIf (StringCompare($build, "e3",0 ) == 0) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\Macchia\bundles\development\Windows\Macchia Toolset Rev 1.0.0 (Build"
					 WinClose("Select Build")
					ExitLoop
				 ElseIf (StringCompare($build, "rc",0 ) > 0) Then
					If (StringLen($version) == 0) Then
						$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\" & $build & "\Event Master Toolset Rev"
						;WinClose("Select Build")
					 Else
						$ftp = "\\10.34.40.5\share1\EventMaster\E2\release\" & $build & "\old\Event Master Toolset Rev " & StringRight($build,3) & " (Build " & $version
					 EndIf
					 WinClose("Select Build")
					ExitLoop
			   ;ElseIf (StringLen($version) > 0) Then
				;	ConsoleWrite(StringRight($build,3))
					;Local $c = StringSplit($build,"rc")[0]
					;$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\rc" & $b & "\Event Master Toolset Rev " & $b & " (Build " & $version
				;	$ftp = "\\10.34.40.5\share1\EventMaster\E2\release\" & $build & "\old\Event Master Toolset Rev " & StringRight($build,3) & " (Build " & $version
				;	WinClose("Select Build")
				;	ExitLoop
			   Else
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\test\Event Master Toolset Rev"
			   EndIf
			   ExitLoop
		EndSwitch
	WEnd
EndFunc

Func copyinstallerfromftp()
	ConsoleWrite($ftp)
	$fileloc = @TempDir & "\InstallerFiles\"
	ConsoleWrite($ftp & "*Setup.exe")
	ConsoleWrite($fileloc)
	FileDelete($fileloc)
	MsgBox($mb_systemmodal, "Info", "Copying Build...", 2)
	If (FileCopy($ftp & "*Setup.exe", $fileloc, $fc_overwrite + $fc_createpath) == 0) Then
		 ConsoleWrite($ftp)
		 If (FileCopy($ftp_old & "*Setup.exe", $fileloc, $fc_overwrite + $fc_createpath) == 0) Then
			ConsoleWrite($ftp_old)
			MsgBox($mb_systemmodal, "Info", "Build Number Not available" & @CRLF & " in Pool folder location", 7)
			Exit
		EndIf
	EndIf
	$apppath = _filelisttoarray(@TempDir & "\InstallerFiles\", Default, Default, True)
	MsgBox($mb_systemmodal, "Info", "Opening Build location" & @CRLF & $apppath[1], 2)
EndFunc

Func installapplication()
	MsgBox($mb_systemmodal, "Info", "Build Installation is in progress...", 3)
	RunWait($apppath[1] & " --mode unattended")
	ShellExecute("C:\Barco")
	MsgBox($mb_systemmodal, "Info", "                            Installation Completed !!" & @CRLF & "           Timeout after 15 seconds or Click the OK button", 15)
EndFunc
