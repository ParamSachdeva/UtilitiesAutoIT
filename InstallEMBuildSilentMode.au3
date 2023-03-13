;#RequireAdmin
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>
#include <File.au3>
#include <AutoItConstants.au3>


;#cs
Global $ftp
Global $ftp_old
Global $apppath
Global $ftp = ""
Global $ftp_old = ""
Global $build
Global $version
Global $path = "C:\Automation\TestSetup\Installer"

DirRemove($path, $DIR_DEFAULT)
DirRemove($path, $DIR_REMOVE)
FileDelete($path)
DirCreate($path)
;selectbuild()


If $CmdLine[0] > 0 Then
			   $build = StringLower($CmdLine[1])
			   ConsoleWrite($build)
			   If (StringCompare($build, "trunk",0 ) == 0) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\test\Event Master Toolset Rev"
					copyinstallerfromftp()
				    installapplication()
			   ElseIf (StringCompare($build, "rc",0) > 0) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\" & $build & "\Event Master Toolset Rev"
					copyinstallerfromftp()
				    installapplication()
			   ElseIf (StringCompare($build, "e3",0 ) == 0) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\Macchia\bundles\development\Windows\Macchia Toolset Rev 1.0.0 (Build"
					copyinstallerfromftp()
				    installapplication()
			    ElseIf (StringCompare($build, "mac",0 ) == 0) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\test\MAC\*osx-installer.tar.gz"
					copyinstallerfromftp()
				    ;installapplication()
			   Else
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\test\Event Master Toolset Rev"
					copyinstallerfromftp()
				    installapplication()
			   EndIf


			   If ($CmdLine[0] > 1) Then
				  $version =  StringLower($CmdLine[2])
				  If (StringIsSpace($version)) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\rc" & $build & "\Event Master Toolset Rev"
					copyinstallerfromftp()
				    installapplication()
				  Else
					;$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\" & $build & "\Event Master Toolset Rev " & $build & " (Build " & $version & ")"
					;$ftp_old = "\\saceng-nas01\Share1\EventMaster\E2\release\" & $build & "\old\Event Master Toolset Rev " & $build & " (Build " & $version & ")"
					$ftp = "\\10.34.40.5\share1\EventMaster\E2\release\rc" & $build & "\old\Event Master Toolset Rev " & StringRight($build,3) & " (Build " & $version
					copyinstallerfromftp()
				    installapplication()
				  EndIf
			   EndIf
Else
       ConsoleWrite("No Parameters passed")
EndIf





Func selectbuild()
	$hgui = GUICreate("EM Installation: Enter build Version", 520, 190)
	GUICtrlSetState(-1, $gui_checked)
	GUICtrlCreateLabel("For Trunk Latest build left everything blank and click OK button this will automatically install the latest trunk build", 20, 10, 500, 18)
	GUICtrlSetFont(-1, 9, 2000)
	GUICtrlCreateLabel("install the latest trunk build", 20, 25, 500, 18)
	GUICtrlSetFont(-1, 9, 2000)
	GUICtrlCreateLabel("For Release mention release number here in X.Y format (Example:7.1) >>>", 20, 60, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("(if left blank then this will install trunk build)", 20, 75, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("Mention here which build number you want to install (Example: 56118) >>>", 20, 115, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlCreateLabel("(if left blank then this will install latest build version)", 20, 130, 500, 18)
	GUICtrlSetFont(-1, 9, 700)
	GUICtrlSetState(-1, $gui_dropaccepted)
	$build = GUICtrlCreateInput("", 440, 60, 30, 18)
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
			   If (StringCompare($build, "trunk",0 ) == 0) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\test\Event Master Toolset Rev"
					WinClose("Select Build")
					ExitLoop
			   ElseIf (StringCompare($build, "e3",0 ) == 0) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\Macchia\bundles\development\Windows\Macchia Toolset Rev 1.0.0 (Build"
					WinClose("Select Build")
					ExitLoop
			   ElseIf (StringCompare($build, "rc",0 ) > 0) Then
					$ftp = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\rc" & $build & "\Event Master Toolset Rev"
					WinClose("Select Build")
					ExitLoop
				ElseIf (StringIsSpace($version)) Then
					$ftp = "\\10.34.40.5\share1\EventMaster\E2\release\rc" & $build & "\old\Event Master Toolset Rev " & StringRight($build,3) & " (Build " & $version
					WinClose("Select Build")
					ExitLoop
				Else
					$ftp = "\\10.34.40.5\share1\EventMaster\E2\release\rc" & $build & "\old\Event Master Toolset Rev " & StringRight($build,3) & " (Build " & $version
					;$ftp_old = "\\saceng-nas01\Share1\EventMaster\E2\release\rc" & $build & "\old\Event Master Toolset Rev " & $build & " (Build " & $version & ")"
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
		If (FileCopy($ftp_old & "*Setup.exe", $fileloc, $fc_overwrite + $fc_createpath) == 0) Then
			MsgBox($mb_systemmodal, "Info", "Build Number Not available" & @CRLF & " in Pool folder location", 7)
			;ExitLoop
		EndIf
	EndIf
	$apppath = _filelisttoarray(@TempDir & "\InstallerFiles\", Default, Default, True)
	MsgBox($mb_systemmodal, "Info", "Opening Build location" & @CRLF & $apppath[1], 2)
EndFunc

Func installapplication()
	MsgBox($mb_systemmodal, "Info", "Build Installation is in progress...", 3)
	ConsoleWrite($apppath[1] & " --mode unattended" & " --prefix " & $path)
	RunWait($apppath[1] & " --mode unattended" & " --prefix " & $path)
	;ShellExecute("C:\Barco")
	MsgBox($mb_systemmodal, "Info", "                            Installation Completed !!" & @CRLF & "           Timeout after 15 seconds or Click the OK button", 15)
EndFunc

;#ce