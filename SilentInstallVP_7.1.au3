#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <AutoItConstants.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ScreenCapture.au3>


Global $IPs = "10.99.125.70"
Global $build = "7.1"
Global $version = ""
Global $oIE
Global $FTP
Global $FTP = ""
Global $FTP_Old = ""
Global $fileLoc = @TempDir & "\InstallerFiles\"
Global $VpBuild
$FTP = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\rc" & $build & "\em_update_vp." & $build & "." & $version
$AppPath = _FileListToArray(@TempDir & "\InstallerFiles\", Default, Default, True)
$VpBuild = $AppPath[1]
ConsoleWrite("$VpBuild  " & $VpBuild)

;SelectBuild()
;CopyInstallerFromFTP()
UpdateVP($IPs)

Func UpdateVP($IPs)
   Opt("WinTitleMatchMode", 2)
   $WinTitle = "Choose File to Upload"
   $IETitle =  "Web App"
   MsgBox($MB_SYSTEMMODAL, "Info", "Update VP on Hardware Intiating...." , 2)
    ;$VpBuild = "C:\Users\parsi\AppData\Local\Temp\InstallerFiles\em_update_vp.7.1.3143.tar.gz"
   $MultipleIPs = StringSplit($IPs,";")
   For $i = 1 To $MultipleIPs[0]
	  $oIE = _IECreate($MultipleIPs[$i])
	  $HWND = _IEPropertyGet($oIE, "hwnd")
	  WinSetState($HWND, "", @SW_MAXIMIZE)
	  _IEAction($oIE, "visible")
	  _IELoadWait($oIE)
	  ClickObjectByName("Tools")
	  ClickObjectByName("Manage Software")
	  Sleep(4000)
	  Local $Pos  = WinGetPos($IETitle)
	  ConsoleWrite("$Pos[0]: " & $Pos[0])
	  ConsoleWrite("$Pos[1]: " & $Pos[1])
	  ConsoleWrite("$Pos[2]: " & $Pos[2])
	  ConsoleWrite("$Pos[3]: " & $Pos[3])
	 ; ConsoleWrite("$Pos[4]: " & $Pos[4])

	  ;$fileUpload = _IEGetObjById($oIE, "fileupload")
	  ;_IEAction($fileUpload, "focus")
	  ;Sleep(3000)
	  ;Send("{SPACE}")
	  ;Sleep(3000)
	  WinWaitActive($WinTitle,"",30)
	  WinActivate($WinTitle,"")
	  Sleep(1000)
	  ControlSend($WinTitle,"","[CLASS:Edit; INSTANCE:1]",$VpBuild);
	  Sleep(2000)
	  ControlClick($WinTitle,"","[CLASS:Button; INSTANCE:1]")
	  Sleep(3000)
	  WinActivate($IETitle,"")
	  Sleep(2000)
	  Local $aPos  = WinGetPos($IETitle)
	  MouseClick("left",$aPos[2] - 1185,$aPos[3] - 160,1)
	  Sleep(10000)
	  ;$hBmp = _ScreenCapture_Capture("")
	  ;_ScreenCapture_Capture($fileLoc &  "/" & $MultipleIPs[$i] & "_" & @MDAY & @MON &  @YEAR &  "_" & @HOUR & @MIN & @SEC &".jpg", 0,0,3000,1500)
      MsgBox($MB_SYSTEMMODAL, "Info", "Refer File for status: " & $fileLoc &  "/" & $MultipleIPs[$i] & "_" & @MDAY & @MON &  @YEAR &  "_" & @HOUR & @MIN & @SEC &".jpg", 2)
	  MsgBox($MB_SYSTEMMODAL, "Info", "File uploaded , process will be done automatically in backgrund...." , 2)
   Next
   ;ShellExecute("taskkill /F /IM iexplore.exe")
   while (WinExists($IETitle,""))
	  WinClose($IETitle)
   WEnd

   ShellExecute($fileLoc)
EndFunc


Func ClickObjectByName($sMyString)
   $sMyString
   $oLinks = _IELinkGetCollection($oIE)
   For $oLink In $oLinks
	  ConsoleWrite("InLoopName...####")
	   Local $sLinkText = _IEPropertyGet($oLink, "innerText")
	   If StringInStr($sLinkText, $sMyString) Then
		   _IEAction($oLink, "click")
		   ExitLoop
	   EndIf
   Next
EndFunc




Func ClickObjectByIDTest($sMyString)
   $sMyString
   $oLinks = _IELinkGetCollection($oIE)
   For $oLink In $oLinks
	  ConsoleWrite("InLoopID...####")
	   Local $sLinkText = _IEPropertyGet($oLink, "id")
	   If StringInStr($sLinkText, $sMyString) Then
		   _IEAction($oLink, "click")
		   ExitLoop
	   EndIf
   Next
EndFunc


Func ClickObjectByID(ByRef $oObject, $sAction)
	If Not IsObj($oObject) Then
		__IEConsoleWriteError("Error", "_IEAction(" & $sAction & ")", "$_IESTATUS_InvalidDataType")
		Return SetError($_IESTATUS_InvalidDataType, 1, 0)
	EndIf
	;
	$sAction = StringLower($sAction)
	Select
		; DOM objects
		Case $sAction = "click"
			If __IEIsObjType($oObject, "documentContainer") Then
				__IEConsoleWriteError("Error", "_IEAction(click)", " $_IESTATUS_InvalidObjectType")
				;Return SetError($_IESTATUS_InvalidObjectType, 1, 0)
			 EndIf
			Sleep(3000)
			$oObject.Click()
			ConsoleWrite("Now")
			;$oObject.Quit()
			Exit
	  ;ExitLoop
   EndSelect
EndFunc


Func SelectBuild()
   $hGUI = GUICreate("Update VP for Real Hardware", 520, 250)
   GUICtrlSetState(-1, $GUI_CHECKED)
   GUICtrlCreateLabel("Enter Real Hardware IP, Use semicolon(;) to provide multiple IP's " , 30, 10, 500, 18)
   GUICtrlSetFont (-1,11, 2000); bold
   GUICtrlCreateLabel("For Trunk Latest build left everything blank and click OK button this will automatically install the latest trunk build" , 20, 60, 500, 18)
   GUICtrlSetFont (-1,9, 2000); bold
   GUICtrlCreateLabel("install the latest trunk build of VP" , 20, 75, 500, 18)
   GUICtrlSetFont (-1,9, 2000); bold
   GUICtrlCreateLabel("For Release mention release number here in X.Y format (Example:7.1) >>>" , 20, 110, 500, 18)
   GUICtrlSetFont (-1,9, 700); bold
   GUICtrlCreateLabel("(if left blank then this will install trunk build)" , 20, 125, 500, 18)
   GUICtrlSetFont (-1,9, 700); bold
   GUICtrlCreateLabel("Mention here which build number you want to install (Example: 56118) >>>" , 20, 155, 500, 18)
   GUICtrlSetFont (-1,9, 700); bold
   GUICtrlCreateLabel("(if left blank then this will install latest build version)" , 20, 170, 500, 18)
   GUICtrlSetFont (-1,9, 700); bold



   GUICtrlSetState(-1, $GUI_DROPACCEPTED)
   $IPs = GUICtrlCreateInput("", 150, 33, 250, 18)
   $build = GUICtrlCreateInput("", 440, 110, 30, 18)
   $version = GUICtrlCreateInput("", 440, 155, 45, 18)
   $hButton = GUICtrlCreateButton("OK", 210, 210, 90, 30)
   GUICtrlSetFont (-1,9, 700); bold


   GUISetState(@SW_SHOW)
   While 1
	  Switch GUIGetMsg()
		 Case $GUI_EVENT_CLOSE
				   Exit
				Case $hButton
					 $IPs     = StringLower(GUICtrlRead($IPs))
					 $build   = StringLower(GUICtrlRead($build))
					 $version = StringLower(GUICtrlRead($version))
				  If (StringIsSpace($build)) Then
					  $FTP = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\test\em_update_vp"
					  WinClose("Select Build")
					  ExitLoop
				   ElseIf(StringIsSpace($version)) Then
					 $FTP = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\rc" & $build & "\em_update_vp"
					 WinClose("Select Build")
					 ExitLoop
				  Else
					  $FTP = "\\folfps07\eng-pool$\Woodbridge\Noida\E2\release\rc" & $build & "\em_update_vp." & $build & "." & $version
					  $FTP_Old = "\\saceng-nas01\Share1\EventMaster\E2\release\rc" & $build & "\old\em_update_vp." & $build & "." & $version
				  EndIf
				  ExitLoop
	  EndSwitch
   WEnd
   MsgBox (1,"Provided list of IPs: ",$IPs, 2)
   ;MsgBox (1,"",$FTP)
   ;MsgBox (1,"",$FTP_Old)

EndFunc




Func CopyInstallerFromFTP()
    ConsoleWrite($FTP)
	$fileLoc = @TempDir & "\InstallerFiles\"
	ConsoleWrite($FTP & "*.tar.gz")
	ConsoleWrite($fileLoc)
	FileDelete($fileLoc)
	MsgBox($MB_SYSTEMMODAL, "Info", "Copying Build..." , 2)
	  If( FileCopy($FTP & "*.tar.gz",$fileLoc, $FC_NOOVERWRITE + $FC_CREATEPATH) == 0) Then
		 MsgBox($MB_SYSTEMMODAL, "Info", "Copying failed due to slow network issue..." , 2)
		 If (FileCopy($FTP_Old & "*.tar.gz",$fileLoc, $FC_NOOVERWRITE + $FC_CREATEPATH) == 0) Then
			MsgBox($MB_SYSTEMMODAL, "Info", "Build Number Not available"& @crlf & " in Pool folder location: " & $FTP_Old,7)
			Exit
		 EndIf
	  EndIf
    $AppPath = _FileListToArray(@TempDir & "\InstallerFiles\", Default, Default, True)
	$VpBuild = $AppPath[1]
    MsgBox($MB_SYSTEMMODAL, "Info", "Opening Build location" & @crlf & $AppPath[1] , 2)
	;ShellExecute($fileLoc)
EndFunc   ;
