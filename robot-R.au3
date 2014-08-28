#cs ----------------------------------------------------------------------------
	robot-R
	AutoIt Version: 3.3.6.1
	Author: JeromeYang

	Script Function:
	My daily work script

#ce ----------------------------------------------------------------------------

#include <WindowsConstants.au3>
#include <Date.au3>
#include <Timers.au3>
#include <Misc.au3>
#include <Audio.au3>
#include <IE.au3>

Opt('MustDeclareVars', 1)
Opt("WinTitleMatchMode", 1)
Opt("TrayAutoPause",0) 

_Singleton("robot-R")

Global $config_file_name, $log_file_name, $username, $password, $retry_pause, $run, $run_work_dir, $reboot, $radio[4], $radio_flag

Dim $set_username, $set_password, $set_login, $set_logout
$retry_pause = 10 * 1000
$config_file_name = "robot-R.ini"
$log_file_name = "robot-R.log"

TraySetToolTip("robot-R")
HotKeySet("{F1}", "_ToggleMuteMaster" )
HotKeySet("^{F1}", "ManualCheck")
HotKeySet("{F2}", "SwitchRadio")

If Not FileExists($config_file_name) Then
	SetConfig()
EndIf

init()

#region Object
Global $oMyError, $oMediaplayer, $oMediaPlayControl, $oMediaPlaySettings

$oMyError = ObjEvent("AutoIt.Error","Quit")
$oMediaplayer = ObjCreate("WMPlayer.OCX.7") 

If Not IsObj($oMediaplayer) Then Exit

$oMediaplayer.Enabled = true
$oMediaplayer.WindowlessVideo= true
$oMediaPlayer.UImode="invisible"
$oMediaPlayControl=$oMediaPlayer.Controls
$oMediaPlaySettings=$oMediaPlayer.Settings
#endregion

#Region Main
While 1
	If Now(1) = $reboot Then
		Shutdown(6)
	ElseIf FileGetTime($config_file_name,0,1)=Now(0) Then 
		init()
	EndIf
	WinSetState ( _Base64Decode("6Zu75a2Q55eF5q2357C956ug5Z+36KGM54uA5oWL"), "", @SW_HIDE )
WEnd

Exit
#Endregion

#Region GeneralFunctions

Func init()
	Dim $i, $j
    Local $run_name_array, $program_name 
	$radio_flag = 0
	$username = IniRead($config_file_name, "DailyLog", "username", 0)
	$password = IniRead($config_file_name, "DailyLog", "password", 0)
	$reboot = IniRead($config_file_name, "Reboot", "time", 70000)
	
	For $j=1 to 3
		$radio[$j] = IniRead($config_file_name, "Radio", "R"&String($j), "")
	Next
	
	$run = IniRead($config_file_name, "ExternalProgram", "run", "")
	$run_work_dir = IniRead($config_file_name, "ExternalProgram", "workdir", "")
	If ($run<>"" AND $run_work_dir<>"") Then
        $run_name_array = StringSplit($run, "\")
        $program_name = $run_name_array[Ubound($run_name_array)-1]
		If NOT FileExists($run) Then 
			MsgBox(4096, $run , "External Program does NOT exists")
		ElseIf NOT FileExists($run_work_dir) Then
			MsgBox(4096, $run_work_dir , "WorkDIR does NOT exists")
		ElseIf NOT ProcessExists($program_name) Then
			Run($run, $run_work_dir)
		Endif 
	Endif
	TrayTip ( "Configuration loaded", "", 10, 1 )
EndFunc   ;==>init

Func SetConfig()
	$set_username = InputBox("Robot-R", "Please input your ID", $username)
	$set_password = InputBox("Robot-R", "Please input your Password", $password, "*")
	IniWrite($config_file_name, "DailyLog", "username", $set_username)
	IniWrite($config_file_name, "DailyLog", "password", $set_password)
	IniWrite($config_file_name, "ExternalProgram", "run", "")
	IniWrite($config_file_name, "ExternalProgram", "workdir", "")
	IniWrite($config_file_name, "Reboot", "time", 70000)
	IniWrite($config_file_name, "Radio", "R1", "mms://bcr.media.hinet.net/RA000018")
	IniWrite($config_file_name, "Radio", "R2", "mms://bcr.media.hinet.net/RA000009")
	IniWrite($config_file_name, "Radio", "R3", "mms://bcr.media.hinet.net/RA000008")
	
EndFunc   ;==>SetConfig
#Endregion

#Region RadioFunctions
Func SwitchRadio()
	$radio_flag=Mod($radio_flag+1,4)
	If $radio_flag=0 Then
		$oMediaPlayControl.Stop
	Else
		$oMediaPlayer.URL=$radio[$radio_flag]
		$oMediaPlayControl.Play
	EndIf
EndFunc

Func _ToggleMuteMaster()
    Local $retVal = 0, $ex = False
    If Not WinExists('[CLASS:Volume Control]') Then
        Run('sndvol32', '', @SW_HIDE)
        $ex = True
    EndIf
    If WinWait('[CLASS:Volume Control]', '', 2) = 0 Then Return -1
    $retVal = ControlCommand('[CLASS:Volume Control]', '', 1000, 'isChecked')
    If @error Then Return -2
    If $retVal Then
    ControlCommand('[CLASS:Volume Control]', '', 1000, 'UnCheck')
    If @error Then Return -2
    Else
        ControlCommand('[CLASS:Volume Control]', '', 1000, 'Check')
        If @error Then Return -2
    EndIf
    If $ex = True Then WinClose('[CLASS:Volume Control]')
    Return
EndFunc
#EndRegion

#Region WorkShiftFunctions

Func ManualCheck ()
	CheckNow()
EndFunc

Func CheckNow()
	Local $try, $oIE, $oForm, $oQuery, $oLogin, $oCheckbox, $oObjs, $aLogTable, $reson_value, $return=0, $element, $lastCheck, $startCheck, $todayHour1, $todayHour2
	TrayTip ( "Checking...", "", 10, 1 )
	$startCheck = Now(0)
	$reson_value=_Base64Decode("55yL54mH5Yik6K6A")


	For $try = 1 To 1
		;Login
		$oIE=_IECreate("http://192.168.200.4/drtime/default.asp", 0, 0, 1, 0)
		$oForm = _IEFormGetCollection($oIE, 0)
		$oQuery = _IEFormElementGetObjByName ($oForm, "Txt_Uid")
		_IEFormElementSetValue($oQuery, $username)
		$oQuery = _IEFormElementGetObjByName ($oForm, "Txt_Pwd")
		_IEFormElementSetValue($oQuery, $password)
		$oLogin = _IEGetObjById($oIE, "submit1")
		_IEAction($oLogin, "click")
		_IELoadWait ($oIE)
		
		;Selective reson
		$oForm = _IEFormGetCollection($oIE, 0)
		_IEFormElementCheckboxSelect ($oForm, $reson_value)
		$oObjs = _IETagNameGetCollection ($oIE, "input")
		For $oObj In $oObjs
			If $oObj.name = "submit" Then
				_IEAction($oObj, "click")
				ExitLoop
			Endif
		Next
		_IELoadWait ($oIE)
		
		;Check
		_IENavigate($oIE, "http://192.168.200.4/drtime/default.asp")
		$oForm = _IEFormGetCollection($oIE, 0)
		$oQuery = _IEFormElementGetObjByName ($oForm, "Txt_Uid")
		_IEFormElementSetValue($oQuery, $username)
		$oQuery = _IEFormElementGetObjByName ($oForm, "Txt_Pwd")
		_IEFormElementSetValue($oQuery, $password)
		$oLogin = _IEGetObjById($oIE, "submit1")
		_IEAction($oLogin, "click")
		_IELoadWait ($oIE)
		_IENavigate($oIE, "http://192.168.200.4/drtime/show_time.asp")
		$aLogTable = _IETableWriteToArray (_IETableGetCollection ($oIE, 0), True)
        ;MsgBox(1, DateNormalize($aLogTable[UBound($aLogTable)-1][0]), "")
		IF DateNormalize($aLogTable[UBound($aLogTable)-1][0])=String(@YEAR-1911)&"/"&String(Number(@MON))&"/"&String(Number(@MDAY)) Then
			$lastCheck = @YEAR * 10000000000 + @MON * 100000000 + @MDAY * 1000000 + Number(StringReplace($aLogTable[UBound($aLogTable)-1][1], ":", ""))
			If $aLogTable[UBound($aLogTable)-1][2] <> "" Then
				$lastCheck = @YEAR * 10000000000 + @MON * 100000000 + @MDAY * 1000000 + Number(StringReplace($aLogTable[UBound($aLogTable)-1][2], ":", ""))
				$todayHour1 = Number($aLogTable[UBound($aLogTable)-1][3])
				$todayHour2 = Number($aLogTable[UBound($aLogTable)-1][4])
			Endif
			If $lastCheck-$startCheck>0 Then
				TrayTip ( "Success", " ("&$todayHour1+$todayHour2&")", 10, 1 )
				TraySetToolTip("robot-R: Success ("&$todayHour1+$todayHour2&")")
				$return = 1
				_IEQuit ($oIE)
				WriteLog("Success ("&$todayHour1&"+"&$todayHour2&")")
				ExitLoop
			Endif
		Endif
		WriteLog("Fail "&$aLogTable[UBound($aLogTable)-1][0]&" "&$lastCheck)
		TrayTip ( "Fail", "", 10, 3 )
		TraySetToolTip("robot-R: Fail "&$aLogTable[UBound($aLogTable)-1][0])
		_IEQuit ($oIE)
		Sleep($retry_pause)
		If $try<3 Then TrayTip ( "Retrying...", "", 10, 2 )
	Next
	Return $return
EndFunc   ;==>CheckNow

Func WriteLog($description)
	Local $file, $line
	$file = FileOpen($log_file_name, 1)
	$line = Now(0) & " " & $description
	FileWriteLine($file, $line)
	FileClose($file)
EndFunc   ;==>WriteLog

Func Now($flag) ;$flag = 0, return datetime, $flag = 1, return time
	If $flag = 0 Then
		Return @YEAR * 10000000000 + @MON * 100000000 + @MDAY * 1000000 + @HOUR * 10000 + @MIN * 100 + @SEC
	ElseIf $flag = 1 Then
		Return @HOUR * 10000 + @MIN * 100 + @SEC
	EndIf
EndFunc   ;==>Now

Func DateNormalize($date)
    Local $array
    $array = StringSplit($date, "/")
    Return String(Number($array[1]))&"/"&String(Number($array[2]))&"/"&String(Number($array[3]))
EndFunc
    
#Endregion

Func _Base64Decode($sData)
    Local $oXml = ObjCreate("Msxml2.DOMDocument")
    If Not IsObj($oXml) Then
        SetError(1, 1, 0)
    EndIf

    Local $oElement = $oXml.createElement("b64")
    If Not IsObj($oElement) Then
        SetError(2, 2, 0)
    EndIf

    $oElement.dataType = "bin.base64"
    $oElement.Text = $sData
    Local $sReturn = BinaryToString($oElement.nodeTypedValue, 4)

    If StringLen($sReturn) = 0 Then
        SetError(3, 3, 0)
    EndIf

    Return $sReturn
EndFunc
