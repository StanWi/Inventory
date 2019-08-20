#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=box1.ico
#AutoIt3Wrapper_Res_Description=Inventory Startup
#AutoIt3Wrapper_Res_Fileversion=0.1.0.1
#AutoIt3Wrapper_Res_Language=1049
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

$dir = @LocalAppDataDir & "\Programs\Inventory\"
$exe = "inventory.exe"
$ex_ = @ScriptDir & "\inventory.ex_"

If Not FileExists($ex_) Then
	MsgBox(0, "Сообщение - Инвентори", "Нет файла " & $ex_, 5)
	Exit
EndIf

If Not FileExists($dir & $exe) Then
	FileCopy($ex_, $dir & $exe, 9)
	ConsoleWrite('Copy (not exists) ' & $ex_ & ' to ' & $dir & $exe & @CRLF)
EndIf

If FileGetVersion($ex_) <> FileGetVersion($dir & $exe) Then
	FileCopy($ex_, $dir & $exe, 9)
	ConsoleWrite('Copy (version missmatch) ' & $ex_ & ' to ' & $dir & $exe & @CRLF)
EndIf

Run($dir & $exe, $dir)
