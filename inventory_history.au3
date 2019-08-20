#include-once
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <GuiComboBox.au3>
#include <StringConstants.au3>

Local $debug = False

Local $hWnd, $ComboVendor, $ComboModule, $ComboSN, $ComboStatus
Local $ComboNSZ, $ComboNE, $ComboNN, $ComboIN, $ComboOptions
Local $ButtonBack, $ButtonForward
$number_options = 20
Global $aHistory
Dim $aHistory[1][8 + $number_options]
; $aHistory[0][0] - History Level
; $aHistory[x][0] - Type or vendor
; $aHistory[x][1] - Module
; $aHistory[x][2] - Serial Number
; $aHistory[x][3] - State
; $aHistory[x][4] - NSZ
; $aHistory[x][5] - Location
; $aHistory[x][6] - Nom Number
; $aHistory[x][7] - Inv Number

Func history_add()
	GUICtrlSetState($ButtonForward, $GUI_DISABLE)
	$value = _GUICtrlComboBox_GetEditText($ComboVendor)
	$value &= "|" & _GUICtrlComboBox_GetEditText($ComboModule)
	$value &= "|" & _GUICtrlComboBox_GetEditText($ComboSN)
	$value &= "|" & _GUICtrlComboBox_GetEditText($ComboStatus)
	$value &= "|" & _GUICtrlComboBox_GetEditText($ComboNSZ)
	$value &= "|" & _GUICtrlComboBox_GetEditText($ComboNE)
	$value &= "|" & _GUICtrlComboBox_GetEditText($ComboNN)
	$value &= "|" & _GUICtrlComboBox_GetEditText($ComboIN)
	For $i = 0 To $number_options - 1
		$value &= "|" & _GUICtrlComboBox_GetEditText($ComboOptions[$i])
	Next
	$items = StringSplit($value, "|", $STR_NOCOUNT)
	If $aHistory[0][0] < UBound($aHistory) - 1 Then
		While $aHistory[0][0] < UBound($aHistory) - 1
			_ArrayDelete($aHistory, UBound($aHistory) - 1)
		WEnd
	EndIf
	$new = False
	For $i = 0 To 7 + $number_options
		If $items[$i] <> $aHistory[UBound($aHistory) - 1][$i] Then
			$new = True
			ExitLoop
		EndIf
	Next
	If $new Then
		_ArrayAdd($aHistory, $value)
		$aHistory[0][0] = UBound($aHistory) - 1
	ElseIf $aHistory[0][0] < UBound($aHistory) - 1 Then
		$aHistory[0][0] += 1
	EndIf
	If UBound($aHistory) > 2 And $aHistory[0][0] > 1 Then
		GUICtrlSetState($ButtonBack, $GUI_ENABLE)
	EndIf
	; If $debug Then (Don't use _ArrayDisplay. Try _SQLite_Displaq2DResult.)
EndFunc   ;==>history_add

Func history_back()
	If $aHistory[0][0] < 3 Then
		GUICtrlSetState($ButtonBack, $GUI_DISABLE)
	EndIf
	$aHistory[0][0] -= 1
	ControlSetText($hWnd, '', $ComboVendor, $aHistory[$aHistory[0][0]][0])
	If _GUICtrlComboBox_GetEditText($ComboVendor) Then
		_SetModule($aHistory[$aHistory[0][0]][1])
	Else
		ControlSetText($hWnd, '', $ComboModule, $aHistory[$aHistory[0][0]][1])
	EndIf
	If _GUICtrlComboBox_GetEditText($ComboModule) Then
		_SetSN($aHistory[$aHistory[0][0]][2])
	Else
		ControlSetText($hWnd, '', $ComboSN, $aHistory[$aHistory[0][0]][2])
	EndIf
	ControlSetText($hWnd, '', $ComboStatus, $aHistory[$aHistory[0][0]][3])
	ControlSetText($hWnd, '', $ComboNSZ, $aHistory[$aHistory[0][0]][4])
	ControlSetText($hWnd, '', $ComboNE, $aHistory[$aHistory[0][0]][5])
	ControlSetText($hWnd, '', $ComboNN, $aHistory[$aHistory[0][0]][6])
	ControlSetText($hWnd, '', $ComboIN, $aHistory[$aHistory[0][0]][7])
	For $i = 8 To 7 + $number_options
		ControlSetText($hWnd, '', $ComboOptions[$i - 8], $aHistory[$aHistory[0][0]][$i])
	Next
	_Filter(True)
	GUICtrlSetState($ButtonForward, $GUI_ENABLE)
	If $debug Then _ArrayDisplay($aHistory)
EndFunc   ;==>history_back

Func history_forward()
	If $aHistory[0][0] > UBound($aHistory) - 3 Then
		GUICtrlSetState($ButtonForward, $GUI_DISABLE)
	EndIf
	$aHistory[0][0] += 1
	ControlSetText($hWnd, '', $ComboVendor, $aHistory[$aHistory[0][0]][0])
	If _GUICtrlComboBox_GetEditText($ComboVendor) Then
		_SetModule($aHistory[$aHistory[0][0]][1])
	Else
		ControlSetText($hWnd, '', $ComboModule, $aHistory[$aHistory[0][0]][1])
	EndIf
	If _GUICtrlComboBox_GetEditText($ComboModule) Then
		_SetSN($aHistory[$aHistory[0][0]][2])
	Else
		ControlSetText($hWnd, '', $ComboSN, $aHistory[$aHistory[0][0]][2])
	EndIf
	ControlSetText($hWnd, '', $ComboStatus, $aHistory[$aHistory[0][0]][3])
	ControlSetText($hWnd, '', $ComboNSZ, $aHistory[$aHistory[0][0]][4])
	ControlSetText($hWnd, '', $ComboNE, $aHistory[$aHistory[0][0]][5])
	ControlSetText($hWnd, '', $ComboNN, $aHistory[$aHistory[0][0]][6])
	ControlSetText($hWnd, '', $ComboIN, $aHistory[$aHistory[0][0]][7])
	For $i = 8 To 7 + $number_options
		ControlSetText($hWnd, '', $ComboOptions[$i - 8], $aHistory[$aHistory[0][0]][$i])
	Next
	_Filter(True)
	GUICtrlSetState($ButtonBack, $GUI_ENABLE)
	If $debug Then _ArrayDisplay($aHistory)
EndFunc   ;==>history_forward
