;~ Check #NoTrayIcon, #AutoIt3Wrapper_Res_Fileversion, $test_mode
#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=box1.ico
#AutoIt3Wrapper_Res_Description=Inventory Database
#AutoIt3Wrapper_Res_Fileversion=0.1.0.6
#AutoIt3Wrapper_Res_Language=1049
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstantsEx.au3>
#include <Array.au3>
#include <Date.au3>
#include <File.au3>
#include <GuiComboBox.au3>
#include <GuiListBox.au3>
#include <WindowsConstants.au3>
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <StaticConstants.au3>
#include <inventory_history.au3>
#include <EzMySql.au3>
#include <secret.au3>

$USER_WRITE = 1
$USER_WORKTIME = 2

If Not _EzMySql_Startup() Then
	MsgBox(0, 'Error Starting MySql', 'Error: ' & @error & @CR & 'Error string: ' & _EzMySql_ErrMsg())
	Exit
EndIf

$test_mode = False
If Not $test_mode Then
	$mysql_database = "idata"
	$mysql_user = $mysql_user_work
	$mysql_pass = $mysql_pass_work
Else
	$mysql_database = "idata_test"
	$mysql_user = $mysql_user_test
	$mysql_pass = $mysql_pass_test
EndIf

If Not _EzMySql_Open($mysql_ip, $mysql_user, $mysql_pass, $mysql_database, $mysql_port) Then
	$error = @error
	MsgBox(0, 'Error - Inventory', 'Error: ' & $error & @CR & 'Description: ' & _EzMySql_ErrMsg())
	Exit
EndIf
_EzMySql_Exec("SET NAMES 'cp1251'")
_EzMySql_Exec("SET CHARACTER SET 'cp1251'")

_EzMySql_Query("SELECT company FROM main WHERE id = 1;")
$aResult = _EzMySql_FetchData()
Global $CompanyName = $aResult[0]

_EzMySql_Query("SELECT program FROM main WHERE id = 1;")
$aResult = _EzMySql_FetchData()
Global $ProgramName = $aResult[0]

Global $vendor_mode = True
Global $moduleList
Global $GMT = 8
Global $hWnd = GUICreate(StringFormat('%s - %s', $CompanyName, $ProgramName), 890, 550, -1, -1, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX, $WS_MAXIMIZEBOX))
Global $autorization_request = 0
Global $g_bSortSense = False

; ===== Первая строка =====
GUICtrlCreateLabel('Тип или производитель', 10, 10, 210, 21)
GUICtrlCreateLabel('Оборудование', 230, 10, 210, 21)
GUICtrlCreateLabel('Серийный №', 450, 10, 210, 21)
GUICtrlCreateLabel('Количество', 670, 10, 190, 21)
$LabelHelp = GUICtrlCreateLabel('?', 870, 10, 10, 21)
GUICtrlSetColor(-1, 0x0000FF)
GUICtrlSetTip(-1, 'Справка')
GUICtrlSetCursor(-1, 0)
; ===== Вторая строка =====
Global $ComboVendor = GUICtrlCreateCombo('', 10, 30, 210, 21) ; Поле "Тип или производитель"
Global $ComboModule = GUICtrlCreateCombo('', 230, 30, 210, 21)
Global $ComboSN = GUICtrlCreateCombo('', 450, 30, 210, 21)
$InputNumber = GUICtrlCreateInput('', 670, 30, 210, 21)
GUICtrlSetState(-1, $GUI_DISABLE)
; ===== Третья строка =====
GUICtrlCreateLabel('Статус', 10, 60, 150, 21)
GUICtrlCreateLabel('НСЗ', 170, 60, 41, 21)
GUICtrlCreateLabel('Площадка', 230, 60, 210, 21)
GUICtrlCreateLabel('Ном. №', 450, 60, 100, 21)
GUICtrlCreateLabel('Инв. №', 560, 60, 100, 21)
; ===== Четвёртая строка =====
Global $ComboStatus = GUICtrlCreateCombo('', 10, 80, 150, 21)
Global $ComboNSZ = GUICtrlCreateCombo('', 170, 80, 50, 21)
GUICtrlSetData($ComboNSZ, 'Да|Нет')
Global $ComboNE = GUICtrlCreateCombo('', 230, 80, 210, 21)
Global $ComboNN = GUICtrlCreateCombo('', 450, 80, 100, 21)
Global $ComboIN = GUICtrlCreateCombo('', 560, 80, 100, 21)
Global $ButtonBack = GUICtrlCreateButton('<', 670, 70, 32, 32)
GUICtrlSetTip(-1, 'Назад')
GUICtrlSetState(-1, $GUI_DISABLE)
$ButtonFilter = GUICtrlCreateButton('Фильтр', 707, 70, 94, 32)
$ButtonForward = GUICtrlCreateButton('>', 806, 70, 32, 32)
GUICtrlSetTip(-1, 'Вперёд')
GUICtrlSetState(-1, $GUI_DISABLE)
$ButtonClearFilter = GUICtrlCreateButton('Х', 848, 70, 32, 32)
GUICtrlSetTip($ButtonClearFilter, 'Очистить форму')
; ===== Область данных =====
$ListViewColumns = 'Дата и время     |Производитель|Оборудование|Серийный №|Ном. №|Инв. №|Комментарий|Статус|НСЗ|Площадка'
Global $ListView = GUICtrlCreateListView($ListViewColumns, 10, 120, 870, 380, BitOR($LVS_SHOWSELALWAYS, $LVS_SINGLESEL))
Global $list_view_options
Dim $ListViewItem[1]
; ===== Опции ===== Поддерживается до 20 опций
Global $number_options = 20
Dim $LabelOptions[$number_options]
Dim $ComboOptions[$number_options]
$k = 0
For $i = 0 To 4
	For $j = 0 To 3
		$LabelOptions[$k] = GUICtrlCreateLabel('', 10 + 220 * $j, 110 + 50 * $i, 210, 21)
		GUICtrlSetData($LabelOptions[$k], $k)
		GUICtrlSetState(-1, $GUI_HIDE)
		$ComboOptions[$k] = GUICtrlCreateCombo('', 10 + 220 * $j, 130 + 50 * $i, 210, 21)
		GUICtrlSetState(-1, $GUI_HIDE)
		$k += 1
	Next
Next
; ===== Нижняя строка =====
$ButtonAutorize = GUICtrlCreateButton('Запрос авторизации', 10, 510, 210, 30)
GUICtrlSetBkColor(-1, 0xefffee)
GUICtrlSetState(-1, $GUI_HIDE)
$ButtonAdd = GUICtrlCreateButton('Добавить', 10, 510, 210, 30)
$ButtonEdit = GUICtrlCreateButton('Редактировать', 230, 510, 210, 30)
GUICtrlSetState(-1, $GUI_DISABLE)
$ButtonTools = GUICtrlCreateLabel('Инструменты', 450, 510, 210, 30, BitOR($SS_CENTER, $SS_CENTERIMAGE)) ;20170315
$ContextMenuTools = GUICtrlCreateContextMenu($ButtonTools)
$ToolsActiveUsers = GUICtrlCreateMenuItem('Активные пользователи', $ContextMenuTools)
$ToolsStatistic = GUICtrlCreateMenuItem('Статистика', $ContextMenuTools)
$tools_equipment_in_work = GUICtrlCreateMenuItem('Распределение: в работе по ЦТЭ', $ContextMenuTools)
$ButtonReport = GUICtrlCreateButton('Отчёт', 670, 510, 210, 30)
GUICtrlSetState(-1, $GUI_DISABLE)

$aResult = _EzMySql_GetTable2d("SELECT name FROM vendor ORDER BY name;")
GUICtrlSetData($ComboVendor, _ArrayToString($aResult, '', 1, -1, '|'))
_log("Start")
_EzMySql_Query(StringFormat("SELECT id, access, cte, name FROM user WHERE comp LIKE '%s/%s';", @ComputerName, @UserName))
$aResult = _EzMySql_FetchData()
Global $user_id = 0
Global $user_access = 0
Global $user_cte = 0
Global $user_name = ''
If Not IsArray($aResult) Then
	GUICtrlSetState($ComboVendor, $GUI_DISABLE)
	GUICtrlSetState($ComboModule, $GUI_DISABLE)
	GUICtrlSetState($ComboSN, $GUI_DISABLE)
	GUICtrlSetState($ComboStatus, $GUI_DISABLE)
	GUICtrlSetState($ComboNSZ, $GUI_DISABLE)
	GUICtrlSetState($ComboNE, $GUI_DISABLE)
	GUICtrlSetState($ComboNN, $GUI_DISABLE)
	GUICtrlSetState($ComboIN, $GUI_DISABLE)
	GUICtrlSetState($ButtonFilter, $GUI_DISABLE)
	GUICtrlSetState($ButtonClearFilter, $GUI_DISABLE)
	GUICtrlSetState($ListView, $GUI_DISABLE)
	GUICtrlSetState($ButtonTools, $GUI_DISABLE)
	GUICtrlSetState($ButtonAdd, $GUI_HIDE)
	GUICtrlSetState($ButtonEdit, $GUI_HIDE)
	GUICtrlSetState($ButtonAutorize, $GUI_SHOW)
	WinSetTitle($hWnd, '', StringFormat('%s - Неавторизованный доступ - %s', $CompanyName, $ProgramName))
Else
	$user_id = $aResult[0]
	$user_access = $aResult[1]
	$user_cte = $aResult[2]
	$user_name = $aResult[3]
	If Not BitAND($USER_WRITE, $user_access) Then
		GUICtrlSetState($ButtonAdd, $GUI_HIDE)
		GUICtrlSetState($ButtonEdit, $GUI_HIDE)
	EndIf
	Local $hName, $hCte
	_EzMySql_Query(StringFormat("SELECT name FROM cte WHERE id = %u;", $user_cte))
	$aCte = _EzMySql_FetchData()
	If $user_cte = 7 Then
		$query = "SELECT DISTINCT name FROM location ORDER BY name;"
	Else
		$query = StringFormat("SELECT DISTINCT name FROM location WHERE cte = %u OR cte = 7 ORDER BY name;", $user_cte)
	EndIf
	$aResult = _EzMySql_GetTable2d($query)
	GUICtrlSetData($ComboNE, _ArrayToString($aResult, '', 1, -1, '|'))
	$aResult = _EzMySql_GetTable2d("SELECT name FROM status ORDER BY name;")
	GUICtrlSetData($ComboStatus, _ArrayToString($aResult, '', 1, -1, '|'))
	WinSetTitle($hWnd, '', StringFormat('%s - %s - %s - %s', $CompanyName, $aCte[0], $user_name, $ProgramName))
EndIf

$tools_working_time = GUICtrlCreateDummy()
If BitAND($USER_WORKTIME, $user_access) Then
	$tools_working_time = GUICtrlCreateMenuItem('Учёт рабочего времени', $ContextMenuTools)
EndIf

GUISetState(@SW_SHOW)
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			_log("Stop")
			_EzMySql_Close()
			_EzMySql_ShutDown()
			Exit
		Case $ComboVendor ; Выбор поля "Тип или производитель"
			_SetModule()
		Case $ComboModule
			_SetSN()
		Case $ComboSN
			_SetInfo()
		Case $ButtonAutorize
			_Authorize()
		Case $ButtonAdd
			_Add()
		Case $ButtonEdit
			_Edit()
		Case $ButtonReport
			_Report()
		Case $ButtonBack
			history_back()
		Case $ButtonFilter
			_Filter()
		Case $ButtonForward
			history_forward()
		Case $ButtonClearFilter ; Очистить фильтр
			_ClearFilter()
		Case $LabelHelp ; Справка
			_Help()
		Case $ToolsActiveUsers ; Активные пользователи
			_ToolsActiveUsers()
		Case $ToolsStatistic ; Статистика
			_ToolsStatistic()
		Case $tools_equipment_in_work ; Распределение оборудования в работе по ЦТЭ
			tools_equipment_in_work()
		Case $tools_working_time ; Учёт рабочего времени
			tools_working_time_gui($user_cte)
	EndSwitch
WEnd

Func _SetModule($item = "") ; Выполняется после выбора поля "Тип или производитель"
	GUICtrlSetData($ButtonAdd, 'Добавить')
	GUICtrlSetState($ButtonEdit, $GUI_DISABLE)
	GUICtrlSetState($ButtonReport, $GUI_DISABLE)
	Local $vendor = _GUICtrlComboBox_GetEditText($ComboVendor) ; Считываем данные в поле "Тип или производитель"
	Local $query = ''
	; ===> Добавлено в релиз 20170315
	Local $wait = 'Ждите...'
	; Some funny
	Local $random = Random(0, 10, 1)
	Switch $random
		Case 0
			$wait = 'Ждите...'
		Case 1
			$wait = 'Пожалуйста, подождите...'
		Case 2
			$wait = 'Одну минутку...'
		Case 3
			$wait = 'Один момент...'
		Case 4
			$wait = 'Минуточку, пожалуйста...'
		Case 5
			$wait = 'Скоро будет готово...'
		Case 6
			$wait = 'Подготовка списка...'
		Case 7
			$wait = 'Всё будет...'
		Case 8
			$wait = 'Ван момент плиз...'
		Case 9
			$wait = 'Почти всё готово...'
		Case 10
			$wait = 'Ждите...'
	EndSwitch
	; Stop funny
	GUICtrlSetData($ComboModule, '') ; Устанавливаем значение в поле "Оборудование"
	GUICtrlSetData($ComboModule, $wait, $wait)
	; <=== 20170315
	GUICtrlSetData($ComboSN, '') ; Поле "Серийный №"
	GUICtrlSetData($InputNumber, '') ; Поле "Количество"
	GUICtrlSetData($ComboNN, '') ; Поле "Ном. №"
	GUICtrlSetData($ComboIN, '') ; Поле "Инв. №"
	_GUICtrlListView_DeleteAllItems($ListView) ; Очищаем список оборудования
	If BitAND($USER_WRITE, $user_access) Then
		GUICtrlSetState($ButtonAdd, $GUI_ENABLE)
	EndIf
	_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name = '%s';", $vendor))
	$vendor_id = _EzMySql_FetchData()
	; Type or vendor
	If IsArray($vendor_id) Then
		If $vendor_id[0] > 1000 Then
			$vendor_mode = False
			layer_type_deactivate()
			layer_type_activate($vendor_id[0])
			$query = StringFormat("SELECT equip_id FROM equipment WHERE type_id = %u ORDER BY equip_id;", $vendor_id[0])
		Else
			$vendor_mode = True
			layer_type_deactivate()
			If $user_cte = 7 Then
				$query = StringFormat("SELECT DISTINCT object FROM inventory WHERE vendor = %u ORDER BY object;", $vendor_id[0])
			Else
				; Запрос типов оборудования с которыми знаком ЦТЭ
				$query = StringFormat("SELECT DISTINCT object FROM inventory WHERE vendor = %u AND location IN (SELECT name FROM location WHERE cte = %u OR cte = 7) ORDER BY object;", $vendor_id[0], $user_cte)
			EndIf
		EndIf
	EndIf
	$aResult = _EzMySql_GetTable2d($query)
	GUICtrlSetData($ComboModule, '')
	If IsArray($aResult) And _EzMySql_Rows() > 0 Then
		GUICtrlSetData($ComboModule, _ArrayToString($aResult, '', 1, -1, '|'), $item)
	EndIf
EndFunc   ;==>_SetModule

Func _SetSN($item = "") ; Run after select "Equipment" field
	GUICtrlSetData($ButtonAdd, 'Добавить')
	If BitAND($USER_WRITE, $user_access) Then
		GUICtrlSetState($ButtonAdd, $GUI_ENABLE)
	EndIf
	_GUICtrlListView_DeleteAllItems($ListView)
	GUICtrlSetState($ButtonEdit, $GUI_DISABLE)
	GUICtrlSetState($ButtonReport, $GUI_DISABLE)
	Local $vendor = _GUICtrlComboBox_GetEditText($ComboVendor)
	Local $module = _GUICtrlComboBox_GetEditText($ComboModule)
	GUICtrlSetData($ComboSN, '')
	GUICtrlSetData($InputNumber, '')
	GUICtrlSetData($ComboNN, '')
	GUICtrlSetData($ComboIN, '')
	_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name LIKE '%s';", $vendor))
	$vendor_id = _EzMySql_FetchData()
	; Type or vendor
	If IsArray($vendor_id) And $vendor_id[0] > 1000 Then
		_EzMySql_Query(StringFormat("SELECT options FROM equipment WHERE equip_id LIKE '%s' AND type_id = %u;", $module, $vendor_id[0]))
		$options = _EzMySql_FetchData()
		If IsArray($options) Then
			$options_array = _options_to_array($options[0])
			$aResult = _EzMySql_GetTable2d(StringFormat("SELECT option_id, name, list FROM options WHERE type_id = %u;", $vendor_id[0]))
			$iRows = _EzMySql_Rows()
			If IsArray($aResult) And $iRows > 0 Then
				For $i = 0 To $iRows - 1
					GUICtrlSetData($ComboOptions[$i], '')
					GUICtrlSetData($ComboOptions[$i], list_to_string($aResult[$i + 1][2]), $options_array[$i + 1])
				Next
			EndIf
			$string = StringSplit($module, ' ')
			$module = _ArrayPop($string)
			$vendor = _ArrayToString($string, ' ', 1)
			_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name LIKE '%s';", $vendor))
			$vendor_id = _EzMySql_FetchData()
			If $vendor_id[0] = '' Then ; Если вендора нет в базе, например при ручном заполнении equipment и пропуске vendor
				Return 0
			EndIf
		EndIf
	EndIf
	If Not IsArray($vendor_id) Then
		Return 0
	EndIf
	; ===== Запрос серийных номеров =====
	If $user_cte = 7 Then
		$query = StringFormat("SELECT DISTINCT serial FROM inventory WHERE vendor = %u AND object LIKE '%s' ORDER BY serial;", $vendor_id[0], $module)
	Else
		$query = StringFormat("SELECT DISTINCT serial FROM inventory WHERE vendor = %u AND object LIKE '%s' AND location IN (SELECT name FROM location WHERE cte = %u OR cte = 7) ORDER BY serial;", $vendor_id[0], $module, $user_cte)
	EndIf
	$aResult = _EzMySql_GetTable2d($query)
	$iRows = _EzMySql_Rows()
	If IsArray($aResult) And $iRows > 0 Then
		GUICtrlSetData($ComboSN, _ArrayToString($aResult, '', 1, -1, '|'), $item)
		GUICtrlSetData($InputNumber, $iRows)
	EndIf
	; ===== Запрос номенклатурных номеров =====
	If $user_cte = 7 Then
		$query = StringFormat("SELECT DISTINCT nomnum FROM inventory WHERE vendor = (SELECT id FROM vendor WHERE name LIKE '%s') AND object LIKE '%s' AND nomnum LIKE '%%' ORDER BY nomnum;", $vendor, $module)
	Else
		$query = StringFormat("SELECT DISTINCT nomnum FROM inventory WHERE vendor = (SELECT id FROM vendor WHERE name LIKE '%s') AND object LIKE '%s' AND location IN (SELECT name FROM location WHERE cte = %u) AND nomnum LIKE '%%' ORDER BY nomnum;", $vendor, $module, $user_cte)
	EndIf
	$aResult = _EzMySql_GetTable2d($query)
	$iRows = _EzMySql_Rows()
	If IsArray($aResult) And $iRows > 0 Then
		GUICtrlSetData($ComboNN, _ArrayToString($aResult, '', 1, -1, '|'), $item)
	EndIf
	; ===== Запрос инвентарных номеров =====
	If $user_cte = 7 Then
		$query = StringFormat("SELECT DISTINCT invnum FROM inventory WHERE vendor = (SELECT id FROM vendor WHERE name LIKE '%s') AND object LIKE '%s' AND invnum LIKE '%%' ORDER BY invnum;", $vendor, $module)
	Else
		$query = StringFormat("SELECT DISTINCT invnum FROM inventory WHERE vendor = (SELECT id FROM vendor WHERE name LIKE '%s') AND object LIKE '%s' AND location IN (SELECT name FROM location WHERE cte = %u) AND invnum LIKE '%%' ORDER BY invnum;", $vendor, $module, $user_cte)
	EndIf
	$aResult = _EzMySql_GetTable2d($query)
	$iRows = _EzMySql_Rows()
	If IsArray($aResult) And $iRows > 0 Then
		GUICtrlSetData($ComboIN, _ArrayToString($aResult, '', 1, -1, '|'), $item)
	EndIf
EndFunc   ;==>_SetSN

Func _SetInfo()
	_GUICtrlListView_DeleteAllItems($ListView)
	$vendor = _GUICtrlComboBox_GetEditText($ComboVendor)
	$module = _GUICtrlComboBox_GetEditText($ComboModule)
	$sn = _GUICtrlComboBox_GetEditText($ComboSN)
	If $module Then
		Local $vendor_id
		_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name LIKE '%s';", $vendor))
		$vendor_id = _EzMySql_FetchData()
		; Type or vendor
		If IsArray($vendor_id) And $vendor_id[0] > 1000 Then
			$string = StringSplit($module, ' ')
			$module = _ArrayPop($string)
			$vendor = _ArrayToString($string, ' ', 1)
			_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name LIKE '%s';", $vendor))
			$vendor_id = _EzMySql_FetchData()
		EndIf
		$query = StringFormat("SELECT time, place, status, location, nomnum, invnum, nsz " & _
				"FROM inventory WHERE vendor = %u AND object LIKE '%s' AND serial LIKE '%s';", $vendor_id[0], $module, $sn)
		$aResult = _EzMySql_GetTable2d($query)
		$iRows = _EzMySql_Rows()
		If IsArray($aResult) And $iRows > 0 Then
			GUICtrlSetData($ComboNN, $aResult[1][4], $aResult[1][4])
			GUICtrlSetData($ComboIN, $aResult[1][5], $aResult[1][5])
			Dim $ListViewItem[$iRows]
			For $i = 1 To $iRows
				$nsz = 'Нет'
				If $aResult[$i][6] = 1 Then
					$nsz = 'Да'
				EndIf
				_EzMySql_Query(StringFormat("SELECT name FROM status WHERE id = %u;", $aResult[$i][2]))
				$status = _EzMySql_FetchData()
				$line = StringFormat('%s|%s|%s|%s|%s|%s|%s|%s|%s|%s', _
						_EPOCH($aResult[$i][0]), _                 ; date and time
						$vendor, _                                 ; vendor
						$module, _                                 ; equipment
						_GUICtrlComboBox_GetEditText($ComboSN), _  ; serial number
						$aResult[$i][4], _                         ; nom number
						$aResult[$i][5], _                         ; inv number
						$aResult[$i][1], _                         ; comment
						$status[0], _                              ; status
						$nsz, _                                    ; nsz
						$aResult[$i][3]) ; location
				$ListViewItem[$i - 1] = GUICtrlCreateListViewItem($line, $ListView)
			Next
			GUICtrlSetData($ButtonAdd, 'Переместить')
			GUICtrlSetState($ButtonReport, $GUI_ENABLE)
			If BitAND($USER_WRITE, $user_access) Then
				GUICtrlSetState($ButtonEdit, $GUI_ENABLE)
			EndIf
		EndIf
	EndIf
EndFunc   ;==>_SetInfo

Func _EPOCH($epoch)
	Return (_DateAdd('s', $epoch + $GMT * 3600, "1970/01/01 00:00:00"))
EndFunc   ;==>_EPOCH

Func _toEPOCH($time)
	Return _DateDiff('s', "1970/01/01 00:00:00", $time) - $GMT * 3600
EndFunc   ;==>_toEPOCH

Func _Authorize()
	If Not $autorization_request Then
		$autorization_request = 1
		_log("Autorization Request")
		MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), _
				'Для полученя доступа, пожалуйста, свяжитесь со' & @CRLF & _
				'Станиславом Сыросенко по тел. (39-52) 79-29-05.')
	EndIf
EndFunc   ;==>_Authorize

Func _Add()
	Local $vendor = _GUICtrlComboBox_GetEditText($ComboVendor)
	Local $query = StringFormat("SELECT id FROM vendor WHERE name = '%s';", $vendor)
	_EzMySql_Query($query)
	$vendor_id = _EzMySql_FetchData()
	If IsArray($vendor_id) And $vendor_id[0] > 1000 Then
		add_type($vendor_id)
	Else
		add_module()
	EndIf
EndFunc   ;==>_Add

Func add_module() ; Add information in database
	#cs
		Function adds information in the database from dialog window
	#ce
	Local $vendor = _GUICtrlComboBox_GetEditText($ComboVendor)
	Local $module = _GUICtrlComboBox_GetEditText($ComboModule)
	If $vendor Then
		If _CheckVendor($vendor) Then
			_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name = '%s';", $vendor))
			$vendor_id = _EzMySql_FetchData()
			; Type or vendor
			$height = 230
			$type_mode = False
			If IsArray($vendor_id) And $vendor_id[0] > 1000 Then
				$type_mode = True
				$string = StringSplit($module, ' ')
				$module = _ArrayPop($string)
				$vendor = _ArrayToString($string, ' ', 1)
				_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name LIKE '%s';", $vendor))
				$vendor_id = _EzMySql_FetchData()
				If $vendor_id = 0 Then
					Local $vendor_id[1]
					$vendor_id[0] = 0
				EndIf
				$height = 310
			EndIf
			Local $hAddWnd = GUICreate(StringFormat('Добавление - %s', $ProgramName), 480, $height, -1, -1, -1, -1, $hWnd)
			GUICtrlCreateLabel('Производитель', 10, 10, 110, 21)
			GUICtrlCreateLabel('Карта/Модуль', 10, 40, 110, 21)
			GUICtrlCreateLabel('Серийный номер', 10, 70, 110, 21)
			GUICtrlCreateLabel('Ном. №', 250, 70, 40, 21)
			GUICtrlCreateLabel('Инв. №', 365, 70, 40, 21)
			GUICtrlCreateLabel('Дата и время', 10, 100, 110, 21)
			GUICtrlCreateLabel('Комментарий', 10, 130, 110, 21)
			GUICtrlCreateLabel('Статус', 10, 160, 110, 21)
			GUICtrlCreateLabel('Площадка', 10, 190, 110, 21)
			GUICtrlCreateGroup('Параметры', 10, 220, 460, 80)
			If Not $type_mode Then
				GUICtrlSetState(-1, $GUI_HIDE)
			EndIf
			Local $params = ''
			For $i = 0 To $number_options - 1
				If _GUICtrlComboBox_GetEditText($ComboOptions[$i]) <> '' Then
					$params &= GUICtrlRead($LabelOptions[$i]) & ': ' & _GUICtrlComboBox_GetEditText($ComboOptions[$i]) & '; '
				EndIf
			Next
			$params = StringTrimRight($params, 2)
			$LabelAddOptions = GUICtrlCreateLabel($params, 20, 240, 440, 50)
			If Not $type_mode Then
				GUICtrlSetState(-1, $GUI_HIDE)
			EndIf
			; Поле вендора
			If $type_mode And $vendor = '' Then
				$LabelAddVendor = GUICtrlCreateCombo("", 130, 7, 225, 21)
				GUICtrlSetData($LabelAddVendor, '', '')
			Else
				$LabelAddVendor = GUICtrlCreateLabel($vendor, 130, 10, 225, 21)
			EndIf
			; Поле модуля
			If _GUICtrlComboBox_GetEditText($ComboModule) = '' Then
				$LabelAddObject = GUICtrlCreateCombo('', 130, 37, 225, 21)
				$aResult = _EzMySql_GetTable2d(StringFormat("SELECT DISTINCT object FROM inventory WHERE vendor = %u ORDER BY object;", $vendor_id[0]))
				$iRows = _EzMySql_Rows()
				If IsArray($aResult) And $iRows > 0 Then
					GUICtrlSetData($LabelAddObject, _ArrayToString($aResult, '', 1, -1, '|'))
				EndIf
			Else
				$LabelAddObject = GUICtrlCreateLabel($module, 130, 40, 340, 21)
			EndIf
			;Поле серийного номера
			If _GUICtrlComboBox_GetEditText($ComboSN) = '' Then
				$InputAddSN = GUICtrlCreateInput('', 130, 67, 110, 21)
			Else
				$InputAddSN = GUICtrlCreateLabel(_GUICtrlComboBox_GetEditText($ComboSN), 130, 70, 110, 21)
			EndIf
			;Поля номенклатурного и инвентарного номеров
			If (_GUICtrlComboBox_GetEditText($ComboNN) = '' And _GUICtrlComboBox_GetEditText($ComboIN) = '') Or _GUICtrlComboBox_GetEditText($ComboSN) = '' Then
				$InputAddNN = GUICtrlCreateInput('', 300, 67, 55, 21)
				$InputAddIN = GUICtrlCreateInput('', 415, 67, 55, 21)
			Else
				$InputAddNN = GUICtrlCreateLabel(_GUICtrlComboBox_GetEditText($ComboNN), 300, 70, 55, 21)
				$InputAddIN = GUICtrlCreateLabel(_GUICtrlComboBox_GetEditText($ComboIN), 415, 70, 55, 21)
			EndIf
			;-----
			$InputAddTime = GUICtrlCreateInput(_NowCalc(), 130, 97, 110, 21)
			$InputAddComment = GUICtrlCreateInput('', 130, 127, 340, 21)
			If _GUICtrlComboBox_GetEditText($ComboSN) Then
				ControlFocus($hAddWnd, '', $InputAddComment)
			EndIf
			$ComboAddStatus = GUICtrlCreateCombo('', 130, 157, 169, 21)
			$checkboxAddNSZ = GUICtrlCreateCheckbox('НСЗ', 309, 157, 41, 21)
			$aResult = _EzMySql_GetTable2d("SELECT name FROM status ORDER BY name;")
			If IsArray($aResult) And _EzMySql_Rows() > 0 Then
				GUICtrlSetData($ComboAddStatus, _ArrayToString($aResult, '', 1, -1, '|'))
			EndIf
			$ComboAddLocation = GUICtrlCreateCombo('', 130, 187, 220, 21)
			If $user_cte = 7 Then
				$query = "SELECT DISTINCT name FROM location ORDER BY name;"
			Else
				$query = StringFormat("SELECT DISTINCT name FROM location WHERE cte = %u OR cte = 7 ORDER BY name;", $user_cte)
			EndIf
			$aResult = _EzMySql_GetTable2d($query)
			If IsArray($aResult) And _EzMySql_Rows() > 0 Then
				GUICtrlSetData($ComboAddLocation, _ArrayToString($aResult, '', 1, -1, '|'))
			EndIf
			$ButtonAddAdd = GUICtrlCreateButton('Добавить', 360, 178, 110, 31)
			GUISetState(@SW_SHOW, $hAddWnd)
			While 1
				$addMsg = GUIGetMsg()
				Switch $addMsg
					Case $GUI_EVENT_CLOSE
						ExitLoop
					Case $ButtonAddAdd
						$nsz = 0
						If GUICtrlRead($checkboxAddNSZ) = $GUI_CHECKED Then
							$nsz = 1
						EndIf
						If GUICtrlRead($LabelAddVendor) <> '' And GUICtrlRead($LabelAddObject) <> '' _
								And GUICtrlRead($InputAddSN) <> '' And GUICtrlRead($InputAddTime) <> '' _
								And _GUICtrlComboBox_GetEditText($ComboAddStatus) <> '' And _GUICtrlComboBox_GetEditText($ComboAddLocation) <> '' Then
							If _CheckLocation(_GUICtrlComboBox_GetEditText($ComboAddLocation)) Then
								If $type_mode Then
									check_params($vendor, $module, $params)
								EndIf
								If _GUICtrlComboBox_GetEditText($ComboSN) = '' Then ; Adding new equipment
									$query = StringFormat("SELECT serial FROM inventory WHERE vendor = (SELECT id FROM vendor WHERE name LIKE '%s') AND serial LIKE '%s';", GUICtrlRead($LabelAddVendor), GUICtrlRead($InputAddSN))
									$aResult = _EzMySql_GetTable2d($query)
									If IsArray($aResult) And _EzMySql_Rows() > 0 And GUICtrlRead($InputAddSN) <> 'б/н' Then
										MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), 'Оборудование с таким серийным номером уже есть в базе.')
									Else
										$query = StringFormat("INSERT INTO inventory (time, vendor, object, serial, place, status, user, location, nomnum, invnum, nsz) " & _
												"VALUES (%u, (SELECT id FROM vendor WHERE name LIKE '%s'), '%s', '%s', '%s', (SELECT id FROM status WHERE name LIKE '%s'), %u, '%s', '%s', '%s', %u);", _
												_toEPOCH(GUICtrlRead($InputAddTime)), _				; time
												GUICtrlRead($LabelAddVendor), _						; vendor
												GUICtrlRead($LabelAddObject), _						; object
												GUICtrlRead($InputAddSN), _							; serial
												GUICtrlRead($InputAddComment), _					; place
												_GUICtrlComboBox_GetEditText($ComboAddStatus), _ 	; status
												$user_id, _											; user
												_GUICtrlComboBox_GetEditText($ComboAddLocation), _ 	; location
												GUICtrlRead($InputAddNN), _							; nomnum
												GUICtrlRead($InputAddIN), _							; invnum
												$nsz) ; nsz
										_log($query)
										_EzMySql_Exec($query)
										_SetInfo()
										If _GUICtrlComboBox_GetEditText($ComboSN) = '' Then
											_SetSN()
										EndIf
										ExitLoop
									EndIf
								Else ; Move the equipment
									$query = StringFormat("INSERT INTO inventory (time, vendor, object, serial, place, status, user, location, nomnum, invnum, nsz) " & _
											"VALUES (%u, (SELECT id FROM vendor WHERE name LIKE '%s'), '%s', '%s', '%s', (SELECT id FROM status WHERE name LIKE '%s'), %u, '%s', '%s', '%s', %u);", _
											_toEPOCH(GUICtrlRead($InputAddTime)), _				; time
											GUICtrlRead($LabelAddVendor), _						; vendor
											GUICtrlRead($LabelAddObject), _						; object
											GUICtrlRead($InputAddSN), _							; serial
											GUICtrlRead($InputAddComment), _					; place
											_GUICtrlComboBox_GetEditText($ComboAddStatus), _ 	; status
											$user_id, _											; user
											_GUICtrlComboBox_GetEditText($ComboAddLocation), _ 	; location
											GUICtrlRead($InputAddNN), _							; nomnum
											GUICtrlRead($InputAddIN), _							; invnum
											$nsz) ; nsz
									_log($query)
									_EzMySql_Exec($query)
									_SetInfo()
									If _GUICtrlComboBox_GetEditText($ComboSN) = '' Then
										_SetSN()
									EndIf
									ExitLoop
								EndIf
							EndIf
						ElseIf GUICtrlRead($InputAddSN) = '' Then
							MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), 'Введите серийный номер устройства.')
						Else
							MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), 'Поля Статус и Площадка должны быть заполнены.')
						EndIf
				EndSwitch
			WEnd
			GUIDelete($hAddWnd)
		EndIf
	Else
		MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), _
				'Пожалуйста, выберите производителя из списка или' & @CRLF & _
				'введите в поле "Тип или производитель" название нового вендора.')
		ControlFocus($hWnd, '', $ComboVendor)
	EndIf
EndFunc   ;==>add_module

Func add_type($vendor_id)
	Local $module = _GUICtrlComboBox_GetEditText($ComboModule)
	If $module Then
		$string = StringSplit($module, ' ')
		$module = _ArrayPop($string)
		$vendor = _ArrayToString($string, ' ', 1)
		_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name = '%s';", $vendor))
		$vendor_id = _EzMySql_FetchData() ; array
	Else
		$module = ''
		$vendor = ''
	EndIf
	Local $hAddWnd = GUICreate(StringFormat('Добавление - %s', $ProgramName), 480, 310, -1, -1, -1, -1, $hWnd)
	GUICtrlCreateLabel('Производитель', 10, 10, 110, 21)
	GUICtrlCreateLabel('Карта/Модуль', 10, 40, 110, 21)
	GUICtrlCreateLabel('Серийный номер', 10, 70, 110, 21)
	GUICtrlCreateLabel('Ном. №', 250, 70, 40, 21)
	GUICtrlCreateLabel('Инв. №', 365, 70, 40, 21)
	GUICtrlCreateLabel('Дата и время', 10, 100, 110, 21)
	GUICtrlCreateLabel('Комментарий', 10, 130, 110, 21)
	GUICtrlCreateLabel('Статус', 10, 160, 110, 21)
	GUICtrlCreateLabel('Площадка', 10, 190, 110, 21)
	GUICtrlCreateGroup('Параметры', 10, 220, 460, 80)
	Local $params = ''
	For $i = 0 To $number_options - 1
		If _GUICtrlComboBox_GetEditText($ComboOptions[$i]) <> '' Then
			$params &= GUICtrlRead($LabelOptions[$i]) & ': ' & _GUICtrlComboBox_GetEditText($ComboOptions[$i]) & '; '
		EndIf
	Next
	If Not $params Then
		If GUICtrlRead($LabelOptions[0]) Then
			MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), StringFormat('Установите парамеры оборудования (%s, %s, ...)', GUICtrlRead($LabelOptions[0]), GUICtrlRead($LabelOptions[1])), 5)
		EndIf
		Return
	EndIf
	$params = StringTrimRight($params, 2)
	$LabelAddOptions = GUICtrlCreateLabel($params, 20, 240, 440, 50)
	; Поле вендора
	If $vendor Then
		$LabelAddVendor = GUICtrlCreateLabel($vendor, 130, 10, 225, 21)
	Else
		$LabelAddVendor = GUICtrlCreateCombo("", 130, 7, 225, 21)
		$query = StringFormat("SELECT equip_id FROM equipment WHERE type_id = %u;", $vendor_id[0])
		$equip_id = _EzMySql_GetTable2d($query)
		Local $equip_id_unique[1]
		$equip_id_unique[0] = ''
		$k = 0
		For $i = 1 To UBound($equip_id) - 1
			$string = StringSplit($equip_id[$i][0], ' ')
			_ArrayPop($string)
			$tmp = _ArrayToString($string, ' ', 1)
			If $equip_id_unique[$k] <> $tmp Then
				_ArrayAdd($equip_id_unique, $tmp)
				$k += 1
			EndIf
		Next
		GUICtrlSetData($LabelAddVendor, _ArrayToString($equip_id_unique, '|', 1), '')
	EndIf
	; Поле модуля
	If $module Then
		$LabelAddObject = GUICtrlCreateLabel($module, 130, 40, 340, 21)
	Else
		$LabelAddObject = GUICtrlCreateInput("", 130, 37, 225, 21)
	EndIf
	; Поле серийного номера
	If _GUICtrlComboBox_GetEditText($ComboSN) = '' Then
		$InputAddSN = GUICtrlCreateInput('', 130, 67, 110, 21)
	Else
		$InputAddSN = GUICtrlCreateLabel(_GUICtrlComboBox_GetEditText($ComboSN), 130, 70, 110, 21)
	EndIf
	; Поля номенклатурного и инвентарного номеров
	If (_GUICtrlComboBox_GetEditText($ComboNN) = '' And _GUICtrlComboBox_GetEditText($ComboIN) = '') Or _GUICtrlComboBox_GetEditText($ComboSN) = '' Then
		$InputAddNN = GUICtrlCreateInput('', 300, 67, 55, 21)
		$InputAddIN = GUICtrlCreateInput('', 415, 67, 55, 21)
	Else
		$InputAddNN = GUICtrlCreateLabel(_GUICtrlComboBox_GetEditText($ComboNN), 300, 70, 55, 21)
		$InputAddIN = GUICtrlCreateLabel(_GUICtrlComboBox_GetEditText($ComboIN), 415, 70, 55, 21)
	EndIf
	$InputAddTime = GUICtrlCreateInput(_NowCalc(), 130, 97, 110, 21)
	$InputAddComment = GUICtrlCreateInput('', 130, 127, 340, 21)
	If _GUICtrlComboBox_GetEditText($ComboSN) Then
		ControlFocus($hAddWnd, '', $InputAddComment)
	EndIf
	$ComboAddStatus = GUICtrlCreateCombo('', 130, 157, 169, 21)
	$checkboxAddNSZ = GUICtrlCreateCheckbox('НСЗ', 309, 157, 41, 21)
	$aResult = _EzMySql_GetTable2d("SELECT name FROM status ORDER BY name;")
	If IsArray($aResult) And _EzMySql_Rows() > 0 Then
		GUICtrlSetData($ComboAddStatus, _ArrayToString($aResult, '', 1, -1, '|'))
	EndIf
	$ComboAddLocation = GUICtrlCreateCombo('', 130, 187, 220, 21)
	If $user_cte = 7 Then
		$query = "SELECT DISTINCT name FROM location ORDER BY name;"
	Else
		$query = StringFormat("SELECT DISTINCT name FROM location WHERE cte = %u OR cte = 7 ORDER BY name;", $user_cte)
	EndIf
	$aResult = _EzMySql_GetTable2d($query)
	If IsArray($aResult) And _EzMySql_Rows() > 0 Then
		GUICtrlSetData($ComboAddLocation, _ArrayToString($aResult, '', 1, -1, '|'))
	EndIf
	Local $ButtonAddAdd = GUICtrlCreateButton('Добавить', 360, 178, 110, 31)
	GUISetState(@SW_SHOW, $hAddWnd)
	While 1
		$addMsg = GUIGetMsg()
		Switch $addMsg
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $ButtonAddAdd
				$nsz = 0
				If GUICtrlRead($checkboxAddNSZ) = $GUI_CHECKED Then
					$nsz = 1
				EndIf
				If GUICtrlRead($LabelAddVendor) <> '' And GUICtrlRead($LabelAddObject) <> '' _
						And GUICtrlRead($InputAddSN) <> '' And GUICtrlRead($InputAddTime) <> '' _
						And _GUICtrlComboBox_GetEditText($ComboAddStatus) <> '' And _GUICtrlComboBox_GetEditText($ComboAddLocation) <> '' Then
					If _CheckLocation(_GUICtrlComboBox_GetEditText($ComboAddLocation)) And _CheckVendor(GUICtrlRead($LabelAddVendor), _GUICtrlComboBox_GetEditText($ComboVendor)) Then
						check_params(GUICtrlRead($LabelAddVendor), StringReplace(GUICtrlRead($LabelAddObject), " ", "_"), $params)
						If _GUICtrlComboBox_GetEditText($ComboSN) = '' Then ; Adding new equipment
							$query = StringFormat("SELECT serial FROM inventory WHERE vendor = (SELECT id FROM vendor WHERE name LIKE '%s') AND serial LIKE '%s';", GUICtrlRead($LabelAddVendor), GUICtrlRead($InputAddSN))
							$aResult = _EzMySql_GetTable2d($query)
							If IsArray($aResult) And _EzMySql_Rows() > 0 And GUICtrlRead($InputAddSN) <> 'б/н' Then
								MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), 'Оборудование с таким серийным номером уже есть в базе.')
							Else
								$query = StringFormat("INSERT INTO inventory (time, vendor, object, serial, place, status, user, location, nomnum, invnum, nsz) " & _
										"VALUES (%u, (SELECT id FROM vendor WHERE name LIKE '%s'), '%s', '%s', '%s', (SELECT id FROM status WHERE name LIKE '%s'), %u, '%s', '%s', '%s', %u);", _
										_toEPOCH(GUICtrlRead($InputAddTime)), _					; time
										GUICtrlRead($LabelAddVendor), _							; vendor
										StringReplace(GUICtrlRead($LabelAddObject), " ", "_"), _; object
										GUICtrlRead($InputAddSN), _								; serial
										GUICtrlRead($InputAddComment), _						; place
										_GUICtrlComboBox_GetEditText($ComboAddStatus), _ 		; status
										$user_id, _												; user
										_GUICtrlComboBox_GetEditText($ComboAddLocation), _ 		; location
										GUICtrlRead($InputAddNN), _								; nomnum
										GUICtrlRead($InputAddIN), _								; invnum
										$nsz) ; nsz
								_log($query)
								_EzMySql_Exec($query)
								_SetInfo()
								If _GUICtrlComboBox_GetEditText($ComboSN) = '' Then
									_SetSN()
								EndIf
								ExitLoop
							EndIf
						Else ; Move the equipment
							$query = StringFormat("INSERT INTO inventory (time, vendor, object, serial, place, status, user, location, nomnum, invnum, nsz) " & _
									"VALUES (%u, (SELECT id FROM vendor WHERE name LIKE '%s'), '%s', '%s', '%s', (SELECT id FROM status WHERE name LIKE '%s'), %u, '%s', '%s', '%s', %u);", _
									_toEPOCH(GUICtrlRead($InputAddTime)), _					; time
									GUICtrlRead($LabelAddVendor), _							; vendor
									StringReplace(GUICtrlRead($LabelAddObject), " ", "_"), _; object
									GUICtrlRead($InputAddSN), _								; serial
									GUICtrlRead($InputAddComment), _						; place
									_GUICtrlComboBox_GetEditText($ComboAddStatus), _ 		; status
									$user_id, _												; user
									_GUICtrlComboBox_GetEditText($ComboAddLocation), _ 		; location
									GUICtrlRead($InputAddNN), _								; nomnum
									GUICtrlRead($InputAddIN), _								; invnum
									$nsz) ; nsz
							_log($query)
							_EzMySql_Exec($query)
							_SetInfo()
							If _GUICtrlComboBox_GetEditText($ComboSN) = '' Then
								_SetSN()
							EndIf
							ExitLoop
						EndIf
					EndIf
				ElseIf GUICtrlRead($InputAddSN) = '' Then
					MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), 'Введите серийный номер устройства.')
				Else
					MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), 'Поля Статус и Площадка должны быть заполнены.')
				EndIf
		EndSwitch
	WEnd
	GUIDelete($hAddWnd)
EndFunc   ;==>add_type

Func _Edit() ; Edit information in database.
	Local $data = GUICtrlRead(GUICtrlRead($ListView)) ; Twice. Once - some number, like '101'. Twice string with '|'.
	If $data <> '' Then
		Dim $dataSplit[11]
		$dataSplit = StringSplit($data, '|')
		Local $eTime = $dataSplit[1]
		Local $eVendor = $dataSplit[2]
		Local $eModule = $dataSplit[3]
		Local $eSerial = $dataSplit[4]
		Local $eNomnum = $dataSplit[5]
		Local $eInvnum = $dataSplit[6]
		Local $eComment = $dataSplit[7]
		Local $eStatus = $dataSplit[8]
		Local $eNSZ = $dataSplit[9]
		$nsz = 0
		If $dataSplit[9] = 'Да' Then
			$nsz = 1
		EndIf
		Local $eLocation = $dataSplit[10]
		Local $hAddWnd = GUICreate(StringFormat('Редактирование - %s', $ProgramName), 480, 230, -1, -1, -1, -1, $hWnd)
		GUICtrlCreateLabel('Производитель', 10, 10, 110, 21)
		GUICtrlCreateLabel('Карта/Модуль', 10, 40, 110, 21)
		GUICtrlCreateLabel('Серийный номер', 10, 70, 110, 21)
		GUICtrlCreateLabel('Ном. №', 250, 70, 40, 21)
		GUICtrlCreateLabel('Инв. №', 365, 70, 40, 21)
		GUICtrlCreateLabel('Дата и время', 10, 100, 110, 21)
		GUICtrlCreateLabel('Комментарий', 10, 130, 110, 21)
		GUICtrlCreateLabel('Статус', 10, 160, 110, 21)
		GUICtrlCreateLabel('Площадка', 10, 190, 110, 21)
		$LabelAddVendor = GUICtrlCreateLabel($eVendor, 130, 10, 225, 21)
		$LabelAddObject = GUICtrlCreateCombo($eModule, 130, 37, 225, 21)
		$aResult = _EzMySql_GetTable2d(StringFormat("SELECT DISTINCT object FROM inventory WHERE vendor = (SELECT id FROM vendor WHERE name LIKE '%s') ORDER BY object;", $eVendor))
		$iRows = _EzMySql_Rows()
		If IsArray($aResult) And $iRows > 0 Then
			GUICtrlSetData($LabelAddObject, '')
			GUICtrlSetData($LabelAddObject, _ArrayToString($aResult, '', 1, -1, '|'), $eModule)
		EndIf
		$InputAddSN = GUICtrlCreateInput($eSerial, 130, 67, 110, 21)
		$InputAddNN = GUICtrlCreateInput($eNomnum, 300, 67, 55, 21)
		$InputAddIN = GUICtrlCreateInput($eInvnum, 415, 67, 55, 21)
		$InputAddTime = GUICtrlCreateInput($eTime, 130, 97, 110, 21)
		$InputAddComment = GUICtrlCreateInput($eComment, 130, 127, 340, 21)
		$ComboAddStatus = GUICtrlCreateCombo('', 130, 157, 169, 21)
		$checkboxAddNSZ = GUICtrlCreateCheckbox('НСЗ', 309, 157, 41, 21)
		If $nsz = 1 Then
			GUICtrlSetState($checkboxAddNSZ, $GUI_CHECKED)
		EndIf
		$aResult = _EzMySql_GetTable2d("SELECT name FROM status ORDER BY name;")
		If IsArray($aResult) And _EzMySql_Rows() > 0 Then
			GUICtrlSetData($ComboAddStatus, _ArrayToString($aResult, '', 1, -1, '|'), $eStatus)
		EndIf
		$ComboAddLocation = GUICtrlCreateCombo('', 130, 187, 220, 21)
		If $user_cte = 7 Then
			$query = "SELECT DISTINCT name FROM location ORDER BY name;"
		Else
			$query = StringFormat("SELECT DISTINCT name FROM location WHERE cte = %u OR cte = 7 ORDER BY name;", $user_cte)
		EndIf
		$aResult = _EzMySql_GetTable2d($query)
		If IsArray($aResult) And _EzMySql_Rows() > 0 Then
			GUICtrlSetData($ComboAddLocation, _ArrayToString($aResult, '', 1, -1, '|'), $eLocation)
		EndIf
		$ButtonAddEdit = GUICtrlCreateButton('Сохранить', 360, 152, 110, 31)
		$ButtonAddDelete = GUICtrlCreateButton('Удалить', 360, 187, 110, 21)
		GUISetState(@SW_SHOW, $hAddWnd)
		While 1
			$addMsg = GUIGetMsg()
			Switch $addMsg
				Case $GUI_EVENT_CLOSE
					ExitLoop
				Case $ButtonAddEdit
					If GUICtrlRead($InputAddTime) <> '' _
							And _GUICtrlComboBox_GetEditText($ComboAddStatus) <> '' _
							And _GUICtrlComboBox_GetEditText($ComboAddLocation) <> '' _
							And GUICtrlRead($LabelAddObject) <> '' _
							And GUICtrlRead($InputAddSN) <> '' Then
						If _CheckLocation(_GUICtrlComboBox_GetEditText($ComboAddLocation)) Then
							If $eComment = "" Then
								$eComment = "IS NOT NULL"
							Else
								$eComment = StringFormat("LIKE '%s'", $eComment)
							EndIf
							$nsz = 0
							If GUICtrlRead($checkboxAddNSZ) = $GUI_CHECKED Then
								$nsz = 1
							EndIf
							$query = StringFormat("UPDATE inventory SET time = %u, object = '%s', serial = '%s', place = '%s', " & _
									"status = (SELECT id FROM status WHERE name LIKE '%s'), user = %u, location = '%s', nomnum = '%s', invnum = '%s', nsz = %u " & _
									"WHERE time = %u AND vendor = (SELECT id FROM vendor WHERE name LIKE '%s') AND object LIKE '%s' AND serial LIKE '%s' " & _
									"AND place %s AND status = (SELECT id FROM status WHERE name LIKE '%s');", _
									_toEPOCH(GUICtrlRead($InputAddTime)), _
									GUICtrlRead($LabelAddObject), _
									GUICtrlRead($InputAddSN), _
									GUICtrlRead($InputAddComment), _
									_GUICtrlComboBox_GetEditText($ComboAddStatus), _
									$user_id, _
									_GUICtrlComboBox_GetEditText($ComboAddLocation), _
									GUICtrlRead($InputAddNN), _
									GUICtrlRead($InputAddIN), _
									$nsz, _
									_toEPOCH($eTime), _
									$eVendor, _
									$eModule, _
									$eSerial, _
									$eComment, _
									$eStatus)
							_log($query)
							_EzMySql_Exec($query)
							_SetInfo()
							ExitLoop
						EndIf
					Else
						MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), 'Необходимо заполнить все поля формы.')
					EndIf
				Case $ButtonAddDelete
					Local $dialog = MsgBox(308, StringFormat('Сообщение - %s', $ProgramName), 'Вы хотите удалить запись?') ;4 Yes and No, 48 Exclamation-point icon, 256 Second button is default button
					If $dialog = 6 Then
						If $eComment = "" Then
							$eComment = "IS NOT NULL"
						Else
							$eComment = StringFormat("LIKE '%s'", $eComment)
						EndIf
						$query = StringFormat("DELETE FROM inventory WHERE " & _
								"time = %u AND " & _
								"vendor = (SELECT id FROM vendor WHERE name LIKE '%s') AND " & _
								"object LIKE '%s' AND " & _
								"serial LIKE '%s' AND " & _
								"place %s AND " & _
								"status = (SELECT id FROM status WHERE name LIKE '%s');", _
								_toEPOCH($eTime), _
								$eVendor, _
								$eModule, _
								$eSerial, _
								$eComment, _
								$eStatus)
						_log($query)
						_EzMySql_Exec($query)
						_SetInfo()
						ExitLoop
					EndIf
			EndSwitch
		WEnd
		GUIDelete($hAddWnd)
	Else
		MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), 'Пожалуйста, выберите строку данных для редактирования.')
	EndIf
EndFunc   ;==>_Edit

Func _Report()
	Local $i, $n
	Local $n = UBound($ListViewItem)
	If $n > 0 Then
		Local $path = FileSaveDialog(StringFormat('Сохрнение отчёта - %s', $ProgramName), @MyDocumentsDir, 'CSV (разделитель - точка с запятой) (*.csv)', 0, StringFormat('%s-%s-%s %s.%s.%s %s.csv', @YEAR, @MON, @MDAY, @HOUR, @MIN, @SEC, $ProgramName), $hWnd)
		If @error Then
			ConsoleWrite('No file was saved.' & @CRLF)
		Else
			ConsoleWrite($path & @CRLF)
			Local $hFileOpen = FileOpen($path, 514)
			If $hFileOpen = -1 Then
				ConsoleWrite('An error occurred when reading the file.' & @CRLF)
				Return False
			EndIf
			FileWriteLine($hFileOpen, 'Дата и время;Производитель;Оборудование;Серийный №;Ном. №;Инв. №;Комментарий;Статус;НСЗ;Площадка')
			For $i = 0 To $n - 1
				FileWriteLine($hFileOpen, StringTrimRight(StringReplace(StringReplace(GUICtrlRead($ListViewItem[$i]), ';', '";"'), '|', ';'), 1))
			Next
			FileClose($hFileOpen)
		EndIf
	EndIf
EndFunc   ;==>_Report

Func _Filter($hist = False)
	$timer = TimerInit()
	If Not $hist Then history_add()
	GUICtrlSetState($ButtonFilter, $GUI_DISABLE)
	GUICtrlSetData($InputNumber, 'Пожалуйста, подождите...')
	GUICtrlSetData($ButtonAdd, 'Добавить')
	GUICtrlSetState($ButtonEdit, $GUI_DISABLE)
	_GUICtrlListView_DeleteAllItems($ListView)
	Local $count = 0
	Local $vendor = _GUICtrlComboBox_GetEditText($ComboVendor)
	_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name = '%s';", $vendor))
	$vendor_id = _EzMySql_FetchData()
	If IsArray($vendor_id) And $vendor_id[0] > 1000 Then
		If _GUICtrlComboBox_GetEditText($ComboModule) Then
			Dim $aRes[2][1]
			$aRes[0][0] = 'equip_id'
			$aRes[1][0] = _GUICtrlComboBox_GetEditText($ComboModule)
			$iRows = 1
		Else
			Dim $options[1]
			For $i = 0 To UBound($ComboOptions) - 1
				$option_value = _GUICtrlComboBox_GetEditText($ComboOptions[$i])
				If $option_value Then
					_EzMySql_Query(StringFormat("SELECT option_id FROM options " & _
							"WHERE name = '%s' AND type_id = %u;", _
							GUICtrlRead($LabelOptions[$i]), $vendor_id[0]))
					$option_id = _EzMySql_FetchData()
					_ArrayAdd($options, $option_id[0] & ': "' & $option_value & '"')
				EndIf
			Next
			If UBound($options) = 1 Then
				$query = "SELECT equip_id FROM equipment;"
			Else
				$query = "SELECT equip_id FROM equipment WHERE "
				For $i = 1 To UBound($options) - 1
					$query &= StringFormat("options LIKE '%%%s%%' AND ", $options[$i])
				Next
				$query = StringTrimRight($query, 5) & ";"
			EndIf
			Local $aRes = _EzMySql_GetTable2d($query)
			$iRows = _EzMySql_Rows()
		EndIf
		$query = StringFormat("SELECT equip_id, options FROM equipment WHERE equip_id IN ('%s') AND type_id = %u", _ArrayToString($aRes, "|", 1, -1, "', '"), $vendor_id[0])
		Local $aOptions = _EzMySql_GetTable2d($query)
		_EzMySql_Query(StringFormat("SELECT COUNT(*) FROM options WHERE type_id = %u", $vendor_id[0]))
		$numbers_options = _EzMySql_FetchData()
		$equip_count = $iRows
		Dim $aResultAll[1][11]
		For $i = 1 To $equip_count
			GUICtrlSetData($InputNumber, 'Пожалуйста, подождите... ' & _
					Round($i / $equip_count * 100) & '%')
			$tmp = StringSplit($aRes[$i][0], ' ')
			$object = _ArrayPop($tmp)
			$vendor = _ArrayToString($tmp, ' ', 1)
			Local $serial = _GUICtrlComboBox_GetEditText($ComboSN)
			If $serial Then
				$serial = StringFormat("WHERE serial LIKE '%%%s%%'", $serial)
			EndIf
			Local $status = _GUICtrlComboBox_GetEditText($ComboStatus)
			If $status Then
				$status = StringFormat("AND tab1.status = (SELECT id FROM status WHERE name = '%s')", $status)
			EndIf
			Local $invnum = _GUICtrlComboBox_GetEditText($ComboIN)
			If $invnum Then
				$invnum = StringFormat("AND tab1.invnum LIKE '%%%s%%'", $invnum)
			EndIf
			Local $nomnum = _GUICtrlComboBox_GetEditText($ComboNN)
			If $nomnum Then
				$nomnum = StringFormat("AND tab1.nomnum LIKE '%%%s%%'", $nomnum)
			EndIf
			Local $nsz = _GUICtrlComboBox_GetEditText($ComboNSZ)
			If $nsz = 'Да' Then
				$nsz = 'AND tab1.nsz = 1'
			ElseIf $nsz = 'Нет' Then
				$nsz = 'AND tab1.nsz = 0'
			EndIf
			Local $location = _GUICtrlComboBox_GetEditText($ComboNE)
			If $location Then
				$location = StringFormat("AND tab1.location = '%s'", $location)
			EndIf
			Local $cte_location = ''
;~ 			If $user_cte <> 7 Then
;~ 				$cte_location = StringFormat("AND location IN (SELECT name FROM location WHERE cte = %u OR cte = 7)", $user_cte)
;~ 			EndIf
			$query = StringFormat("SELECT tab1.time, " & _
					"(SELECT name FROM vendor WHERE id = tab1.vendor), " & _
					"tab1.object, " & _
					"tab1.serial, " & _
					"tab1.place, " & _
					"(SELECT name FROM status WHERE id = tab1.status), " & _
					"tab1.user, " & _
					"tab1.location, " & _
					"tab1.invnum, " & _
					"tab1.nomnum, " & _
					"tab1.nsz FROM inventory tab1 " & _
					"INNER JOIN (SELECT MAX(time) as now, serial FROM inventory %s GROUP BY serial, object) tab2 " & _
					"ON (tab1.serial = tab2.serial AND tab1.time = tab2.now) " & _
					"WHERE tab1.vendor = (SELECT id FROM vendor WHERE name = '%s') " & _
					"AND tab1.object = '%s' %s %s %s %s %s %s " & _ ; Точное совпадение. Раньше было LIKE '%%%s%%'. Но бывало задвоение при похожих моделях оборудования.
					"ORDER BY tab1.object, tab1.serial;", _
					$serial, $vendor, $object, $status, $invnum, $nomnum, $nsz, $location, $cte_location)
			$aResult = _EzMySql_GetTable2d($query)
			$iRows = _EzMySql_Rows()
			If IsArray($aResult) And $iRows > 0 Then
				_ArrayConcatenate($aResultAll, $aResult, 1)
			EndIf
		Next
		GUICtrlSetData($InputNumber, 'Пожалуйста, подождите...')
		Local $vQuery, $sQuery
		Dim $ListViewItem[UBound($aResultAll)]
		For $i = 1 To UBound($aResultAll) - 1
			$nsz = 'Нет'
			If $aResultAll[$i][10] = 1 Then
				$nsz = 'Да'
			EndIf
			For $j = 1 To UBound($aOptions) - 1
				If $aOptions[$j][0] = $aResultAll[$i][1] & ' ' & $aResultAll[$i][2] Then
					$options = $aOptions[$j][1]
					ExitLoop
				EndIf
			Next
			$aOptionsValues = _options_to_array($options)
			$sOptions = ''
			For $j = 1 To $numbers_options[0]
				$sOptions &= '|' & $aOptionsValues[$j]
			Next
			$ListViewItem[$i - 1] = GUICtrlCreateListViewItem( _
					_EPOCH($aResultAll[$i][0]) & '|' & _
					$aResultAll[$i][1] & '|' & _
					$aResultAll[$i][2] & '|' & _
					$aResultAll[$i][3] & '|' & _
					$aResultAll[$i][9] & '|' & _
					$aResultAll[$i][8] & '|' & _
					$aResultAll[$i][4] & '|' & _
					$aResultAll[$i][5] & '|' & _
					$nsz & '|' & _
					$aResultAll[$i][7] & _
					$sOptions, $ListView)
		Next
		GUICtrlSetData($InputNumber, UBound($aResultAll) - 1)
		GUICtrlSetState($ButtonReport, $GUI_ENABLE)
	Else
		If $vendor = '' Then
			$vendor = 'IS NOT NULL'
		Else
			$vendor = '= (SELECT id FROM vendor WHERE name LIKE ''' & $vendor & ''')'
		EndIf
		Local $object = _GUICtrlComboBox_GetEditText($ComboModule)
		If $object = '' Then
			$object = 'IS NOT NULL'
		Else
			$object = 'LIKE ''%' & $object & '%'''
		EndIf
		Local $serial = _GUICtrlComboBox_GetEditText($ComboSN)
		If $serial = '' Then
			$serial = 'IS NOT NULL'
		Else
			$serial = 'LIKE ''%' & $serial & '%'''
		EndIf
		Local $location = _GUICtrlComboBox_GetEditText($ComboNE)
		If $location = '' Then
			$location = 'IS NOT NULL'
		Else
			$location = 'LIKE ''' & $location & ''''
		EndIf
		Local $status = _GUICtrlComboBox_GetEditText($ComboStatus)
		If $status = '' Then
			$status = 'IS NOT NULL'
		Else
			$status = '= (SELECT id FROM status WHERE name LIKE ''' & $status & ''')'
		EndIf
		Local $nsz = _GUICtrlComboBox_GetEditText($ComboNSZ)
		If $nsz = 'Да' Then
			$nsz = '= 1'
		ElseIf $nsz = 'Нет' Then
			$nsz = '= 0'
		Else
			$nsz = 'IS NOT NULL'
		EndIf
		Local $invnum = _GUICtrlComboBox_GetEditText($ComboIN)
		If $invnum = '' Then
			$invnum = 'IS NOT NULL'
		Else
			$invnum = 'LIKE ''%' & $invnum & '%'''
		EndIf
		Local $nomnum = _GUICtrlComboBox_GetEditText($ComboNN)
		If $nomnum = '' Then
			$nomnum = 'IS NOT NULL'
		Else
			$nomnum = 'LIKE ''%' & $nomnum & '%'''
		EndIf
		Local $request
		If $user_cte = 7 Then
			$request = 'SELECT tab1.time, (SELECT name FROM vendor WHERE id = tab1.vendor), tab1.object, tab1.serial, tab1.place, (SELECT name FROM status WHERE id = tab1.status), tab1.user, tab1.location, tab1.invnum, tab1.nomnum, tab1.nsz FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial, object FROM inventory WHERE serial ' & $serial & ' AND object ' & $object & ' GROUP BY serial, object) tab2 ON (tab1.serial = tab2.serial AND tab1.time = tab2.now AND tab1.object = tab2.object) WHERE tab1.vendor ' & $vendor & ' AND tab1.object ' & $object & ' AND tab1.status ' & $status & ' AND tab1.invnum ' & $invnum & ' AND tab1.nomnum ' & $nomnum & ' AND tab1.nsz ' & $nsz & ' AND tab1.location ' & $location & ' ORDER BY tab1.object;'
		Else
			$request = 'SELECT tab1.time, (SELECT name FROM vendor WHERE id = tab1.vendor), tab1.object, tab1.serial, tab1.place, (SELECT name FROM status WHERE id = tab1.status), tab1.user, tab1.location, tab1.invnum, tab1.nomnum, tab1.nsz FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial, object FROM inventory WHERE serial ' & $serial & ' AND object ' & $object & ' GROUP BY serial, object) tab2 ON (tab1.serial = tab2.serial AND tab1.time = tab2.now AND tab1.object = tab2.object) WHERE tab1.vendor ' & $vendor & ' AND tab1.object ' & $object & ' AND tab1.status ' & $status & ' AND tab1.invnum ' & $invnum & ' AND tab1.nomnum ' & $nomnum & ' AND tab1.nsz ' & $nsz & ' AND tab1.location ' & $location & ' AND location IN (SELECT name FROM location WHERE cte = ' & $user_cte & ' OR cte = 7) ORDER BY tab1.object;'
		EndIf
;~ 		print($request)
		$aResult = _EzMySql_GetTable2d($request)
		$iRows = _EzMySql_Rows()
		If IsArray($aResult) And $iRows > 0 Then
			Local $vQuery, $sQuery
			Dim $ListViewItem[$iRows]
			For $i = 1 To $iRows
				$nsz = 'Нет'
				If $aResult[$i][10] = 1 Then
					$nsz = 'Да'
				EndIf
				$ListViewItem[$i - 1] = GUICtrlCreateListViewItem(_EPOCH($aResult[$i][0]) & '|' & $aResult[$i][1] & '|' & $aResult[$i][2] & '|' & $aResult[$i][3] & '|' & $aResult[$i][9] & '|' & $aResult[$i][8] & '|' & $aResult[$i][4] & '|' & $aResult[$i][5] & '|' & $nsz & '|' & $aResult[$i][7], $ListView)
			Next
			GUICtrlSetData($InputNumber, $iRows)
			GUICtrlSetState($ButtonReport, $GUI_ENABLE)
		Else
			GUICtrlSetData($InputNumber, 0)
		EndIf
	EndIf
	GUICtrlSetState($ButtonFilter, $GUI_ENABLE)
;~ 	print(TimerDiff($timer) & ' < _Filter()')
EndFunc   ;==>_Filter

Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $iMsg, $wParam
	Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR, $hWndListView, $tInfo
	; Local $tBuffer
	$hWndListView = $ListView
	If Not IsHWnd($ListView) Then $hWndListView = GUICtrlGetHandle($ListView)
	;$tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	$tNMHDR = DllStructCreate($tagNMLISTVIEW, $lParam)
	$hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	$iIDFrom = DllStructGetData($tNMHDR, "IDFrom")
	$iCode = DllStructGetData($tNMHDR, "Code")
	Switch $hWndFrom
		Case $hWndListView
			Switch $iCode
				Case $NM_DBLCLK ; Sent by a list-view control when the user double-clicks an item with the left mouse button
					$tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
					_GoToObject(GUICtrlRead(GUICtrlRead($ListView)))
					history_add() ; Don't True debug history
					_DebugPrint("$NM_DBLCLK" & @CRLF & "--> hWndFrom:" & @TAB & $hWndFrom & @CRLF & _
							"-->IDFrom:" & @TAB & $iIDFrom & @CRLF & _
							"-->Code:" & @TAB & $iCode & @CRLF & _
							"-->Index:" & @TAB & DllStructGetData($tInfo, "Index") & @CRLF & _
							"-->SubItem:" & @TAB & DllStructGetData($tInfo, "SubItem") & @CRLF & _
							"-->NewState:" & @TAB & DllStructGetData($tInfo, "NewState") & @CRLF & _
							"-->OldState:" & @TAB & DllStructGetData($tInfo, "OldState") & @CRLF & _
							"-->Changed:" & @TAB & DllStructGetData($tInfo, "Changed") & @CRLF & _
							"-->ActionX:" & @TAB & DllStructGetData($tInfo, "ActionX") & @CRLF & _
							"-->ActionY:" & @TAB & DllStructGetData($tInfo, "ActionY") & @CRLF & _
							"-->lParam:" & @TAB & DllStructGetData($tInfo, "lParam") & @CRLF & _
							"-->KeyFlags:" & @TAB & DllStructGetData($tInfo, "KeyFlags"))
				Case $LVN_COLUMNCLICK ; A column was clicked
					_GUICtrlListView_SimpleSort($hWndListView, $g_bSortSense, DllStructGetData($tNMHDR, "SubItem")) ; Sort direction for next sort toggled by default
					ConsoleWrite(DllStructGetData($tNMHDR, "SubItem") & @CRLF)
			EndSwitch
	EndSwitch
	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

Func _DebugPrint($s_Text, $sLine = @ScriptLineNumber)
	ConsoleWrite( _
			"!===========================================================" & @CRLF & _
			"+======================================================" & @CRLF & _
			"-->Line(" & StringFormat("%04d", $sLine) & "):" & @TAB & $s_Text & @CRLF & _
			"+======================================================" & @CRLF)
EndFunc   ;==>_DebugPrint

Func _GoToObject($line) ; Переход к оборудованию из списка фильтра.
	#cs
		Функция загружает в список информацию по одному выбраному (двойным щелчком) блоку
		Обновление 02.11.2016 - Перенос инвентарных и номенклатурных номеров (при их наличии) при добавлении состояния.
	#ce
	Local $split = StringSplit($line, '|')
	If $split[0] > 4 And @error <> 1 Then
		If $vendor_mode Then
			GUICtrlSetData($ComboVendor, $split[2])
			_SetModule($split[3])
		Else
			_SetModule($split[2] & ' ' & $split[3])
		EndIf
		_SetSN($split[4])
		ControlSetText($hWnd, '', $ComboStatus, '')
		ControlSetText($hWnd, '', $ComboNE, '')
		ControlSetText($hWnd, '', $ComboIN, $split[5]) ; up02.11.2016
		ControlSetText($hWnd, '', $ComboNN, $split[6]) ; up02.11.2016
		_SetInfo()
	EndIf
EndFunc   ;==>_GoToObject

Func _CheckLocation($location) ; Проверка существования узла связи и вопрос на добавление нового.
	Local $result = False
	Local $dialog
	If $user_cte = 7 Then
		$query = StringFormat("SELECT name FROM location WHERE name LIKE '%s';", $location)
	Else
		$query = StringFormat("SELECT name FROM location WHERE name LIKE '%s' AND (cte = %u OR cte = 7);", $location, $user_cte)
	EndIf
	$aResult = _EzMySql_GetTable2d($query)
	$iRows = _EzMySql_Rows()
	If IsArray($aResult) And $iRows = 0 Then
		$dialog = MsgBox(292, StringFormat('Сообщение - %s', $ProgramName), 'Площадки "' & $location & '" нет в базе данных.' & @CRLF & 'Вы хотите добавить новую площадку?') ;4 Yes and No, 32 Question-mark icon, 256 Second button is default button
		If $dialog = 6 Then
			Local $query = StringFormat("INSERT INTO location (ne, name, cte) VALUES ('%s', '%s', '%u');", $location, $location, $user_cte)
			_log($query)
			_EzMySql_Exec($query)
			If $user_cte = 7 Then
				$query = "SELECT DISTINCT name FROM location ORDER BY name;"
			Else
				$query = StringFormat("SELECT DISTINCT name FROM location WHERE cte = %u OR cte = 7 ORDER BY name;", $user_cte)
			EndIf
			$aResult = _EzMySql_GetTable2d($query)
			GUICtrlSetData($ComboNE, '')
			GUICtrlSetData($ComboNE, _ArrayToString($aResult, '', 1, -1, '|'))
			$result = True
		EndIf
	Else
		$result = True
	EndIf
	Return $result
EndFunc   ;==>_CheckLocation

Func _CheckVendor($vendor, $default = '') ; Проверка существования производителя оборудования и вопрос на добавление нового.
	Local $result = False
	Local $dialog
	$aResult = _EzMySql_GetTable2d(StringFormat("SELECT name FROM vendor WHERE name = '%s';", $vendor))
	$iRows = _EzMySql_Rows()
	If IsArray($aResult) And $iRows = 0 Then
		$dialog = MsgBox(292, StringFormat('Сообщение - %s', $ProgramName), 'Производителя "' & $vendor & '" нет в базе данных.' & @CRLF & 'Вы хотите добавить нового производителя?') ; 4 Yes and No, 32 Question-mark icon, 256 Second button is default button
		If $dialog = 6 Then
			_EzMySql_Query("SELECT MAX(id) FROM vendor WHERE id <= 1000;")
			$max_id_vendor = _EzMySql_FetchData()
			If $max_id_vendor[0] > 1000 Then
				MsgBox(0, StringFormat('Сообщение - %s', $ProgramName), 'В базе не может быть больше 1000 вендоров.')
			Else
				Local $query = StringFormat("INSERT INTO vendor (id, name) VALUES (%u, '%s');", $max_id_vendor[0] + 1, $vendor)
				_log($query)
				_EzMySql_Exec($query)
				$aResult = _EzMySql_GetTable2d("SELECT name FROM vendor ORDER BY name;")
				GUICtrlSetData($ComboVendor, '')
				GUICtrlSetData($ComboVendor, _ArrayToString($aResult, '', 1, -1, '|'), $default)
				$result = True
			EndIf
		EndIf
	Else
		$result = True
	EndIf
	Return $result
EndFunc   ;==>_CheckVendor

Func _ClearFilter() ; Очистка фильтра.
	GUICtrlSetData($ButtonAdd, 'Добавить')
	_GUICtrlComboBox_SetCurSel($ComboModule)
	_GUICtrlComboBox_SetCurSel($ComboSN)
	GUICtrlSetData($InputNumber, '')
	_GUICtrlComboBox_SetCurSel($ComboStatus)
	_GUICtrlComboBox_SetCurSel($ComboNSZ)
	_GUICtrlComboBox_SetCurSel($ComboNE)
	_GUICtrlComboBox_SetCurSel($ComboIN)
	_GUICtrlComboBox_SetCurSel($ComboNN)
	For $i = 0 To UBound($ComboOptions) - 1
		_GUICtrlComboBox_SetCurSel($ComboOptions[$i])
	Next
EndFunc   ;==>_ClearFilter

Func _Help() ; Справка о программе.
	If FileExists($help_file) Then
		ShellExecute($help_file)
	EndIf
EndFunc   ;==>_Help

Func _ToolsActiveUsers() ; Список активных пользователей или тех, кто коректно не закыл программу.
	Local $msg = 'Нет активных пользователей.'
	Local $query = "SELECT user.name, cte.name FROM user JOIN cte ON user.cte = cte.id  WHERE comp IN (SELECT comp FROM log WHERE id IN (SELECT MAX(id) FROM log GROUP BY comp) AND request NOT LIKE 'Stop') ORDER BY user.name;"
	$aResult = _EzMySql_GetTable2d($query)
	$iRows = _EzMySql_Rows()
	If IsArray($aResult) And $iRows > 0 Then
		$msg = $aResult[1][0] & ' (' & $aResult[1][1] & ')'
		For $i = 2 To $iRows
			$msg &= @CRLF & $aResult[$i][0] & ' (' & $aResult[$i][1] & ')'
		Next
	EndIf
	MsgBox(0, StringFormat('Активные пользователи - %s', $ProgramName), $msg)
EndFunc   ;==>_ToolsActiveUsers

Func _ToolsStatistic()
	Local $num_users ;Кол-во пользователей
	Local $num_vendors ;Кол-во вендоров
	Local $num_items ;Кол-во карт, модулей, блоков
	Local $num_locations ;Кол-во площадок
	Local $msg = '' ;Результирущая строка в сообщении
	_EzMySql_Query("SELECT COUNT(DISTINCT name) FROM user;")
	$num_users = _EzMySql_FetchData()
	_EzMySql_Query("SELECT COUNT(DISTINCT name) FROM vendor;")
	$num_vendors = _EzMySql_FetchData()
;~ 	_EzMySql_Query("SELECT COUNT(DISTINCT serial) FROM inventory;")
	_EzMySql_Query("SELECT COUNT(*) FROM (SELECT * FROM inventory GROUP BY serial, object) AS inv;")
	$num_items = _EzMySql_FetchData()
	_EzMySql_Query("SELECT COUNT(DISTINCT name) FROM location;")
	$num_locations = _EzMySql_FetchData()
	$msg = $num_users[0] & ' пользователей' & @CRLF & $num_vendors[0] & ' вендоров' & @CRLF & $num_items[0] & ' единиц оборудования' & @CRLF & $num_locations[0] & ' площадок'
	MsgBox(0, StringFormat('Статистика - %s', $ProgramName), $msg)
EndFunc   ;==>_ToolsStatistic

Func tools_equipment_in_work()
	$timer = TimerInit()
	$cte = _EzMySql_GetTable2d("SELECT name FROM cte ORDER BY id;")
	Local $query = 'SELECT (SELECT name FROM vendor WHERE id = vendor) as "Производитель", object as "Оборудование", SUM(cte1) as "' & $cte[1][0] & '", SUM(cte2) as "' & $cte[2][0] & '", SUM(cte3) as "' & $cte[3][0] & '", SUM(cte4) as "' & $cte[4][0] & '", SUM(cte5) as "' & $cte[5][0] & '",  SUM(cte8) as "' & $cte[8][0] & '",  SUM(cte9) as "' & $cte[9][0] & '", SUM(cte6) as "' & $cte[6][0] & '"' & @CRLF & _
			'FROM (' & @CRLF & _
			'SELECT vendor, object, COUNT(*) as cte1, 0 as cte2, 0 as cte3, 0 as cte4, 0 as cte5, 0 as cte8, 0 as cte9, 0 as cte6' & @CRLF & _
			'FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial FROM inventory GROUP BY serial, object) tab2 ' & _
			'ON (tab1.serial = tab2.serial AND tab1.time = tab2.now) WHERE tab1.status = 6 AND location IN (SELECT name FROM location WHERE cte = 1) GROUP BY tab1.object' & @CRLF & _
			'UNION' & @CRLF & _
			'SELECT vendor, object, 0 as cte1, COUNT(*) as cte2, 0 as cte3, 0 as cte4, 0 as cte5, 0 as cte8, 0 as cte9, 0 as cte6' & @CRLF & _
			'FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial FROM inventory GROUP BY serial, object) tab2 ' & _
			'ON (tab1.serial = tab2.serial AND tab1.time = tab2.now) WHERE tab1.status = 6 AND location IN (SELECT name FROM location WHERE cte = 2) GROUP BY tab1.object' & @CRLF & _
			'UNION' & @CRLF & _
			'SELECT vendor, object, 0 as cte1, 0 as cte2, COUNT(*) as cte3, 0 as cte4, 0 as cte5, 0 as cte8, 0 as cte9, 0 as cte6' & @CRLF & _
			'FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial FROM inventory GROUP BY serial, object) tab2 ' & _
			'ON (tab1.serial = tab2.serial AND tab1.time = tab2.now) WHERE tab1.status = 6 AND location IN (SELECT name FROM location WHERE cte = 3) GROUP BY tab1.object' & @CRLF & _
			'UNION' & @CRLF & _
			'SELECT vendor, object, 0 as cte1, 0 as cte2, 0 as cte3, COUNT(*) as cte4, 0 as cte5, 0 as cte8, 0 as cte9, 0 as cte6' & @CRLF & _
			'FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial FROM inventory GROUP BY serial, object) tab2 ' & _
			'ON (tab1.serial = tab2.serial AND tab1.time = tab2.now) WHERE tab1.status = 6 AND location IN (SELECT name FROM location WHERE cte = 4) GROUP BY tab1.object' & @CRLF & _
			'UNION' & @CRLF & _
			'SELECT vendor, object, 0 as cte1, 0 as cte2, 0 as cte3, 0 as cte4, COUNT(*) as cte5, 0 as cte8, 0 as cte9, 0 as cte6' & @CRLF & _
			'FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial FROM inventory GROUP BY serial, object) tab2 ' & _
			'ON (tab1.serial = tab2.serial AND tab1.time = tab2.now) WHERE tab1.status = 6 AND location IN (SELECT name FROM location WHERE cte = 5) GROUP BY tab1.object' & @CRLF & _
			'UNION' & @CRLF & _
			'SELECT vendor, object, 0 as cte1, 0 as cte2, 0 as cte3, 0 as cte4, 0 as cte5, COUNT(*) as cte8, 0 as cte9, 0 as cte6' & @CRLF & _
			'FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial FROM inventory GROUP BY serial, object) tab2 ' & _
			'ON (tab1.serial = tab2.serial AND tab1.time = tab2.now) WHERE tab1.status = 6 AND location IN (SELECT name FROM location WHERE cte = 8) GROUP BY tab1.object' & @CRLF & _
			'UNION' & @CRLF & _
			'SELECT vendor, object, 0 as cte1, 0 as cte2, 0 as cte3, 0 as cte4, 0 as cte5, 0 as cte8, COUNT(*) as cte9, 0 as cte6' & @CRLF & _
			'FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial FROM inventory GROUP BY serial, object) tab2 ' & _
			'ON (tab1.serial = tab2.serial AND tab1.time = tab2.now) WHERE tab1.status = 6 AND location IN (SELECT name FROM location WHERE cte = 9) GROUP BY tab1.object' & @CRLF & _
			'UNION' & @CRLF & _
			'SELECT vendor, object, 0 as cte1, 0 as cte2, 0 as cte3, 0 as cte4, 0 as cte5, 0 as cte8, 0 as cte9, COUNT(*) as cte6' & @CRLF & _
			'FROM inventory tab1 INNER JOIN (SELECT MAX(time) as now, serial FROM inventory GROUP BY serial, object) tab2 ' & _
			'ON (tab1.serial = tab2.serial AND tab1.time = tab2.now) WHERE tab1.status = 6 AND location IN (SELECT name FROM location WHERE cte = 6) GROUP BY tab1.object' & @CRLF & _
			') tall' & @CRLF & _
			'GROUP BY object ORDER BY `Производитель`, `Оборудование`;'
	$aResult = _EzMySql_GetTable2d($query)
;~ 	print(TimerDiff($timer) & ' < tools_equipment_in_work()')
	If IsArray($aResult) Then
		_ArrayDisplay($aResult, StringFormat('Распределение: в работе по ЦТЭ - %s', $ProgramName))
	EndIf
EndFunc   ;==>tools_equipment_in_work

Func tools_working_time_gui($user_cte)
	Local $hWorkTimeWnd = GUICreate(StringFormat('Учёт рабочего времени - %s', $ProgramName), 800, 460, -1, -1, -1, -1, $hWnd)
	GUICtrlCreateLabel('Период (дней)', 10, 16, 80, 21)
	Local $PeriodWorkTime = GUICtrlCreateInput('31', 90, 13, 50, 22)
	GUICtrlCreateUpdown($PeriodWorkTime)
	Local $WorkTimeButton = GUICtrlCreateButton("Показать", 150, 9, 100, 28)
	Local $WorkTimeListColumns = 'N|Подразделение|Имя|Кол-во сессий|Общее время|Внесено|Обновлено|Удалено'
	Local $WorkTimeList = GUICtrlCreateListView($WorkTimeListColumns, 10, 50, 780, 400, BitOR($LVS_SHOWSELALWAYS, $LVS_SINGLESEL))

	GUISetState(@SW_SHOW, $hWorkTimeWnd)
	While 1
		$workTimeMsg = GUIGetMsg()
		Switch $workTimeMsg
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $WorkTimeButton
				GUICtrlSetState($WorkTimeButton, $GUI_DISABLE)
				If Int(GUICtrlRead($PeriodWorkTime)) <= 0 Then
					GUICtrlSetData($PeriodWorkTime, 1)
				EndIf
				_GUICtrlListView_DeleteAllItems($WorkTimeList)
				Local $result = tools_working_time($user_cte, Int(GUICtrlRead($PeriodWorkTime)))
				For $k = 1 To $result[0][0]
					$line = StringFormat('%s|%s|%s|%s|%s|%s|%s|%s', _
							$result[$k][0], _
							$result[$k][1], _
							$result[$k][2], _
							$result[$k][3], _
							$result[$k][4], _
							$result[$k][5], _
							$result[$k][6], _
							$result[$k][7])
					GUICtrlCreateListViewItem($line, $WorkTimeList)
				Next
				GUICtrlSetState($WorkTimeButton, $GUI_ENABLE)
		EndSwitch
	WEnd
	GUIDelete($hWorkTimeWnd)
EndFunc   ;==>tools_working_time_gui

Func tools_working_time($user_cte, $period_days)
	Local $period = $period_days * 24 * 60 * 60 ; 31 day to seconds
	Local $stop_time = _DateDiff('s', "1970/01/01 00:00:00", _NowCalc())
	Local $start_time = $stop_time - $period
	Local $query = StringFormat("SELECT DISTINCT comp FROM log WHERE time > %u;", $start_time)
	If $user_cte <> 7 Then
		$query = StringFormat("SELECT DISTINCT comp FROM log WHERE time > %u AND comp IN (SELECT comp FROM user WHERE cte = %u);", $start_time, $user_cte)
	EndIf
	Local $comp = _EzMySql_GetTable2d($query)
	Local $comp_len = _EzMySql_Rows()
	Local $i, $update, $user, $sessions, $sessions_len, $sessions_count, $sessions_time, $sessions_mode
	Local $g_iHour, $g_iMins, $g_iSecs
	If IsArray($aResult) And $comp_len > 0 Then
;~ 		_ArrayDisplay($comp)
		Local $result[$comp_len + 1][8]
		$result[0][0] = 'N'
		$result[0][1] = 'Подразделение'
		$result[0][2] = 'Имя'
		$result[0][3] = 'Кол-во сессий'
		$result[0][4] = 'Общее время'
		$result[0][5] = 'Внесено'
		$result[0][6] = 'Обновлено'
		$result[0][7] = 'Удалено'
		Local $k = 0
		For $i = 1 To $comp_len
			$query = StringFormat("SELECT name, (SELECT name FROM cte WHERE id = cte) FROM user WHERE comp = '%s';", $comp[$i][0])
			_EzMySql_Query($query)
			$user = _EzMySql_FetchData()
			If $user <> 0 Then
				$k += 1
				$result[$k][0] = $k
				$result[$k][1] = $user[1]
				$result[$k][2] = $user[0]
				$query = StringFormat("SELECT time, request FROM log WHERE time > %u AND comp = '%s' AND (request = 'Start' OR request = 'Stop');", $start_time, $comp[$i][0])
				$sessions = _EzMySql_GetTable2d($query)
				$sessions_len = _EzMySql_Rows()
				If $sessions[1][1] <> 'Start' Then
					$sessions_len = _ArrayInsert($sessions, 1, StringFormat('%u|Start', $start_time))
					$sessions_len -= 1
				EndIf
				If $sessions[$sessions_len][1] <> 'Stop' Then
					$sessions_len = _ArrayAdd($sessions, StringFormat('%u|Stop', $stop_time))
				EndIf
				$sessions_time = $sessions[$sessions_len][0]
				$sessions_count = 0
				$sessions_mode = 'Stop'
				For $j = $sessions_len To 1 Step -1
					If $sessions[$j][1] = $sessions_mode Then
						ContinueLoop
					Else
						If $sessions[$j][1] = 'Start' Then
							$sessions_time -= $sessions[$j][0]
							$sessions_count += 1
						Else
							$sessions_time += $sessions[$j][0]
						EndIf
						$sessions_mode = $sessions[$j][1]
					EndIf
				Next
				_TicksToTime($sessions_time * 1000, $g_iHour, $g_iMins, $g_iSecs)
				$result[$k][3] = $sessions_count
				$result[$k][4] = StringFormat("%02i:%02i:%02i", $g_iHour, $g_iMins, $g_iSecs)
				$query = StringFormat("SELECT COUNT(*) FROM log WHERE time > %u AND comp = '%s' AND request LIKE 'INSERT %';", $start_time, $comp[$i][0])
				_EzMySql_Query($query)
				$update = _EzMySql_FetchData()
				If $update <> 0 Then
					$result[$k][5] = $update[0]
				EndIf
				$query = StringFormat("SELECT COUNT(*) FROM log WHERE time > %u AND comp = '%s' AND request LIKE 'UPDATE %';", $start_time, $comp[$i][0])
				_EzMySql_Query($query)
				$update = _EzMySql_FetchData()
				If $update <> 0 Then
					$result[$k][6] = $update[0]
				EndIf
				$query = StringFormat("SELECT COUNT(*) FROM log WHERE time > %u AND comp = '%s' AND request LIKE 'DELETE %';", $start_time, $comp[$i][0])
				_EzMySql_Query($query)
				$update = _EzMySql_FetchData()
				If $update <> 0 Then
					$result[$k][7] = $update[0]
				EndIf
			Else
				Sleep(0)
;~ 				print($comp[$i][0]) ; Not registerd users
			EndIf
		Next
	EndIf
	$result[0][0] = $k
	Return $result
;~ 	_ArrayDisplay($result)
EndFunc   ;==>tools_working_time

Func layer_type_activate($type_id)
	$aResult = _EzMySql_GetTable2d(StringFormat("SELECT option_id, name, list FROM options WHERE type_id = %u ORDER BY option_id;", $type_id))
	$iRows = _EzMySql_Rows()
	If IsArray($aResult) And $iRows > 0 Then
		For $i = 0 To $iRows - 1
			GUICtrlSetData($LabelOptions[$i], $aResult[$i + 1][1])
			GUICtrlSetData($ComboOptions[$i], list_to_string($aResult[$i + 1][2]), '')
			GUICtrlSetState($LabelOptions[$i], $GUI_SHOW)
			GUICtrlSetState($ComboOptions[$i], $GUI_SHOW)
			_GUICtrlListView_AddColumn($ListView, $aResult[$i + 1][1])
		Next
	EndIf
	Local $aPos = WinGetPos("[ACTIVE]")
	Local $x = Round(10 * $aPos[2] / 906)
	Local $y = Round(120 * $aPos[3] / 588 + 50 * $aPos[3] / 588 * (Ceiling($iRows / 4)))
	Local $w = Round(870 * $aPos[2] / 906)
	Local $h = Round(380 * $aPos[3] / 588 - 50 * $aPos[3] / 588 * (Ceiling($iRows / 4)))
	GUICtrlSetPos($ListView, $x, $y, $w, $h)
EndFunc   ;==>layer_type_activate

Func layer_type_deactivate()
	For $i = 0 To 19
		GUICtrlSetState($LabelOptions[$i], $GUI_HIDE)
		GUICtrlSetState($ComboOptions[$i], $GUI_HIDE)
		GUICtrlSetData($LabelOptions[$i], '')
		GUICtrlSetData($ComboOptions[$i], '', '')
	Next
	For $i = _GUICtrlListView_GetColumnCount($ListView) To 10 Step -1
		_GUICtrlListView_DeleteColumn($ListView, $i)
	Next
	Local $aPos = WinGetPos("[ACTIVE]")
	Local $x = Round(10 * $aPos[2] / 906)
	Local $y = Round(120 * $aPos[3] / 588)
	Local $w = Round(870 * $aPos[2] / 906)
	Local $h = Round(380 * $aPos[3] / 588)
	GUICtrlSetPos($ListView, $x, $y, $w, $h)
EndFunc   ;==>layer_type_deactivate

Func list_to_string($string)
	If StringLeft($string, 2) = '["' Then
		$string = StringTrimLeft($string, 2)
	EndIf
	If StringRight($string, 2) = '"]' Then
		$string = StringTrimRight($string, 2)
	EndIf
	Return StringReplace($string, '", "', '|')
EndFunc   ;==>list_to_string

Func _options_to_array($string)
	Dim $options_array[$number_options + 1]
	$string = StringTrimLeft($string, 1)
	$string = StringTrimRight($string, 2)
	$string = StringSplit($string, '", ', 1)
	For $i = 1 To $string[0]
		$item = StringSplit($string[$i], ': "', 1)
		For $j = 1 To $number_options
			If $j = $item[1] Then
				$options_array[$j] = $item[2]
			EndIf
		Next
	Next
	Return $options_array
EndFunc   ;==>_options_to_array

Func options_to_dict($string)
	; dict is a 2D array starts from row 1
	Dim $options_dict[1][2]
	$string = StringTrimLeft($string, 1)
	$string = StringTrimRight($string, 2)
	$string = StringSplit($string, '", ', 1)
	For $i = 1 To $string[0]
		$tmp = _ArrayExtract(StringSplit($string[$i], ': "', 1), 1, 2)
		_ArrayTranspose($tmp)
		_ArrayAdd($options_dict, $tmp)
	Next
	Return $options_dict
EndFunc   ;==>options_to_dict

Func string_to_dict_string($string, $type_id)
	Local $result = '{'
	$string = StringSplit($string, "; ", 1)
	Local $aRow
	For $i = 1 To $string[0]
		$tmp = StringSplit($string[$i], ': ', 1)
		$query = "SELECT option_id FROM options WHERE name LIKE '" & $tmp[1] & "' AND type_id = " & $type_id & ";"
		_EzMySql_Query($query)
		$aRow = _EzMySql_FetchData()
		$result &= $aRow[0] & ': "' & $tmp[2] & '", '
	Next
	$result = StringTrimRight($result, 2) & '}'
	Return $result
EndFunc   ;==>string_to_dict_string

Func check_params($vendor, $module, $params)
	Local $result = False
	If $params Then
		Local $query = StringFormat("SELECT type_id, options FROM equipment WHERE equip_id LIKE '%s %s';", $vendor, $module)
		Local $aRow, $option_name
		_EzMySql_Query($query)
		$aRow = _EzMySql_FetchData()
		If IsArray($aRow) Then
			Local $type_id = $aRow[0]
			Local $options = $aRow[1]
			Local $options_dict = options_to_dict($options)
			Local $string = ''
			For $i = 1 To UBound($options_dict) - 1
				$query = "SELECT name FROM options WHERE option_id = " & $options_dict[$i][0] & ";"
				_EzMySql_Query($query)
				$option_name = _EzMySql_FetchData()
				$string &= $option_name[0] & ": " & $options_dict[$i][1] & "; "
			Next
			$string = StringTrimRight($string, 2)
			If $string = $params Then
				$result = True
			Else
				$message = 'Новые параметры: ' & @CRLF & $params & @CRLF & 'отличаются от данных в базе:' & @CRLF & _
						$string & @CRLF & 'Вы хотите изменить параметры для оборудования ' & $vendor & " " & $module & '?'
				$dialog = MsgBox(292, StringFormat('Сообщение - %s', $ProgramName), $message) ; 4 Yes and No, 32 Question-mark icon, 256 Second button is default button
				If $dialog = 6 Then
					$params_dict = string_to_dict_string($params, $type_id)
					$query = StringFormat("UPDATE equipment SET options = '%s' WHERE equip_id = '%s %s' AND type_id = %u;", _
							$params_dict, $vendor, $module, $type_id)
					_log($query)
					_EzMySql_Exec($query)
				EndIf
			EndIf
		Else
			_EzMySql_Query(StringFormat("SELECT id FROM vendor WHERE name = '%s';", _GUICtrlComboBox_GetEditText($ComboVendor)))
			$type_id = _EzMySql_FetchData()
			$params_dict = string_to_dict_string($params, $type_id[0])
			If _CheckVendor($vendor) Then
				$query = StringFormat("INSERT INTO equipment (equip_id, type_id, options) VALUES ('%s %s', %u, '%s');", $vendor, $module, $type_id[0], $params_dict)
;~ 				print($query)
				_log($query)
				_EzMySql_Exec($query)
			EndIf
		EndIf
	EndIf
	Return $result
EndFunc   ;==>check_params

Func print($string)
	ConsoleWrite($string & @CRLF)
EndFunc   ;==>print

Func _log($query)
	;
	; Inserts request into log.
	;
	_EzMySql_Exec(StringFormat("INSERT INTO log " & _
			"VALUES (NULL, %u, '%s/%s', '%s');", _
			_toEPOCH(_NowCalc()), @ComputerName, @UserName, _
			StringReplace($query, "'", "\'")))
EndFunc   ;==>_log
