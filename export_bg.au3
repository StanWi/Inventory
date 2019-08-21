#include <Array.au3>
#include <Date.au3>
#include <EzMySql.au3>
#include <File.au3>
#include <secret.au3>

Global $aInventory
Global $fileInventory = @ScriptDir & '\Export\BG\CardInventory20190805135143.csv'
Global $GMT = 8
;~ Workmode full access or readonly
Global $mode = True
Global $mode = False ; Readonly

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

$aInventory = file_to_array($fileInventory)
Global $time = time($fileInventory)
check_curent_equipment($aInventory)
check_active_equipment($aInventory)
_EzMySql_Close()
_EzMySql_ShutDown()

Func file_to_array($file)
	If Not FileExists($file) Then
		print('File not found.')
		Exit
	EndIf
	Dim $aLines
	_FileReadToArray($file, $aLines)
	Dim $aInv[$aLines[0] + 1][25]
	$aInv[0][0] = $aLines[0]
	For $i = 1 To $aLines[0]
		$string = StringSplit($aLines[$i], ',')
		For $j = 1 To $string[0]
			$aInv[$i][$j - 1] = $string[$j]
		Next
	Next
	_ArrayDisplay($aInv)
	Return $aInv
EndFunc   ;==>file_to_array

Func time($file)
	;
	; Returns EPOCH in sec
	;
	$i = StringInStr($file, "CardInventory20")
	$year = StringMid($file, $i + 13, 4)
	$month = StringMid($file, $i + 17, 2)
	$day = StringMid($file, $i + 19, 2)
	$hh = StringMid($file, $i + 21, 2)
	$mm = StringMid($file, $i + 23, 2)
	$ss = StringMid($file, $i + 25, 2)
	Return _DateDiff('s', '1970/01/01 00:00:00', _
			StringFormat('%s/%s/%s %s:%s:%s', $year, $month, $day, $hh, $mm, $ss)) - 3600 * $GMT
EndFunc   ;==>time

Func check_curent_equipment($inv)
	For $i = 2 To $inv[0][0]
		$ne = $inv[$i][1]
		$type = $inv[$i][6]
		$sn = $inv[$i][7]
		While StringLeft($sn, 1) = '0'
			$sn = StringTrimLeft($sn, 1)
		WEnd
		$place = StringFormat('%s_%s', $ne, $inv[$i][4])
		If $sn = '.0' Then
			$sn = 'NA'
		EndIf
		$query = StringFormat("SELECT * FROM inventory WHERE vendor = 1 AND serial = '%s' AND object = '%s' ORDER BY time;", $sn, $type)
		$aResult = _EzMySql_GetTable2d($query)
		$iRows = _EzMySql_Rows()
		If $iRows = 0 Then ; Check concatenated serial numbers
			$query = StringFormat("SELECT * FROM inventory WHERE vendor = 1 AND serial LIKE '%s@%' AND object = '%s' ORDER BY time DESC;", $sn, $type)
			$aResult = _EzMySql_GetTable2d($query)
			$iRows = _EzMySql_Rows()
			If $iRows = 0 Then ; Object not found even in concatenated numbers
				$location = location($ne)
				If $location Then
					$query = StringFormat("INSERT INTO inventory (time, vendor, object, serial, place, status, user, location, invnum, nomnum, nsz) " & _
							"VALUES(%u, 1, '%s', '%s', '%s', 6, 100, '%s', '', '', 0);", $time, $type, $sn, $place, $location)
					print($query)
					If $mode Then _EzMySql_Exec($query)
				EndIf
			Else ; Found concateneted serial number
				$serials = ''
				$found = False
				For $k = 1 To $iRows
					$serial_db = $aResult[$k][3]
					$place_db = $aResult[$k][4]
					If $place = $place_db And Not StringInStr($serial_db, $serials) Then
						$found = True
						ExitLoop
					Else
						$serials &= $serial_db
					EndIf
				Next
				If Not $found Then
					print(StringFormat("Not found matches by concatenated numbers! " & _
							"Equipment %s installed in %s serial number %s", $type, $place, $sn))
				EndIf
			EndIf
		Else
			If $time > $aResult[$iRows][0] Then
				If $place <> $aResult[$iRows][4] Then
					print(StringFormat("New record: '%s' not equal '%s'", $place, $aResult[$iRows][4]))
					$location = location($ne)
					If $location Then
						$query = StringFormat("INSERT INTO inventory (time, vendor, object, serial, place, status, user, location, invnum, nomnum, nsz) " & _
								"VALUES(%u, 1, '%s', '%s', '%s', 6, 100, '%s', '', '', 0);", $time, $type, $sn, $place, $location)
						print($query)
						If $mode Then _EzMySql_Exec($query)
					EndIf
				ElseIf $aResult[$iRows][5] <> 6 Then
					print('Equipment in work, but status not. Rare case.')
					$query = StringFormat("UPDATE inventory SET status = 6 " & _
							"WHERE time = %u AND vendor = 1 AND object = '%s' AND serial = '%s' AND place = '%s' AND status = %u;", _
							$aResult[$iRows][0], $type, $sn, $aResult[$iRows][4], $aResult[$iRows][5])
					print($query)
					If $mode Then _EzMySql_Exec($query)
				EndIf
			ElseIf $time < $aResult[$iRows][0] Then
				print('New data after ' & $time)
				_ArrayDisplay($aResult)
			Else
				print('Data already inserted')
			EndIf
		EndIf
	Next
EndFunc   ;==>check_curent_equipment

Func check_active_equipment($inv)
	;
	; Check equipment in work and in database with status in work
	;
	Local $xdm[] = ['AOC_T', 'BG20_B_Shelf', 'BG30_B_Shelf', 'BG30_E_Shelf', 'CHTR_B', 'DIOB', 'D_MD_H_M1', 'HLXC_768', 'M2_84', 'MECP', 'MO_DW16MDR', 'MO_DW40DC', 'MO_DW40MC', _
			'MO_OFA_M', 'MO_OFA_MH', 'MO_OFA_PHBC', 'MO_ROADM8', 'MO_ROADM8D', 'OFA_2', 'OM01_4', 'OM16_1', 'OMCD_8', 'OMS01_4', 'OMS04_1', _
			'OMTX10', 'OMTX10_EF', 'OMTX10_LAN', 'OM_ILC', 'OM_MO_OFA_M', 'OM_OFA_B', 'OM_OFA_MH', 'OM_OFA_P', 'OM_OW_OSC', 'OTGbE_LX', 'OTR1_L3', 'OTR1_L5', 'OTR1_S3', 'OTR1_S3BD', 'OTR1_S5BD', 'OTR1_VL5', _
			'OTR103_LR', 'OTR10Txx_AL', 'OTR16_L5', 'OTR16_S3', 'OTR4_L5', 'OTR4_S3', 'OTR64_PI3', 'OTX10xx', 'OTX10_AT', 'PIO2_84', 'PIO2_84H', 'PSFU', 'SHELF ARTEMIS 1P ASSEMBLED', _
			'SIO164F', 'SIO16M', 'SIO16_2B', 'SIO16_4B', 'SIO1n4B', 'SIO1n4M', 'TMU_L', 'TRP10_4M', 'xFCU40', 'xFCU_H', 'xINF', 'xINF40', 'xINF_H', _
			'XIO384F', 'xMCPB', 'XDM-1000', 'XDM-500', 'XDM-40', 'xRAP-D']
	$query = "SELECT tab1.* FROM inventory tab1 " & _
			"INNER JOIN (SELECT MAX(time) as now, serial, object FROM inventory GROUP BY serial, object) tab2 " & _
			"ON (tab1.serial = tab2.serial AND tab1.time = tab2.now AND tab1.object = tab2.object) " & _
			"WHERE tab1.vendor = 1 AND tab1.status = 6 ORDER BY tab1.object;"
	$aResult = _EzMySql_GetTable2d($query)
	$iRows = _EzMySql_Rows()
	If IsArray($aResult) And $iRows > 0 Then
		;_ArrayDisplay($aResult)
		Local $i, $j, $k
		For $i = 1 To $iRows
			For $j = 1 To $inv[0][0]
				$sn = StringSplit($aResult[$i][3], '@') ; Separate concatenated serial numbers
				$sn_file = $inv[$j][7]
				While StringLeft($sn_file, 1) = '0'
					$sn_file = StringTrimLeft($sn_file, 1)
				WEnd
				If $sn_file = '.0' Then $sn_file = 'NA'
				If $sn[1] = $sn_file Then
					ExitLoop
				EndIf
			Next
			If $j = $inv[0][0] + 1 Then
				For $k = 0 To UBound($xdm) - 1
					If $aResult[$i][2] = $xdm[$k] Then
						ExitLoop
					EndIf
					If $k = UBound($xdm) - 1 Then
						print('Equipment not in work SN ' & $aResult[$i][3] & ' module ' & $aResult[$i][2])
					EndIf
				Next
			EndIf
		Next
	EndIf
EndFunc   ;==>check_active_equipment

Func location($ne)
	$query = StringFormat("SELECT name FROM location WHERE ne = '%s'", $ne)
	_EzMySql_Query($query)
	$name = _EzMySql_FetchData()
	If IsArray($name) Then
		Return $name[0]
	EndIf
	print('Location not found for NE ' & $ne)
	Return False
EndFunc   ;==>location

Func print($s = '', $end = @CRLF)
	ConsoleWrite($s & $end)
EndFunc   ;==>print
