#include <Array.au3>
#include <Date.au3>
#include <EzMySql.au3>
#include <File.au3>
#include <secret.au3>

Global $aInventory
$fileInventory = @ScriptDir & '\Export\BG\SFPInventory20190912162407.csv'
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
	Dim $aInv[$aLines[0] + 1][17]
	$aInv[0][0] = $aLines[0]
	$k = 1
	For $i = 1 To $aLines[0]
		$string = StringSplit($aLines[$i], ',')
		If $string[8] Then
			For $j = 1 To $string[0]
				$aInv[$k][$j - 1] = $string[$j]
			Next
			$k += 1
		Else
			$aInv[0][0] -= 1
		EndIf
	Next
	Return $aInv
EndFunc   ;==>file_to_array

Func time($file)
	;
	; Returns EPOCH in sec
	;
	$i = StringInStr($file, "SFPInventory20")
	$year = StringMid($file, $i + 12, 4)
	$month = StringMid($file, $i + 16, 2)
	$day = StringMid($file, $i + 18, 2)
	$hh = StringMid($file, $i + 20, 2)
	$mm = StringMid($file, $i + 22, 2)
	$ss = StringMid($file, $i + 24, 2)
	Return _DateDiff('s', '1970/01/01 00:00:00', _
			StringFormat('%s/%s/%s %s:%s:%s', $year, $month, $day, $hh, $mm, $ss)) - 3600 * $GMT
EndFunc   ;==>time

Func check_curent_equipment($inv)
	Local $sfp[] = ['OTR1', 'OTR4', 'OTR16', 'OTGBE']
	For $i = 2 To $inv[0][0]
		If $inv[$i][7] Then ; Serial
			$ne = $inv[$i][1]
			$type = $inv[$i][6] ; Type (OTR1, OTR4, OTR16, OTGBE)
			If Not (_ArraySearch($sfp, $type) = -1) Then
				$type = $inv[$i][11]
			EndIf
			If StringLeft($type, 5) = "OTGBE" Then StringReplace($type, "OTGBE", "OTGbE", 1, 1)
			$sn = $inv[$i][7]
			While StringLeft($sn, 1) = '0'
				$sn = StringTrimLeft($sn, 1)
			WEnd
			$place = StringFormat('%s_%s_%s', $ne, $inv[$i][3], $inv[$i][5])
			$query = StringFormat("SELECT * FROM inventory WHERE time <= %u AND vendor = 1 AND serial = '%s' AND object = '%s' ORDER BY time;", $time, $sn, $type)
			$aResult = _EzMySql_GetTable2d($query)
			$iRows = _EzMySql_Rows()
			If $iRows = 0 Then ; Check concatenated serial numbers
				$query = StringFormat("SELECT * FROM inventory " & _
						"WHERE time <= %u AND vendor = 1 AND serial LIKE '%s@%' AND object = '%s' ORDER BY time DESC;", $time, $sn, $type)
				$aResult = _EzMySql_GetTable2d($query)
				$iRows = _EzMySql_Rows()
				If $iRows = 0 Then ; Object not found even in concatenated numbers
					$location = location($ne)
					If $location Then
						$query = StringFormat("INSERT INTO inventory " & _
								"(time, vendor, object, serial, place, status, user, location, invnum, nomnum, nsz) " & _
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
							If $aResult[$k][5] <> 6 Then
								print(StringFormat("Not in work. Equipment %s installed into %s serial number %s", $type, $place, $sn))
								$query = StringFormat("UPDATE inventory SET status = 6 " & _
										"WHERE time = %u AND vendor = 1 AND object = '%s' AND serial = '%s' AND place = '%s' AND status = %u;", _
										$aResult[$k][0], $aResult[$k][2], $aResult[$k][3], $aResult[$k][4], $aResult[$k][5])
								print($query)
								If $mode Then _EzMySql_Exec($query)
							EndIf
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
						_ArrayDisplay($aResult)
;~ 						$location = location($ne)
;~ 						If $location Then
;~ 							$query = StringFormat("INSERT INTO inventory " & _
;~ 									"(time, vendor, object, serial, place, status, user, location, invnum, nomnum, nsz) " & _
;~ 									"VALUES(%u, 1, '%s', '%s', '%s', 6, 100, '%s', '', '', 0);", $time, $type, $sn, $place, $location)
;~ 							print($query)
;~ 							If $mode Then _EzMySql_Exec($query)
;~ 						EndIf
					ElseIf $aResult[$iRows][5] <> 6 Then
						print('Equipment in work, but status not. Rare case.')
						$query = StringFormat("UPDATE inventory SET status = 6 " & _
								"WHERE time = %u AND vendor = 1 AND object = '%s' AND serial = '%s' AND place = '%s' AND status = %u;", _
								$aResult[$iRows][0], $type, $sn, $aResult[$iRows][4], $aResult[$iRows][5])
						print($query)
						If $mode Then _EzMySql_Exec($query)
					EndIf
				ElseIf $time < $aResult[$iRows][0] Then
					print(StringFormat("New data after file. Equipment %s installed in %s serial number %s", $type, $place, $sn))
;~ 					_ArrayDisplay($aResult)
				ElseIf $time = $aResult[$iRows][0] And $place <> $aResult[$iRows][4] Then
					print(StringFormat("Duplicate equipment. Equipment %s installed in %s serial number %s", $type, $place, $sn))
				Else
					print(StringFormat("Data already inserted. Equipment %s installed in %s serial number %s", $type, $place, $sn))
				EndIf
			EndIf
		EndIf
	Next
EndFunc   ;==>check_curent_equipment

Func check_active_equipment($inv)
	;
	; Check equipment in work and in database with status in work
	;
	Local $sfp[] = ['OTR1_S3', 'OTR1_S3BD', 'OTR1_S5BD', 'OTR1_L3', 'OTR1_L5', 'OTR4_S3', 'OTR4_L5', 'OTR16_S3', 'OTR16_L5', 'OTGbE_LX']
;~ 	time, vendor, object, serial, place, status, user, location, invnum, nomnum, nsz
	$query = StringFormat("SELECT tab1.time, tab1.object, tab1.serial, tab1.place FROM inventory tab1 " & _
			"INNER JOIN (SELECT MAX(time) as now, serial, object FROM inventory WHERE time <= %u GROUP BY serial, object ) tab2 " & _
			"ON (tab1.serial = tab2.serial AND tab1.time = tab2.now AND tab1.object = tab2.object) " & _
			"WHERE tab1.vendor = 1 AND tab1.status = 6 AND tab1.object IN ('%s') AND place LIKE '%%ort %%' ORDER BY tab1.object;", $time, _ArrayToString($sfp, "', '"))
	$aResult = _EzMySql_GetTable2d($query)
	$iRows = _EzMySql_Rows()
	$aResult[0][0] = $iRows
;~ 	_ArrayDisplay($aResult)
;~ 	_ArrayDisplay($inv)
	print('Equipment in database: ' & $aResult[0][0])
	print('Equipment in file:     ' & $inv[0][0] - 1)
	If IsArray($aResult) And $iRows > 0 Then
;~ 		Local $i, $j, $k
		For $i = 1 To $iRows
			For $j = 1 To $inv[0][0]
				$sn = StringSplit($aResult[$i][2], '@') ; Separate concatenated serial numbers
				$sn_file = $inv[$j][7]
				While StringLeft($sn_file, 1) = '0'
					$sn_file = StringTrimLeft($sn_file, 1)
				WEnd
				If $sn_file = '.0' Then $sn_file = 'NA'
				If $sn[1] = $sn_file And $aResult[$i][3] = StringFormat('%s_%s_%s', $inv[$j][1], $inv[$j][3], $inv[$j][5]) Then
					ExitLoop
				EndIf
			Next
			If $j = $inv[0][0] + 1 Then
;~ 				For $k = 0 To UBound($xdm) - 1
;~ 					If $aResult[$i][2] = $xdm[$k] Then
;~ 						ExitLoop
;~ 					EndIf
;~ 					If $k = UBound($xdm) - 1 Then
				print('Equipment not in work SN ' & $aResult[$i][2] & ' module ' & $aResult[$i][1])
;~ 					EndIf
;~ 				Next
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

Func replace_module_types($inv)
	For $i = 2 To $inv[0][0] ; Strarts from 2nd row
		$old_object = $inv[$i][6]
		If $old_object = "OTGBE" Then $old_object = "OTGbE"
		$new_object = $inv[$i][11]
		If StringLeft($new_object, 5) = "OTGBE" Then StringReplace($new_object, "OTGBE", "OTGbE", 1, 1)
		$serial = $inv[$i][7]
		While StringLeft($serial, 1) = '0'
			$serial = StringTrimLeft($serial, 1)
		WEnd
		$query = StringFormat("UPDATE inventory SET object = '%s' " & _
				"WHERE vendor = 1 AND object = '%s' AND (serial = '%s' OR serial LIKE '%s@%');", _
				$new_object, $old_object, $serial, $serial)
		If $mode Then _EzMySql_Exec($query)
	Next
EndFunc   ;==>replace_module_types
