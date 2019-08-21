#include <Array.au3>
#include <Date.au3>
#include <EzMySql.au3>
#include <File.au3>
#include <secret.au3>

Global $aInventory
Global $fileInventory = @ScriptDir & '\Export\XDM\Inventory_190807'
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
	Dim $aInv[$aLines[0] + 1][10]
	$aInv[0][0] = 0
	Local $text = ''
	Local $string = ''
	Local $name = ''
	Local $blacklist = ['', '1234', 'N.A', '3', 'Serial']
	$k = 1
	For $i = 1 To $aLines[0]
		$text = StringStripWS($aLines[$i], 4)
		$text = StringReplace($text, ' |', '|')
		$text = StringReplace($text, '| ', '|')
		$string = StringSplit($text, '|')
		If $string[0] = 1 And StringLeft($string[1], 1) = ' ' And Not StringInStr($string[1], '=') Then
			$name = StringStripWS($string[1], 8)
		EndIf
		If ($string[0] > 5) And (_ArraySearch($blacklist, $string[5]) = -1) Then ; Serial not in blacklist
			For $j = 1 To $string[0]
				$aInv[$k][$j - 1] = $string[$j]
			Next
			$aInv[$k][0] = $name
			$k += 1
		EndIf
		$aInv[0][0] = $k - 1
	Next
	Return $aInv
EndFunc   ;==>file_to_array

Func time($file)
	;
	; Returns EPOCH UTC in sec
	;
	Dim $aLines
	_FileReadToArray($file, $aLines)
	$k = 1
	While Not StringInStr($aLines[$k], 'Date : ')
		$k += 1
	WEnd
	Local $i
	Local $string
	Local $year = '1970'
	Local $month = '01'
	Local $day = '01'
	Local $hhmm = '00:00'
	$string = StringSplit(StringStripWS($aLines[$k], 4), ' ')
	$year = $string[7]
	Switch $string[4]
		Case 'Jan'
			$month = '01'
		Case 'Feb'
			$month = '02'
		Case 'Mar'
			$month = '03'
		Case 'Apr'
			$month = '04'
		Case 'May'
			$month = '05'
		Case 'Jun'
			$month = '06'
		Case 'Jul'
			$month = '07'
		Case 'Aug'
			$month = '08'
		Case 'Sep'
			$month = '09'
		Case 'Oct'
			$month = '10'
		Case 'Nov'
			$month = '11'
		Case 'Dec'
			$month = '12'
	EndSwitch
	$day = $string[5]
	$hhmm = $string[6]
	print(StringFormat('%s/%s/%s %s:00', $year, $month, $day, $hhmm))
	Return _DateDiff('s', '1970/01/01 00:00:00', _
			StringFormat('%s/%s/%s %s:00', $year, $month, $day, $hhmm)) - 3600 * $GMT
EndFunc   ;==>time

Func check_curent_equipment($inv)
;~ $aResult
;~ Row|Col 0     |Col 1 |Col 2 |Col 3                      |Col 4           |Col 5 |Col 6|Col 7   |Col 8 |Col 9 |Col 10
;~ [0]|time      |vendor|object|serial                     |place           |status|user |location|invnum|nomnum|nsz
;~ [1]|1433726940|1     |OTR64 |1116122023@RTPC_I5_1-3     |RTPC_I5_1-3     |6     |1    |RTPC    |      |      |0
;~ [2]|1437103320|1     |OTR64 |1116122023@Taishet_2_I9_1-1|Taishet_2_I9_1-1|6     |1    |ALAZ    |      |      |0
;~ $iRows = 2
	Local $i, $j
	Local $aRecords, $aRecordsTemp
	Local $file
	Local $path
	Local $request = ''
	Local $place = ''
	Local $percent = 0.1
	Local $aResult, $iRows, $iColumns
	Local $blacklist = ['', '1234', 'N.A', '3', 'Serial']
	Local $sfp[] = ['OTR1', 'OTR4', 'OTR16', 'OTR64', 'OTGbE', 'OTR103', 'OTR10Txx']
	For $i = 1 To $inv[0][0]
		If ($i / $inv[0][0]) >= $percent Then
			print('.')
			$percent += 0.1
		EndIf
		$ne = $inv[$i][0]
		$name = $inv[$i][1]
		$Type = $inv[$i][2]
		$Serial = $inv[$i][4]
		If _ArraySearch($blacklist, $Serial) = -1 Then ; Serial not in blacklist
			$Type = StringSplit($Type, ' ') ; Type 'Object[ SubSlot]'
			$query = StringFormat("SELECT * FROM inventory WHERE time <= %u AND vendor = 1 AND serial = '%s' AND object = '%s' ORDER BY time;", $time, $Serial, $Type[1])
			If Not (_ArraySearch($sfp, $Type[1]) = -1) Then
				If $inv[$i][6] Then
					$Type[1] &= "_" & $inv[$i][6]
					$query = StringFormat("SELECT * FROM inventory WHERE time <= %u AND vendor = 1 AND serial = '%s' AND object = '%s' ORDER BY time;", $time, $Serial, $Type[1])
				Else
					print(StringFormat('Application code not found for %s %s %s', $Type[1], $Serial, place($ne, $name)))
					$query = StringFormat("SELECT * FROM inventory WHERE time <= %u AND vendor = 1 AND serial = '%s' AND object LIKE '%s_%' ORDER BY time;", $time, $Serial, $Type[1])
				EndIf
			EndIf
			$aResult = _EzMySql_GetTable2d($query)
			$iRows = _EzMySql_Rows()
			If $iRows = 0 Then ; Check concatenated serial numbers
				$query = StringFormat("SELECT * FROM inventory WHERE time <= %u AND vendor = 1 AND serial LIKE '%s@%' AND object = '%s' ORDER BY time DESC;", $time, $Serial, $Type[1])
				If Not (_ArraySearch($sfp, $Type[1]) = -1) Then
					$query = StringFormat("SELECT * FROM inventory WHERE time <= %u AND vendor = 1 AND serial LIKE '%s@%' AND object LIKE '%s_%' ORDER BY time DESC;", $time, $Serial, $Type[1])
				EndIf
				$aResult = _EzMySql_GetTable2d($query)
				$iRows = _EzMySql_Rows()
				If $iRows = 0 Then ; Object not found even in concatenated numbers
					$location = location($ne)
					If $location Then
						$query = StringFormat("INSERT INTO inventory (time, vendor, object, serial, place, status, user, location, invnum, nomnum, nsz) " & _
								"VALUES (%u, 1, '%s', '%s', '%s', 6, 100, '%s', '', '', 0);", $time, $Type[1], $Serial, place($ne, $name), $location)
						print($query)
						If $mode Then _EzMySql_Exec($query)
					EndIf
				Else ; Found concateneted serial number
					$place = place($ne, $name)
					$serials = ''
					$found = False
					For $k = 1 To $iRows
						$serial_db = $aResult[$k][3]
						$place_db = $aResult[$k][4]
						If $place = $place_db And Not StringInStr($serial_db, $serials) Then
							If $aResult[$k][5] <> 6 Then
								print(StringFormat("Not in work. Equipment %s installed in %s serial number %s", $Type[1], $place, $Serial))
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
								"Equipment %s installed in %s serial number %s", $Type[1], $place, $Serial))
					EndIf
				EndIf
			Else
				$place = place($ne, $name)
				If $time > $aResult[$iRows][0] Then
					If $place <> $aResult[$iRows][4] Then
						print(StringFormat("New record: '%s' not equal '%s'", $place, $aResult[$iRows][4]))
						$location = location($ne)
						If $location Then
							$query = StringFormat("INSERT INTO inventory (time, vendor, object, serial, place, status, user, location, invnum, nomnum, nsz) " & _
									"VALUES(%u, 1, '%s', '%s', '%s', 6, 100, '%s', '', '', 0);", $time, $Type[1], $Serial, $place, $location)
							print($query)
							If $mode Then _EzMySql_Exec($query)
						EndIf
					ElseIf $aResult[$iRows][5] <> 6 Then
						print('Equipment in work, but status not. Rare case.')
						$query = StringFormat("UPDATE inventory SET status = 6 " & _
								"WHERE time = %u AND vendor = 1 AND object = '%s' AND serial = '%s' AND place = '%s' AND status = %u;", _
								$aResult[$iRows][0], $Type[1], $Serial, $aResult[$iRows][4], $aResult[$iRows][5])
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
		EndIf
	Next
EndFunc   ;==>check_curent_equipment

Func check_active_equipment($inv)
	;
	; Check equipment in work and in database with status in work
	;
	Local $bg[] = ['BG20_B_Shelf', 'BG20_E_Shelf', 'DMGE_4_L2', 'FCU_30B', 'FCU_30BH', 'FCU_30E', 'INF_30B', _
			'INF_30BH', 'INF_30E', 'MBP_30B', 'MBP_30E', 'MCP30B', 'MPS_2G_8F', 'MPS_4F', 'MPS_6F', _
			'MXC-20', 'MXC-20C-DC', 'PME1_21', 'S1_4', 'SM10E BASE CARD', 'SMD1H', 'SMQ1&4', _
			'XDM-1000', 'XDM-500', 'XDM-40', 'XIO30-16', 'XIO30-4', 'xRAP-D', 'D_MD_H_M1', 'SHELF ARTEMIS 1P ASSEMBLED']
	$query = StringFormat("SELECT tab1.time, tab1.object, tab1.serial, tab1.place FROM inventory tab1 " & _
			"INNER JOIN (SELECT MAX(time) as now, serial, object FROM inventory WHERE time <= %u GROUP BY serial, object) tab2 " & _
			"ON (tab1.serial = tab2.serial AND tab1.time = tab2.now AND tab1.object = tab2.object) " & _
			"WHERE tab1.vendor = 1 AND tab1.status = 6 AND tab1.object NOT IN ('%s') AND place NOT LIKE '%%ort _' ORDER BY tab1.object;", $time, _ArrayToString($bg, "', '"))
	$aResult = _EzMySql_GetTable2d($query)
	$iRows = _EzMySql_Rows()
	$aResult[0][0] = $iRows
	print('Equipment in database:  ' & $aResult[0][0])
	print('Equipment in file:      ' & $inv[0][0])
;~ 	_ArrayDisplay($aResult)
;~ 	_ArrayDisplay($inv)
	If IsArray($aResult) And $iRows > 0 Then
		Local $i, $j, $k
		$percent = 0.1
		For $i = 1 To $iRows
			If ($i / $iRows) >= $percent Then
				print('.', '')
				$percent += 0.1
			EndIf
			For $j = 1 To $inv[0][0]
				$sn = StringSplit($aResult[$i][2], '@') ; Separate concatenated serial numbers
				$place = place($inv[$j][0], $inv[$j][1])
				If $sn[1] = $inv[$j][4] And $aResult[$i][3] = $place Then
					ExitLoop
				EndIf
			Next
			If $j = $inv[0][0] + 1 Then
				print('Equipment not in work SN ' & $aResult[$i][2] & ' module ' & $aResult[$i][1])
			EndIf
		Next
		print('')
		$percent = 0.1
		For $j = 1 To $inv[0][0]
			If ($j / $inv[0][0]) >= $percent Then
				print('.', '')
				$percent += 0.1
			EndIf
			$place = place($inv[$j][0], $inv[$j][1])
			For $i = 1 To $iRows
				$sn = StringSplit($aResult[$i][2], '@')
				If $sn[1] = $inv[$j][4] And $aResult[$i][3] = $place Then
					ExitLoop
				EndIf
			Next
			If $i = $iRows + 1 Then
				print('Equipment not in database SN ' & $inv[$j][4] & ' module ' & $inv[$j][1])
			EndIf
		Next
		print('')
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
;~ 	_ArrayDisplay($inv)
	Local $sfp[] = ['OTR1', 'OTR4', 'OTR16', 'OTR64', 'OTGbE', 'OTR103', 'OTR10Txx']
	For $i = 1 To $inv[0][0]
		$old_object_array = StringSplit($inv[$i][2], ' ')
		$old_object = $old_object_array[1]
		If Not (_ArraySearch($sfp, $old_object) = -1) And $inv[$i][6] Then
			$new_object = $old_object & "_" & $inv[$i][6]
			$Serial = $inv[$i][4]
			$query = StringFormat("UPDATE inventory SET object = '%s' " & _
					"WHERE vendor = 1 AND object = '%s' AND (serial = '%s' OR serial LIKE '%s@%');", _
					$new_object, $old_object, $Serial, $Serial)
			print($query)
			If $mode Then _EzMySql_Exec($query)
		EndIf
	Next
EndFunc   ;==>replace_module_types

Func place($ne, $name) ; Name 'Slot Object[ SubSlot]'
	Local $place
	Local $tmp = StringSplit($name, ' ')
	If $tmp[0] = 3 Then
		$place = $ne & '_' & $tmp[1] & '_' & $tmp[3]
	ElseIf $tmp[0] = 2 Then
		$place = $ne & '_' & $tmp[1]
	Else
		print('Problem with NE "%s" and Name "%s"' & $ne, $name)
		$place = ''
	EndIf
	Return $place
EndFunc   ;==>place
