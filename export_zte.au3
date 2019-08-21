#include <Array.au3>
#include <Date.au3>
#include <EzMySql.au3>
#include <File.au3>
#include <secret.au3>

Global $aInventory
Global $fileInventory = @ScriptDir & '\Export\ZTE\Engineering Equipment Information Export_20190814083734.csv'
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
Global $time
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
	Dim $aInv[$aLines[0] + 1][21]
	$aInv[0][0] = $aLines[0]
	For $i = 1 To $aLines[0]
		$string = StringSplit($aLines[$i], ',')
		For $j = 1 To $string[0]
			$aInv[$i][$j - 1] = $string[$j]
		Next
	Next
	Return $aInv
EndFunc   ;==>file_to_array

Func time($t)
	;
	; Returns EPOCH in sec
	;
	Return _DateDiff('s', '1970/01/01 00:00:00', _
			StringReplace($t, '-', '/')) - 3600 * $GMT
EndFunc   ;==>time

Func check_curent_equipment($inv)
	For $i = 3 To $inv[0][0]
		$time = time($inv[$i][0])
		$ne = $inv[$i][2]
		$type = $inv[$i][8]
		$sn = $inv[$i][9]
		$place = StringFormat('%s_[%u-%u-%u]', $ne, $inv[$i][10], $inv[$i][11], $inv[$i][12])
		If $sn = '""' Then
			$sn = 'NA'
		EndIf
		$query = StringFormat("SELECT * FROM inventory WHERE vendor = 2 AND serial = '%s' AND object = '%s' ORDER BY time;", $sn, $type)
		$aResult = _EzMySql_GetTable2d($query)
		$iRows = _EzMySql_Rows()
		If $iRows = 0 Then ; Check concatenated serial numbers
			$query = StringFormat("SELECT * FROM inventory WHERE vendor = 2 AND serial LIKE '%s@%' AND object = '%s' ORDER BY time DESC;", $sn, $type)
			$aResult = _EzMySql_GetTable2d($query)
			$iRows = _EzMySql_Rows()
			If $iRows = 0 Then ; Object not found even in concatenated numbers
				$location = location($ne)
				If $location Then
					$query = StringFormat("INSERT INTO inventory (time, vendor, object, serial, place, status, user, location, invnum, nomnum, nsz) " & _
							"VALUES(%u, 2, '%s', '%s', '%s', 6, 100, '%s', '', '', 0);", $time, $type, $sn, $place, $location)
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
								"VALUES(%u, 2, '%s', '%s', '%s', 6, 100, '%s', '', '', 0);", $time, $type, $sn, $place, $location)
						print($query)
						If $mode Then _EzMySql_Exec($query)
					EndIf
				ElseIf $aResult[$iRows][5] <> 6 Then
					print('Equipment in work, but status not. Rare case.')
					$query = StringFormat("UPDATE inventory SET status = 6 " & _
							"WHERE time = %u AND vendor = 2 AND object = '%s' AND serial = '%s' AND place = '%s' AND status = %u;", _
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
	Local $zte[] = ['@XFP_FTLX6614MCC-Z1', '@XFP_JXP01TMAC1CZ5PGA', '@XFP_LTX1305-BC+', '@XFP_XFP-10G-10LRDP', 'G652 120km.LC', 'NX41-21', _
					'MTRS-2E35-01', 'TR-PX13L-NSN', 'WDM Subrack NX41-21B', 'WXTRPGEAS1E', 'ZTE_SM-10km-1310-10G-C']
	$query = StringFormat("SELECT tab1.* FROM inventory tab1 " & _
			"INNER JOIN (SELECT MAX(time) as now, serial, object FROM inventory WHERE time <= %u GROUP BY serial, object) tab2 " & _
			"ON (tab1.serial = tab2.serial AND tab1.time = tab2.now AND tab1.object = tab2.object) " & _
			"WHERE tab1.vendor = 2 AND tab1.status = 6 AND tab1.object NOT IN ('%s') ORDER BY tab1.object;", $time, _ArrayToString($zte, "', '"))
	$aResult = _EzMySql_GetTable2d($query)
	$iRows = _EzMySql_Rows()
	If IsArray($aResult) And $iRows > 0 Then
		Local $i, $j, $k
		For $i = 1 To $iRows
			For $j = 1 To $inv[0][0]
				$sn = StringSplit($aResult[$i][3], '@') ; Separate concatenated serial numbers
				If $sn[1] = 'NA' Then $sn[1] = '""'
				If $sn[1] = $inv[$j][9] Then
					ExitLoop
				EndIf
			Next
			If $j = $inv[0][0] + 1 Then
				print('Equipment not in work SN ' & $aResult[$i][3] & ' module ' & $aResult[$i][2])
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
