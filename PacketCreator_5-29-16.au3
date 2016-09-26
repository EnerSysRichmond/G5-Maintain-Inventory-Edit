#include-once
#include <File.au3>
#include <array.au3>
#include <GuiListBox.au3>
#include <TreeViewConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiScrollBars.au3>
#include <Excel.au3>
#include <string.au3>
Opt("MustDeclareVars", 1)
Opt("PixelCoordMode", 0)
Opt("TrayIconDebug", 1)
Opt("WinDetectHiddenText", 1)
Opt("WinTitleMatchMode", 2)

;*****************************************************************************************
HotKeySet("{ESC}", "TogglePause") ;user can press escape at any time to pause the script
Local $g_bPaused = False
Func TogglePause()
	$g_bPaused = Not $g_bPaused
	While $g_bPaused
		Sleep(100)
		ToolTip('Script is "Paused."', 0, 0)
	WEnd
	ToolTip("")
EndFunc   ;==>TogglePause
;*****************************************************************************************

;this script will:
;read all order numbers and line items from the week closing excel spread sheet
;print all orders to files
;trim each text file to only include the header (shipping info) and the line item, with add ons
;adjust the inventory in Baan, if needed, for the connector
;indicate on the spread sheet if a connector was found
;indicate on spreadsheet if the order needs to ship complete so that it can be given a sticker
;deposits the trimmed order ack, with production number and connector specified, into the configurator folder
Local $oExcel, $ExcelFileName, $oClosing ;$oClosing = Excel spreadsheet obj, Week Closing Report
Local $extrasPage, $oWorksheet
Local $populated = False, $aConnectors, $addOnArray[0][2], $exportFlag = false
Local $TypeOfOrderColumn, $OrderTypeColumn, $orderNumberColumn, $PositionColumn, $ItemColumn, $DateColumn, $QttyColumn, $printedColumn, $prodOrderColumn
Local Const $mfgCompany = 705, $salesCompany = 701, $baanPath = @DesktopDir & "\baan.bwc"
Func delete($window, $control)
	If ControlGetFocus($window) <> $control Then
		ControlClick($window, "", $control)
		Sleep(150)
	Else
		ControlSend($window, "", $control, "{END}")
		Sleep(100)
		ControlSend($window, "", $control, "{SHIFTDOWN}{HOME}{SHIFTUP}")
		Sleep(150)
	EndIf
	ControlSend($window, "", $control, "{BS}")
	Sleep(250)
EndFunc
Func putText($window, $control, $newText)
	If StringCompare(ControlGetFocus($window), $control) <> 0 Then
		ControlClick($window, "", $control)
		Sleep(150)
	Else
		ControlSend($window, "", $control, "{END}")
		Sleep(100)
		ControlSend($window, "", $control, "{SHIFTDOWN}{HOME}{SHIFTUP}")
		Sleep(150)
	EndIf
	Opt("SendKeyDelay", 15)
	ControlSend($window, "", $control, $newText, 1) ;send raw keystrokes when entering text!
	Sleep(600)
	WinActivate($window)
	ControlSend($window, "", "", "{TAB}")
	Do
		Sleep(50)
	Until StringCompare(ControlGetFocus($window), $control) <> 0
	Opt("SendKeyDelay", 5)
EndFunc
Func _WinGetCaretPos()
    Local $iXAdjust = 5
    Local $iYAdjust = 40
    Local $iOpt = Opt("CaretCoordMode", 0) ; Set "CaretCoordMode" to relative mode and store the previous option.
    Local $aGetCaretPos = WinGetCaretPos() ; Retrieve the relative caret coordinates.
    Local $aGetPos = WinGetPos("[ACTIVE]") ; Retrieve the position as well as height and width of the active window.
    Local $sControl = ControlGetFocus("[ACTIVE]") ; Retrieve the control name that has keyboard focus.
    Local $aControlPos = ControlGetPos("[ACTIVE]", "", $sControl) ; Retrieve the position as well as the size of the control.
    $iOpt = Opt("CaretCoordMode", $iOpt) ; Reset "CaretCoordMode" to the previous option.
    Local $aReturn[2] = [0, 0] ; Create an array to store the x, y position.
    If IsArray($aGetCaretPos) And IsArray($aGetPos) And IsArray($aControlPos) Then
        $aReturn[0] = $aGetCaretPos[0] + $aGetPos[0] + $aControlPos[0] + $iXAdjust
        $aReturn[1] = $aGetCaretPos[1] + $aGetPos[1] + $aControlPos[1] + $iYAdjust
        Return $aReturn ; Return the array.
    Else
        Return SetError(1, 0, $aReturn) ; Return the array and set @error to 1.
    EndIf
EndFunc

main()

Func main()
	MsgBox(0, "Don't forget", "You haven't checked the RPO trimming functions or seens what the connector addition looks like")
	Local $aConnector[0][2]
  	Local $orderArray = openExcel()
 	printOrdersToTextFiles($orderArray)
	splitFiles($orderArray)
	checkMissedFiles($orderArray)
	trimFiles($orderArray, $aConnector)
	processConnectorFile($orderArray)
	deleteAllOrderFiles()
	removeFluff()
EndFunc

Func removeFluff()
	Local $fileList = _FileListToArray("C:\printFolder"), $line, $zipCode
	For $i = 1 To $fileList[0]
		For $j = 7 To 15 ;should never be more lines than this unless the orders change drastically. then we have way more problems than this.
			$line = FileReadLine("C:\printFolder\" & $fileList[$i], $j)
			If StringLen(StringStripWS($line, 8)) == 0 Then
				While StringCompare(StringLeft(FileReadLine("C:\printFolder\" & $fileList[$i], $j), 4), "----") <> 0
					_FileWriteToLine("C:\printFolder\" & $fileList[$i], $j, "", True)
				WEnd
				ExitLoop
			EndIf
		Next
	Next
EndFunc

Func deleteAllOrderFiles()
	Local $fileList = _FileListToArray("C:\printFolder")
	For $i = 1 To $fileList[0]
		If StringCompare(StringLeft($fileList[$i], 5), "order") == 0 Then
			FileDelete("C:\printFolder\" & $fileList[$i])
		EndIf
	Next
	If FileExists("C:\printFolder\connectorFile.txt") Then FileDelete("C:\printFolder\connectorFile.txt")
EndFunc

Func trimFiles($orderArray, ByRef $aConnector)
	Local $connector = "", $headerString = "", $success
	Local $type = "", $productionOrderNumber
	ToolTip("Trimming order files now...", 750, 0)
	For $i = 0 To UBound($orderArray) - 1
		ConsoleWrite("trim: " & $orderArray[$i][2] & @CRLF)
		Local $display = 1
		Local $orderNumber = $orderArray[$i][2], $orderAndLine = $orderNumber & $orderArray[$i][3]
		If StringCompare($orderArray[$i][8], "Processed") == 0 Then ContinueLoop
		If StringInStr($orderArray[$i][0], "Sales") Then
			$type = "sales"
		Else
			$type = "rpo"
		EndIf
		If StringCompare($type, "sales") == 0 Then
			$connector = trimSalesOrder($orderArray, $i)
			If $connector Then
				$productionOrderNumber = $orderArray[$i][7]
				writeToConnectorFile($productionOrderNumber, $connector)
			Else
				_Excel_RangeWrite($oClosing, $oWorksheet, "NO Conn", "T" & $i + 2)
				$oClosing.Sheets("Main").Range("A1:T" & String($i + 2)).Interior.ColorIndex = 7
			EndIf
		Else
			If StringCompare($type, "rpo") == 0 Then
				trimRPO($orderArray, $i)
			EndIf
		EndIf
	Next
	ToolTip("")
EndFunc

Func writeToConnectorFile($productionOrderNumber, $connector)
	Local $connectorFile = "C:\printFolder\connectorFile.txt"
	If Not FileExists($connectorFile) Then
		If Not FileWrite($connectorFile, $productionOrderNumber & "|" & $connector & "|" & @CRLF) Then MsgBox($MB_SYSTEMMODAL, "", "An error occurred when writing the file.")
	Else
		Local $fileHandle = FileOpen($connectorFile)
		If $fileHandle == -1 Then
			MsgBox($MB_SYSTEMMODAL, "", "An error occurred when opening the file.")
		Else
			FileWriteLine($connectorFile, $productionOrderNumber & "|" & $connector & "|")
			FileClose($fileHandle)
		EndIf
	EndIf
EndFunc

Func processConnectorFile($orderArray)
	Local $connectorFile = "C:\printFolder\connectorFile.txt", $success = False, $aConnector[0][3]
	Local $fileHandle = FileOpen($connectorFile)
	If $fileHandle == -1 Then Exit MsgBox(0, "Error", "The connector file could not be read. Terminating")
	For $i = 1 To _FileCountLines($connectorFile)
		_ArrayAdd($aConnector, FileReadLine($connectorFile, $i))
	Next
	FileClose($fileHandle)
	For $i = 0 To UBound($aConnector) - 1
		If $aConnector[$i][2] Then ContinueLoop
		$success = maintainInventory($aConnector[$i][0], $aConnector[$i][1], $orderArray)
		ConsoleWrite("success: " & $success & @CRLF)
		;write the result of maintaining inventory to the proper line in excel. reference orderArray to find the right line, using the production order number
		Local $row = String(_ArraySearch($orderArray, $aConnector[$i][0], 0, 0, 0, 0, 1, 7) + 2)
		Local $range = "T" & $row
		If $success Then $oClosing.Sheets("Main").Range($range ).Value = "adjusted"
		_FileWriteToLine($connectorFile, $i + 1, $aConnector[$i][0] & "|" & $aConnector[$i][1] & "|" & $success, 1)
		$range = "U" & $row
		$oClosing.Sheets("Main").Range($range).Value = $aConnector[$i][1]
		_Excel_BookSave($oClosing)
	Next
EndFunc

Func maintainInventory($productionOrderNumber, $connector, $orderArray)
	Local $ChargersWithG1Connectors = ["GN", "HN", "IN", "JN", "IP", "JP", "KP", "LP"]
	Local Static $firstTime = True
	ToolTip("processing for conn: " & $connector, 750, 0)
	ControlSetText("D:\Users\", "", "Scintilla2", "")
	ConsoleWrite($productionOrderNumber & @CRLF)
	If StringCompare("X039-6320", $connector) == 0 Then
		Local $secondPart = StringSplit($partNumber, "-")[1]
		For $i = 0 To UBound($ChargersWithG1Connectors) - 1
			If StringCompare($ChargersWithG1Connectors, $secondPart)
				$connector = $connector & "G5"
				ExitLoop
			EndIf
		Next
		If StringCompare("X039-6320", $connector) == 0 Then
			Return True;
		EndIf
	EndIf
	Local $memWindow = "tisfc0110m000", $find = "Maintain Estimated Materials - Find", $found = False
	Local $controlID = ""
	If Not WinExists($memWindow) Then expandMenu(0, 0, 0, 0, 1)
	searchForProdOrder($productionOrderNumber)
	WinActivate($memWindow)
	WinWaitActive($memWindow)
	For $i = 1 To 10
		ConsoleWrite("TsTextWinClass" & $productionOrderNumber & "     text: " & ControlGetText($memWindow, "", "TsTextWinClass" & $i) & @CRLF)
		If StringCompare(ControlGetText($memWindow, "", "TsTextWinClass" & $i), $productionOrderNumber) == 0 Then
			$found = True
			ExitLoop
		EndIf
	Next
	If  Not $found Then $found = tryRelease($orderArray, $productionOrderNumber)
	If Not $found Then
		MsgBox(0, "Error", "Production order number: " & $productionOrderNumber & " not found. Noting on spreadsheet.", 10)
		Local $index = _ArraySearch($orderArray, $productionOrderNumber, 0, 0, 0, 0, 1, 7)
		_Excel_RangeWrite($oClosing, $oWorksheet, "!!!!!!", "T" & $index + 2)
		$oClosing.Sheets("Main").Range("A" & $index + 2 & "1:T" & $index + 2).Interior.ColorIndex = 45
		_Excel_BookSave($oClosing)
		Return False
	EndIf
	Local $scrollControlID = ControlGetHandle($memWindow, "", "ScrollBar1")
	Local $maxScroll = _GUIScrollBars_GetScrollRange($scrollControlID, $SB_CTL)[1], $pos = $maxScroll, $attempts
	Do
		ControlSend($memWindow, "", $scrollControlID, "{DOWN}")
	Until _GUIScrollBars_GetScrollPos($scrollControlID, $SB_CTL) == $maxScroll
	While $controlID = ""
		If $pos == 0 Then
			$attempts += 1
			Do
				ControlSend($memWindow, "", $scrollControlID, "{DOWN}")
			Until _GUIScrollBars_GetScrollPos($scrollControlID, $SB_CTL) == $maxScroll
		EndIf
		For $i = 0 To 190
			Local $text = ControlGetText($memWindow, "", "TsTextWinClass" & $i)
			If $text == "" Then ContinueLoop
			If StringCompare($text, $connector) == 0 Then Return True
			If StringInStr($text, "X039-") Then
				$controlID = "TsTextWinClass" & $i
				ExitLoop
			EndIf
		Next
		If $controlID == "" Then
			ControlSend($memWindow, "", $scrollControlID, "{UP 6}")
		Else
			ExitLoop
		EndIf
		Sleep(1200)
		$pos = _GUIScrollBars_GetScrollPos($scrollControlID, $SB_CTL)
		If $attempts > 1 Then
			Local $index = _ArraySearch($orderArray, $productionOrderNumber, 0, 0, 0, 0, 1, 7)
			If FileExists("C:\printFolder\" & $orderArray[$index][2] & "_" & $orderArray[$index][3] & ".txt") Then
				_Excel_RangeWrite($oClosing, $oWorksheet, "no conn?", "T" & $index + 2)
				$oClosing.Sheets("Main").Range(String($prodOrderColumn & $index + 2)).Select
				MsgBox(0, "Manual Entry Required", "Could not find a connector for production order: " & $productionOrderNumber & ". Adjust manually. Worksheet will save after this dialog is closed.")
				ShellExecute("notepad.exe", $orderArray[$index][2] & "_" & $orderArray[$index][3] & ".txt")
				_Excel_BookSave($oClosing)
			Else
				MsgBox(0, "Error", "Could not read order: " & $orderArray[$index][2] & "_" & $orderArray[$index][3] & " with notepad. Does file exist?")
			EndIf
			ToolTip("")
			Return False
		EndIf
	WEnd
	ConsoleWrite("id: " & $controlID & @TAB & "current id: " & ControlGetFocus($memWindow) & @CRLF)
	If StringInStr($connector, "X039-A", 2) Or StringInStr($connector, "X039-E", 2) And $found Then Return setToZero($controlID)
	ControlClick($memWindow, "", $controlID)
	WinActivate($memWindow)
	If $firstTime = True Then
		$firstTime = False
		Sleep(200)
		$controlID = ControlGetFocus($memWindow)
	EndIf
	While StringCompare(ControlGetFocus($memWindow), $controlID) <> 0
		Sleep(10)
	WEnd
	ControlSend($memWindow, "", $controlID, $connector, 1)
	While StringCompare(ControlGetText($memWindow, "", $controlID), $connector) <> 0
		Sleep(10)
	WEnd
	ControlClick($memWindow, "", "TsDrawnButtonWinClass2")
	ToolTip("")
	Return True
EndFunc
Func searchForProdOrder($productionOrderNumber)
	ConsoleWrite("search..." & @CRLF)
	Local $memWindow = "tisfc0110m000", $find = "Maintain Estimated Materials - Find"
	WinActivate($memWindow)
	ControlClick($memWindow, "", "TsDrawnButtonWinClass8")
	Sleep(100)
	While Not WinActive($find)
		WinActivate($find)
		Sleep(250)
	WEnd
	WinActivate($memWindow)
	ControlClick($find, "", "TsTextWinClass1")
	ControlSetText($find, "", "TsTextWinClass1", $productionOrderNumber, 1)
	Sleep(200)
	ControlSend($find, "", "", "{TAB}")
	ControlSetText($find, "", "TsTextWinClass2", "80", 1)
	ControlClick($find, "", "OK")
	While WinActive($find)
		Sleep(50)
	WEnd
EndFunc
Func tryRelease($orderArray, $productionOrderNumber)
	Local $browserWin = "Menu browser", $find = "Maintain Production Orders - Find", $mpoWindow = "tisfc0101m000", $zoom = "Maintain Production Orders - Zoom"
	Local $index = _ArraySearch($orderArray, $productionOrderNumber)
	If Not WinExists($browserWin) Then openBaan()
	If Not WinExists("tisfc0101m000") Then
		WinActivate($browserWin)
		ControlSend($browserWin, "", "", "{ALT}{i}{d}")
		WinWaitActive("ttdsk1300m000")
		ControlSetText("ttdsk1300m000", "", "TsTextWinClass1", "Maintain Production Orders", 1)
		ControlClick("ttdsk1300m000", "", "OK")
		While WinActive("ttdsk1300m000")
			Sleep(50)
		WEnd
		ControlSend($browserWin, "", "", "{ENTER}")
	EndIf
	WinActivate($mpoWindow)
	WinWaitActive($mpoWindow)
	ControlClick($mpoWindow, "", "TsDrawnButtonWinClass8")
	WinWaitActive($find)
	ControlSetText($find, "", "TsTextWinClass1", $productionOrderNumber, 1)
	ControlClick($find, "", "OK")
	While WinActive($find)
		Sleep(50)
	WEnd
	WinActivate($mpoWindow)
	ControlSend($mpoWindow, "", "", "{ALT}{s}{z}")
	WinWaitActive($zoom)
	ControlSend($zoom, "", "", "{DOWN 2}{ENTER}")
	;now look again for the production order
	searchForProdOrder($productionOrderNumber)
	For $i = 0 To 10
		If StringCompare(ControlGetText($mpoWindow, "", "TsTextWinClass" & $i), $productionOrderNumber) == 0 Then Return True
	Next
	Return False
EndFunc
Func setToZero($controlID)
	Local Static $firstTime = True
	Local $memWindow = "tisfc0110m000"
	ConsoleWrite("setting to zero" & @TAB & "control: " & $controlID & @CRLF)
	ControlClick($memWindow, "", $controlID)
	If $firstTime Then
		$firstTime = False
		Sleep(300)
		$controlID = ControlGetFocus($memWindow)
	EndIf
	While StringCompare(ControlGetFocus($memWindow), $controlID) <> 0
		Sleep(10)
	WEnd
	Local $yPosition = _WinGetCaretPos()[1]
	ControlSend($memWindow, "", "", "{TAB 3}")
	Sleep(150)
	Do
		Sleep(50)
		WinActivate($memWindow)
		$controlID = ControlGetFocus($memWindow)
	Until _WinGetCaretPos()[1] == $yPosition
	ControlSend($memWindow, "", $controlID, "0", 1); set the qtty to 0
	Sleep(100)
	ControlClick($memWindow, "", "TsDrawnButtonWinClass2"); save
	Sleep(100)
	Return True
EndFunc
Func trimSalesOrder($orderArray, $index)
	ConsoleWrite("trim sales order." & @CRLF)
	Local $trimmedOrder[0], $connector
	Local $orderNumber = $orderArray[$index][2], $lineNumber = Number($orderArray[$index][3]), $orderAndLine = $orderNumber & "_" & $lineNumber
	Local $entireSalesOrder = FileReadToArray("C:\printFolder\order_" & $orderAndLine & ".txt")
	ConsoleWrite("trim sales..." & @CRLF)
	If UBound($entireSalesOrder) > 30 Then
		getHeader($trimmedOrder, $entireSalesOrder)
		_ArrayAdd($trimmedOrder, " ")
	Else
		MsgBox(0, "Error", "The order " & $orderNumber & " was not found or did not read to an array correctly.")
		Return 0
	EndIf
	_ArrayAdd($trimmedOrder, getLineItem($entireSalesOrder, $lineNumber, $connector, true))
	if $exportFlag Then
		_Excel_RangeWrite($oClosing, $oWorksheet, "Export", "U" & $index)
		$exportFlag = False
	EndIf
	Local $shipCompleteItems = shipComplete($orderArray, $index)
	Local $addOnIndex = _ArraySearch($addOnArray, $orderNumber, 0, 0, 0, 0, 1, 0)
	If $addOnIndex >= 0 Then
		_ArrayAdd($trimmedOrder, getLineItem($entireSalesOrder, $addOnArray[$addOnIndex][1], $connector, False)) ;don't change the $connector value
	EndIf
	If UBound($shipCompleteItems) Then _ArrayAdd($trimmedOrder, $shipCompleteItems)
	If StringLen($connector) > 3 Then
		_ArrayAdd($trimmedOrder, "Connector: " & $connector)
	Else
		$connector = 0
		ShellExecute("notepad.exe", "C:\printFolder\order_" & $orderAndLine & ".txt")
		While $connector == 0
			$connector = InputBox("Is the connector on the order?", "If connector found, enter now. If not found, click 'Cancel'. Order will be noted on spread sheet.", "X039-")
			If @error == 1 Then
				Return 0
			EndIf
		WEnd
	EndIf
	_ArrayAdd($trimmedOrder, "Production: " & $orderArray[$index][7] & @CRLF)
	ConsoleWrite("added all lines to array..." & @CRLF)
	_FileWriteFromArray("C:\printFolder\" & $orderNumber & "_" & $lineNumber & ".txt", $trimmedOrder)
	If FileExists("C:\printFolder\" & $orderAndLine & ".txt") Then $orderArray[$index][8] = "Processed"
	$oWorksheet = $oClosing.Sheets("Main")
	_Excel_RangeWrite($oClosing, $oWorksheet, "Processed", "S" & $index + 2)
	_Excel_RangeWrite($oClosing, $oWorksheet, $connector, "T" & $index + 2)
	_Excel_BookSave($oClosing)
	ConsoleWrite("file written." & @CRLF)
	Return $connector
EndFunc

Func trimRPO($orderArray, $index)
	ConsoleWrite("Trim RPO" & @CRLF)
	Local $trimmedRPO[0], $orderNumber = $orderArray[$index][2], $lineNumber = Number($orderArray[$index][3]), $orderAndLine = $orderNumber & "_" & $lineNumber
	Local $filePath = "C:\printFolder\order_" & $orderAndLine & ".txt", $fileIndexOfLineNumber, $aRPO[0], $connector = "", $line
	If Not FileExists($filePath) Then
		MsgBox(0, "Error", "The RPO " & $orderNumber & " does not exist in the print folder. Noting as 'error' on spreadsheet.")
		$oClosing.Sheets("Main").Range("S" & $index + 2).Value = "file error"
		Return
	EndIf
	While StringLen(StringStripWS(FileReadLine($filePath, 1), 8)) == 0 ;remove all blank space above "Replenishment Acknowledgement"
		_FileWriteToLine($filePath, 1, "", True)
	WEnd
	Local $firstLineNumber = getFirstLineItemOnRPO($filePath)
	Local $eof = _FileCountLines($filePath)
	For $j = 1 To $firstLineNumber - 1
		$line = FileReadLine($filePath, $j) ;get all text up to the first line number
		_ArrayAdd($aRPO, $line)
	Next
	For $i = $firstLineNumber To $eof
		$line = FileReadLine($filePath, $i)
		If isRPOlineItem($line) == $lineNumber Then ;look for first line number. once found, put everything associated with it in the array
			$fileIndexOfLineNumber = $i
			ConsoleWrite("file index " & $fileIndexOfLineNumber & @CRLF)
			While 1
				_ArrayAdd($aRPO, $line)
				If $i + 1 >= $eof Then ExitLoop
				$line = FileReadLine($filePath, $i + 1)
				If isRPOlineItem($line) Then ExitLoop
				$i += 1
			WEnd
		EndIf
		If isRPOlineItem($line) == $lineNumber + 1 Then ;we should already be at the next line item if one exists.
			If (Not chargerPartNumber(getPartNumber($line))) And (getQtty(FileReadLine($filePath, $fileIndexOfLineNumber), 1) == getQtty($line, 2)) Then
				While 1  ;the above boolean expression tests to see if the next line item is NOT a charger, AND if it has the same qtty as the charger line item. if it does, then add it to the array
					_ArrayAdd($aRPO, $line)
					If $i + 1 >= $eof Then ExitLoop
					$line = FileReadLine($filePath, $i + 1)
					If isRPOlineItem($line) Then ExitLoop
					$i += 1
				WEnd
			EndIf
			For $k = 0 To UBound($aRPO) - 1
				If StringInStr($aRPO[$k], "Connector:") Then
					$connector = StringSplit(StringStripWS(StringTrimLeft($aRPO[$k], 10), 1), " ")[0]
				EndIf
			Next
			If StringLen($connector) == 0 Then $connector = "X039-6320"
			_ArrayAdd($aRPO, "Connector: " & $connector)
			_ArrayAdd($aRPO, "Production: " & $orderArray[$index][8] & @CRLF)
			_FileWriteFromArray("C:\printFolder\" & $orderAndLine & ".txt", $aRPO)
			Return
		EndIf
		ContinueLoop ;if it's not the line number, don't worry about it
	Next
	$aRPO = _ArrayUnique($aRPO)
	_ArrayAdd($aRPO, "Connector: " & $connector)
	_ArrayAdd($aRPO, "Production: " & $orderArray[$index][8])
	_FileWriteFromArray("C:\printFolder\" & $orderAndLine & ".txt", $aRPO)
	If FileExists("C:\printFolder\" & $orderAndLine & ".txt") Then $orderArray[$index][8] = "Processed"
	$oWorksheet = $oClosing.Sheets("Main")
	_Excel_RangeWrite($oClosing, $oWorksheet, "Processed", "S" & $index + 2)
	_Excel_RangeWrite($oClosing, $oWorksheet, $connector, "T" & $index + 2)
	_Excel_BookSave($oClosing)
	ConsoleWrite("file written." & @CRLF)
EndFunc

Func getQtty($line, $count = 0)
	Local $qtty = StringSplit($line, " ")
	For $i = 0 To UBound($qtty) - 1
		If StringInStr($qtty[$i], ".000") Then Return Number($qtty[$i])
	Next
	Return 0
EndFunc

Func getPartNumber($line)
	Local $partNumber = StringRegExp($line, "(?:\s{1,2}[0-9]{1,2}\s{2})(\S{4,18})", 3)
	If UBound($partNumber) > 0 Then	Return $partNumber[0]
EndFunc

Func chargerPartNumber($partNumber)
	Local $Impaq1 = "EI1-"
	Local $Impaq3 = "EI3-"
	Local $ImpaqPlus1 = "EIP1"
	Local $ImpaqPlus3 = "EIP3"
	Local $Nexus1 = "NI1-"
	Local $Nexus3 = "NI3-"
	Local $NexusPlus1 = "NIP1"
	Local $NexusPlus3 = "NIP3"
	Local $Douglas1 = "DL1-"
	Local $Douglas2 = "DL2-"
	If StringLen($partNumber) < 10 Or StringLen($partNumber) >= 16 Then Return False
	Local $modelArray[10] = [$Impaq1, $Impaq3, $ImpaqPlus1, $ImpaqPlus3, $Nexus1, $Nexus3, $NexusPlus1, $NexusPlus3, $Douglas1, $Douglas2]
	For $i = 0 To 7 Step 1
		If StringInStr($partNumber, $modelArray[$i], 1, 1, 1, 4) >= 1 Then Return True
	Next
	Return False
EndFunc

Func getFirstLineItemOnRPO($filePath)
	Local $line
	For $i = 0 To _FileCountLines($filePath)
		$line = FileReadLine($filePath, $i)
		If Not isRPOlineItem($line) Then ContinueLoop
		Return $i
	Next
	MsgBox(0, "error", "First line item of RPO was never found.")
	Return -1
EndFunc

Func isRPOlineItem($line)
	If Not StringInStr($line, ".000   ea") Then Return 0
	Local $lineNumber = StringRegExp($line, "(?:\s{2})([0-9]{1,2})", 3)
	If UBound($lineNumber) And ($lineNumber[0] < 100 And $lineNumber[0] >= 1) Then
		Return Number($lineNumber[0])
	EndIf
	Return 0
EndFunc

Func getAddOnArray()
	ConsoleWrite("In getAddon" & @CRLF)
	Local $aParts[0], $endRow = 0
	If $extrasPage == "" Then Return $addOnArray
	If IsObj($oClosing.Sheets($extrasPage)) Then $oClosing.Sheets($extrasPage).Activate
	$endRow = $oClosing.Sheets($extrasPage).Range("A1").SpecialCells($xlCellTypeLastCell).Row
	For $i = 2 To $endRow
		Local $cellContents = $oClosing.Sheets($extrasPage).Range($orderNumberColumn & $i).Value
		If $cellContents = "" Then ContinueLoop
		_ArrayAdd($addOnArray, $cellContents & "|" & $oClosing.Sheets($extrasPage).Range($PositionColumn & $i).Value)
	Next
	Return $addOnArray
EndFunc
Func shipComplete($orderArray, $index)
	Local $aLineItems[0], $chargerIndex, $aLinesToAdd[0], $row = $index + 2, $warehouse
	ConsoleWrite("ship complete" & @CRLF)
	Local $aFile = FileReadToArray("C:\printFolder\order_" & $orderArray[$index][2] & "_" & $orderArray[$index][3] & ".txt")
	For $i = 30 To UBound($aFile) - 1
		$warehouse = StringRegExp($aFile[$i], '(?:\d{2}-\d{2}-\d{4})(?:\s{3})([0-9]{3})', 1)
		if UBound($warehouse) > 0 Then
			if ($warehouse[0] == 264) Then
				_ArrayAdd($aLinesToAdd, $aFile[$i])
				ExitLoop
				EndIf
		EndIf
	Next
	If UBound($aLinesToAdd) > 0 Then
		markExcel($row)
		Return $aLinesToAdd
	EndIf
	Return 0
EndFunc
Func warehouseMatches264($line, $date)
	ConsoleWrite("wrhse match" & @TAB & "date: " & $date & @CRLF)
	Local $warehouse = StringRegExp($line, '(?:\d{2}-\d{2}-\d{4})(?:\s{3})([0-9]{3})', 1)
	If UBound($warehouse) > 0 Then
		If $wareHouse[0] == 264 Then Return True
	EndIf
	Return False
EndFunc
Func markExcel($row)
	$oClosing.Sheets(1).Range("A" & $row & ":V"& $row).Interior.ColorIndex = 7
	$oClosing.Sheets(1).Range("V" & $row).Value = "Ship Complete!"
EndFunc
Func getHeader(ByRef $aHeader, $entireSalesOrder)
	For $line In $entireSalesOrder
		If getLineNumber($line) >= 5 Then Return
		$line = StringReplace($line, "|", " ")
		_ArrayAdd($aHeader, $line)
		_ArrayDelete($entireSalesOrder, 0)
	Next
	ConsoleWrite("get header")
EndFunc
Func getLineNumber($line)
	Local $splitLine = _StringBetween($line, "[", "]")
	If UBound($splitLine) >= 1 Then
		For $entry In $splitLine
			If Number($entry) Then Return Number($entry)
		Next
	EndIf
EndFunc
Func getLineItem($entireSalesOrder, $lineNumber, ByRef $connector, $find = True, $getExport = false)
	Local $linesLeft = UBound($entireSalesOrder) - 1, $aLineItem[0], $currentIndex
	If $find Then $connector = 0
	ConsoleWrite("getting line item..." & @CRLF)
	For $i = 0 To $linesLeft
		If getLineNumber($entireSalesOrder[$i]) == Number($lineNumber) Then
			Do
				If StringLen(StringStripWS($entireSalesOrder[$i], 8)) > 0 Then
					$entireSalesOrder[$i] = StringReplace($entireSalesOrder[$i], "|", " ")
					If headerLines($entireSalesOrder[$i]) Then
						$i += 1
						ContinueLoop
					EndIf
					_ArrayAdd($aLineItem, $entireSalesOrder[$i])
					If $find And StringLen(String($connector)) < 4 Then
						Local $_connector = StringRegExp($entireSalesOrder[$i], "(?:CONNECTOR:\s{0,10})(\S{4,20})", 3)
						If UBound($_connector) > 0 Then $connector = $_connector[0]
						If Number($connector) Then $connector = "X039-" & Number($connector)
					EndIf
				EndIf
				$i += 1
				If StringInStr($entireSalesOrder[$i], "[") And StringInStr($entireSalesOrder[$i], "]") Then
					$currentIndex = $i
					ExitLoop
				EndIf
			Until $i >= $linesLeft
			If $getExport Then
				For $j = 0 To $linesLeft
					If StringInStr($entireSalesOrder[$j], "EXPORT") Then
						_ArrayAdd($aLineItem, $entireSalesOrder[$j])
						$exportFlag = true
					EndIf
				Next
			EndIf
		EndIf
	Next
	Return $aLineItem
EndFunc

Func headerLines($line)
	If StringInStr($line, "-------", 2) Then Return True
	If StringInStr($line, "Quantity Unit  Item") Then Return True
	If StringInStr($line, 'ORDER ACKNOWLEDGEMENT', 2) Then Return True
	If StringInStr($line, "Customer    :", 2, 1, 1, 13) Then Return True
	If Number(StringStripWS($line, 8)) > 6000 Then Return True
	If StringInStr($line, "Delivery  :", 2) Then Return True
	If StringInStr($line, "Payment", 2, 1, 1, 7) Then Return True
	Return False
EndFunc
Func checkMissedFiles($orderArray)
	Local $orderFiles = _FileListToArray("C:\printFolder"), $orders[0], $missedPrints[0]
	_ArrayDelete($orderFiles, 0)
	For $i = 0 To UBound($orderFiles) - 1
		If Not StringInStr($orderFiles[$i], "order_") Then ContinueLoop
		$orderFiles[$i] = StringTrimRight(StringTrimLeft($orderFiles[$i], 6), 4)
	Next
	For $i = 0 To UBound($orderArray) - 1
		If _ArraySearch($orderFiles, $orderArray[$i][2] & "_" & $orderArray[$i][3]) == -1 Then
			_ArrayAdd($missedPrints, $orderArray[$i][2] & "_" & $orderArray[$i][3])
			$oWorksheet = $oClosing.Sheets("Main")
			_Excel_RangeWrite($oClosing, $oWorksheet, "", "S" & $i + 2)
			_Excel_BookSave($oClosing)
			$orderArray[$i][8] = "failed"
		EndIf
	Next
	_Excel_BookSave($oClosing)
	If UBound($missedPrints) > 0 Then
		Local $display = MsgBox(4, "Display Array?", 'If you want to display the array of missed orders, click "YES"', 20)
		If $display == 6 Then _ArrayDisplay($missedPrints)
		Local $newArray[0][9]
		For $i = 0 To UBound($missedPrints) - 1
			Local $index =  _ArraySearch($orderArray, $missedPrints[$i])
			If $index <> -1 Then _ArrayAdd($newArray, $orderArray[$index][0] & "|" & $orderArray[$index][1] & "|" & $orderArray[$index][2] & "|" & $orderArray[$index][3] & "|" & $orderArray[$index][4] & "|" & $orderArray[$index][5] & "|" & $orderArray[$index][6] & "|" & $orderArray[$index][7] & "|" & $orderArray[$index][8])
		Next
		printOrdersToTextFiles($newArray)
	EndIf
EndFunc
Func splitFiles($orderArray)
	Local $fileList = _FileListToArray("C:\printfolder"), $orderNumber, $lineItems[0], $i = 0
	_ArrayDelete($fileList, 0)
	While $i < UBound($fileList)
		If Not StringInStr($fileList[$i], "batch") Then
			_ArrayDelete($fileList, $i)
			ContinueLoop
		EndIf
		$i += 1
	WEnd
	For $fileName In $fileList
		Local $order[0]
		Local $file = FileReadToArray("C:\printFolder\" & $fileName)
		For $i = 0 To UBound($file)
			If StringLen(StringStripWS($file[$i], 8)) == 0 Then ContinueLoop
			If StringInStr($file[$i], "order acknowledgement", 2) And StringInStr($file[$i], "original", 2) Then
				Do
					_ArrayAdd($order, $file[$i])
					$i += 1
				Until (StringInStr($file[$i], "order acknowledgement", 2) And StringInStr($file[$i], "original", 2)) Or ($i == UBound($file) - 1)
			EndIf
			$i -= 1
			For $line In $order
				If StringInStr($line, "Sales Order  :", 2) Then
					$orderNumber = Number(StringTrimLeft($line, StringInStr($line, "Sales Order  :", 2) + 14))
					ExitLoop
				EndIf
			Next
			Local $indeces = _ArrayFindAll($orderArray, $orderNumber, 0, 0, 0, 0, 2)
			For $index In $indeces
				_FileWriteFromArray("C:\printFolder\order_" & $orderNumber & "_" & $orderArray[$index][3] & ".txt", $order)
			Next
			_ArrayDelete($lineItems, "0-" & UBound($lineItems) - 1)
			_ArrayDelete($order, "0-" & UBound($order) - 1)
			$orderNumber = ""
			If $i > UBound($file) - 5 Then ExitLoop
		Next
		FileDelete("C:\printFolder\" & $fileName)
	Next
EndFunc
Func isValid($item)
	Local $Impaq1 = "EI1-"
	Local $Impaq3 = "EI3-"
	Local $ImpaqPlus1 = "EIP1"
	Local $ImpaqPlus3 = "EIP3"
	Local $Nexus1 = "NI1-"
	Local $Nexus3 = "NI3-"
	Local $NexusPlus1 = "NIP1"
	Local $NexusPlus3 = "NIP3"
	If StringLen($item) < 10 Or StringLen($item) >= 16 Then Return False
	Local $modelArray[8] = [$Impaq1, $Impaq3, $ImpaqPlus1, $ImpaqPlus3, $Nexus1, $Nexus3, $NexusPlus1, $NexusPlus3]
	For $i = 0 To 7 Step 1
		If StringInStr($item, $modelArray[$i], 1, 1, 1, 4) >= 1 Then Return True
	Next
	Return False
EndFunc
Func openExcel()
	$oExcel = _Excel_Open()
	$ExcelFileName = FileOpenDialog("Week Closing Report", @DesktopDir & "\Reports\", "Microsoft Excel Document(*.xlsx)")
	If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Sales Order Printer", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	$oClosing = _Excel_BookOpen($oExcel, $ExcelFileName, False)
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "Packet Creator Error", "Error opening workbook '" & $ExcelFileName & @CRLF & "@error = " & @error & ", @extended = " & @extended)
		_Excel_BookClose($oExcel)
		Exit
	EndIf
	Local $findExtras = MsgBox(4, "Extras?", "Is there a sheet that contains non-charger line items?")
	If $findExtras == 6 Then
		$extrasPage = InputBox("Any extra line items?", "If there is a sheet that lists extra parts, what sheet is it? Enter the name with quotes around it or the page's index: ", 'PARTS')
		If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Packet Creator Error", "Error reading from workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	Else
		$extrasPage = 0
	EndIf
	$oClosing.Sheets(1).Name = "Main"
	_Excel_BookSave($oClosing)
	Return assignColumns()
EndFunc
Func assignColumns()
	Local $headingArray[0]
	Local $foundCount = 0, $columnLetterInASCI = 65 ;letter 'A'
	Local $endRow = $oClosing.Sheets("Main").Range("A1").SpecialCells($xlCellTypeLastCell).Row
	$oWorksheet = $oClosing.Sheets("Main")
	_Excel_RangeWrite($oClosing, $oWorksheet, "Printed?", "S1")
	_Excel_RangeWrite($oClosing, $oWorksheet, "ProdOrd", "R1")
	_Excel_RangeSort($oClosing, $oWorksheet, "A2:S" & $endRow, "E:E")
	_Excel_BookSave($oClosing)
	For $columnLetterInASCI = 65 To 120
		If  StringCompare(_Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1"), "TYPE OF ORDER") == 0 Then
			$TypeOfOrderColumn = Chr($columnLetterInASCI)
			_ArrayAdd($headingArray, "TYPE OF ORDER")
			ContinueLoop
		EndIf
		If StringCompare(_Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1"), "ORDER TYPE") == 0 Then
			$OrderTypeColumn = Chr($columnLetterInASCI)
			_ArrayAdd($headingArray, "ORDER TYPE")
			ContinueLoop
		EndIf
		If StringCompare(_Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1"), "ORDER") == 0 Then
			$orderNumberColumn = Chr($columnLetterInASCI)
			_ArrayAdd($headingArray, "ORDER")
			ContinueLoop
		EndIf
		If StringCompare(_Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1"), "POS") == 0 Then
			$PositionColumn = Chr($columnLetterInASCI)
			_ArrayAdd($headingArray, "POS")
			ContinueLoop
		EndIf
		If StringCompare(_Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1"), "Item") == 0 Then
			$ItemColumn = Chr($columnLetterInASCI)
			_ArrayAdd($headingArray, "Item")
			ContinueLoop
		EndIf
		If StringCompare(_Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1"), "FRIDAY DATE") == 0 Then
			$DateColumn = Chr($columnLetterInASCI)
			_ArrayAdd($headingArray, "FRIDAY DATE")
			ContinueLoop
		EndIf
		If StringCompare(_Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1"), "QTY") == 0 Then
			$QttyColumn = Chr($columnLetterInASCI)
			_ArrayAdd($headingArray, "QTY")
			ContinueLoop
		EndIf
		If StringCompare(_Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1"), "ProdOrd") == 0 Then
			$prodOrderColumn = Chr($columnLetterInASCI)
			_ArrayAdd($headingArray, "ProdOrd")
			ContinueLoop
		EndIf
		If StringCompare(_Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1"), "Printed?") == 0 Then
			$printedColumn = Chr($columnLetterInASCI)
			_ArrayAdd($headingArray, "Printed?")
			ContinueLoop
		EndIf
		If _Excel_RangeRead($oClosing, $oWorksheet, Chr($columnLetterInASCI) & "1") = "" And $columnLetterInASCI > 84 Then ExitLoop
	Next
	$oClosing.Sheets("Main").Activate
	_Excel_BookSaveAs($oClosing, @DesktopDir & "\ClosingCSV.csv", $xlCSV, True)
	_Excel_BookClose($oClosing)
	$oExcel = _Excel_Open()
	$oClosing = _Excel_BookOpen($oExcel, $ExcelFileName, False)
	Local $orderArray = readCSVfileToArray()
	If FileExists( @DesktopDir & "\Closing.csv") Then FileDelete( @DesktopDir & "\Closing.csv")
	$addOnArray = getAddOnArray() ;declared outside of all functions
	Local $numberOfColumns = UBound($orderArray, $UBOUND_COLUMNS) - 1
	Local $aBool[$numberOfColumns + 1]
	For $i = 0 To $numberOfColumns
		If _ArraySearch($headingArray, $orderArray[0][$i]) == -1 Then
			$aBool[$i] = 0
		Else
			$aBool[$i] = 1
		EndIf
	Next
	Local $arrayLine = "", $newArray[0][9], $splitArrayLine
	;cleaning up the order array since code was written only for valid columns. it's easier than altering the entire code base
	;and paring down the information used is better to do in an array than it is to do on the excel spreadsheet (stability of script is compromised with Excel reading. reading from an array is also much faster)
	For $i = 1 To UBound($orderArray) - 1 ;skip the first line in the array since it only contains headings
		For $j = 0 To $numberOfColumns
			If $aBool[$j] Then $arrayLine = $arrayLine & $orderArray[$i][$j] & "|"
		Next
		$arrayLine = StringReplace($arrayLine, ",", "|") ;concatenate with the "|"
		$arrayLine = StringTrimRight($arrayLine, StringLen($arrayLine) - StringInStr($arrayLine, "|", Default, 9) + 1) ;trim the last concatenation marker
		_ArrayAdd($newArray, $arrayLine)
		$arrayLine = ""
	Next
	Return $newArray
EndFunc
Func readCSVfileToArray()
	Local $file = @DesktopDir & "\ClosingCSV.csv", $splitLine, $newLine
	Local $columnCount = UBound(StringSplit(FileReadLine($file, 2), ",", 2))
	Local $fileArray[0][$columnCount]
	For $i = 1 To _FileCountLines($file)
		Local $line = FileReadLine($file, $i)
		If StringInStr($line, '"') Then
			$splitLine = StringSplit($line, '"', 2);there is a description that contains commas. this messes up the column count.
			$newLine = $splitLine[0] & $splitLine[2] ;this relies on there only being one set of parenthesis.
		Else
			$newLine = $line
		EndIf
		$newLine = StringReplace($newLine, ",", "|")
		_ArrayAdd($fileArray, $newLine)
	Next
	Return $fileArray
EndFunc

Func printOrdersToTextFiles($orderArray, $missedPrints = 0)
	Local $browserWin = WinGetTitle("Menu browser"), $expandPOAWindow = True, $tenOrderArray[0][2]
	$oWorksheet = $oClosing.Sheets("Main")
	If Not WinExists("Menu browser") Then openBaan()
	If StringInStr($browserWin, $mfgCompany) Then changeCompany()
	If Not WinExists("Print Order Acknowledgements") Then expandMenu($expandPOAWindow)
	For $i = 0 To UBound($orderArray, 1) - 1
		If StringCompare($orderArray[$i][8], "failed") == 0 Then printWithAdditionalDetail($orderArray, $i)
		If StringLen($orderArray[$i][8]) > 1 Then ContinueLoop
		If StringCompare($orderArray[$i][0], "Sales Order") == 0 Then
			ConsoleWrite($orderArray[$i][2] & @CRLF)
			If StringCompare($orderArray[$i][1], "USP") == 0 Then
				printWithAdditionalDetail($orderArray, $i)
				dontSleep()
				ContinueLoop
			EndIf
			If UBound($tenOrderArray,$UBOUND_ROWS) == 10 Then
				fillPrintOrderScreen($tenOrderArray, $orderArray)
				If $i == UBound($orderArray) - 1 Then Return
			EndIf
			If _ArraySearch($tenOrderArray, $orderArray[$i][2]) == -1 Then
				_ArrayAdd($tenOrderArray, $orderArray[$i][2] & "|" & $i)
			Else
				_Excel_RangeWrite($oClosing, $oWorksheet, "duplicate", "S" & $i + 2)
			EndIf
		EndIf
		If StringInStr($orderArray[$i][0], "Replenish") Then
			Local $success = printRPOtoTextFile($orderArray[$i][2], $orderArray[$i][3])
			If $success Then
				_Excel_RangeWrite($oClosing, $oWorksheet, "printed", "S" & $i + 2)
			Else
				_Excel_RangeWrite($oClosing, $oWorksheet, "failed", "S" & $i + 2)
			EndIf
			_Excel_BookSave($oClosing)
		EndIf
	Next
	If UBound($tenOrderArray) > 0 Then
		fillPrintOrderScreen($tenOrderArray, $orderArray)
	EndIf
EndFunc

Func fillPrintOrderScreen(ByRef $tenOrderArray, $orderArray)
	Local $zeroArray[0][2], $successHandle = "", $window = "tdsls4401m000"
	expandMenu(1)
	dontSleep()
	If UBound($tenOrderArray) > 0 Then
	Do
		setPOAwindow()
		set_ESSONwindowToZero()
		Local $success = fill_ESSONwindow($tenOrderArray)
	Until $success
		ControlClick($window, "", "TsPushButtonWinClass1")
		$successHandle = fillSelectDeviceWindow()
	EndIf
	If StringLen(WinGetTitle($successHandle)) == 0 Then
		ControlClick($successHandle, "", "TsPushButtonWinClass1")
		For $j = 0 To UBound($tenOrderArray) - 1
			_Excel_RangeWrite($oClosing, $oWorksheet, "printed", "S" & $tenOrderArray[$j][1] + 2)
			If $oClosing.Sheets("Main").Range("A" & $tenOrderArray[$j][1] + 2).Interior.ColorIndex == 28 Then $oClosing.Sheets("Main").Range("A" & $tenOrderArray[$j][1] + 2 & ":R"& $tenOrderArray[$j][1] + 2).Interior.ColorIndex = -4142
		Next
	EndIf
	If WinExists("ttstps0014") Then
		MsgBox(0, "Error printing", "One of the orders in the queue did not print. Marking all orders on the spreadsheet, highlighting in teal. Will attempt to find rogue order.", 10)
		ControlClick("ttstps0014", "", "TsPushButtonWinClass1")
		For $j = 0 To UBound($tenOrderArray) - 1
			_Excel_RangeWrite($oClosing, $oWorksheet, "failed", "S" & $tenOrderArray[$j][1] + 2)
			$oClosing.Sheets("Main").Range("A" & $tenOrderArray[$j][1] + 2 & ":R"& $tenOrderArray[$j][1] + 2).Interior.ColorIndex = 28
			_Excel_BookSave($oClosing)
		Next
		singleOutBadOrder($tenOrderArray, $orderArray)
		Return
	EndIf
	_Excel_BookSave($oClosing)
	$tenOrderArray = $zeroArray
EndFunc
Func singleOutBadOrder(ByRef $tenOrderArray, $orderArray)
	Local $POAwindow = "tdsls4401m000", $sleepCount = 0, $errorWindow = "ttstps0014"
	ControlClick($errorWindow, "", "OK")
	Sleep(500)
	If Not (ControlGetText($POAwindow, "", "TsTextWinClass1") == "No") Then
		ControlClick($POAwindow, "", "TsTextWinClass1")
		Sleep(500)
		ControlSend($POAwindow, "", "TsTextWinClass1", "n", 1)
		Sleep(250)
		ControlSend($POAwindow, "", "TsTextWinClass1", "{TAB 3}")
	EndIf
	For $i = 0 To UBound($tenOrderArray) - 1
		ControlClick($POAwindow, "", "TsTextWinClass4")
		Do
			Sleep(200)
			WinActivate($POAwindow)
		Until StringInStr(ControlGetFocus($POAwindow), "4")
		Do
			If $sleepCount == 0 Then ControlSend($POAwindow, "", "TsTextWinClass4", $tenOrderArray[$i][0], 1)
			Sleep(200)
			$sleepCount += 1
			If $sleepCount >= 15 Then
				$sleepCount = 0
				ControlSend($POAwindow, "", "TsTextWinClass4", "{SHIFTDOWN}{HOME}{BS}{SHIFTUP}")
				Sleep(200)
			EndIf
			ConsoleWrite("get text: " & ControlGetText($POAwindow, "", "TsTextWinClass4") & "." & @CRLF)
		Until StringCompare(ControlGetText($POAwindow, "", "TsTextWinClass4"), $tenOrderArray[$i][0]) == 0
		$sleepCount = 0
		ControlSend($POAwindow, "", "", "{TAB}")
		Sleep(200)
		ControlClick($POAwindow, "", "Continue")
		Do
			Sleep(100)
			WinActivate("ttstpsplopen")
		Until StringCompare(ControlGetFocus("ttstpsplopen"), "TsTextWinClass11")
		Sleep(1000)
		Do
			If $sleepCount == 0 Then ControlSend("ttstpsplopen", "", "TsTextWinClass11", "D", 1)
			If WinExists("ttstpopen7") Then ControlClick("ttstpopen7", "", "OK")
			If WinExists("ttaadd3001") Then ControlClick("ttaadd3001", "", "OK")
			Sleep(100)
			$sleepCount += 1
			If $sleepCount >= 30 Then
				$sleepCount = 0
				ControlSend("ttstpsplopen", "", "TsTextWinClass11", "{SHIFTDOWN}{HOME}{BS}{SHIFTUP}")
				Sleep(200)
			EndIf
		Until StringCompare(ControlGetText("ttstpsplopen", "", "TsTextWinClass11"), "D") == 0
		Sleep(250)
		ControlSend("ttstpsplopen", "", "", "{TAB}")
		Do
			Sleep(100)
			WinActivate("ttstpsplopen")
		Until StringInStr(ControlGetFocus("ttstpsplopen"), "TsPushButtonWinClass")
		Local $currentHandle = WinGetHandle("ttstpsplopen"), $newHandle = ""
		ControlClick("ttstpsplopen", "", "Continue")
		Do
			Sleep(100)
			$newHandle = WinGetHandle("ttstpsplopen")
		Until $newHandle > 0 And StringCompare($currentHandle, $newHandle) <> 0
		ControlClick("ttstpsplopen", "", "Continue")
		Do
			If WinExists("ttstps0014") Then
				WinClose("ttstps0014")
				_Excel_RangeWrite($oClosing, Default, "FAILED", $tenOrderArray[$i][1] + 2)
				Local $temp = $tenOrderArray[$i][0]
				_ArrayDelete($tenOrderArray, $i)
				Local $index = _ArraySearch($orderArray, $temp)
				$orderArray[$index][8] = "failed"
				Return
			EndIf
			If WinExists("Display Browser") Then
				WinClose("Display Browser")
				ExitLoop
			EndIf
		Until 0
	Next
EndFunc
Func fillSelectDeviceWindow()
	Local Static  $nextNumber = 0, $firstTime = True, $window = "ttstpsplopen : Select"
	Local $company = 701
	If $firstTime Then
		$firstTime = False
		Local $aList = _FileListToArray("C:\printFolder"),$batchNumbersUsed[0]
		If UBound($aList) > 1 Then
			_ArrayDelete($aList, 0)
			For $file In $aList
				If StringInStr($file, "batch") Then _ArrayAdd($batchNumbersUsed, StringMid($file, 7, StringLen($file) - 10))
			Next
			_ArraySort($batchNumbersUsed)
			If UBound($batchNumbersUsed) > 0 Then $nextNumber = Number($batchNumbersUsed[UBound($batchNumbersUsed) - 1])
		EndIf
	EndIf
	$nextNumber += 1
	WinWaitActive($window)
	Do
		Sleep(100)
	Until StringCompare("TsTextWinClass11", ControlGetFocus($window)) == 0
	Do
		ControlSend($window, "", "TsTextWinClass11", "DOWNLOAD", 1)
		Do
			Sleep(250)
		Until StringCompare(StringRight(ControlGetText($window, "", "TsTextWinClass11"), 1), "D") == 0
		If StringCompare(ControlGetText($window, "", "TsTextWinClass11"), "DOWNLOAD") == 0 Then ExitLoop
		ControlSend($window, "", "TsTextWinClass11", "{SHIFTDOWN}{HOME}{BS}{SHIFTUP}")
		Sleep(1500)
	Until 0
	ControlSend($window, "", "", "{TAB}")
	Do
		Sleep(100)
	Until StringCompare("TsTextWinClass9", ControlGetFocus($window)) == 0
	Do
		ControlSend($window, "", "TsTextWinClass9", "C:\printFolder\batch_" & $nextNumber & ".txt", 1)
		Do
			Sleep(200)
		Until StringCompare(StringRight(ControlGetText($window, "", "TsTextWinClass9"), 4), ".txt") == 0
		If StringCompare(ControlGetText($window, "", "TsTextWinClass9"), "C:\printFolder\batch_" & $nextNumber & ".txt") == 0 Then ExitLoop
		ControlSend($window, "", "TsTextWinClass11", "{SHIFTDOWN}{HOME}{DEL}{SHIFTUP}")
		Sleep(500)
	Until 0
	Local $currentHandle = WinGetHandle($window), $newHandle = ""
	ControlClick($window, "", "Continue")
	Do
		Sleep(50)
		$newHandle = WinGetHandle($window)
	Until $newHandle > 0 And StringCompare($newHandle, $currentHandle) <> 0
	ControlClick($window, "", "Continue")
	Do
		Sleep(50)
		$newHandle = WinGetHandle("[CLASS:TsShellWinClass]", "OK")
	Until $newHandle > 0
	Return $newHandle
EndFunc
Func setPOAwindow()
	Local $window = "tdsls4401m000"
	Local $allText = ControlGetText($window, "", "TsTextWinClass8")
	If Not (StringCompare($allText, "All Lines") == 0) Then
		Do
			ControlClick($window, "", "TsTextWinClass8")
			Sleep(500)
			ControlSend($window, "", "TsTextWinClass8", "a", 1)
			ControlSend($window, "", "", "{TAB}")
		Until StringCompare(ControlGetText($window, "", "TsTextWinClass8"), "All Lines") == 0
	EndIf
	WinActivate($window)
	ControlClick($window, "", "TsTextWinClass1")
	Sleep(500)
	If StringCompare("Yes", ControlGetText($window, "", "TsTextWinClass1")) <> 0 Then
	   ControlSetText($window, "", "TsTextWinClass1", "Yes", 1)
    EndIf
	Sleep(100)
	ControlSend($window, "", "", "{TAB}")
	Sleep(400)
	Do
		Sleep(100)
		WinActivate("tdsls4820s000")
	Until StringInStr(ControlGetFocus("tdsls4820s000"), "TsTextWinClass")
EndFunc
Func set_ESSONwindowToZero()
	Local $window = "tdsls4820s000"
	Do
		If Not StringCompare(ControlGetFocus($window), "TsTextWinClass1") == 0 Then ControlClick($window, "", "TsTextWinClass1")
		ControlSend($window, "", "TsTextWinClass1", "0", 1)
		Sleep(400)
		ControlSend($window, "", "", "{TAB}")
		Sleep(350)
		ControlSend($window, "", "", "{SHIFTDOWN}{TAB}{SHIFTUP}")
		Sleep(350)
		Sleep(100)
	Until StringCompare(ControlGetFocus($window), "TsTextWinClass1") == 0 And StringCompare(ControlGetText($window, "", "TsTextWinClass1"), "0") == 0
EndFunc
Func fill_ESSONwindow(ByRef $tenOrderArray)
	Local $window = "tdsls4820s000", $control = "TsTextWinClass"
	WinActivate($window)
	ControlClick($window, "", $control & 1, "left", 2)
	For $i = 0 To UBound($tenOrderArray, $UBOUND_ROWS) - 1
		Do
			ControlSetText($window, "", $control & String($i + 1), $tenOrderArray[$i][0], 1)
		Until StringCompare($tenOrderArray[$i][0], ControlGetText($window, "", $control & String($i + 1))) == 0
		Sleep(100)
		ControlSend($window, "", "", "{TAB}")
		While StringCompare(ControlGetFocus($window), $control & String($i + 1)) == 0
			Sleep(20)
		WEnd
				Do
				If WinExists("tdslss0004") Then
					ControlClick("tdslss0004", "", "TsPushButtonWinClass1")
					ControlClick("tdsls4820s000", "", "TsPushButtonWinClass2")
					_Excel_RangeWrite($oClosing, $oWorksheet, "not valid", "S" & $tenOrderArray[$i][1] + 2)
					MsgBox(0, "Error", "The order " & $tenOrderArray[$i] & " is not found. Highlighting spreadsheet and removing from print queue.", 20)
					WinActivate($window)
					_ArrayDelete($tenOrderArray, $i)
					Return 0
				EndIf
				WinActivate($window)
			Until ((StringCompare(ControlGetFocus($window), $control & $i + 2) == 0) Or (StringCompare(ControlGetFocus($window), "TsPushButtonWinClass1") == 0))
	Next
	ControlClick($window, "", "TsPushButtonWinClass1")
	Do
		Sleep(100)
	Until WinActive("tdsls4401m000")
	Return True
EndFunc
Func changeCompany()
	Local $browserWin = WinGetTitle("Menu browser"), $company = $mfgCompany
	If StringInStr($browserWin, $mfgCompany) Then $company = $salesCompany
	Opt("SendKeyDelay", 50)
	WinActivate($browserWin)
	ControlSend($browserWin, "", "", "{ALT}{o}{c}")
	Do
		WinActivate("ttdsk2003m000")
		Sleep(100)
	Until ControlGetFocus("ttdsk2003m000") <> ""
	Sleep(500)
	ControlSend("ttdsk2003m000", "", "TsTextWinClass1", $company, 1)
	Sleep(500)
	ControlClick("ttdsk2003m000", "", "TsPushButtonWinClass1")
	Opt("SendKeyDelay", 10)
EndFunc

;this should be redone using controlsettext and redrawing the window every time
Func expandMenu($poa = False, $special = False, $RPO = False, $notes = False, $maintain = False)
	openBaan()
	Local $browserWin = WinGetTitle("Menu browser"), $delay = 150, $windowOfInterest
	WinActivate($browserWin)
	WinMove($browserWin, "", 0, 0)
	If Not $maintain And StringInStr($browserWin, $mfgCompany) Then changeCompany()
	Select
		Case $poa
			If WinExists("tdsls4401m000") Then Return
		Case $special
			If WinExists("tdsls9449m000") Then Return
		Case $RPO
			If WinExists("tdrpl0411m000") Then Return
		Case $notes
			If WinExists("tdsls4102s000") Then Return
		Case $maintain
			If WinExists("tisfc0110m000") Then Return
	EndSelect
	WinMove($browserWin, "", 0, 0)
	Opt("SendKeyDelay", 50)
	WinActivate("Browser menu")
	Local $oldHandle = WinGetHandle("Browser menu")
	Do
		If $poa Then
			$windowOfInterest = "tdsls4401m000"
			searchBrowserGUI("Distribution")
			searchBrowserGUI("Sales Control")
			searchBrowserGUI("Sales Orders")
			searchBrowserGUI("Procedure")
			searchBrowserGUI("4401", 1)
		EndIf
		If $special Then
			$windowOfInterest = "tdsls9449m000"
			searchBrowserGUI("Distribution")
			searchBrowserGUI("Sales Control")
			searchBrowserGUI("Sales Orders")
			searchBrowserGUI("Procedure")
			searchBrowserGUI("Reports")
			searchBrowserGUI("9449", 1)
		EndIf
		If $RPO Then
			$windowOfInterest = "tdrpl0411m000"
			searchBrowserGUI("Distribution")
			searchBrowserGUI("Replenishment Order Control")
			searchBrowserGUI("Replenishment Orders")
			searchBrowserGUI("Procedure")
			searchBrowserGUI("0411", 1)
		EndIf
		If $notes Then
			$windowOfInterest = "tdsls4101m000"
			searchBrowserGUI("Distribution")
			searchBrowserGUI("Sales Control")
			searchBrowserGUI("Sales Orders")
			searchBrowserGUI("Procedure")
			searchBrowserGUI("Maintain Sales Orders")
			WinWait("tdsls4101m000")
			ControlSend("tdsls4101m000", "", "", "{CTRLDOWN}{y}{CTRLUP}")
		EndIf
		If $maintain Then
			If StringInStr(WinGetTitle($browserWin), $salesCompany) Then changeCompany()
			$windowOfInterest = "tisfc0110m000"
			searchBrowserGUI("Manufacturing")
			searchBrowserGUI("Shop Floor Control")
			searchBrowserGUI("Production Order Control")
			searchBrowserGUI("Procedure")
			searchBrowserGUI("Maintain Estimated Materials")
		EndIf
		Local $hHandle = WinWaitActive($windowOfInterest, "", 5) ;if not found, returns 0
		If $hHandle == 0 Then
			If (WinGetHandle("[ACTIVE]") <> $oldHandle) Then WinClose("[ACTIVE]") ;if there is a new window, it will be active. If the new, active window is not one we wanted, close it.
		EndIf
		WinActivate($windowOfInterest)
	Until ControlGetFocus($windowOfInterest) <> ""
	Opt("SendKeyDelay", 10)
EndFunc
Func searchBrowserGUI($searchTerm = "General", $type = 0)
	Local $byDescrOrCode[2] = ["{ALT}{i}{d}","{ALT}{i}{b}"]
	While Not WinActive("Menu browser")
		WinActivate("Menu browser")
	WEnd
	Opt("SendKeyDelay", 50)
	ControlSend("Menu browser", "", "", $byDescrOrCode[$type] )
	Opt("SendKeyDelay", 10)
	WinWaitActive("ttdsk1300m000")
	ControlSetText("ttdsk1300m000", "", "TsTextWinClass1", $searchTerm, 1)
	ControlClick("ttdsk1300m000", "", "OK")
	WinActivate("Menu browser")
	WinWaitActive("Menu browser")
	ControlSend("Menu browser", "", "", "{ENTER}")
EndFunc

Func printWithAdditionalDetail($orderArray, $index)
	Local $browserWin = WinGetTitle("Menu browser"), $additionalDetailsWin = "tdsls9449m000", $selectWin = "ttstpsplopen"
	If StringInStr($browserWin, $mfgCompany) Then changeCompany()
	If Not WinExists($additionalDetailsWin) Then expandMenu(0, 1)
	WinActivate($additionalDetailsWin)
	Do
		ControlSend($additionalDetailsWin, "", "", "{TAB}")
		Sleep(250)
	Until StringCompare(ControlGetFocus($additionalDetailsWin), "TsTextWinClass3") == 0
	ControlSend($additionalDetailsWin, "", "TsTextWinClass3", $orderArray[$index][2], 1)
	Sleep(750)
	ControlSend($additionalDetailsWin, "", "TsTextWinClass3", "{TAB}")
	ControlClick($additionalDetailsWin, "", "TsPushButtonWinClass1")
	Do
		Sleep(200)
		WinActivate($selectWin)
		Local $deviceID = ControlGetFocus($selectWin)
		ControlClick("ttstpsplopen", "", "TsTextWinClass11")
	Until StringInStr($deviceID, "TsTextWinClass")
	Do
		Sleep(500)
		If StringCompare(ControlGetText($selectWin, "", $deviceID), "DOWNLOAD") == 0 Then
			ControlSend($selectWin, "", $deviceID, "{TAB}")
			ExitLoop
		EndIf
		ControlClick($selectWin, "", $deviceID)
		ControlSend($selectWin, "", $deviceID, "DOWNLOAD", 1)
		ControlSend($selectWin, "", "", "{TAB}")
		Sleep(500)
	Until StringCompare(ControlGetText($selectWin, "", $deviceID), "DOWNLOAD") == 0
	Do
		Sleep(250)
		Local $outputFileID = ControlGetFocus($selectWin)
	Until StringCompare($outputFileID, "TsTextWinClass9") == 0
	Do
		ControlClick($selectWin, "", $outputFileID, Default, 2)
		Sleep(1000)
		ControlSend($selectWin, "", $outputFileID, "C:\printFolder\order_" & $orderArray[$index][2] & "_" & $orderArray[$index][3] & ".txt", 1)
		Sleep(1000)
		ControlSend($selectWin, "", $outputFileID, "{TAB}")
		Sleep(1000)
	Until StringCompare(ControlGetText($selectWin, "", $outputFileID), "C:\printFolder\order_" & $orderArray[$index][2] & "_" & $orderArray[$index][3] & ".txt") == 0
	Do
		Sleep(250)
		Local $controlID = ControlGetFocus($selectWin)
	Until StringCompare($controlID, $outputFileID) <> 0
	ControlClick($selectWin, "", "Continue")
	Do
		Sleep(300)
		Local $successWindow = WinGetTitle("[ACTIVE]", "OK")
	Until $successWindow == ""
	ControlClick("[ACTIVE]", "OK", "TsPushButtonWinClass1")
	_Excel_RangeWrite($oClosing, $oWorksheet, "printed", "S" & $index + 2)
	_Excel_BookSave($oClosing)
EndFunc

Func printRPOtoTextFile($orderNumber, $row)
	Local $orderNumberID = "TsTextWinClass2", $deviceWindow = "ttstpsplopen", $firstTime = True, $RPOwindow = "tdrpl0411m000 : Print Replenishment"
	ConsoleWrite("Replenish..." & @CRLF)
	dontSleep()
	If Not WinExists($RPOwindow) Then expandMenu(0, 0, 1)
	WinMove($RPOwindow, "", 0, 0)
	WinActivate($RPOwindow)
	ControlClick($RPOwindow, "", $orderNumberID)
	Do
		Sleep(50)
	Until StringCompare(ControlGetFocus($RPOwindow), $orderNumberID) == 0
	Sleep(400)
	Do
		putText($RPOwindow, $orderNumberID, $orderNumber)
	Until StringCompare(ControlGetText($RPOwindow, "", $orderNumberID), $orderNumber) == 0
	Sleep(250)
	If $firstTime Then
		$firstTime = False
		ControlClick($RPOwindow, "", "TsArrowButtonWinClass2")
		Do
			Sleep(50)
		Until StringCompare(ControlGetFocus($RPOwindow), "TsListBoxWinClass1") == 0
		ControlSend($RPOwindow, "", "TsListBoxWinClass1", "a")
		Sleep(400)
		ControlSend($RPOwindow, "", "TsListBoxWinClass1", "{ENTER}")
		Do
			Sleep(50)
		Until StringCompare(ControlGetFocus($RPOwindow), "TsTextWinClass12") == 0
	EndIf
	ControlClick($RPOwindow, "", "TsPushButtonWinClass1")
	WinWait($deviceWindow)
	fillRPOdeviceWindow($orderNumber, $row)
	Sleep(1000)
	Local $sleepCount = 0, $hHandle
	Do
		Sleep(1000)
		$hHandle = WinGetTitle("[ACTIVE]")
		$sleepCount += 1
	Until $hHandle == "" Or $sleepCount == 240 ;this might take a long time to print! I have given four minutes
	If Not $hHandle == "" Then
		MsgBox(0, "Timeout", "The print confirmation dialog never opened. You should investigate why. Order number " & $orderNumber& ".")
		Return False
	Else
		ControlClick($hHandle, "", "TsPushButtonWinClass1")
		Return True
	EndIf
EndFunc

Func fillRPOdeviceWindow($orderNumber, $row)
	ToolTip("fill device window...", 750, 0)
	Local $deviceWindow = "ttstpsplopen", $firstTime = True, $outputFileID
	WinActivate($deviceWindow)
	Local $deviceID = "TsTextWinClass11"
	WinWaitActive($deviceWindow)
	ControlClick($deviceWindow, "", $deviceID)
	Do
		Sleep(50)
	Until StringCompare(ControlGetFocus($deviceWindow), $deviceID) == 0
	Do
		If Not $firstTime Then delete($deviceWindow, $deviceID)
		putText($deviceWindow, $deviceID, "DOWNLOAD")
		Sleep(250)
		$outputFileID = ControlGetFocus($deviceWindow)
		$firstTime = False
	Until StringCompare(ControlGetText($deviceWindow, "", $deviceID), "DOWNLOAD") == 0
	Local $filenm = ("C:\printFolder\order_" & $orderNumber & "_" & $row & ".txt")
	If FileExists($filenm) Then FileDelete($filenm)
	Do
		delete($deviceWindow, $outputFileID)
		putText($deviceWindow, $outputFileID, $filenm)
	Until StringCompare(ControlGetText($deviceWindow, "", $outputFileID), $filenm) == 0
	ControlClick($deviceWindow, "", "TsPushButtonWinClass2")
	Do
		Sleep(25)
	Until Not WinExists($deviceWindow)
	ConsoleWrite("Finished with fillDeviceWindow()" & @CRLF)
EndFunc
Func openBaan()
	Local $list = "", $found = False, $loginWin = ""
	Opt("WinTitleMatchMode", 2)
	If Not WinExists("Menu browser") Then
		ShellExecute($baanPath)
		If WinWait("Menu browser", "", 30) == 0 Then Exit MsgBox(0, "Error", "Baan never opened! Terminating.")
	EndIf
	Opt("WinTitleMatchMode", 1)
	ConsoleWrite("Exiting openBaan function." & @CRLF)
EndFunc
Func dontSleep()
    Local $CurPos = MouseGetPos( )
    MouseMove ( $CurPos[0] + 25, $CurPos[1] )
    MouseMove ( $CurPos[0] - 25, $CurPos[1] )
EndFunc