#pragma compile(ProductVersion, 0.1)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Скрипт для сбора статистики по печати для всех филиалов)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555 - nn-admin@nnkk.budzdorov.su)
#pragma compile(ProductName, create_print_statistics_for_all)

#include <File.au3>
#include <FileConstants.au3>
#include <Excel.au3>
#include <Constants.au3>
#include <Excel.au3>
#include <String.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <ListViewConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <StaticConstants.au3>
#include <ColorConstants.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include <GuiScrollBars.au3>
#include <FontConstants.au3>

#Region ==========================    Variables    ==========================
Local $isGui = True
If $CmdLine[0] > 0 Then
   If $CmdLine[1] = "silent" Then $isGui = False
EndIf

Local $toLogView = 0
Local $currentLabel = 0
Local $currentProgress = 0
Local $currentNameLabel = 0
Local $totalLabel = 0
Local $totalProgress = 0

Local $error = False
Local $oMyError = ObjEvent("AutoIt.Error","HandleComError")
Local $generalSection = "general"
Local $current_pc_name = @ComputerName
Local $errStr = "===ERROR=== "
Local $messageToSend = ""
Local $iniFile = @ScriptDir & "\create_print_statistics_for_all.ini"
Local $logFilePath = @ScriptDir & "\create_print_statistics_for_all.log"
Local $logFile = FileOpen($logFilePath, $FO_OVERWRITE)

Local $mailSection = "mail"
Local $server_backup = ""
Local $login_backup = ""
Local $password_backup = ""
Local $to_backup = ""
Local $send_email_backup = "1"

Local $server = IniRead($iniFile, $mailSection, "server", $server_backup)
Local $login = IniRead($iniFile, $mailSection, "login", $login_backup)
Local $password = IniRead($iniFile, $mailSection, "password", $password_backup)
Local $to = IniRead($iniFile, $mailSection, "to", $to_backup)
Local $send_email = IniRead($iniFile, $mailSection, "send_email", $send_email_backup)

If Not FileExists($iniFile) Then
   ToLog($errStr & "Cannot find the settings file: " & $iniFile)
   SendEmail()
EndIf

Local $paths_to_logs = IniReadSection($iniFile, $generalSection)
If @error Then
   ToLog("Cannot read the section: '" & $generalSection & "' in the ini file: " & $iniFile)
   SendEmail()
EndIf

If UBound($paths_to_logs, $UBOUND_ROWS) > 0 Then _ArrayDelete($paths_to_logs, 0)

Local $storagePath = IniRead($iniFile, "storage", "path", "")
If Not FileExists($storagePath) Then $storagePath = @ScriptDir & "\"
ToLog("Reports storage: " & $storagePath & @CRLF)

Local $cur_year = @YEAR
Local $cur_month = @MON
If $cur_month == "01" Then
   $cur_month = "12"
   $cur_year -= 1
Else
   $cur_month -= 1
   If $cur_month < 10 Then
	  $cur_month = "0" & $cur_month
   EndIf
EndIf

Local $currentPeriod = $cur_year & "-" & $cur_month

If Not $isGui Then
   SilentModeCreateReports($paths_to_logs, $currentPeriod, $storagePath)
   SendEmail()
EndIf
#EndRegion

#Region ==========================    GUI    ==========================
Local $mainGui = GUICreate("Create print statistics", 400, 600)
GUISetFont(10)

Local $periodLabel = GUICtrlCreateLabel("Time period:", 10, 12, 75, 20)
Local $periodInput = GUICtrlCreateInput($currentPeriod, 115, 10, 60, 20, BitOR($ES_READONLY, $ES_CENTER))
GUICtrlSetFont(-1, 9, $FW_BOLD, Default, "Courier New", Default)
GUICtrlSetBkColor(-1, $COLOR_WHITE)

Local $prevMonth = GUICtrlCreateButton("<<", 85, 10, 25, 20)
Local $nextMonth = GUICtrlCreateButton(">>", 180, 10, 25, 20)

Local $saveToLabel = GUICtrlCreateLabel("Save to:", 10, 40, 50, 20, $SS_LEFTNOWORDWRAP)
Local $saveLabel = GUICtrlCreateEdit("", 62, 39, 259, 20, $ES_READONLY)
GUICtrlSetFont(-1, 9, $FW_BOLD, Default, "Courier New", Default)
GUICtrlSetBkColor(-1, $COLOR_WHITE)
Local $changeButton = GUICtrlCreateButton("Change", 330, 33, 60, 30)
SetDataForPathLabel()

Local $listView = GUICtrlCreateListView("Name|Path", 10, 70, 380, 470, BitOr($LVS_SHOWSELALWAYS, $LVS_REPORT))
GUICtrlSetState(-1, $GUI_FOCUS)
_GUICtrlListView_SetColumnWidth($listView, 0, 80)
_GUICtrlListView_SetColumnWidth($listView, 1, 295)
_GUICtrlListView_AddArray($listView, $paths_to_logs)
_GUICtrlListView_SetSelectionMark(-1, -1)

Local $helpLabel = GUICtrlCreateLabel("Press and hold the ctrl key to select several lines", 10, 540, 380, 15, $SS_CENTER)
GUICtrlSetFont(-1, 8)
GUICtrlSetColor(-1, $COLOR_GRAY)

Local $selectAllButton = GUICtrlCreateButton("Select all", 10, 560, 120, 30)
Local $runReport = GUICtrlCreateButton("Create report(s)", 270, 560, 120, 30)
GUICtrlSetState(-1, $GUI_DISABLE)

GUISetState(@SW_SHOW)
GUIRegisterMsg($WM_NOTIFY, "MY_WM_NOTIFY")

While 1
   Switch GUIGetMsg()
	  Case $GUI_EVENT_CLOSE
		 Exit
	  Case $prevMonth
		 AddValueToMonth(-1)
	  Case $nextMonth
		 AddValueToMonth(1)
	  Case $changeButton
		 Local $selected = FileSelectFolder("Path to save reports", $storagePath)
		 If $selected Then
			$storagePath = $selected
			SetDataForPathLabel()
		 EndIf
	  Case $selectAllButton
		 SelectAll()
	  Case $runReport
		 CreateReports()
   EndSwitch
   Sleep(20)
WEnd
#EndRegion

#Region ==========================    Functions    ==========================
Func MY_WM_NOTIFY($hWnd, $Msg, $wParam, $lParam)
   If  $wParam <> $listView Then Return

   #forceref $hWndGUI, $MsgID, $wParam
   Local $tagNMHDR, $event, $hwndFrom, $code
   $tagNMHDR = DllStructCreate("int;int;int", $lParam);NMHDR (hwndFrom, idFrom, code)
   If @error Then Return
   $event = DllStructGetData($tagNMHDR, 3)

   If  $event = $NM_CLICK Or $event = -12 Then
	  CheckSelected()
   EndIf

   $tagNMHDR = 0
   $event = 0
   $lParam = 0
EndFunc

Func CreateReports()
   Local $selectedItems[0][2]
   For $i = 0 To _GUICtrlListView_GetItemCount($listView)
	  If _GUICtrlListView_GetItemSelected($listView, $i) Then
		 Local $item = StringSplit(_GUICtrlListView_GetItemTextString($listView, $i), "|", $STR_NOCOUNT)
		 _ArrayTranspose($item)
		 _ArrayAdd($selectedItems, $item)
	  EndIf
   Next

   Local $prevPosX = WinGetPos($mainGui)[0]
   Local $prevPosY = WinGetPos($mainGui)[1]

   GUISetState(@SW_HIDE)
   Local $logGui = GUICreate("Creating report(s)", 400, 622, $prevPosX, $prevPosY, $WS_OVERLAPPED, $WS_EX_DLGMODALFRAME, $mainGui)
   GUISwitch($logGui)
   GUISetFont(10)

   Local $logView = GUICtrlCreateEdit("", 10, 10, 380, 440, BitOR($ES_AUTOVSCROLL, $ES_READONLY, $WS_VSCROLL))
   GUICtrlSetBkColor(-1, $COLOR_WHITE)
   GUICtrlSetFont(-1, 8)
   _GUICtrlEdit_SetLimitText($logView, 999999)

   $currentLabel = GUICtrlCreateLabel("Current: ", 10, 460, 50, 20)
   $currentNameLabel = GUICtrlCreateLabel("", 60, 460, 330, 20)
   $currentProgress = GUICtrlCreateProgress(10, 480, 380, 20)
   $totalLabel = GUICtrlCreateLabel("Total: ", 10, 510, 60, 20)
   $totalProgress = GUICtrlCreateProgress(10, 530, 380,20)

   Local $closeButton = GUICtrlCreateButton("Close", 270, 560, 120, 30)
   GUICtrlSetState(-1, $GUI_DISABLE)
   Local $openButton = GUICtrlCreateButton("Open folder", 10, 560, 120, 30)
   GUICtrlSetState(-1, $GUI_DISABLE)

   GUISetState()

   $toLogView = $logView
   SilentModeCreateReports($selectedItems, $currentPeriod, $storagePath)

   MsgBox($MB_ICONINFORMATION, "Creating report(s)", "All is done")

   GUICtrlSetState($closeButton, $GUI_ENABLE)
   GUICtrlSetState($openButton, $GUI_ENABLE)

   While 1
	  Local $msg = GUIGetMsg()
	  If $msg = $GUI_EVENT_CLOSE Or $msg = $closeButton Then
		 $prevPosX = WinGetPos($logGui)[0]
		 $prevPosY = WinGetPos($logGui)[1]
		 GUIDelete($logGui)
		 GUISwitch($mainGui)
		 SelectAll(False)
		 GUICtrlSetState($runReport, $GUI_DISABLE)
		 WinMove($mainGui, Default, $prevPosX, $prevPosY)
		 GUISetState(@SW_SHOW)
		 ExitLoop
	  ElseIf $msg = $openButton Then
		 Run("explorer.exe " & $storagePath & $currentPeriod)
	  EndIf
	  Sleep(20)
   WEnd
EndFunc

Func SilentModeCreateReports($paths, $mask, $saveTo)
   For $i = 0 To Ubound($paths, $UBOUND_ROWS) - 1
	  If $paths[$i][1] = "" Then
		 ToLog("--- There are no logs path for : " & $paths[$i][0])
		 ContinueLoop
	  EndIf

	  ParseLogs($paths[$i][0], $paths[$i][1], $mask, $saveTo)
	  If $isGui And $totalProgress Then GUICtrlSetData($totalProgress, (($i + 1) / Ubound($paths, $UBOUND_ROWS)) * 100)
	  If $isGui And $currentNameLabel Then GUICtrlSetData($currentNameLabel, "")
	  If $isGui And $totalProgress Then GUICtrlSetData($currentProgress, 100)
   Next
EndFunc

Func SelectAll($set = True)
   _GUICtrlListView_BeginUpdate($listView)

   For $i = 0 To _GUICtrlListView_GetItemCount($listView)
	  _GUICtrlListView_SetItemSelected($listView, $i, $set, $set)
   Next

   GUICtrlSetState($listView, $GUI_FOCUS)
   _GUICtrlListView_EndUpdate($listView)
   GUICtrlSetState($runReport, $GUI_ENABLE)
EndFunc

Func CheckSelected()
   Local $ret = False
   For $i = 0 To _GUICtrlListView_GetItemCount($listView)
	  If _GUICtrlListView_GetItemSelected($listView, $i) Then
		 $ret = True
		 GUICtrlSetState($runReport, $GUI_ENABLE)
		 ExitLoop
	  EndIf
   Next
   If Not $ret Then GUICtrlSetState($runReport, $GUI_DISABLE)
EndFunc

Func SetDataForPathLabel()
   If StringRight($storagePath, 1) <> "\" Then
	  $storagePath &= "\"
   EndIf
   Local $toLabel = $storagePath
   If StringLen($storagePath) > 35 Then $toLabel = StringLeft($toLabel, 32) & "..."
   GUICtrlSetTip($saveLabel, $storagePath)
   GUICtrlSetData($saveLabel, $toLabel)
EndFunc

Func AddValueToMonth($val)
   Local $year = Int(StringLeft($currentPeriod, 4))
   Local $month = Int(StringRight($currentPeriod, 2))
   $month += $val
   If $month = 0 Then
	  $month = 12
	  $year -= 1
   ElseIf $month = 13 Then
	  $month = 1
	  $year += 1
   EndIf
   $currentPeriod = $year & "-" & (($month < 10) ? "0" & $month : $month)
   GUICtrlSetData($periodInput, $currentPeriod)
EndFunc

Func ParseLogs($name, $path, $maskToSearch, $saveTo)
   If $isGui And $currentNameLabel Then GUICtrlSetData($currentNameLabel, $name)
   If $isGui And $totalProgress Then GUICtrlSetData($currentProgress, 0)

   If Not FileExists($saveTo & $maskToSearch & "\") Then
	  If Not DirCreate($saveTo & $maskToSearch & "\") Then
		 ToLog($errStr & "Cannot create the directory: " & $saveTo & $maskToSearch & "\")
		 Return
	  EndIf
   EndIf

   ToLog("--- Parsing logs for: '" & $name & "' in path: " & $path)

   Local $grandTotal = 0
   Local $users[0][2]
   Local $computers[0][4]
   Local $printers[0][2]

   If Not FileExists($path) Then
	  ToLog($errStr & "Cannot open the directory: " & $path)
	  Return
   EndIf

   Local $mask = "papercut-print-log-*.csv"
   $mask = StringReplace($mask, "*", $maskToSearch)
   If $isGui And $currentNameLabel Then GUICtrlSetData($currentNameLabel, $name & ", searching for files")
   Local $result = _FileListToArrayRec($path, $mask, $FLTAR_FILES, $FLTAR_RECUR, Default, $FLTAR_FULLPATH)

   If Not IsArray($result) Then
	  ToLog("    No files matching the mask: " & $mask)
	  Return
   EndIf

   Local $excel = _Excel_Open(False, False, False, False, True)
   If $excel = 0 Then
	  ToLog($errStr & "Cannot open the Excel application")
	  Return
   EndIf

   Local $workbook = _Excel_BookNew($excel, 1)
   If $workbook = 0 Then
	  ToLog($errStr & "Cannot create the new workbook")
	  Return
   EndIf
   _Excel_SheetAdd($workbook, 1, Default, Default, "Details")
   _Excel_SheetDelete($workbook, 2)

   Local $counter = 1
   Local $sendToExcel = 0
   Local $excelFact = 0

   For $i = 1 To UBound($result) - 1
	  Local $currentComp = StringReplace($result[$i], $path, "")
	  $currentComp = StringSplit($currentComp, "\", $STR_NOCOUNT)[0]

	  ToLog("    " & StringReplace($result[$i], $path, "..\"), False)

	  Local $currentCompTotal = 0
	  Local $printersCount[0][2]
	  Local $file = FileOpen($result[$i], $FO_READ & $FO_FULLFILE_DETECT)
	  If $file = -1 Then
		 ToLog($errStr & "Cannot open the file: " & $result[$i])
		 ContinueLoop
	  EndIf

	  Local $fileContent = FileReadToArray($file)
	  If @error = 2 Then
		 ToLog($errStr & "The file is empty: " & $result[$i])
	  ElseIf @error = 1 Then
		 ToLog($errStr & "2Cannot open the file: " & $result[$i])
	  EndIf

	  Local $detailedData[0][16]
	  If $counter = 1 Then
		 Local $title = ["Time", "User", "Pages", "Copies", "Total", "Printer", "Document Name", "Client", "Paper Size", _
			"Language", "Height", "Width", "Duplex", "Grayscal", "Size", ""]
		 _ArrayTranspose($title)
		 _ArrayAdd($detailedData, $title)
	  EndIf

	  For $x in $fileContent
		 Local $originalString = $x
		 If $x = "" Or StringInStr($x, "PaperCut") Or StringInStr($x, "Time,User") Then ContinueLoop
		 If StringLeft($x, 1) = '"' Then $x = StringRight($x, StringLen($x) - 1)
		 If StringRight($x, 10) = '";;;;;;;;;' Then $x = StringLeft($x, StringLen($x) - 10)
		 If StringInStr($x, '"') Then
			Local $firstQuote = StringInStr($x, '"')
			Local $lastQuote = StringInStr($x, '"', Default, -1)
			Local $docName = StringMid($x, $firstQuote, $lastQuote - $firstQuote + 1)
			Local $docNameRight = StringRegExpReplace($docName, "[!@#$%^&*=,]", "")
			$x = StringReplace($x, $docName, $docNameRight)
		 EndIf

		 $x = StringSplit($x, ",", $STR_NOCOUNT)
		 _ArrayTranspose($x)

		 If UBound($x, $UBOUND_COLUMNS) <> 15 Then
			ToLog($errStr & "Columns quantity doesn't equal 15: " & _ArrayToString($x) & @CRLF)
			ContinueLoop
		 EndIf

		 If Not StringIsInt($x[0][2]) OR Not StringIsInt($x[0][3]) Then
			ToLog($errStr & "2nd or 3rd column not an int: " & _ArrayToString($x) & @CRLF)
			ContinueLoop
		 EndIf

		 Local $total = $x[0][2] * $x[0][3]
		 _ArrayColInsert($x, 4)
		 $x[0][4] = $total

		 $grandTotal += $total
		 $currentCompTotal += $total

		 Local $tmp = _ArraySearch($printersCount, $x[0][5])
		 If $tmp >= 0 Then
			$printersCount[$tmp][1] += $total
		 Else
			_ArrayAdd($printersCount, $x[0][5] & "," & $total, Default, ",")
		 EndIf

		 Local $tmp = _ArraySearch($users, $x[0][1])
		 If $tmp >= 0 Then
			$users[$tmp][1] += $total
		 Else
			_ArrayAdd($users, $x[0][1] & "," & $total, Default, ",")
		 EndIf

		 Local $tmp = _ArraySearch($printers, $x[0][5])
		 If $tmp >= 0 Then
			$printers[$tmp][1] += $total
		 Else
			_ArrayAdd($printers, $x[0][5] & "," & $total, Default, ",")
		 EndIf

		 If _ArrayAdd($detailedData, $x) = -1 Then
			ToLog($errStr & "Cannot add line to array: " & @error & " - " & _ArrayToString($x) & @CRLF)
		 EndIf

		 If UBound($detailedData, $UBOUND_ROWS) > 500 Or $originalString = $fileContent[UBound($fileContent) - 1] Then
			Local $toRange = "a" & $counter & ":p" & $counter + UBound($detailedData, $UBOUND_ROWS) - 1
			Local $range = _Excel_RangeWrite($workbook, Default, $detailedData, $toRange)
			If @error Then ToLog($errStr & "Written range doesn't equal initial data: " & @error & @CRLF)
			If $range.Rows.Count - UBound($detailedData, $UBOUND_ROWS) <> 0 Then ConsoleWrite($errStr & @CRLF)
			$counter += UBound($detailedData, $UBOUND_ROWS)
			_ArrayDelete($detailedData, "0-" & UBound($detailedData, $UBOUND_ROWS) - 1)
			If $isGui And $currentNameLabel Then GUICtrlSetData($currentNameLabel, $name & ", processed lines: " & $counter)
		 EndIf
	  Next

	  For $x = 0 To UBound($printersCount, $UBOUND_ROWS) - 1
		 _ArrayAdd($computers, $currentComp & "," & $currentCompTotal & "," & _
			   $printersCount[$x][0] & "," & $printersCount[$x][1], Default, ",")
	  Next

	  If $isGui And $totalProgress Then GUICtrlSetData($currentProgress, (($i + 1) / Ubound($result, $UBOUND_ROWS)) * 90)
   Next

   _Excel_RangeSort($workbook, Default, Default, "E1", $xlDescending, Default, $xlYes)

   $excel.ActiveSheet.UsedRange.Columns.AutoFit
   $excel.ActiveSheet.Range("A1").AutoFilter
   $excel.ActiveSheet.Columns("G").ColumnWidth = 60

   _Excel_SheetAdd($workbook, Default, False, Default, "Report")

   _Excel_RangeWrite($workbook, Default, $name & " print report for the period: " & $maskToSearch & ". Total pages printed: " & $grandTotal, _
	  "a1", True, False)
   $excel.ActiveSheet.Range("A1", "J1").Merge
   $excel.ActiveSheet.Cells(1.1).HorizontalAlignment = -4108
   $excel.ActiveSheet.Range("A1").Font.Bold = True

   Local $currentRange = "a" & 3 & ":b" & 3 + UBound($users, $UBOUND_ROWS) - 1
   _Excel_RangeWrite($workbook, Default, $users, $currentRange, True, False)
   $excel.ActiveSheet.Range($currentRange).Borders.LineStyle = 2
   _Excel_RangeSort($workbook, Default, $currentRange, "b3", $xlDescending)

   $currentRange = "d" & 3 & ":g" & 3 + UBound($computers, $UBOUND_ROWS) - 1
   _Excel_RangeWrite($workbook, Default, $computers, $currentRange, True, False)
   $excel.ActiveSheet.Range($currentRange).Borders.LineStyle = 2
   _Excel_RangeSort($workbook, Default, $currentRange, "e3", $xlDescending)

   For $i = 3 To 3 + UBound($computers, $UBOUND_ROWS) - 2
	  If $excel.ActiveSheet.Range("d" & $i).Value = $excel.ActiveSheet.Range("d" & $i+1).Value Then
		 Local $firstMerge = $i
		 Local $lastMerge = $i+1

		 For $x = $i+2 To $i+2 + UBound($computers, $UBOUND_ROWS) - 2
			If $excel.ActiveSheet.Range("d" & $i).Value = $excel.ActiveSheet.Range("d" & $x).Value Then
			   $lastMerge = $x
			Else
			   $i = $x
			   ExitLoop
			EndIf
		 Next

		 $excel.ActiveSheet.Range("d" & $firstMerge & ":d" & $lastMerge).Merge
		 $excel.ActiveSheet.Range("d" & $firstMerge & ":d" & $lastMerge).VerticalAlignment = $xlCenter
		 $excel.ActiveSheet.Range("e" & $firstMerge & ":e" & $lastMerge).Merge
		 $excel.ActiveSheet.Range("e" & $firstMerge & ":e" & $lastMerge).VerticalAlignment = $xlCenter
	  EndIf
   Next

   $currentRange = "i" & 3 & ":j" & 3 + UBound($printers, $UBOUND_ROWS) - 1
   _Excel_RangeWrite($workbook, Default, $printers, $currentRange, True, False)
   $excel.ActiveSheet.Range($currentRange).Borders.LineStyle = 2
   _Excel_RangeSort($workbook, Default, $currentRange, "j3", $xlDescending)

   $excel.ActiveSheet.UsedRange.Columns.AutoFit

   ToLog("    Total pages printed: " & $grandTotal)
   Local $pathToSave = $saveTo & $maskToSearch & "\" & $name & "-" & $maskToSearch & ".xlsx"
   If _Excel_BookSaveAs($workbook, $pathToSave, Default, True) Then
	  ToLog("    The workbook successfully saved at: " & $pathToSave)
   Else
	  ToLog($errStr & " failed to save workbook at: " & $pathToSave)
   EndIf

   _Excel_Close($excel)
EndFunc

Func ToLog($message, $toMail = True)
   $message &= @CRLF
   If $toMail Then $messageToSend &= $message
   ConsoleWrite($message)
   _FileWriteLog($logFile, $message)

   If $isGui And $toLogView Then
	  Local $currentData = GUICtrlRead($toLogView)
	  $message = @HOUR & ":" & @MIN & ":" & @SEC & ": " & $message
	  _GUICtrlEdit_AppendText($toLogView, $message)
   EndIf
EndFunc

Func SendEmail()
   If Not $send_email Then
	  FileClose($logFile)
	  Exit
   EndIf

   Local $from = "Create print statistics for all"
   Local $title = "All reports successfully created"
   If StringInStr($messageToSend, $errStr) Then
	  $title = "Error(s) occured while creating the reports"
	  $to = $to_backup = ""
   EndIf

   ToLog(@CRLF & "--- Sending email", False)
   If _INetSmtpMailCom($server, $from, $login, $to, _
		 $title, $messageToSend, "", "", "", $login, $password) <> 0 Then

	  _INetSmtpMailCom($server_backup, $from, $login_backup, $to_backup, _
		 $title, $messageToSend, "", "", "", $login_backup, $password_backup)
   EndIf

   FileClose($logFile)
   Exit
EndFunc

Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, _
   $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", _
   $s_BccAddress = "", $s_Username = "", $s_Password = "",$IPPort=25, $ssl=0)

   Local $objEmail = ObjCreate("CDO.Message")
   Local $i_Error = 0
   Local $i_Error_desciption = ""

   $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
   $objEmail.To = $s_ToAddress

   If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
   If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

   $objEmail.Subject = $s_Subject

   If StringInStr($as_Body,"<") and StringInStr($as_Body,">") Then
	  $objEmail.HTMLBody = $as_Body
   Else
	  $objEmail.Textbody = $as_Body & @CRLF
   EndIf

   If $s_AttachFiles <> "" Then
	  Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
	  For $x = 1 To $S_Files2Attach[0] - 1
		 $S_Files2Attach[$x] = _PathFull ($S_Files2Attach[$x])
		 If FileExists($S_Files2Attach[$x]) Then
			$objEmail.AddAttachment ($S_Files2Attach[$x])
		 Else
			$i_Error_desciption = $i_Error_desciption & @lf & 'File not found to attach: ' & $S_Files2Attach[$x]
			SetError(1)
			return 0
		 EndIf
	  Next
   EndIf

   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

   If $s_Username <> "" Then
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
   EndIf

   If $Ssl Then
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
   EndIf

   $objEmail.Configuration.Fields.Update
   $objEmail.Send

   if @error then
	  SetError(2)
   EndIf

   Return @error
EndFunc

Func HandleComError()
   ToLog($errStr & @ScriptName & " (" & $oMyError.scriptline & ") : ==> COM Error intercepted!" & @CRLF & _
            @TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oMyError.number) & @CRLF & _
            @TAB & "err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
            @TAB & "err.description is: " & @TAB & $oMyError.description & @CRLF & _
            @TAB & "err.source is: " & @TAB & @TAB & $oMyError.source & @CRLF & _
            @TAB & "err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
            @TAB & "err.helpcontext is: " & @TAB & $oMyError.helpcontext & @CRLF & _
            @TAB & "err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
            @TAB & "err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
            @TAB & "err.retcode is: " & @TAB & "0x" & Hex($oMyError.retcode) & @CRLF & @CRLF)
Endfunc
#EndRegion