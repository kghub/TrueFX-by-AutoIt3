; use your IE, sign in to TrueFX.com, be sure to check "remember me" box
; adjust according to your needs === begin
$iLastYear = 2016
$iLastMonth = 8
$iFirstYear = 2009
$iFirstMonth = 5
$sDnlDir = @WorkingDir & "\zip\" ;enter target sub-directory for zip files, be sure to put "\" at the end
$bDontDownload = False ;dry test if True: click cancel instead of save in "Save As..." dialog
$iNumInstruments = 15 ;adjust when removing or adding pairs below
Local $aInstruments[$iNumInstruments] = ["AUDJPY", "AUDNZD", "AUDUSD", "CADJPY", "CHFJPY", "EURCHF", "EURGBP", "EURJPY", "EURUSD", "GBPJPY", "GBPUSD", "NZDUSD", "USDCAD", "USDCHF", "USDJPY"]
; adjust according to your needs === end

#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <Math.au3>
Opt("SendKeyDownDelay", 50) ;my IE don't get SHIFT prior to F10 otherwise

Local $aMonthNames[12] = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
DirCreate($sDnlDir)

; open IE
_IEErrorNotify(False)
Do
	$oIE = _IECreate("about:blank")
	If Not IsObj($oIE) Then
		Sleep(5000)
	EndIf
Until IsObj($oIE) ; https://autoit.de/index.php/Thread/83663-IE-au3-T3-0-2-Error-from-function-IECreate-Browser-Object-Creation-Failed/
If @error <> 0 Then
	Exit
EndIf
_IEErrorNotify(True)

; open truefx download page
$sWebPage = "http://truefx.com/?page=downloads"
_IENavigate($oIE, $sWebPage)
$hWndBrowser = WinWaitActive("[CLASS:IEFrame]","",10)
If _IEPropertyGet($oIE, "locationurl") <> $sWebPage Then
	ConsoleWrite("Can't open ...page=downloads or got redirected to a different URL. Please check if signed in with remember_me option." & @CRLF)
	Exit
EndIf

$iFirstYear  = _Min($iFirstYear, $iLastYear)
$iFirstYear  = _Max($iFirstYear, 2009) ;earliert files avaialbe are from 2009 when I'm typing this
$iLastYear   = _Max($iFirstYear, $iLastYear)
$iFirstMonth = _Max(1, _Min(12, $iFirstMonth))
$iLastMonth  = _Max(1, _Min(12, $iLastMonth))
If $iFirstYear=$iLastYear Then
	$iFirstMonth = _Min($iFirstMonth, $iLastMonth)
EndIf
ConsoleWrite("Collecting files from " & $iFirstYear & "/" & $iFirstMonth & " till " & $iLastYear & "/" & $iLastMonth & @CRLF)

For $iYear = $iLastYear To $iFirstYear Step -1
	if $iYear < $iLastYear Then
		$iToMonth = 12
	Else
		$iToMonth = $iLastMonth
	EndIf
	if $iYear <= $iFirstYear Then
		$iFromMonth = $iFirstMonth
	Else
		$iFromMonth = 1 ; January
	EndIf
	If $iYear <= 2009 Then
		$iFromMonth = _Max($iFromMonth, 5) ; 2009 is available from May when I'm typing this
	EndIf
	For $iMonth = $iToMonth To $iFromMonth Step -1
		For $iInstrument = 0 To $iNumInstruments-1 Step 1
			$sZipFileName = $aInstruments[$iInstrument] & "-" & $iYear & "-" & StringFormat("%02u", $iMonth)
			If -1=FileFindFirstFile($sDnlDir & $sZipFileName & ".*") Then
				$sZipFileName = $sZipFileName & ".zip"
				if $iYear > 2009 Then
					$sWebPage = "http://truefx.com/?page=download&description=" & StringLower($aMonthNames[$iMonth-1]) & $iYear & "&dir=" & $iYear & "/" & StringUpper($aMonthNames[$iMonth-1]) & "-" & $iYear
					;http://truefx.com/?page=download&description=december2010&dir=2010/DECEMBER-2010
				Else
					$sWebPage = "http://truefx.com/?page=download&description=" & StringLower($aMonthNames[$iMonth-1])          & "&dir=" & $iYear & "/" & StringUpper($aMonthNames[$iMonth-1]) & "-" & $iYear
					; http://truefx.com/?page=download&description=december&dir=2009/DECEMBER-2009
				EndIf
				If _IEPropertyGet($oIE, "locationurl") <> $sWebPage Then
					ConsoleWrite("Navigating to " & $sWebPage & @CRLF)
					_IENavigate($oIE, $sWebPage)
				EndIf
				$hWndBrowser = WinWaitActive("[CLASS:IEFrame]","",10)
				ConsoleWrite("Looking for " & $sZipFileName & @CRLF)
				FindAndSaveAs($hWndBrowser, $sDnlDir, $sZipFileName)
			Else
				ConsoleWrite("File " & $sZipFileName & ".* was found in " & $sDnlDir & " , skipping download." & @CRLF)
			EndIf
		Next
	Next ; Month
Next ; Year
ShellExecute($sDnlDir)
Exit

Func FindAndSaveAs($hWndBrowser, $sDnlDir, $sZipFileName)
	If FileExists($sDnlDir & $sZipFileName) Then
		Return 2
	EndIf
	if Not WinActive($hWndBrowser) Then
		Return 1
	EndIf
	Send("^f") ; find
	Send($sZipFileName)
	Sleep(100)
	Send("{TAB}")
	Send("{TAB}")
	Send("+{TAB}")
	Send("+{F10}") ; like mouse right button
	Sleep(100) ; let context menu open
	Send("{DOWN}")
	Send("{DOWN}")
	Send("{DOWN}")
	Send("{DOWN}")
	Send("{ENTER}") ; save target as...
	Local $hWndSaveAs = WinWaitActive("[CLASS:#32770]","",10) ; used Au3Info.exe to find out the classes
	ControlSetText($hWndSaveAs, "", "Edit1", $sDnlDir & $sZipFileName)
	If $bDontDownload Then
		ControlClick($hWndSaveAs, "", "Button2") ;cancel button
		ConsoleWrite("Dry test mode, clicked Cancel instead of Save for " & $sZipFileName & @CRLF)
		WinWaitActive($hWndBrowser,"",10)
	Else
		ControlClick($hWndSaveAs, "", "Button1") ;save button
		ConsoleWrite("Clicked Save, now waiting...") ;no CRLF
		Do
			ConsoleWrite(".") ;no CRLF
			Sleep(10000)
		Until FileExists($sDnlDir & $sZipFileName) ; wait until download completes
		ConsoleWrite(@CRLF & $sZipFileName & " downloaded" & @CRLF)
	EndIf
	Return 0
EndFunc