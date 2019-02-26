#include <Excel.au3>

Global $Aktualy_czas, $Excel_full_name, $Plik_data, $Excel_full_name
Global $Sciezka_excela, $Program_Excel_open, $Plik_Excel_open
Date()
Func Date()
	Global $Dzien = @MDAY, $Miesiac_name, $Miesiac = int(@MON)
	Global $Rok_skr = StringRight(Int(@YEAR), 2), $Rok = Int(@YEAR), $h = Int(@HOUR), $m = Int(@MIN)
	Global $Date[13] = _
			[0, 'Styczeń', 'Luty', 'Marzec', 'Kwiecień', _
			'Maj', 'Czerwiec', 'Lipiec', 'Sierpień', _
			'Wrzesień', 'Październik', 'Listopad', 'Grudzień']
	$Miesiac_name = $Date[$Miesiac]
EndFunc   ;==>Date

If $Miesiac < 10 Then $Miesiac = "0" & $Miesiac

$Aktuala_data = String($Dzien & '.' & $Miesiac & '.' & $Rok)
$Plik_data = '2.' & $Miesiac & '.' & $Rok_skr & ' ' & $Miesiac_name
$Excel_full_name = '\' & $Plik_data & '.xls'
$Sciezka_excela = @ScriptDir & $Excel_full_name

$Program_Excel_open = _Excel_Open()
$Plik_Excel_open = _Excel_BookOpen($Program_Excel_open, $Sciezka_excela)
If $Plik_Excel_open = 0 Then
	MsgBox(0,'Godzinka do Excela', 'Aktualny plik excel nie istnieje' & @CRLF & 'Utwórz plik i spróbuj ponownie',5)
	Exit
	EndIf
WinSetState($Plik_data, '', @SW_MAXIMIZE)

WinWait($Plik_data)
WinActivate($Plik_data)
Sleep(300)
ControlSend($Plik_data, '', 'NetUIHWND2', '^{HOME}')
Sleep(300)
ControlSend($Plik_data, '', 'NetUIHWND2', '+{F5}')
WinWaitActive('Find and Replace')
ControlSend('Find and Replace', '', 18, $Aktuala_data)
Sleep(300)
ControlSend('Find and Replace', '', 18, '{Enter}')
ControlSend('Find and Replace', '', 18, '{Esc}')

;~ $h = 6 ; test
;~ $m = 45 ; test

If $h < 12 Then
	ControlSend($Plik_data, '', 'NetUIHWND2', '{RIGHT}')
Else
	ControlSend($Plik_data, '', 'NetUIHWND2', '{RIGHT 2}')
EndIf

$Aktualy_czas = $h & ':' & $m

If $h < 12 And ($m - 10) < 0 Then
	$ho = ($h - 1)
	$mo = (60 - 10 + $m)
	$Czas_start = $ho & ':' & $mo
Else
$Czas_start = $h & ':' & ($m - 10)
EndIf

If $h >= 12 And $m >= 51 Then
	$ho = ($h + 1)
	$mo = ((10 - (60 - $m)) + 10)
	$Czas_stop = $ho & ':' & $mo
	Else
$Czas_stop = $h & ':' & ($m + 10)
EndIf

Sleep(300)

WinActivate($Plik_data)
WinWaitActive($Plik_data)

If $h < 12 Then
	Send($Czas_start)
Else
	Send($Czas_stop)
EndIf

Sleep(500)
ControlSend($Plik_data, '', 'NetUIHWND2', '{Enter}')
Sleep(800)
ControlSend('Microsoft Excel', '', 2, '{Enter}')
Sleep(300)
ControlSend($Plik_data, '', 'NetUIHWND2', '^{s}')
Sleep(3000)
WinKill($Plik_data)


