#pragma compile(Console, True)
#pragma compile(ExecLevel, RequireAdministrator)
#RequireAdmin
#Include <File.au3>
#include <WinAPIFiles.au3>
#include <FileConstants.au3>

;Получаем список локальных дисков
Local $driveArray = DriveGetDrive("FIXED")
Local $resultArray

If @error Then
    ConsoleWrite("Error while attempting to get drives.")
Else
    ConsoleWrite("Drive List:" & @CRLF)
	For $i = 1 To $driveArray[0]
        ; Показывает все найденные диски и переводит букву в верхний регистр.
		ConsoleWrite( $i & "-" & StringUpper($driveArray[$i]) & @CRLF)
    Next
EndIf

For $i = 1 To $driveArray[0]
	ConsoleWrite("Searching on" & " " & StringUpper($driveArray[$i])& @CRLF)
	$resultArray = _FindFiles($driveArray[$i], "stat.mdb")
	If IsArray($resultArray) Then
		For $n = 1 To $resultArray[0]
			Local $sSourceFilePath = @ScriptDir & "\stat_new.mdb"
			Local $sDestPath = $driveArray[$i] & "\"& $resultArray[$n]

			ConsoleWrite("$sSourceFilePath "&$sSourceFilePath& @CRLF)
			ConsoleWrite("$sDestPath " & $sDestPath & @CRLF)
			ConsoleWrite(FileCopy($sSourceFilePath,  $sDestPath	, $FC_OVERWRITE) & @CRLF)
		Next
	EndIf
Next
MsgBox(4096, "Результат", "Я закончил")

Func _FindFiles($sRoot, $sFile)

    Local $FileList

    $FileList = _FileListToArrayRec($sRoot, $sFile, $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_SORT)
    If Not @error Then
        For $i = 1 To $FileList[0]
            ConsoleWrite($sRoot & '\' & $FileList[$i] & @CR)
        Next
    EndIf
    Return $FileList
;~ 	_ArrayDisplay($FileList)
EndFunc