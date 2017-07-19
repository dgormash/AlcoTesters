#pragma compile(Console, True)
#pragma compile(ExecLevel, RequireAdministrator)
#RequireAdmin
#Include <File.au3>
#include <WinAPIFiles.au3>
#include <FileConstants.au3>

;Получаем список локальных дисков
Local $driveArray = DriveGetDrive("FIXED")
Local $resultArray
Local $SourceFilePath
Local $DestFilePath
If Not FileExists (@DesktopDir & "\log.txt") Then
	_FileCreate(@DesktopDir & "\log.txt")
 EndIf

If @error Then
    ConsoleWrite("Error while attempting to get drives.")
Else
    ConsoleWrite("Drive List:" & @CRLF)
	LogWrite("Drive List:" & @CRLF)
	For $i = 1 To $driveArray[0]
        ; Показывает все найденные диски и переводит букву в верхний регистр.
		ConsoleWrite( $i & "-" & StringUpper($driveArray[$i]) & @CRLF)
		LogWrite($i & "-" & StringUpper($driveArray[$i]) & @CRLF)
    Next
EndIf

For $i = 1 To $driveArray[0]
	ConsoleWrite("Searching on" & " " & StringUpper($driveArray[$i])& @CRLF)
	LogWrite("Searching on" & " " & StringUpper($driveArray[$i])& @CRLF)
	$resultArray = _FindFiles($driveArray[$i], "stat.mdb")
	If IsArray($resultArray) Then
		For $n = 1 To $resultArray[0]
			Local $jupExe = StringReplace($driveArray[$i] & "\" & $resultArray[$n], "stat.mdb", "StatisticAlc.exe")
			Local $lionExe = StringReplace($driveArray[$i] & "\" & $resultArray[$n], "stat.mdb", "Statistic400.exe")
			If(FileExists($jupExe)) Then
				$SourceFilePath = @ScriptDir & "\jupiter_stat.mdb"
			ElseIf (FileExists($lionExe)) Then
				$SourceFilePath = @ScriptDir & "\lion_stat.mdb"
			EndIf
			Local $DestFilePath = $driveArray[$i] & "\"& $resultArray[$n]

			If(ChangeFile($SourceFilePath, $DestFilePath) = 1) then
				ConsoleWrite ("File " & $driveArray[$i] & "\"& $resultArray[$n] & " change status is OK" & @CRLF)
				LogWrite("File " & $driveArray[$i] & "\"& $resultArray[$n] & " change status is OK" & @CRLF)
			Else
				ConsoleWrite ("File " & $driveArray[$i] & "\"& $resultArray[$n] & " change status is NOT OK" & @CRLF)
				LogWrite("File " & $driveArray[$i] & "\"& $resultArray[$n] & " change status is NOT OK" & @CRLF)
			EndIf
		Next
	EndIf
Next
MsgBox(4096, "Результат", "Я закончил")
Run("notepad.exe" & " " & @DesktopDir & "\log.txt", @WindowsDir, @SW_MAXIMIZE)
WinWait("[CLASS:Notepad]")
$hNotepad = WinGetHandle("[CLASS:Notepad]")
WinWaitClose($hNotepad)
FileDelete(@DesktopDir & "\log.txt")


Func _FindFiles($sRoot, $sFile)

    Local $FileList

    $FileList = _FileListToArrayRec($sRoot, $sFile, $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_SORT)
    If Not @error Then
        For $i = 1 To $FileList[0]
            ConsoleWrite($sRoot & '\' & $FileList[$i] & @CRLF)
			LogWrite($sRoot & '\' & $FileList[$i] & @CRLF)
        Next
    EndIf
    Return $FileList
EndFunc

Func LogWrite($sendedText)
   $file = FileOpen(@DesktopDir & "\log.txt", $FO_APPEND)
   FileWriteLine($file, $sendedText)
   FileClose($file)
EndFunc

Func ChangeFile($sourceFile, $destFile)
	Local $result = FileCopy($sourceFile,  $destFile, $FC_OVERWRITE)
	Return $result
EndFunc