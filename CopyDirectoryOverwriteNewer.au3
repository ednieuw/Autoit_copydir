#include <Array.au3>
#include <GuiEdit.au3>
#include <GuiConstantsEx.au3>

;MsgBox(0, "EJN V1.0  DirectoryCopy", "Started",1)

_CopyFolderNoOverwrite("C:\Users\Ed\Documents\Files2016\" , "A:\DataEd\Files2016\")    ; kopieer en overschrijf oudere bestanden

; MsgBox(0, "EJN V1.0  DirectoryCopy", "Copy Ended",4)

Func _CopyFolderNoOverwrite($SourceFolder, $DestinationFolder)
    Local $Array = _FileListToArray_Recursive($SourceFolder, "*", 1, 2, 1)
    If Not IsArray($Array) Or Not $Array[1] Then Return MsgBox(4096, 'No Files Found', 'Cannot Find any Files in ' & $SourceFolder)
    Local $TCopyArray[$Array[0] + 1][2]
	Local $CopyArray[$Array[0] + 1][2]
    $TCopyArray[0][0] = $Array[0] ; de geselecteerde files opslaan
    For $i = 1 To $Array[0]
        $TCopyArray[$i][0] = $Array[$i]
        $TCopyArray[$i][1] = StringRegExpReplace($Array[$i], StringReplace($SourceFolder, "\", "\\"), StringReplace($DestinationFolder, "\", "\\"))
    Next
	; Alle bestanden hebben nu de juiste source en destination strings. nu kunnen wij de tijden gaan vergelijken
	local $j
	$j=1
	$CopyArray[0][0]=0
	    For $i = 1 To $Array[0]
		if FileGetTime($TCopyArray[$i][0],0,1) > FileGetTime($TCopyArray[$i][1],0,1) Then
        $CopyArray[$j][0] = $Array[$i]
        $CopyArray[$j][1] = StringRegExpReplace($Array[$i], StringReplace($SourceFolder, "\", "\\"), StringReplace($DestinationFolder, "\", "\\"))
		$CopyArray[0][0]=$j
		$j=$j+1
		endif
    Next
	If $CopyArray[0][0]=0 Then Return  MsgBox(4096, 'No Files to Copy Found', 'Cannot Find any Files to Copy in ' & $SourceFolder,1)
    ;_ArrayDisplay($CopyArray); 2 Dimensional Array with Source File Path and Destination File Path

	Local $hEdit

	; Create GUI
	GUICreate("EJN V1.0  DesktopCopy", 800, 300)
	$hEdit = GUICtrlCreateEdit("Copied files" & @CRLF, 2, 2, 794, 268)
	GUISetState()

    For $i = 1 To $CopyArray[0][0]
		 ;MsgBox(4096, 'File copied', 'File: ' & $CopyArray[$i][0],2)
		FileCopy($CopyArray[$i][0], $CopyArray[$i][1], 9)
		_GUICtrlEdit_AppendText($hEdit,  $CopyArray[$i][0] & @CRLF)
    Next

	;Do
	;Until GUIGetMsg() = $GUI_EVENT_CLOSE
	_GUICtrlEdit_AppendText($hEdit, @CRLF & "Last file copied. Will close in 5 seconds"& @CRLF)
	Sleep(5000)
	GUIDelete()
EndFunc   ;==>_CopyFolderNoOverwrite

;===============================================================================
; $iRetItemType: 0 = Files and folders, 1 = Files only, 2 = Folders only
; $iRetPathType: 0 = Filename only, 1 = Path relative to $sPath, 2 = Full path/filename
Func _FileListToArray_Recursive($sPath, $sFilter = "*", $iRetItemType = 0, $iRetPathType = 0, $bRecursive = False)
    Local $sRet = "", $sRetPath = ""
    $sPath = StringRegExpReplace($sPath, "[\\/]+\z", "")
    If Not FileExists($sPath) Then Return SetError(1, 1, "")
    If StringRegExp($sFilter, "[\\/ :> <\|]|(?s)\A\s*\z") Then Return SetError(2, 2, "")
    $sPath &= "\|"
    $sOrigPathLen = StringLen($sPath) - 1
    While $sPath
        $sCurrPathLen = StringInStr($sPath, "|") - 1
        $sCurrPath = StringLeft($sPath, $sCurrPathLen)
        $Search = FileFindFirstFile($sCurrPath & $sFilter)
        If @error Then
            $sPath = StringTrimLeft($sPath, $sCurrPathLen + 1)
            ContinueLoop
        EndIf
        Switch $iRetPathType
            Case 1 ; relative path
                $sRetPath = StringTrimLeft($sCurrPath, $sOrigPathLen)
            Case 2 ; full path
                $sRetPath = $sCurrPath
        EndSwitch
        While 1
            $File = FileFindNextFile($Search)
            If @error Then ExitLoop
            If ($iRetItemType + @extended = 2) Then ContinueLoop
            $sRet &= $sRetPath & $File & "|"
        WEnd
        FileClose($Search)
        If $bRecursive Then
            $hSearch = FileFindFirstFile($sCurrPath & "*")
            While 1
                $File = FileFindNextFile($hSearch)
                If @error Then ExitLoop
                If @extended Then $sPath &= $sCurrPath & $File & "\|"
            WEnd
            FileClose($hSearch)
        EndIf
        $sPath = StringTrimLeft($sPath, $sCurrPathLen + 1)
    WEnd
    If Not $sRet Then Return SetError(4, 4, "")
    Return StringSplit(StringTrimRight($sRet, 1), "|")
EndFunc   ;==>_FileListToArray_Recursive