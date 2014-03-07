' =============================================================================
' VBS swapping script
' 
' Created by Arthur in 2012
' =============================================================================

'
' Set strRootFolder to the top level of your enumeration
'
strRootFolder = "C:\Users\Arthur\Desktop\0"
strPathSwapFile = "C:\Users\Arthur\Desktop\1"

' ==============================================
' Save to an array all lines from SwapNames
' ==============================================
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (strPathSwapFile & "\SwapNames.txt", ForReading)
Do Until objTextFile.AtEndOfStream
    strNextLine = objTextFile.ReadAll
    arrServiceList = Split(strNextLine , vbCrLf)
Loop

' ==============================================
' Enumerates a folder tree and returns all files
' ==============================================
Set objFso = CreateObject("Scripting.FileSystemObject")

GetFiles strRootFolder
'GetSubFolders strRootFolder

Sub GetFiles(sFolder)

	On Error Resume Next

	Set oFolder = objFso.GetFolder(sFolder)
	Set cf = oFolder.Files

	i = 0
	For Each oFile In cF
	'
	' Do file checking stuff here
	'
		'WScript.Echo "Alt: " & oFile & vbCrLf & "Neu: " & oFolder & "\" & arrServiceList(i) & ".avi"
		'oFile.Rename(oFolder & "\" & arrServiceList(i) & ".avi")
		objFSO.MoveFile oFile, oFolder & "\" & arrServiceList(i) & ".avi"
		i = i + 1
		DoEvents
	Next

	Set oFolder = Nothing
	Set cF = Nothing

End Sub