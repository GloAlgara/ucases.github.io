strScriptFile = Wscript.ScriptFullName ' C:\download.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
strFolder = objFSO.GetParentFolderName(objFile) ' C:

Set objShell = CreateObject("WScript.Shell")
 
strLink = "http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab"
' Use strFolder to save on the same location of this script.
strSaveTo = "C:\"
 
' WGet saves file always on the actual folder. So, change the actual folder for C:\, where we want to save file
objShell.CurrentDirectory = strSaveTo
 
' "C:\wget.exe" "http://download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab" -N
objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(strLink) & " -N",1,True
' -N: Continue download only if the local version is outdated.

objShell.CurrentDirectory = strFolder
 
' Add Quotes to string
' http://stackoverflow.com/questions/2942554/vbscript-adding-quotes-to-a-string
Function Quotes(strQuotes)
	Quotes = chr(34) & strQuotes & chr(34)
End Function
