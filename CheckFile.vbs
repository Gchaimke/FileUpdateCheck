'
' VBSScript for checking last time file updated from now by Chaim Gorbov
' if  (Now - DateLastModified in minutes(n)) more then 10 minutes then show warning popap for 10 minutes
' else (Now - DateLastModified in minutes(n)) less then 10 minutes then show no errors popap for 10 minutes
Option Explicit 

Dim oFSO 				'Object of file system
Dim PathToFile			'Path to file string
Dim file				'File object
Dim StopNow 			'Boolin true = quit

SET oFSO = CreateObject("Scripting.FileSystemObject") 
PathToFile = "P:\Integral\IntOut\HEV007\SITEEXP\Items.CSV"
'-- determine if file exist
If Not oFSO.FileExists(PathToFile) Then 
CreateObject("WScript.Shell").Popup "File not found, check file path.", 600, "Warning",48 'worning popap
StopNow = True 'quit from job with error
else
SET file = oFSO.GetFile(PathToFile) 'set object of file from string
	If DateDiff("n", file.DateLastModified, Now) > 10 Then 'if  (Now - DateLastModified in minutes(n)) more then 10 minutes then show warning popap for 10 minutes
            CreateObject("WScript.Shell").Popup "Warning last time file Items.csv modified at : " & file.DateLastModified, 600, "Warning",48
			else
			'else (Now - DateLastModified in minutes(n)) less then 10 minutes then show no errors popap for 10 minutes
			CreateObject("WScript.Shell").Popup "File Items.csv up to date!", 600, "NO ERRORS",64
	End If
End If 
If StopNow Then 
Wscript.Quit(16) 'terminate script
End If 
