projectFolderName = "\prj"
documentsFolderName = "\doc"
rtlFolderName = "\rtl"
simulateFolderName = "\sim"



Set objFSO = CreateObject("Scripting.FileSystemObject")
currentFolderPath = objFSO.GetAbsolutePathName(".")
projectName = InputBox("current path is: " & currentFolderPath , "input name")


If projectName <> "" Then
    WScript.Echo "The project name is:" & projectName
Else
    WScript.Echo "You don't input name!"
    WScript.Quit 0
End If

IF objFSO.FolderExists(currentFolderPath & "\" & projectName) Then
    WScript.Echo "Folder existed!"
Else
    objFSO.CreateFolder projectName
    objFSO.CreateFolder projectName & projectFolderName
    objFSO.CreateFolder projectName & documentsFolderName
    objFSO.CreateFolder projectName & rtlFolderName
    objFSO.CreateFolder projectName & simulateFolderName
End if

Set objShell = CreateObject("WScript.Shell")

objShell.Run "code ./ & exit 0" & projectName, 1, True

WScript.Quit 





