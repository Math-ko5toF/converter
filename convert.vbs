inputName = InputBox("input format Math-ko5toF")
outName = InputBox("out format")

Set WshShell = CreateObject("WScript.Shell")
FolderPath = WshShell.CurrentDirectory

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(FolderPath)
Set colFiles = objFolder.Files

Set objRegExp = CreateObject("VBScript.RegExp")
objRegExp.Pattern = "."& inputName


For Each objFile in colFiles
    if objRegExp.test(objFile.Name) then
	ast = ast & objFile.Name & vblf
    	Set objfile = objFSO.GetFile(FolderPath & "\" & objFile.Name)
	ShortName = objRegExp.Replace(FolderPath & "\" & objFile.Name, "")

        objfile.move ShortName &"."& outName
    End If 
next
