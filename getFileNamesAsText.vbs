Dim objFSO, objFile, objTextFile, strText, strPath, strProjectPath

strPath = inputbox("Input the target directory.", "INPUT BOX")

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
strProjectPath = objFSO.getParentFolderName(WScript.ScriptFullName) & "\files.txt"
Set objTextFile = objFSO.CreateTextFile(strProjectPath)

For Each objFile In objFSO.GetFolder(strPath).Files
    ' File Name
    If strText <> "" Then
        strText = strText & vbCrLf & objFile.Name
    Else
        strText = objFile.Name
    End If
    
Next

objTextFile.WriteLine(strText)

Set objFSO = Nothing
