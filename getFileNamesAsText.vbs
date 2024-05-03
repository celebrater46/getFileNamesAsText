' https://oshiete.goo.ne.jp/qa/5661099.html

Dim objFSO
Dim objFile
Dim strText
Dim strPath
Dim strProjectPath

' type "J:\Dropbox\PC5_cloud\pg\VB\testVBS\test"
strPath = inputbox("type the target file directory (includes the file name).", "INPUT BOX")
strProjectPath = fso.getParentFolderName(WScript.ScriptFullName) & "\files.txt"

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.CreateTextFile(strProjectPath)

For Each objFile In objFSO.GetFolder(strPath).Files
    ' File Name
    If strText <> "" Then
        strText = strText & vbCrLf & objFile.Name
    Else
        strText = objFile.Name
    End If
    
    ' Add a line break
    ' strText = strText & "\r\n"
    objFile.WriteLine(strText)
Next

' WScript.Echo strText

Set objFSO = Nothing




' ' Get the current directory
' Dim fso
' Set fso = createObject("Scripting.FileSystemObject")
' ' Msgbox fso.getParentFolderName(WScript.ScriptFullName)

' Dim strPath, objFS, objFile
' strPath = fso.getParentFolderName(WScript.ScriptFullName) & "\test\testCTF.txt"

' Set objFS = CreateObject("Scripting.FileSystemObject")
' Set objFile = objFS.CreateTextFile(strPath)

' objFile.WriteLine("Hello world!!!!!!!!")
' objFile.Close