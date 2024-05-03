' https://oshiete.goo.ne.jp/qa/5661099.html

Dim objFSO
Dim objFile
Dim strText
Dim strPath

' type "J:\Dropbox\PC5_cloud\pg\VB\testVBS\test"
strPath = inputbox("type the target file directory (includes the file name).", "INPUT BOX")

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
For Each objFile In objFSO.GetFolder(strPath).Files
    ' File Name
    If strText <> "" Then
        strText = strText & vbCrLf & objFile.Name
    Else
        strText = objFile.Name
    End If
    
    ' modify date
    strText = strText & ":" & objFile.DateLastModified
Next

WScript.Echo strText

Set objFSO = Nothing