Public Sub LoadGitHubModule()
    Dim url As String
    Dim tempFile As String
    Dim http As Object
    Dim stream As Object

    url = "https://raw.githubusercontent.com/angelgreyesrosado/Miscellaneous/refs/heads/main/VisualBasic/FileDirectoryMagazines.bas"
    tempFile = Environ$("TEMP") & "\FileDirectoryMagazines.bas"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False

    On Error Resume Next
    http.send
    On Error GoTo 0

    ' If download failed or returned non-200 status, keep existing module
    If http.readyState <> 4 Or http.Status <> 200 Then
        MsgBox "Could not update VBA module from GitHub — using last saved version.", vbExclamation, "GitHub Update Failed"
        Debug.Print "GitHub module NOT updated — using last imported version."
        Exit Sub
    End If

    ' Save downloaded file
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile tempFile, 2
    stream.Close

    ' Replace contents of Module1
    Dim vbModule As Object
    Dim codeModule As Object
    Dim newCode As String
    Dim lineCount As Long

    On Error Resume Next
    Set vbModule = ThisWorkbook.VBProject.VBComponents("Module1")
    On Error GoTo 0

    If vbModule Is Nothing Then
        MsgBox "Module1 not found in this workbook.", vbCritical, "Error"
        Exit Sub
    End If

    ' Read the downloaded file
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text mode
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile tempFile
    newCode = stream.ReadText
    stream.Close

    ' Clear existing code and add new code
    Set codeModule = vbModule.CodeModule
    lineCount = codeModule.CountOfLines
    
    If lineCount > 0 Then
        codeModule.DeleteLines 1, lineCount
    End If

    codeModule.AddFromString newCode

    MsgBox "Module1 updated successfully from GitHub.", vbInformation, "Update Complete"
    Debug.Print "Module1 updated successfully from GitHub."
End Sub


