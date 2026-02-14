Private g_Folder As Outlook.MAPIFolder
Private g_Find As String
 
Public Sub FindFolder()
Dim xFldName As String
Dim xFolders As Outlook.Folders
Dim xYesNo As Integer
On Error Resume Next
Set g_Folder = Nothing
g_Find = ""
xFldName = InputBox("Folder name:", "Kutools for Outlook")
If Trim(xFldName) = "" Then Exit Sub
g_Find = xFldName
g_Find = UCase(g_Find)
Set xFolders = Application.Session.Folders
LoopFolders xFolders
If Not g_Folder Is Nothing Then
    xYesNo = MsgBox("Activate folder: " & vbCrLf & g_Folder.FolderPath, vbQuestion Or vbYesNo, "Kutools for Outlook")
    If xYesNo = vbYes Then
        Set Application.ActiveExplorer.CurrentFolder = g_Folder
    End If
Else
    MsgBox "Not found", vbInformation, "Kutools for Outlook"
End If
End Sub
 
Private Sub LoopFolders(Folders As Outlook.Folders)
Dim xFolder As Outlook.MAPIFolder
Dim xFound As Boolean
On Error Resume Next
xFound = False
For Each xFolder In Folders
    If UCase(xFolder.Name) = g_Find Then xFound = True
    If xFound Then
        Set g_Folder = xFolder
        Exit For
    Else
        LoopFolders xFolder.Folders
        If Not g_Folder Is Nothing Then Exit For
    End If
Next
End Sub