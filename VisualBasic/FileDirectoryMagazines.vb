Option Explicit

Sub ReadOneDriveMagazines()

    Dim fso As Object
    Dim root As Object
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim startPath As String
    
    startPath = "C:\Users\angel\OneDrive\Magazines"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(startPath) Then
        MsgBox "Folder not found: " & startPath, vbCritical
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("OneDrive Files")
    ws.Cells.ClearContents
    
    ws.Range("A1:F1").Value = Array("Index", "Full Path", "File Name", "Year", "Month", "Date Modified")
    
    nextRow = 2
    Set root = fso.GetFolder(startPath)
    
    Call ScanFolder(root, ws, nextRow, startPath)
    
    ' Update the Magazine Inventory sheet to only populate PDFs for matching publications
    Call UpdateMagazineInventory(startPath)
    
    MsgBox "Completed. " & nextRow - 2 & " files processed.", vbInformation

End Sub


Private Sub ScanFolder(ByVal folder As Object, ByVal ws As Worksheet, ByRef nextRow As Long, ByVal rootPath As String)

    Dim file As Object
    Dim subFolder As Object
    Dim yearVal As String, monthVal As String, indexVal As String, topFolder As String
    
    For Each file In folder.Files
        
        Call ExtractYearMonth(file.Name, yearVal, monthVal)
        
        If yearVal <> "" And monthVal <> "" Then
            indexVal = yearVal & "|" & monthVal
        Else
            indexVal = ""
        End If
        
        topFolder = GetTopLevelFolder(folder.Path, rootPath)
        If topFolder <> "" Then
            If indexVal <> "" Then
                indexVal = topFolder & "|" & indexVal
            Else
                indexVal = topFolder
            End If
        End If
        
        ws.Cells(nextRow, 1).Value = indexVal
        
        ws.Hyperlinks.Add Anchor:=ws.Cells(nextRow, 2), _
            Address:=file.Path, TextToDisplay:=file.Path
        
        ws.Cells(nextRow, 3).Value = file.Name
        ws.Cells(nextRow, 4).Value = yearVal
        ws.Cells(nextRow, 5).Value = monthVal
        ws.Cells(nextRow, 6).Value = file.DateLastModified
        
        nextRow = nextRow + 1
    Next file
    
    For Each subFolder In folder.SubFolders
        Call ScanFolder(subFolder, ws, nextRow, rootPath)
    Next subFolder

End Sub


Private Sub ExtractYearMonth(ByVal fileName As String, ByRef yr As String, ByRef mn As String)

    Dim parts() As String
    Dim p As Variant
    Dim cleaned As String
    Dim foundMonth As Boolean
    Dim fourDigitYear As String
    
    yr = ""
    mn = ""
    foundMonth = False
    
    cleaned = LCase(fileName)
    cleaned = Replace(cleaned, ".pdf", "")
    cleaned = Replace(cleaned, ".jpg", "")
    cleaned = Replace(cleaned, ".jpeg", "")
    cleaned = Replace(cleaned, ".png", "")
    cleaned = Replace(cleaned, ".zip", "")
    cleaned = Replace(cleaned, "_", " ")
    cleaned = Replace(cleaned, "-", " ")
    cleaned = Replace(cleaned, ".", " ")   ' <-- FIX para casos como 04.2022
    
    parts = Split(cleaned, " ")
    
    ' Detectar año de 4 dígitos en cualquier parte
    fourDigitYear = GetYearFromString(cleaned)
    If fourDigitYear <> "" Then yr = fourDigitYear
    
    For Each p In parts
        
        ' Holiday explícito
        If InStr(p, "hol") > 0 Then
            mn = "Holiday"
            Dim y As String
            y = Trim(Replace(p, "hol", ""))
            If IsNumeric(y) Then
                If CInt(y) > 60 Then yr = "19" & y Else yr = "20" & y
            End If
            Exit Sub
        End If
        
        ' Año de 4 dígitos si no se detectó antes
        If yr = "" Then
            If IsNumeric(p) And Len(p) = 4 Then
                If CInt(p) >= 1900 And CInt(p) <= Year(Date) + 1 Then yr = p
            End If
        End If
        
        ' Año de 2 dígitos si no hay año de 4 dígitos
        If yr = "" Then
            If IsNumeric(p) And Len(p) = 2 Then
                If CInt(p) > 60 Then yr = "19" & p Else yr = "20" & p
            End If
        End If
        
        ' MES POR PALABRA (PRIORIDAD MÁXIMA)
        Select Case p
            Case "jan", "january": mn = "January": foundMonth = True
            Case "feb", "february": mn = "February": foundMonth = True
            Case "mar", "march": mn = "March": foundMonth = True
            Case "apr", "april": mn = "April": foundMonth = True
            Case "may": mn = "May": foundMonth = True
            Case "jun", "june": mn = "June": foundMonth = True
            Case "jul", "july": mn = "July": foundMonth = True
            Case "aug", "august": mn = "August": foundMonth = True
            Case "sep", "sept", "september": mn = "September": foundMonth = True
            Case "oct", "october": mn = "October": foundMonth = True
            Case "nov", "november": mn = "November": foundMonth = True
            Case "dec", "december": mn = "December": foundMonth = True
        End Select
        
        ' Si ya se detectó mes por palabra, ignorar meses numéricos
        If foundMonth Then GoTo SkipNumeric
        
        ' MES NUMÉRICO SOLO SI NO HAY MES POR PALABRA
        If IsNumeric(p) And Len(p) <= 2 Then
            If CInt(p) >= 1 And CInt(p) <= 12 Then
                mn = MonthName(CInt(p))
                foundMonth = True
            End If
        End If
        
SkipNumeric:
    Next p
    
    ' 13 o 14 = Holiday si no hay mes
    If Not foundMonth Then
        For Each p In parts
            If p = "13" Or p = "14" Then
                mn = "Holiday"
                Exit For
            End If
        Next p
    End If

End Sub


Private Function GetYearFromString(ByVal s As String) As String
    Dim i As Long
    Dim candidate As String
    Dim y As Long
    
    For i = 1 To Len(s) - 3
        candidate = Mid$(s, i, 4)
        If IsNumeric(candidate) Then
            y = CLng(candidate)
            If y >= 1900 And y <= Year(Date) + 1 Then
                GetYearFromString = CStr(y)
                Exit Function
            End If
        End If
    Next i
    
    GetYearFromString = ""
End Function

Private Function GetTopLevelFolder(ByVal folderPath As String, ByVal rootPath As String) As String
    Dim rel As String

    rel = folderPath
    If LCase(Left(rel, Len(rootPath))) = LCase(rootPath) Then
        rel = Mid(rel, Len(rootPath) + 1)
    End If
    If Left(rel, 1) = "\" Then rel = Mid(rel, 2)

    If rel = "" Then
        GetTopLevelFolder = ""
        Exit Function
    End If

    GetTopLevelFolder = Split(rel, "\")(0)
End Function

Private Sub UpdateMagazineInventory(ByVal startPath As String)
    Dim ws As Worksheet
    Dim pubCol As Long, docCol As Long
    Dim lastCol As Long, lastRow As Long
    Dim c As Long
    Dim pub As String
    Dim pubMap As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Magazine Inventory")
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        Dim header As String
        header = Trim(CStr(ws.Cells(1, c).Value))
        If LCase(header) = "publication" Then pubCol = c
        If LCase(header) = "pdf document" Or LCase(header) = "document" Then docCol = c
    Next c

    If pubCol = 0 Or docCol = 0 Then Exit Sub

    Set pubMap = GetPublicationPdfMap(startPath)
    If pubMap Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, pubCol).End(xlUp).Row

    For c = 2 To lastRow
        pub = Trim(CStr(ws.Cells(c, pubCol).Value))
        If pub <> "" And pubMap.Exists(pub) Then
            ws.Cells(c, docCol).Value = pubMap(pub)
        Else
            ws.Cells(c, docCol).ClearContents
        End If
    Next c
End Sub

Private Function GetPublicationPdfMap(ByVal startPath As String) As Object
    Dim fso As Object
    Dim root As Object
    Dim subFolder As Object
    Dim file As Object
    Dim bestPdfPath As String
    Dim bestDate As Date
    Dim pubMap As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(startPath) Then
        Set GetPublicationPdfMap = Nothing
        Exit Function
    End If

    Set root = fso.GetFolder(startPath)
    Set pubMap = CreateObject("Scripting.Dictionary")
    pubMap.CompareMode = vbTextCompare

    For Each subFolder In root.SubFolders
        bestPdfPath = ""
        bestDate = #1/1/1900#
        For Each file In subFolder.Files
            If LCase(Right(file.Name, 4)) = ".pdf" Then
                If file.DateLastModified > bestDate Then
                    bestDate = file.DateLastModified
                    bestPdfPath = file.Path
                End If
            End If
        Next file
        pubMap(subFolder.Name) = bestPdfPath
    Next subFolder

    Set GetPublicationPdfMap = pubMap
End Function
