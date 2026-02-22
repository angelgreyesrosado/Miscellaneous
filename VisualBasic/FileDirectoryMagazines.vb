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
    
    ' Index column FIRST
    ws.Range("A1:F1").Value = Array("Index", "Full Path", "File Name", "Year", "Month", "Date Modified")
    
    nextRow = 2
    Set root = fso.GetFolder(startPath)
    
    Call ScanFolder(root, ws, nextRow)
    
    MsgBox "Completed. " & nextRow - 2 & " files processed.", vbInformation

End Sub


Private Sub ScanFolder(ByVal folder As Object, ByVal ws As Worksheet, ByRef nextRow As Long)

    Dim file As Object
    Dim subFolder As Object
    Dim yearVal As String, monthVal As String, indexVal As String
    
    For Each file In folder.Files
        
        Call ExtractYearMonth(file.Name, yearVal, monthVal)
        
        If yearVal <> "" And monthVal <> "" Then
            indexVal = yearVal & "|" & monthVal
        Else
            indexVal = ""
        End If
        
        ' Index
        ws.Cells(nextRow, 1).Value = indexVal
        
        ' Full Path as hyperlink
        ws.Hyperlinks.Add _
            Anchor:=ws.Cells(nextRow, 2), _
            Address:=file.Path, _
            TextToDisplay:=file.Path
        
        ' File Name
        ws.Cells(nextRow, 3).Value = file.Name
        
        ' Year
        ws.Cells(nextRow, 4).Value = yearVal
        
        ' Month
        ws.Cells(nextRow, 5).Value = monthVal
        
        ' Date Modified
        ws.Cells(nextRow, 6).Value = file.DateLastModified
        
        nextRow = nextRow + 1
    Next file
    
    For Each subFolder In folder.SubFolders
        Call ScanFolder(subFolder, ws, nextRow)
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
    cleaned = Replace(cleaned, ".", " ")   ' <-- key fix for "04.2022"
    
    parts = Split(cleaned, " ")
    
    ' FIRST PRIORITY: detect ANY 4-digit year anywhere in the filename
    fourDigitYear = GetYearFromString(cleaned)
    If fourDigitYear <> "" Then
        yr = fourDigitYear
    End If
    
    For Each p In parts
        
        ' Holiday edition (explicit "Hol")
        If InStr(p, "hol") > 0 Then
            mn = "Holiday"
            
            Dim y As String
            y = Trim(Replace(p, "hol", ""))
            
            If IsNumeric(y) Then
                If CInt(y) > 60 Then
                    yr = "19" & y
                Else
                    yr = "20" & y
                End If
            End If
            
            Exit Sub
        End If
        
        ' If no 4-digit year found, allow 4-digit token detection
        If yr = "" Then
            If IsNumeric(p) And Len(p) = 4 Then
                If CInt(p) >= 1900 And CInt(p) <= Year(Date) + 1 Then
                    yr = p
                End If
            End If
        End If
        
        ' Only use 2-digit year if NO 4-digit year was found
        If yr = "" Then
            If IsNumeric(p) And Len(p) = 2 Then
                If CInt(p) > 60 Then
                    yr = "19" & p
                Else
                    yr = "20" & p
                End If
            End If
        End If
        
        ' Month detection (numeric 1–12)
        If IsNumeric(p) And Len(p) <= 2 Then
            If CInt(p) >= 1 And CInt(p) <= 12 Then
                mn = MonthName(CInt(p))
                foundMonth = True
            End If
        End If
        
        ' Month detection (word)
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
        
    Next p
    
    ' 13 or 14 = Holiday edition if no month found
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


