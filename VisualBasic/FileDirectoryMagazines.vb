Option Explicit

' ============================================================
' MAIN ENTRY POINT
' ============================================================
Sub ReadOneDriveMagazines()

    Dim fso As Object
    Dim root As Object
    Dim ws As Worksheet
    Dim startPath As String
    Dim fileList As Collection
    Dim dataArr() As Variant
    Dim i As Long

    '-----------------------------------------
    ' SPEED BOOST: Turn off expensive features
    '-----------------------------------------
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

    On Error GoTo Cleanup

    startPath = "C:\Users\angel\OneDrive\Magazines"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(startPath) Then
        MsgBox "Folder not found: " & startPath, vbCritical
        GoTo Cleanup
    End If

    Set ws = ThisWorkbook.Sheets("OneDrive Files")
    ws.Cells.ClearContents

    ws.Range("A1:F1").Value = Array("Index", "Full Path", "File Name", "Year", "Month", "Date Modified")

    '-----------------------------------------
    ' SCAN FILES INTO MEMORY (FAST)
    '-----------------------------------------
    Set fileList = New Collection
    Set root = fso.GetFolder(startPath)

    Call CollectFiles(root, fileList, startPath)

    '-----------------------------------------
    ' BUILD ARRAY FOR FAST WRITE
    '-----------------------------------------
    ReDim dataArr(1 To fileList.Count, 1 To 6)

    For i = 1 To fileList.Count
        Dim info As Variant
        info = fileList(i)

        dataArr(i, 1) = info(0) ' Index
        dataArr(i, 2) = info(1) ' Full Path
        dataArr(i, 3) = info(2) ' File Name
        dataArr(i, 4) = info(3) ' Year
        dataArr(i, 5) = info(4) ' Month
        dataArr(i, 6) = info(5) ' Date Modified
    Next i

    '-----------------------------------------
    ' WRITE ALL ROWS IN ONE SHOT (FASTEST)
    '-----------------------------------------
    ws.Range("A2").Resize(fileList.Count, 6).Value = dataArr

    '-----------------------------------------
    ' ADD HYPERLINKS (FASTER VIA FORMULA)
    '-----------------------------------------
    Dim r As Long  
    For r = 2 To fileList.Count + 1  
        Dim linkPath As String  
        linkPath = Replace(ws.Cells(r, 2).Value, "#", "%23")  
        ws.Cells(r, 2).Formula = "=HYPERLINK(""file:///" & linkPath & """,""" & "file:///" & linkPath & """)"  
    Next r  
    '-----------------------------------------
    ' UPDATE INVENTORY
    '-----------------------------------------
    Call UpdateMagazineInventory(startPath)

    MsgBox "Completed. " & fileList.Count & " files processed.", vbInformation

Cleanup:
    '-----------------------------------------
    ' RESTORE EXCEL SETTINGS
    '-----------------------------------------
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True

End Sub


' ============================================================
' COLLECT FILES INTO MEMORY (NO SHEET WRITES)
' ============================================================
Private Sub CollectFiles(ByVal folder As Object, ByRef fileList As Collection, ByVal rootPath As String)

    Dim file As Object
    Dim subFolder As Object
    Dim yearVal As String, monthVal As String
    Dim indexVal As String, topFolder As String

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

        Dim info(0 To 5) As Variant
        info(0) = indexVal
        info(1) = file.Path
        info(2) = file.Name
        info(3) = yearVal
        info(4) = monthVal
        info(5) = file.DateLastModified

        fileList.Add info
    Next file

    For Each subFolder In folder.SubFolders
        Call CollectFiles(subFolder, fileList, rootPath)
    Next subFolder

End Sub


' ============================================================
' YEAR/MONTH EXTRACTION (UNCHANGED)
' ============================================================
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
    cleaned = Replace(cleaned, ".", " ")

    parts = Split(cleaned, " ")

    fourDigitYear = GetYearFromString(cleaned)
    If fourDigitYear <> "" Then yr = fourDigitYear

    For Each p In parts

        If InStr(p, "hol") > 0 Then
            mn = "Holiday"
            GoTo SkipLoop
        End If

        If InStr(p, "anni") > 0 Then
            mn = "Anniversary"
            GoTo SkipLoop
        End If

        If yr = "" Then
            If IsNumeric(p) And Len(p) = 4 Then
                If CInt(p) >= 1900 And CInt(p) <= Year(Date) + 1 Then yr = p
            End If
        End If

        If yr = "" Then
            If IsNumeric(p) And Len(p) = 2 Then
                If CInt(p) > 60 Then yr = "19" & p Else yr = "20" & p
            End If
        End If

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

        If foundMonth Then GoTo SkipNumeric

        If IsNumeric(p) And Len(p) <= 2 Then
            If CInt(p) >= 1 And CInt(p) <= 12 Then
                mn = MonthName(CInt(p))
                foundMonth = True
            End If
        End If

SkipNumeric:
    Next p

    If Not foundMonth Then
        For Each p In parts
            If p = "13" Or p = "14" Then
                mn = "Holiday"
                Exit For
            End If
        Next p
    End If

SkipLoop:
End Sub


' ============================================================
' YEAR DETECTION
' ============================================================
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


' ============================================================
' TOP-LEVEL FOLDER
' ============================================================
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


' ============================================================
' UPDATE INVENTORY (UNCHANGED)
' ============================================================
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


' ============================================================
' MAP PUBLICATION → MOST RECENT PDF
' ============================================================
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
