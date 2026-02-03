Sub SelectSheetByFilter()
    Dim ws As Worksheet
    Dim visibleSheets As Collection
    Dim filteredSheets As Collection
    Dim i As Long
    Dim userFilter As String
    Dim userChoice As Variant
    
    ' Step 1: Collect visible sheets in workbook order
    Set visibleSheets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            visibleSheets.Add ws
        End If
    Next ws
    
    ' Step 2: Show all visible sheets
    Dim msg As String
    msg = "Visible sheets:" & vbCrLf
    For i = 1 To visibleSheets.Count
        msg = msg & "- " & visibleSheets(i).name & vbCrLf
    Next i
    'MsgBox msg, vbInformation, "Step 1: All Visible Sheets"
    
    ' Step 3: Ask user for filter text
    userFilter = InputBox("Enter text to filter sheet names (case-insensitive):", "Filter Sheets")
    If userFilter = "" Then Exit Sub
    
    ' Step 4: Collect filtered sheets
    Set filteredSheets = New Collection
    For i = 1 To visibleSheets.Count
        If InStr(1, visibleSheets(i).name, userFilter, vbTextCompare) > 0 Then
            filteredSheets.Add visibleSheets(i)
        End If
    Next i
    
    ' Step 5: Show filtered list
    If filteredSheets.Count = 0 Then
        MsgBox "No sheets match your filter.", vbExclamation, "No Match"
        Exit Sub
    End If
    
    msg = "Filtered sheets:" & vbCrLf
    For i = 1 To filteredSheets.Count
        msg = msg & i & ") " & filteredSheets(i).name & vbCrLf
    Next i
    
    userChoice = InputBox(msg & vbCrLf & "Enter the number of the sheet to activate:", "Select Sheet")
    
    ' Step 6: Activate chosen sheet
    If IsNumeric(userChoice) Then
        i = CLng(userChoice)
        If i >= 1 And i <= filteredSheets.Count Then
            filteredSheets(i).Activate
        Else
            MsgBox "Invalid selection.", vbCritical, "Error"
        End If
    Else
        MsgBox "Invalid input.", vbCritical, "Error"
    End If
End Sub

Sub SelectWorkbookByName()
    Dim wb As Workbook
    Dim msg As String
    Dim i As Long
    Dim userChoice As Variant
    
    ' Step 1: Build list of open workbooks
    msg = "Open workbooks:" & vbCrLf
    For i = 1 To Application.Workbooks.Count
        msg = msg & i & ") " & Application.Workbooks(i).name & vbCrLf
    Next i
    
    ' Step 2: Ask user to choose
    userChoice = InputBox(msg & vbCrLf & "Enter the number of the workbook to activate:", "Select Workbook", 1)
    
    ' Step 3: Activate chosen workbook

    If IsNumeric(userChoice) Then
        i = CLng(userChoice)
        If i >= 1 And i <= Application.Workbooks.Count Then
            Application.Workbooks(i).Activate
        Else
            MsgBox "Invalid selection.", vbCritical, "Error"
        End If
    Else
        MsgBox "Invalid input.", vbCritical, "Error"
    End If
End Sub

Sub SelectRecentByName()
    Dim wb As Workbook
    Dim msg As String
    Dim i As Long
    Dim userChoice As Variant
    Dim filteredWB As Collection
    
    
    ' Step 1: Ask user for filter text
    userFilter = InputBox("Enter text to filter recent workbook path (case-insensitive):", "Filter Workbook")
    If userFilter = "" Then Exit Sub
    
    ' Step 2: Collect filtered WB
    Set filteredWB = New Collection
    For i = 1 To Application.RecentFiles.Count
        If InStr(1, Application.RecentFiles(i).Path, userFilter, vbTextCompare) > 0 Then
            filteredWB.Add Application.RecentFiles(i)
        End If
    Next i
    
    
    ' Step 3: Show filtered list
    If filteredWB.Count = 0 Then
        MsgBox "No recent workbook match your filter.", vbExclamation, "No Match"
        Exit Sub
    End If
    
    msg = "Filtered workbooks:" & vbCrLf
    For i = 1 To filteredWB.Count
        
        msg = msg & i & ") " & filteredWB(i).Path & vbCrLf
    Next i
    
    userChoice = InputBox(msg & vbCrLf & "Enter the number of the workbook to open:", "Select Workbook")
    
    ' Step 4: Activate chosen workbook

    If IsNumeric(userChoice) Then
        i = CLng(userChoice)
        If i >= 1 And i <= Application.RecentFiles.Count Then
            arr = Split(filteredWB(i).Path, "\")
            Debug.Print arr(0)
            Debug.Print arr(1)
            resp = MsgBox("Open following path ?" & vbCrLf & filteredWB(i).Path, vbYesNo, "Open Workbook")
            If resp = vbYes Then filteredWB(i).Open
        Else
            MsgBox "Invalid selection.", vbCritical, "Error"
        End If
    Else
        MsgBox "Invalid input.", vbCritical, "Error"
    End If
End Sub



