Option Explicit
'widgets: lblStatus cmdOpen cmdOpenFolder cmdCancel txtFilter lstFiles
' Dictionary: display text → RecentFile object
Private MasterList As Object

Private Sub UserForm_Initialize()
    Set MasterList = CreateObject("Scripting.Dictionary")
    LoadRecentFiles
    PopulateList MasterList
    Me.Caption = "Select Recent Workbook"
    txtFilter.SetFocus
End Sub

Private Sub LoadRecentFiles()
    Dim rf As Object    ' RecentFile
    Dim display As String
    
    MasterList.RemoveAll
    
    On Error Resume Next
    For Each rf In Application.RecentFiles
        ' Create a readable display string
        display = rf.name
        If Len(rf.Path) > 0 Then
            display = rf.Path
        End If
        
        ' Avoid duplicate keys (rare but possible)
        If Not MasterList.Exists(display) Then
            Set MasterList(display) = rf
        Else
            ' If duplicate name, add index or path fragment
            Dim i As Long: i = 2
            Do While MasterList.Exists(display & " #" & i)
                i = i + 1
            Loop
            display = display & " #" & i
            Set MasterList(display) = rf
        End If
    Next rf
    On Error GoTo 0
End Sub

Private Sub txtFilter_Change()
    Dim temp As Object
    Dim key As Variant
    Dim filter As String
    
    Set temp = CreateObject("Scripting.Dictionary")
    filter = LCase(Trim(txtFilter.text))
    
    If Len(filter) = 0 Then
        PopulateList MasterList
        Exit Sub
    End If
    
    For Each key In MasterList.Keys
        If InStr(1, LCase(key), filter, vbTextCompare) > 0 Then
            Set temp(key) = MasterList(key)
        End If
    Next
    
    PopulateList temp
End Sub

Private Sub PopulateList(ByVal dict As Object)
    Dim key As Variant
    
    lstFiles.Clear
    
    If dict.Count = 0 Then
        lstFiles.AddItem "(no matching recent files)"
        cmdOpen.Enabled = False
        Exit Sub
    End If
    
    For Each key In dict.Keys
        lstFiles.AddItem key
    Next key
    
    If lstFiles.ListCount > 0 Then
        lstFiles.ListIndex = 0
        cmdOpen.Enabled = True
    End If
    
End Sub

Private Sub cmdOpen_Click()
    If lstFiles.ListIndex < 0 Then Exit Sub
    
    Dim selectedDisplay As String
    selectedDisplay = lstFiles.List(lstFiles.ListIndex)
    
    If Not MasterList.Exists(selectedDisplay) Then
        MsgBox "Selected item is no longer available.", vbExclamation
        Exit Sub
    End If
    
    Dim rf As Object
    Set rf = MasterList(selectedDisplay)
    
    Dim msg As String
    msg = "Open this file?" & vbCrLf & rf.Path
    
    If MsgBox(msg, vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    On Error Resume Next
    rf.Open
    If Err.Number <> 0 Then
        MsgBox "Could not open the file." & vbCrLf & vbCrLf & Err.Description, vbCritical
    Else
        Unload Me
    End If
    On Error GoTo 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub lstFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOpen_Click
End Sub

Private Sub UserForm_Terminate()
    Set MasterList = Nothing
End Sub

''' UX optimization'''
Private Sub lstFiles_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then   ' Enter or Space
        KeyCode = 0
        cmdOpen_Click   ' or cmdActivate_Click in the other form
    End If
End Sub
Private Sub txtFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0  ' ← very important: cancel the key so it doesn't insert characters or move cursor
        If lstFiles.ListIndex >= 0 Then
            cmdOpen_Click
        End If
    End If

    If KeyCode = vbKeyUp Then
        KeyCode = 0   ' ← very important: cancel the key so it doesn't insert characters or move cursor
        If lstFiles.ListIndex >= 0 Then
            If lstFiles.ListIndex > 0 Then lstFiles.ListIndex = lstFiles.ListIndex - 1
        End If
    End If

    If KeyCode = vbKeyDown Then
        KeyCode = 0   ' ← very important: cancel the key so it doesn't insert characters or move cursor
        If lstFiles.ListIndex >= 0 Then
            If lstFiles.ListIndex < lstFiles.ListCount - 1 Then lstFiles.ListIndex = lstFiles.ListIndex + 1
        End If
    End If
End Sub
Private Sub cmdOpenFolder_Click()
    If lstFiles.ListIndex < 0 Then
        MsgBox "Please select a file first.", vbExclamation
        Exit Sub
    End If
    
    Dim selectedDisplay As String
    selectedDisplay = lstFiles.List(lstFiles.ListIndex)
        

        
        
    If Not MasterList.Exists(selectedDisplay) Then
        MsgBox "Selected item is no longer available.", vbExclamation
        Exit Sub
    End If
    
    Dim rf As Object   ' RecentFile
    Set rf = MasterList(selectedDisplay)
    
    Dim folderPath As String
    Dim pos As Integer
    pos = InStrRev(selectedDisplay, "\")
    folderPath = Left(selectedDisplay, pos - 1)
    
    If Len(folderPath) = 0 Or Dir(folderPath, vbDirectory) = "" Then
        MsgBox "The containing folder cannot be found or no longer exists." & vbCrLf & vbCrLf & _
               "Path: " & folderPath, vbExclamation
        Exit Sub
    End If
    
    ' Open the folder in Explorer
    On Error Resume Next
    If MsgBox("Open this folder?" & vbCrLf & folderPath, vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "Could not open the folder." & vbCrLf & Err.Description, vbCritical
    End If
    On Error GoTo 0
End Sub

