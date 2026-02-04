' Bookmark system - per workbook (not shared across different files)
Option Explicit

Private Type BookmarkType
    WorkbookName As String
    SheetName As String
    Address As String
End Type

Private Bookmarks(1 To 9) As BookmarkType     ' Module-level array → persists while workbook is open

Sub BookmarkSet()
    Dim inputStr As String
    Dim slot As Long
    Dim action As String
    Dim msg As String
    Dim i As Long
    Dim ws As Worksheet
    Dim rng As Range
    ' Build current bookmark list for display
    msg = "Current bookmarks (1-9):" & vbCrLf & vbCrLf
    
    For i = 1 To 9
        If Bookmarks(i).SheetName = "" Then
            msg = msg & i & ": (empty)" & vbCrLf
        ElseIf Bookmarks(i).WorkbookName <> ActiveWorkbook.name Then
            msg = msg & i & ": " & "[used by other workbook]" & vbCrLf
        Else
            msg = msg & i & ": " & Bookmarks(i).SheetName & "!" & Bookmarks(i).Address & vbCrLf
        End If
    Next i
    
    msg = msg & vbCrLf & _
          "enter bookmark to set: 1 to 9" & vbCrLf
    
    inputStr = InputBox(msg, "Bookmark Manager", "")
    If Trim(inputStr) = "" Then Exit Sub

    If Len(inputStr) = 1 And IsNumeric(inputStr) Then
        slot = CLng(inputStr)
    Else
        MsgBox "Invalid format can only input 1 to 9", vbExclamation, "Invalid Input"
        Exit Sub
    End If
    
    
    ' SET bookmark using current active sheet and selection
    Bookmarks(slot).WorkbookName = ActiveWorkbook.name
    Bookmarks(slot).SheetName = ActiveSheet.name
    Bookmarks(slot).Address = Selection.Address   ' e.g. $A$1 or A1:B10
    

    
End Sub

Sub BookmarkGet()
    Dim inputStr As String
    Dim slot As Long
    Dim action As String
    Dim msg As String
    Dim i As Long
    Dim ws As Worksheet
    Dim rng As Range
    
    ' Build current bookmark list for display
    msg = "Current bookmarks (1-9):" & vbCrLf & vbCrLf
    
    For i = 1 To 9
        If Bookmarks(i).SheetName = "" Then
            msg = msg & i & ": (empty)" & vbCrLf
        ElseIf Bookmarks(i).WorkbookName <> ActiveWorkbook.name Then
            msg = msg & i & ": " & "[used by other workbook]" & vbCrLf
        Else
            msg = msg & i & ": " & Bookmarks(i).SheetName & "!" & Bookmarks(i).Address & vbCrLf
        End If
    Next i
    
    msg = msg & vbCrLf & _
          "enter bookmark to go to: 1 to 9" & vbCrLf
    
    inputStr = InputBox(msg, "Bookmark Manager", "")
    If Trim(inputStr) = "" Then Exit Sub
    
    ' Detect action
    If Len(inputStr) = 1 And IsNumeric(inputStr) Then
        slot = CLng(inputStr)
    Else
        MsgBox "Invalid format can only input 1 to 9", vbExclamation, "Invalid Input"
        Exit Sub
    End If
    

    ' GO TO bookmark
    If Bookmarks(slot).SheetName = "" Then
        MsgBox "Bookmark " & slot & " is empty.", vbExclamation, "No Bookmark"
        Exit Sub
    ElseIf Bookmarks(slot).WorkbookName <> ActiveWorkbook.name Then
        MsgBox "Bookmark " & slot & " is used by other workbook.", vbExclamation, "Forbidden Bookmark"
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Bookmarks(slot).SheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    On Error Resume Next
    ws.Activate
    Set rng = ws.Range(Bookmarks(slot).Address)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Invalid range address for bookmark " & slot & "." & vbCrLf & _
               "It may have been deleted or the sheet structure changed.", vbCritical, "Invalid Range"
        Exit Sub
    End If
    
    rng.Select
    Application.Goto rng
    
End Sub

