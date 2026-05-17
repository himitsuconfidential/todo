Option Explicit

'------------------------------------------------------------------------------
' Import / Export macro
'
' Reads external-link formulas on the "Import" sheet, copies matching values
' from the "Export" sheet (same row/column), and writes them into the linked
' workbooks.
'
' Import sheet: formulas pointing at external files, e.g.
'   ='C:\Folder\[Book.xlsx]Sheet1'!A1
' Export sheet: plain values to push (must not contain formulas in those cells)
'------------------------------------------------------------------------------

' Parsed parts of one external-link formula
Private Type ExternalRef
    fullPath As String      ' e.g. C:\Folder\Book.xlsx
    wbName As String        ' e.g. Book.xlsx  (file name only)
    sheetName As String     ' e.g. Sheet1
    CellAddress As String   ' e.g. A1
End Type



'------------------------------------------------------------------------------
' Validates each distinct (workbook path, sheet name) once after workbooks are open.
' Keys are stored as FullPath & Chr(1) & SheetName.
'------------------------------------------------------------------------------
Private Function AllTargetSheetsExist( _
    workbookList As Object, _
    targetWorkbookSheetKeys As Object) As Boolean

    Dim k As Variant
    Dim delimPos As Long
    Dim fp As String
    Dim tgtSheet As String
    Dim wb As Workbook
    Dim ws As Worksheet

    AllTargetSheetsExist = False

    For Each k In targetWorkbookSheetKeys.Keys

        delimPos = InStr(1, CStr(k), Chr(1))
        fp = Left(CStr(k), delimPos - 1)
        tgtSheet = Mid(CStr(k), delimPos + 1)

        Set wb = WorkbookFromFullPath(fp, workbookList)

        If wb Is Nothing Then

            MsgBox _
                "Workbook is not open:" & vbCrLf & vbCrLf & fp, _
                vbCritical

            Exit Function

        End If

        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Sheets(tgtSheet)
        On Error GoTo 0

        If ws Is Nothing Then

            MsgBox _
                "Target sheet not found:" & vbCrLf & tgtSheet & vbCrLf & fp, _
                vbExclamation

            Exit Function

        End If

    Next k

    AllTargetSheetsExist = True

End Function


'------------------------------------------------------------------------------
' Resolves workbook by FullName match; fallback to opened file with matching name.
'------------------------------------------------------------------------------
Private Function WorkbookFromFullPath( _
    ByVal fullPath As String, _
    workbookList As Object) As Workbook

    Dim wb As Workbook
    Dim shortName As String

    Set WorkbookFromFullPath = Nothing

    For Each wb In Workbooks

        If (wb.FullName) = (fullPath) Then
            Set WorkbookFromFullPath = wb
            Exit Function
        End If

    Next wb



    For Each wb In Workbooks

        If (wb.Name) = (fullPath) Then
            Set WorkbookFromFullPath = wb
            Exit Function
        End If

    Next wb

End Function


'------------------------------------------------------------------------------
' BackupSheetValuesOnly
'
' Copies wsSource to a new sheet named  <namePrefix>_Backup_<timestamp>
' and converts the copy to values only (no formulas).
'------------------------------------------------------------------------------
Private Sub BackupSheetValuesOnly( _
    wsSource As Worksheet, _
    ByVal namePrefix As String)

    Dim wsBackup As Worksheet
    Dim wb As Workbook
    Dim backupName As String

    backupName = namePrefix & "_Backup_" & Format(Now, "yyyymmdd_hhnnss")

    Set wb = wsSource.Parent
    wsSource.Copy After:=wb.Sheets(wb.Sheets.Count)
    Set wsBackup = wb.Sheets(wb.Sheets.Count)
    wsBackup.Name = backupName

    With wsBackup.UsedRange
        .Value = .Value
    End With

End Sub


'------------------------------------------------------------------------------
' Returns True when formula looks like an external workbook link
' (contains both [ and ] around the file name)
'------------------------------------------------------------------------------
Function IsExternalReference(formulaText As String) As Boolean

    IsExternalReference = _
        (InStr(formulaText, "[") > 0 And _
         InStr(formulaText, "]") > 0)

End Function


'------------------------------------------------------------------------------
' Returns True if a workbook with the given file name is already open
'------------------------------------------------------------------------------
Function WorkbookIsOpen(wbName As String) As Boolean

    Dim wb As Workbook
    
    WorkbookIsOpen = False
    
    For Each wb In Workbooks
    
        If (wb.Name) = (wbName) Then
            WorkbookIsOpen = True
            Exit Function
        End If
        
    Next wb

End Function


'------------------------------------------------------------------------------
' ParseExternalReference
'
' Splits an external-link formula into path, workbook, sheet, and cell.
'
' Example input:
'   ='C:\Folder\[Template.xlsx]Sheet1'!K14
'
' Example output:
'   FullPath     = C:\Folder\Template.xlsx
'   WbName       = Template.xlsx
'   SheetName    = Sheet1
'   CellAddress  = K14
'
' Quoted paths with brackets in the file name are supported, e.g.
'   ='C:\t\[2.xlsx]sheet1'!A1  ->  C:\t\2.xlsx
'------------------------------------------------------------------------------
Function ParseExternalReference(ByVal formulaText As String) As ExternalRef

    Dim result As ExternalRef
    Dim temp As String
    Dim pathPrefix As String
    Dim sheetPart As String
    Dim exclPos As Long
    Dim bracketStart As Long
    Dim bracketEnd As Long
    
    ' Drop leading "="
    temp = Replace(formulaText, "=", "")
    
    ' Positions of !  [  ]  in the link string
    exclPos = InStr(temp, "!")
    bracketStart = InStr(temp, "[")
    bracketEnd = InStr(temp, "]")
    
    If exclPos = 0 Or bracketStart = 0 Or bracketEnd = 0 Or bracketEnd < bracketStart Then
        Err.Raise vbObjectError + 1, "ParseExternalReference", "Invalid external reference: " & formulaText
    End If
    
    ' Everything after "!" is the remote cell address
    result.CellAddress = Mid(temp, exclPos + 1)
    
    ' Text inside [ ] is the workbook file name
    result.wbName = Mid(temp, bracketStart + 1, bracketEnd - bracketStart - 1)
    
    ' Text between "]" and "!" is the sheet name (may end with a quote)
    sheetPart = Mid(temp, bracketEnd + 1, exclPos - bracketEnd - 1)
    If Right(sheetPart, 1) = "'" Then
        result.sheetName = Left(sheetPart, Len(sheetPart) - 1)
    Else
        result.sheetName = sheetPart
    End If
    
    ' Folder path = text before "["; strip wrapping quotes; append file name
    ' e.g. 'C:\t\[2.xlsx]sheet1'  ->  prefix 'C:\t\'  ->  C:\t\2.xlsx
    pathPrefix = Left(temp, bracketStart - 1)
    If Len(pathPrefix) > 0 And Left(pathPrefix, 1) = "'" Then
        pathPrefix = Mid(pathPrefix, 2)
    End If
    If Len(pathPrefix) > 0 And Right(pathPrefix, 1) = "'" Then
        pathPrefix = Left(pathPrefix, Len(pathPrefix) - 1)
    End If
    
    result.fullPath = pathPrefix
    If Len(result.fullPath) > 0 And Right(result.fullPath, 1) <> "\" Then
        result.fullPath = result.fullPath & "\"
    End If
    result.fullPath = result.fullPath & result.wbName
    
    ParseExternalReference = result

End Function


'------------------------------------------------------------------------------
' True when workbookSheets contains a tab whose .Name equals sheetName
' (worksheet or chart sheet; case-insensitive).
' Example: ElementInTableName ThisWorkbook.Sheets, "Import"
'------------------------------------------------------------------------------
Private Function ElementInTableName( _
    ByVal workbookSheets As Sheets, _
    ByVal sheetName As String) As Boolean

    Dim i As Long

    If Len(sheetName) = 0 Then Exit Function
    
    For i = 1 To workbookSheets.Count
        
        If LCase(workbookSheets(i).Name) = LCase(sheetName) Then
            ElementInTableName = True
            Exit Function
        End If
        
    Next i

End Function




'------------------------------------------------------------------------------
' Main entry point — run from a button or macro listt
' mode can be exportValue, exportFormula, dryRun
'------------------------------------------------------------------------------
Sub RunImportExport(ByVal mode As String)

    Dim wsImport As Worksheet
    Dim wsExport As Worksheet
    
    Dim c As Range
    Dim formulaText As String
    Dim extRef As ExternalRef
    
    Dim fullPath As String
    Dim wbName As String
    
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    
    Dim exportValue As Variant
    
    ' fullPath -> wbName (Scripting.Dictionary)
    Dim workbookList As Object
    
    ' Unique "fullPath" & Chr(1) & sheetName pairs to validate before export
    Dim targetWorkbookSheetKeys As Object
    
    Dim workbookDisplay As String
    
    Dim response As VbMsgBoxResult
    
    If Not ElementInTableName(ThisWorkbook.Sheets, "Import") _
        Or Not ElementInTableName(ThisWorkbook.Sheets, "Export") Then
        
        MsgBox "Missing required sheet(s): Import and/or Export.", vbCritical, "Missing Sheet"
        Exit Sub
        
    End If
    
    Set wsImport = ThisWorkbook.Sheets("Import")
    Set wsExport = ThisWorkbook.Sheets("Export")
    
    Set workbookList = CreateObject("Scripting.Dictionary")
    Set targetWorkbookSheetKeys = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    
    '==================================================
    ' STEP 1 — Scan Import sheet; verify files existt
    '==================================================
    ' Build a unique list of external workbooks before changing anything.
    
    For Each c In wsImport.UsedRange.Cells
    
        If c.HasFormula Then
        
            formulaText = c.Formula
            
            If IsExternalReference(formulaText) Then
            
                extRef = ParseExternalReference(formulaText)
                fullPath = extRef.fullPath
                wbName = extRef.wbName
                
                ' Dir() returns "" when the file is already open — check bothh
                If Dir(fullPath) = "" And Not WorkbookIsOpen(wbName) Then
                
                    MsgBox _
                        "File not found:" & vbCrLf & vbCrLf & _
                        fullPath, _
                        vbCritical
                    
                    GoTo CleanExit
                    
                End If
                
                ' Dictionary key = full path, value = workbook file name
                If Not workbookList.Exists(fullPath) Then
                    workbookList.Add fullPath, wbName
                End If
                
                If Not targetWorkbookSheetKeys.Exists(extRef.fullPath & Chr(1) & extRef.sheetName) Then
                    targetWorkbookSheetKeys.Add extRef.fullPath & Chr(1) & extRef.sheetName, True
                End If
                
            End If
            
        End If
        
    Next c
    
    If workbookList.Count = 0 Then
        MsgBox "No external links found on the Import sheet.", vbExclamation
        GoTo CleanExit
    End If
    
    '==================================================
    ' STEP 2 — Ask user to confirm before writingg
    '==================================================
    
    workbookDisplay = ""
    
    Dim key As Variant
    
    For Each key In workbookList.Keys
    
        ' Show full path for each target workbook
        workbookDisplay = workbookDisplay & key & vbCrLf & vbCrLf
        
    Next key
    
    response = MsgBox( _
        "The following workbooks will be updated:" & _
        vbCrLf & vbCrLf & _
        workbookDisplay & _
        "Continue?", _
        vbYesNo + vbQuestion, _
        "Confirm Update")
    
    If response <> vbYes Then
        GoTo CleanExit
    End If
    
    '==================================================
    ' STEP 3 — Backup Import and Export sheets (values only))
    '==================================================
    
    BackupSheetValuesOnly wsImport, "Import"
    BackupSheetValuesOnly wsExport, "Export"
    
    '==================================================
    ' STEP 4 — Open linked workbooks that are not open yett
    '==================================================
    ' key = full path, workbookList(key) = file name
    
    For Each key In workbookList.Keys
    
        wbName = workbookList(key)
        
        If Not WorkbookIsOpen(wbName) Then
            Workbooks.Open key
        End If
        
    Next key
    
    If Not AllTargetSheetsExist(workbookList, targetWorkbookSheetKeys) Then
        GoTo CleanExit
    End If
    
    '==================================================
    ' STEP 5 — Copy Export values into external cellss
    '==================================================
    ' Import and Export use the same row/column for each link pair.
    wsImport.Cells.Font.Underline = xlUnderlineStyleNone
    wsImport.Cells.Font.Underline = xlUnderlineStyleNone
    For Each c In wsImport.UsedRange.Cells
        
        If c.HasFormula Then
            
            formulaText = c.Formula
            
            If IsExternalReference(formulaText) Then
                c.Font.Underline = xlUnderlineStyleSingle
                wsExport.Cells(c.Row, c.Column).Font.Underline = xlUnderlineStyleSingle
                
                extRef = ParseExternalReference(formulaText)
                
                Set targetWb = WorkbookFromFullPath(extRef.fullPath, workbookList)
                Set targetWs = targetWb.Sheets(extRef.sheetName)
                
                If mode = "exportValue" Then ' wsExport.Cells(c.Row, c.Column).HasFormula Then
                
                
                    exportValue = wsExport.Cells(c.Row, c.Column).Value
                    targetWs.Range(extRef.CellAddress).Value = exportValue
                    
                End If
                
                If mode = "exportFormula" Then ' wsExport.Cells(c.Row, c.Column).HasFormula Then
                
                
                    exportValue = wsExport.Cells(c.Row, c.Column).Formula2
                    targetWs.Range(extRef.CellAddress).Formula2 = exportValue
                    
                End If
                
                'If mode = "dryRun" Then do nothing
            End If
            
        End If
        
    Next c
    
    MsgBox "Update completed.", vbInformation

CleanExit:
    Application.ScreenUpdating = True

End Sub

Sub RunImportExportValue()
    RunImportExport "exportValue"
End Sub
Sub RunImportExportFormula()
    RunImportExport "exportFormula"
End Sub
Sub DryRun()
    RunImportExport "dryRun"
End Sub

Sub ExtractLinksShort()
    Dim links As Variant
    links = ThisWorkbook.LinkSources(xlExcelLinks)
    
    Range("B2:B99").ClearContents
    For i = 1 To UBound(links)
        Range("B" & i + 1).Value = links(i)
        i = i + 1
    Next i
End Sub
Sub UpdateLinks()
    Dim i As Long

    For i = 1 To 99
        If .Range("A" & i).Value = "Yes" And .Range("C" & i).Value <> "" Then
            ThisWorkbook.ChangeLink .Range("B" & i).Value, .Range("C" & i).Value, xlExcelLinks
        End If
    Next i

End Sub

