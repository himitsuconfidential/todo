' macro in this page:
' Part1: xopen, xclose, xsaveclose, xsaveas
' Part2: xrange, xrangeSelect, xrangeSet, xcopy, xpaste, xpasteA/F/T/V/W/L  <- meaning same as ctrl+alt+v shortcut
' Part3: xgoalseek, xreplace, xchangeLink, xchangeLinkOpen, xrun, xlog


' ------------------------------------------------------------------
'  Part1: xopen, xclose, xsaveclose, xsaveas
'  File control
' ------------------------------------------------------------------

Sub xopen(WBpath As String)
    Workbooks.Open (WBpath)
End Sub

Sub xclose(WBpath As String, Optional SaveChanges As Boolean = False)
    Workbooks(GetFileNameOnly(WBpath)).Close SaveChanges:=SaveChanges
End Sub

Sub xsaveclose(WBpath As String)
    xclose WBpath, True
End Sub

Function GetFileNameOnly(fullPath As String) As String
    Dim parts() As String
    parts = Split(Replace(fullPath, "/", "\"), "\")
    If UBound(parts) >= 0 Then
        GetFileNameOnly = parts(UBound(parts))
    End If
End Function
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'  Part2: xrange, xrangeSelect, xrangeSet, xcopy, xpaste, xpasteA/F/T/V/W/L  <- meaning same as ctrl+alt+v shortcut
'  Basic Excel
'  Core range parser ─ uses ^ v < >  and  ^^ vv << >>, e.g. A1>> simulate selected A1 then ctrl+right, A1:A9&vv simulate selected A1:A9 then ctrl+shift+down
'  Format: [address][& or nothing][single arrow][number] or [address][& or nothing][double arrow]
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'use ^v>< to extend range
Function xrange(Optional expr As String = "", Optional Sheetname As String = "", Optional WBname As String = "") As Range
    Dim wb As Workbook, ws As Worksheet
    Dim parts() As String
    Dim arr() As String
    Dim baseCell As Range, targetCell As Range
    If expr = "" Then
        Dim s As Range
        Set s = Selection
        Set xrange = Selection
        Exit Function
    End If
    ' Decide workbook
    If WBname = "" Then
        Set wb = ThisWorkbook
    Else
        WBname = GetFileNameOnly(WBname)
        Set wb = Workbooks(WBname)
    End If
    
    ' Decide sheet
    If Sheetname = "" Then
        Set ws = wb.ActiveSheet
    Else
        Set ws = wb.Sheets(Sheetname)
    End If
    
    ' Split by "&" if range extension
    If InStr(expr, "&") > 0 Then
        parts = Split(expr, "&")
        Set baseCell = ParseExpr(parts(0), ws)
        Set targetCell = ParseExpr(parts(0) & parts(1), ws)
        Set xrange = ws.Range(baseCell, targetCell)
    Else
        Set xrange = ParseExpr(expr, ws)
    End If
End Function
'use ^v>< to extend range
Function ParseExpr(expr As String, ws As Worksheet) As Range
    Dim addr As String, suffix As String
    Dim i As Long, ch As String
    Dim baseCell As Range
    Dim offsetCount As Long
    
    ' Extract base address (letters+digits until arrow symbols)
    i = 1
    Do While i <= Len(expr)
        ch = Mid(expr, i, 1)
        If ch = "^" Or ch = "v" Or ch = "<" Or ch = ">" Then Exit Do
        i = i + 1
    Loop
    addr = Left(expr, i - 1)
    If addr = "" Then addr = Selection.address
    Set baseCell = ws.Range(addr)
    suffix = Mid(expr, i)
    ' Handle suffix
    If suffix <> "" Then
        correct_format_help = "Format: [address][& or nothing][single arrow][number] or [address][& or nothing][double arrow]" & vbCrLf & "single arrow: ^ v < >" & vbCrLf & "double arrow:  ^^ vv << >>"
        ' Double arrows (Ctrl+Arrow)
        If Left(suffix, 2) = "^^" Or Left(suffix, 2) = "vv" Or Left(suffix, 2) = "<<" Or Left(suffix, 2) = ">>" Then
            If Len(suffix) > 2 Then MsgBox "Warning: Wrong address format " & expr & ctrl & correct_format_help
            Select Case Left(suffix, 2)
                Case "^^": Set baseCell = baseCell.End(xlUp)
                Case "vv": Set baseCell = baseCell.End(xlDown)
                Case "<<": Set baseCell = baseCell.End(xlToLeft)
                Case ">>": Set baseCell = baseCell.End(xlToRight)
            End Select
        
        ' Single arrows
        Else
            If Len(suffix) > 1 Then
                If IsNumeric(Mid(suffix, 2)) = False Then MsgBox "Warning: Wrong address format " & expr & vbCrLf & correct_format_help
                offsetCount = Val(Mid(suffix, 2))
            Else
                offsetCount = 1
            End If
            Select Case Left(suffix, 1)
                Case "^": Set baseCell = baseCell.Offset(-offsetCount, 0)
                Case "v": Set baseCell = baseCell.Offset(offsetCount, 0)
                Case "<": Set baseCell = baseCell.Offset(0, -offsetCount)
                Case ">": Set baseCell = baseCell.Offset(0, offsetCount)
            End Select
        End If
    End If
    Set ParseExpr = baseCell
End Function
'use ^v>< to extend range
' Paste normally into target range
Sub xrangeSelect(address As String, Optional Sheetname As String = "", Optional WBname As String = "")
    If WBname = "" Then
        ThisWorkbook.Activate
    Else
        WBname = GetFileNameOnly(WBname)
        Workbooks(WBname).Activate
    End If
    If Sheetname <> "" Then ActiveWorkbook.Sheets(Sheetname).Select
    xrange(address, Sheetname, WBname).Select
End Sub
'use ^v>< to extend range
' Paste normally into target range
Sub xrangeSet(Optional address As String = "", Optional Sheetname As String = "", Optional WBname As String = "", Optional formula As String = "")
    If formula = "" Then
        xrange(address, Sheetname, WBname).ClearContents
    Else
        xrange(address, Sheetname, WBname).Formula2 = formula
    End If
End Sub
'use ^v>< to extend range
' Copy a range to clipboard
Sub xcopy(Optional address As String = "", Optional Sheetname As String = "", Optional WBname As String = "")
    xrange(address, Sheetname, WBname).Copy
End Sub
'use ^v>< to extend range
' Paste all (values, formats, formulas) into target range
Sub xpasteA(Optional address As String = "", Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteAll, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^v>< to extend range
' Paste formula (formulas) into target range
Sub xpasteF(Optional address As String = "", Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteFormulas, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^v>< to extend range
' Paste value (values) into target range
Sub xpasteV(Optional address As String = "", Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteValues, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^v>< to extend range
' Paste format (formats) into target range
Sub xpasteT(Optional address As String = "", Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteFormats, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^v>< to extend range
' Paste column width (width) into target range
Sub xpasteW(Optional address As String = "", Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteColumnWidths, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^v>< to extend range
' Paste link (link) into target range
Sub xpasteL(Optional address As String = "", Optional Sheetname As String = "", Optional WBname As String = "")
    Call xrangeSelect(address, Sheetname, WBname)
    ActiveSheet.Paste Link:=True
End Sub
'use ^v>< to extend range
' Paste normally into target range
Sub xpaste(Optional address As String = "", Optional Sheetname As String = "", Optional WBname As String = "")
    Call xrangeSelect(address, Sheetname, WBname)
    ActiveSheet.Paste
End Sub
Sub xsaveas(filepath As String)
    ChDir ActiveWorkbook.Path
    ActiveWorkbook.SaveAs filepath
End Sub
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
'  Part3: xgoalseek, xreplace, xchangeLink, xchangeLinkOpen, xrun, xlog
'  Less frequently used
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub xgoalseek(goal, address1 As String, Sheetname1 As String, WBname1 As String, address2 As String, Sheetname2 As String, WBname2 As String)
    'address to is goal, address2 is changing cell
    xrange(address1, Sheetname1, WBname1).GoalSeek goal:=0, ChangingCell:=xrange(address2, Sheetname2, WBname2)
End Sub

Function textreplace(context As String, ParamArray args() As Variant) As String
    Dim dict As Object
    Dim i As Long
    Dim result As String
    Dim key As Variant
    
    ' Create dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Validate argument count (must be odd: 3,5,7,...)

    If UBound(args) < 1 Or (UBound(args) + 1) Mod 2 <> 0 Then
        textreplace = CVErr(xlErrValue)
        Exit Function
    End If
    
    ' Build dictionary from pairs
    For i = 0 To UBound(args) Step 2
        dict(args(i)) = args(i + 1)
    Next i
    
    ' Start with cell value
    result = context
    
    ' Apply replacements
    For Each key In dict.Keys
        result = Replace(result, key, dict(key))
    Next key
    
    ' Return modified string
    textreplace = result
End Function
Sub xreplace(targetRange As Range, ParamArray args() As Variant)
    Dim dict As Object
    Dim i As Long
    Dim result As String
    Dim key As Variant
    For Each cell In targetRange
        ' Create dictionary
        Set dict = CreateObject("Scripting.Dictionary")
        
        ' Validate argument count (must be odd: 3,5,7,...)

        If UBound(args) < 1 Or (UBound(args) + 1) Mod 2 <> 0 Then
            MsgBox "error: argument count must 3,5,7... but got " & UBound(args)
            Exit Sub
        End If
        
        ' Build dictionary from pairs
        For i = 0 To UBound(args) Step 2
            dict(args(i)) = args(i + 1)
        Next i
        
        ' Start with cell value
        result = cell.Formula2
        
        ' Apply replacements
        For Each key In dict.Keys
            result = Replace(result, key, dict(key))
        Next key
        cell.Formula2 = result
    Next cell
End Sub

Sub xchangeLink(WBname As String, oldWBfullname As String, newWBfullname As String)
    
    If WBname = "" Then
        ThisWorkbook.Activate
    Else
        WBname = GetFileNameOnly(WBname)
        Workbooks(WBname).Activate
    End If
    
    ActiveWorkbook.ChangeLink oldWBfullname, newWBfullname
End Sub

Sub xchangeLinkOpen(WBname As String, oldWBfullname As String, newWBfullname As String)
    xopen newWBfullname
    xchangeLink WBname, oldWBfullname, newWBfullname
    xclose newWBfullname
End Sub

Sub xrun(macroname As String, Optional Sheetname As String = "", Optional WBname As String = "")
    If WBname = "" Then
        ThisWorkbook.Activate
    Else
        arr = Split(WBname, "\")
        WBname = arr(UBound(arr))
        Workbooks(WBname).Activate
    End If
    If Sheetname <> "" Then ActiveWorkbook.Sheets(Sheetname).Select
    
    Application.Run WBname & "!" & macroname
End Sub

Sub xlog()
    Dim address, Sheetname, WBname, theFormula As String
    Dim rng As Range
    Dim q As String
    q = """"
    address = Selection.address
    Sheetname = Selection.Parent.name
    WBname = Selection.Parent.Parent.FullName   ' Use .Name, not .FullName
    
    Dim firstFormula As String
    Dim cell As Range
    
    
    ' If only one cell
    If Selection.Count = 1 Then
        theFormula = Selection.Formula2
        
    Else
    
        ' Multiple cells them check consistency
        firstFormula = Selection.Cells(1, 1).Formula2R1C1
        For Each cell In Selection
            If cell.Formula2 <> firstFormula Then
                MsgBox "Error: Selected cells do not all have the same formula.", vbCritical
                Exit Sub
            End If
        Next cell
        
        ' All formulas match then return first
        theFormula = Selection.Cells(1, 1).Formula2

    End If
    ' sepcial treatment for double quote
    theFormula = Replace(theFormula, q, q & q)
    Debug.Print Join(Array("xrangeSet ", q, address, q, ",", q, Sheetname, q, ",", q, WBname, q, ",", q, theFormula, q), "")
End Sub
