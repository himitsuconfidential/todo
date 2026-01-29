Sub xopen(WBpath As String)
    Set wb = Workbooks.Open(WBpath)
End Sub

Sub xclose(WBname As String)
    arr = Split(WBname, "\")
    WBname = arr(UBound(arr))
    Workbooks(WBname).Close Savechanges:=False
End Sub
Sub xsaveclose(WBname As String)
    arr = Split(WBname, "\")
    WBname = arr(UBound(arr))
    Workbooks(WBname).Close Savechanges:=True
End Sub
'use ^ˇ>< to extend range
Function xrange(expr As String, Optional Sheetname As String = "", Optional WBname As String = "") As Range
    Dim wb As Workbook, ws As Worksheet
    Dim parts() As String
    Dim arr() As String
    Dim baseCell As Range, targetCell As Range
    
    ' Decide workbook
    If WBname = "" Then
        Set wb = ThisWorkbook
    Else
        arr = Split(WBname, "\")
        WBname = arr(UBound(arr))
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
'use ^ˇ>< to extend range
Function ParseExpr(expr As String, ws As Worksheet) As Range
    Dim addr As String, suffix As String
    Dim i As Long, ch As String
    Dim baseCell As Range
    Dim offsetCount As Long
    
    ' Extract base address (letters+digits until arrow symbols)
    i = 1
    Do While i <= Len(expr)
        ch = Mid(expr, i, 1)
        If ch = "^" Or ch = "ˇ" Or ch = "<" Or ch = ">" Then Exit Do
        i = i + 1
    Loop
    addr = Left(expr, i - 1)
    If addr = "" Then addr = Selection.address
    Set baseCell = ws.Range(addr)
    suffix = Mid(expr, i)
    ' Handle suffix
    If suffix <> "" Then
        
        ' Double arrows (Ctrl+Arrow)
        If Left(suffix, 2) = "^^" Or Left(suffix, 2) = "ˇˇ" Or Left(suffix, 2) = "<<" Or Left(suffix, 2) = ">>" Then
            offsetCount = IIf(Len(suffix) > 2, Val(Mid(suffix, 3)), 0)
            Select Case Left(suffix, 2)
                Case "^^": Set baseCell = baseCell.End(xlUp): If offsetCount <> 0 Then Set baseCell = baseCell.Offset(-offsetCount, 0)
                Case "ˇˇ": Set baseCell = baseCell.End(xlDown): If offsetCount <> 0 Then Set baseCell = baseCell.Offset(offsetCount, 0)
                Case "<<": Set baseCell = baseCell.End(xlToLeft): If offsetCount <> 0 Then Set baseCell = baseCell.Offset(0, -offsetCount)
                Case ">>": Set baseCell = baseCell.End(xlToRight): If offsetCount <> 0 Then Set baseCell = baseCell.Offset(0, offsetCount)
            End Select
        
        ' Single arrows
        Else
            offsetCount = IIf(Len(suffix) > 1, Val(Mid(suffix, 2)), 1)
            Select Case Left(suffix, 1)
                Case "^": Set baseCell = baseCell.Offset(-offsetCount, 0)
                Case "ˇ": Set baseCell = baseCell.Offset(offsetCount, 0)
                Case "<": Set baseCell = baseCell.Offset(0, -offsetCount)
                Case ">": Set baseCell = baseCell.Offset(0, offsetCount)
            End Select
        End If
    End If
    Set ParseExpr = baseCell
End Function
'use ^ˇ>< to extend range
Function Extendrange(baseCell As Range, suffix As String) As Range
    If suffix <> "" Then
        
        ' Double arrows (Ctrl+Arrow)
        If Left(suffix, 2) = "^^" Or Left(suffix, 2) = "ˇˇ" Or Left(suffix, 2) = "<<" Or Left(suffix, 2) = ">>" Then
            offsetCount = IIf(Len(suffix) > 2, Val(Mid(suffix, 3)), 0)
            Select Case Left(suffix, 2)
                Case "^^": Set baseCell = baseCell.End(xlUp): If offsetCount <> 0 Then Set baseCell = baseCell.Offset(-offsetCount, 0)
                Case "ˇˇ": Set baseCell = baseCell.End(xlDown): If offsetCount <> 0 Then Set baseCell = baseCell.Offset(offsetCount, 0)
                Case "<<": Set baseCell = baseCell.End(xlToLeft): If offsetCount <> 0 Then Set baseCell = baseCell.Offset(0, -offsetCount)
                Case ">>": Set baseCell = baseCell.End(xlToRight): If offsetCount <> 0 Then Set baseCell = baseCell.Offset(0, offsetCount)
            End Select
        
        ' Single arrows
        Else
            offsetCount = IIf(Len(suffix) > 1, Val(Mid(suffix, 2)), 1)
            Select Case Left(suffix, 1)
                Case "^": Set baseCell = baseCell.Offset(-offsetCount, 0)
                Case "ˇ": Set baseCell = baseCell.Offset(offsetCount, 0)
                Case "<": Set baseCell = baseCell.Offset(0, -offsetCount)
                Case ">": Set baseCell = baseCell.Offset(0, offsetCount)
            End Select
        End If
    End If
    Set Extendrange = baseCell
End Function
'use ^ˇ>< to extend range
' Copy a range to clipboard
Sub xcopy(address As String, Optional Sheetname As String = "", Optional WBname As String = "")
    xrange(address, Sheetname, WBname).Copy
End Sub
'use ^ˇ>< to extend range
' Paste all (values, formats, formulas) into target range
Sub xpasteA(address As String, Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteAll, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^ˇ>< to extend range
' Paste formula (formulas) into target range
Sub xpasteF(address As String, Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteFormulas, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^ˇ>< to extend range
' Paste value (values) into target range
Sub xpasteV(address As String, Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteValues, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^ˇ>< to extend range
' Paste format (formats) into target range
Sub xpasteT(address As String, Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteFormats, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^ˇ>< to extend range
' Paste column width (width) into target range
Sub xpasteW(address As String, Optional Sheetname As String = "", Optional WBname As String = "", Optional Transpose As Variant = False, Optional SkipBlanks As Variant = False)
    xrange(address, Sheetname, WBname).PasteSpecial Paste:=xlPasteColumnWidths, SkipBlanks:=SkipBlanks, Transpose:=Transpose
End Sub
'use ^ˇ>< to extend range
' Paste link (link) into target range
Sub xpasteL(address As String, Optional Sheetname As String = "", Optional WBname As String = "")
    If WBname = "" Then
        ThisWorkbook.Activate
    Else
        arr = Split(WBname, "\")
        WBname = arr(UBound(arr))
        Workbooks(WBname).Activate
    End If
    xrange(address, Sheetname, WBname).Select
    ActiveSheet.Paste Link:=True
End Sub
'use ^ˇ>< to extend range
' Paste normally into target range
Sub xpaste(address As String, Optional Sheetname As String = "", Optional WBname As String = "")
    If WBname = "" Then
        ThisWorkbook.Activate
    Else
        arr = Split(WBname, "\")
        WBname = arr(UBound(arr))
        Workbooks(WBname).Activate
    End If
    xrange(address, Sheetname, WBname).Select
    ActiveSheet.Paste
End Sub
Sub xsaveas(filepath As String)
    ChDir ActiveWorkbook.Path
    ActiveWorkbook.SaveAs filepath
End Sub

