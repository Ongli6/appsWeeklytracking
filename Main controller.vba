' --- Refresh all data connections first ---
Sub RefreshAllDataConnections()
    Dim conn As WorkbookConnection
    Dim isRefreshing As Boolean

    ThisWorkbook.RefreshAll

    ' Wait until all connections finish refreshing
    Do
        isRefreshing = False
        For Each conn In ThisWorkbook.Connections
            ' Check if this connection supports background refresh and if it's still refreshing
            On Error Resume Next
            If conn.Type = xlConnectionTypeOLEDB Or conn.Type = xlConnectionTypeODBC Then
                If conn.OLEDBConnection.BackgroundQuery Then
                    If conn.OLEDBConnection.Refreshing Then
                        isRefreshing = True
                        Exit For
                    End If
                End If
            ElseIf conn.Type = xlConnectionTypeWORKSHEET Then
                ' You can add checks for other connection types if needed
            End If
            On Error GoTo 0
        Next conn
        DoEvents
        Application.Wait Now + TimeValue("0:00:01") ' short pause to avoid hogging CPU
    Loop While isRefreshing
End Sub

' --- 6-block copy macro for multiple campus sheets ---
Sub CopyBlocksAllCampuses()
    ' Refresh data first
    RefreshAllDataConnections

    Dim campusSheets As Variant
    campusSheets = Array("Berkshire 16-18", "Berkshire 19+", "Oxfordshire 16-18", "Oxfordshire 19+", "Surrey 16-18", "Surrey 19+")

    Dim i As Long
    For i = LBound(campusSheets) To UBound(campusSheets)
        CopyBlockWithSelectiveFormulasForSheet CStr(campusSheets(i))
    Next i

    MsgBox "All campus sheets updated!", vbInformation
End Sub

Sub CopyBlockWithSelectiveFormulasForSheet(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim headerRow1 As Long: headerRow1 = 1
    Dim headerRow2 As Long: headerRow2 = 2
    Dim headerRow3 As Long: headerRow3 = 3
    Dim blockStartCol As Long: blockStartCol = 3 ' Column C
    Dim blockWidth As Long: blockWidth = 6

    ' Find the next empty block
    Dim col As Long: col = blockStartCol
    Do While ws.Cells(headerRow2, col).Value <> ""
        col = col + blockWidth
    Loop

    Dim srcCol As Long: srcCol = col - blockWidth
    Dim tgtCol As Long: tgtCol = col

    ' Find the total row by looking for "Total" in column A
    Dim totalRow As Long
    totalRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Do While totalRow > 1 And ws.Cells(totalRow, 1).Value <> "Total"
        totalRow = totalRow - 1
    Loop

    ' 1. Copy formatting and headers (including rows 1–3)
    Dim lastDataRow As Long
    lastDataRow = totalRow - 1

    Dim srcRange As Range, tgtRange As Range
    Set srcRange = ws.Range(ws.Cells(headerRow1, srcCol), ws.Cells(totalRow, srcCol + blockWidth - 1))
    Set tgtRange = ws.Cells(headerRow1, tgtCol)
    srcRange.Copy
    tgtRange.PasteSpecial Paste:=xlPasteFormats
    tgtRange.PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False

    ' 2. Merge row 1 and 2 in target block and fill labels
    With ws.Range(ws.Cells(headerRow1, tgtCol), ws.Cells(headerRow1, tgtCol + blockWidth - 1))
        .Merge
        .Value = "Date"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(headerRow2, tgtCol), ws.Cells(headerRow2, tgtCol + blockWidth - 1))
        .Merge
        .Value = Date
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 3. Copy row 3 headings (not merged)
    ws.Range(ws.Cells(headerRow3, srcCol), ws.Cells(headerRow3, srcCol + blockWidth - 1)).Copy
    ws.Cells(headerRow3, tgtCol).PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    ' 4. Copy formulas and clear contents (except difference cols)
    Dim r As Long, c As Long
    For c = 0 To blockWidth - 1
        For r = headerRow3 + 1 To lastDataRow
            Select Case c + 1
                Case 3 ' third column in block (difference formula)
                    ws.Cells(r, tgtCol + c).FormulaR1C1 = "=RC[-2]-RC[-1]"
                Case 6 ' sixth column in block (difference formula)
                    ws.Cells(r, tgtCol + c).FormulaR1C1 = "=RC[-2]-RC[-1]"
                Case Else
                    ws.Cells(r, tgtCol + c).formula = ws.Cells(r, srcCol + c).formula
            End Select
        Next r
    Next c

    ' ** Replace absolute date references in formulas to point to new block date cell **
    Dim originalDateCell As String
    Dim newDateCell As String
    Dim formulaText As String

    originalDateCell = ws.Cells(headerRow2, srcCol).Address(True, True)
    newDateCell = ws.Cells(headerRow2, tgtCol).Address(True, True)

    For c = 0 To blockWidth - 1
        For r = headerRow3 + 1 To lastDataRow
            With ws.Cells(r, tgtCol + c)
                If .HasFormula Then
                    formulaText = .formula
                    formulaText = Replace(formulaText, originalDateCell, newDateCell)
                    .formula = formulaText
                End If
            End With
        Next r
    Next c

    ' 5. Copy formulas only from the TOTAL row across all columns
    For c = 0 To blockWidth - 1
    ws.Cells(totalRow, tgtCol + c).FormulaR1C1 = ws.Cells(totalRow, srcCol + c).FormulaR1C1
    Next c

    ' 6. Replace formulas with values in new block except col 3, 6, and total row
    For r = headerRow3 + 1 To lastDataRow
        If r <> totalRow Then
            For c = 0 To blockWidth - 1
                Select Case c + 1
                    Case 3, 6 ' Skip difference columns
                    Case Else
                        ws.Cells(r, tgtCol + c).Value = ws.Cells(r, tgtCol + c).Value
                End Select
            Next c
        End If
    Next r
End Sub



' --- 8-block copy macro for weekly report sheets ---
Sub CopyBlocksWeeklyReports()
    ' Refresh data first
    RefreshAllDataConnections

    Dim reportSheets As Variant
    reportSheets = Array("Weekly Report 16-18", "Weekly Report 19+")

    Dim i As Long
    For i = LBound(reportSheets) To UBound(reportSheets)
        CopyBlockWithSelectiveFormulasForSheet_8block ThisWorkbook.Sheets(reportSheets(i))
    Next i

    MsgBox "Weekly Report sheets updated!", vbInformation
End Sub

Sub CopyBlockWithSelectiveFormulasForSheet_8block(ws As Worksheet)
    ' Declare all variables once here
    Dim headerRow1 As Long: headerRow1 = 1
    Dim headerRow2 As Long: headerRow2 = 2
    Dim headerRow3 As Long: headerRow3 = 3
    Dim dataStartRow As Long: dataStartRow = 4
    Dim blockWidth As Long: blockWidth = 8

    Dim r As Long, c As Long
    Dim srcCell As Range, tgtCell As Range
    Dim startCol As Long: startCol = 3 ' Column C
    Dim tgtCol As Long, srcCol As Long

    Dim totalRow As Long
    Dim totalCell As Range
    Dim lastRow As Long

    Dim oldDateRef As String, newDateRef As String
    Dim f As String
    Dim srcColLetter As String, tgtColLetter As String

    ' Find the next empty block
    tgtCol = startCol
    Do While ws.Cells(headerRow2, tgtCol).Value <> ""
        tgtCol = tgtCol + blockWidth
    Loop
    srcCol = tgtCol - blockWidth

    ' Find total row by searching "Totals" in column A
    Set totalCell = ws.Columns(1).Find("Totals", LookIn:=xlValues, LookAt:=xlWhole)
    If Not totalCell Is Nothing Then
        totalRow = totalCell.Row
    Else
        MsgBox "Could not find 'Totals' in column A on sheet '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    lastRow = totalRow - 1

    ' Copy formatting from header rows in the source block
    With ws.Range(ws.Cells(headerRow1, srcCol), ws.Cells(headerRow1, srcCol + blockWidth - 1))
        .Copy
        ws.Cells(headerRow1, tgtCol).PasteSpecial Paste:=xlPasteFormats
        ws.Cells(headerRow1, tgtCol).PasteSpecial Paste:=xlPasteColumnWidths
    End With

    With ws.Range(ws.Cells(headerRow2, srcCol), ws.Cells(headerRow2, srcCol + blockWidth - 1))
        .Copy
        ws.Cells(headerRow2, tgtCol).PasteSpecial Paste:=xlPasteFormats
        ws.Cells(headerRow2, tgtCol).PasteSpecial Paste:=xlPasteColumnWidths
    End With
    Application.CutCopyMode = False

    ' Re-merge and assign header values
    With ws.Range(ws.Cells(headerRow1, tgtCol), ws.Cells(headerRow1, tgtCol + blockWidth - 1))
        .Merge
        .Value = "Date"
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(headerRow2, tgtCol), ws.Cells(headerRow2, tgtCol + blockWidth - 1))
        .Merge
        .Value = Date
        .HorizontalAlignment = xlCenter
    End With

    ' Copy formatting for header row 1 again (just to be sure)
    ws.Range(ws.Cells(headerRow1, srcCol), ws.Cells(headerRow1, srcCol + blockWidth - 1)).Copy
    ws.Cells(headerRow1, tgtCol).PasteSpecial Paste:=xlPasteFormats

    ' Copy formatting for totals row
    ws.Range(ws.Cells(totalRow, srcCol), ws.Cells(totalRow, srcCol + blockWidth - 1)).Copy
    ws.Cells(totalRow, tgtCol).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Copy row 3 headings (values and number formats)
    ws.Range(ws.Cells(headerRow3, srcCol), ws.Cells(headerRow3, srcCol + blockWidth - 1)).Copy
    ws.Cells(headerRow3, tgtCol).PasteSpecial Paste:=xlPasteValuesAndNumberFormats

    ' Copy formatting for data rows
    ws.Range(ws.Cells(headerRow3, srcCol), ws.Cells(lastRow, srcCol + blockWidth - 1)).Copy
    ws.Cells(headerRow3, tgtCol).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Prepare old/new date references for formula adjustments
    oldDateRef = "'" & ws.Name & "'!" & ws.Cells(headerRow2, srcCol).Address(ReferenceStyle:=xlA1)
    newDateRef = "'" & ws.Name & "'!" & ws.Cells(headerRow2, tgtCol).Address(ReferenceStyle:=xlA1)

    ' Get column letters for source and target (for absolute references)
    srcColLetter = Split(ws.Cells(1, srcCol).Address(True, False), "$")(0)
    tgtColLetter = Split(ws.Cells(1, tgtCol).Address(True, False), "$")(0)

    ' Copy formulas row by row with special handling
    For r = dataStartRow To lastRow
        For c = 0 To blockWidth - 1
            Set tgtCell = ws.Cells(r, tgtCol + c)
            Set srcCell = ws.Cells(r, srcCol + c)

            Select Case c + 1
                Case 3, 6 ' Difference columns with fixed formula
                    tgtCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
                Case 7, 8 ' Merged columns with formulas or values
                    If srcCell.MergeCells Then tgtCell.MergeArea.Merge
                    If srcCell.HasFormula Then
                        tgtCell.FormulaR1C1 = srcCell.FormulaR1C1
                    Else
                        tgtCell.Value = srcCell.Value
                    End If
                Case Else ' Other columns - adjust formula references
                    If srcCell.HasFormula Then
                        f = srcCell.formula
                        f = Replace(f, oldDateRef, newDateRef, , , vbTextCompare)
                        f = Replace(f, "$" & srcColLetter & "$2", "$" & tgtColLetter & "$2", , , vbTextCompare)
                        tgtCell.formula = f
                    Else
                        tgtCell.Value = srcCell.Value
                    End If
            End Select
        Next c
    Next r

    ' Copy totals row formulas/values preserving relative references
    For c = 0 To blockWidth - 1
        Set srcCell = ws.Cells(totalRow, srcCol + c)
        Set tgtCell = ws.Cells(totalRow, tgtCol + c)

        If srcCell.HasFormula Then
            tgtCell.FormulaR1C1 = srcCell.FormulaR1C1
        Else
            tgtCell.Value = srcCell.Value
        End If
    Next c

    ' Convert all non-special columns' formulas to values (excluding cols 3,6,7,8)
    For r = dataStartRow To lastRow
        If r <> totalRow Then
            For c = 0 To blockWidth - 1
                Select Case c + 1
                    Case 3, 6, 7, 8
                        ' Skip these special columns
                    Case Else
                        With ws.Cells(r, tgtCol + c)
                            If .HasFormula Then .Value = .Value
                        End With
                End Select
            Next c
        End If
    Next r
End Sub

