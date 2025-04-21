Option Explicit

Sub 年計表()
    ' Purpose: Create formatted annual summary table of overtime and transportation expenses
    
    Dim wsMain As Worksheet
    Dim lastRow As Long, i As Long, j As Long, rowIndex As Long
    Dim totalRowText As String
    Dim sheetCount As Integer
    
    Application.ScreenUpdating = False
    
    Set wsMain = Sheets("年計表")
    sheetCount = Worksheets.Count
    
    ' Format header
    With wsMain.Range("C1:I1")
        .MergeCells = True
        .Value = "車資夜點加班費年計表"
        .Font.Name = "微軟正黑體"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(154, 0, 54)
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
    End With
    
    ' Set column headers
    With wsMain
        .Range("D2") = "加班費"
        .Range("E2") = "夜點費"
        .Range("F2") = "小計"
        .Range("G2") = "交通補助費"
        .Range("H2") = "總計"
        .Range("I2") = "MoM"
        
        ' Format column headers
        .Range("C2:I2").Font.Color = RGB(154, 0, 54)
        .Range("C2:I2").Font.Name = "微軟正黑體"
        .Range("C2:I2").Font.Bold = True
        
        ' Column width adjustments
        .Columns("B").ColumnWidth = 5
        .Columns("B").ShrinkToFit = True
        .Columns("F").ColumnWidth = 13
    End With
    
    ' Populate data from other sheets
    For i = 2 To sheetCount
        rowIndex = i + 1
        wsMain.Range("B" & rowIndex) = i - 1
        wsMain.Range("C" & rowIndex) = Sheets(i).Range("N1").Value
        
        ' Find totals in source sheet
        For j = 40 To 50
            totalRowText = Sheets(i).Range("A" & j).Value
            If totalRowText = "總計" Or totalRowText = "總   計" Then
                With wsMain
                    .Range("D" & rowIndex) = Sheets(i).Range("D" & j).Value
                    .Range("E" & rowIndex) = Sheets(i).Range("E" & j).Value
                    .Range("F" & rowIndex) = Sheets(i).Range("F" & j).Value
                    .Range("G" & rowIndex) = Sheets(i).Range("G" & j).Value
                    .Range("H" & rowIndex) = Sheets(i).Range("I" & j).Value
                End With
                Exit For
            End If
        Next j
    Next i
    
    ' Alternate row coloring
    lastRow = wsMain.Cells(wsMain.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRow Step 2
        wsMain.Range("C" & i & ":I" & i).Interior.Color = RGB(255, 201, 220)
    Next i
    
    ' Calculate Month-over-Month percentages
    With wsMain
        For i = 4 To sheetCount + 1
            If .Range("H" & (i - 1)).Value <> 0 Then
                .Range("I" & i).Value = (.Range("H" & i).Value - .Range("H" & (i - 1)).Value) / .Range("H" & (i - 1)).Value
                .Range("I" & i).NumberFormat = "0.00%"
            End If
        Next i
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub 總合年計明細()
    ' Purpose: Create detailed yearly summary with lookups from all sheets
    
    Dim wsMain As Worksheet
    Dim lookupRange As Range, resultRange As Range
    Dim sheetCount As Integer, lastRow As Long
    Dim i As Long, j As Long, sourceColumn As Long
    Dim cellValue As Variant
    
    Application.ScreenUpdating = False
    
    Set wsMain = Sheets("年計表")
    sheetCount = Worksheets.Count
    
    ' Clear and set up the header area
    wsMain.Cells(2, 12) = "車資夜點加班費合計"
    wsMain.Range("M2:AA150").ClearContents
    
    ' Format the header row
    With wsMain.Range(wsMain.Cells(2, "L"), wsMain.Cells(2, sheetCount + 11))
        .Interior.Color = RGB(154, 0, 54)
        .Font.Color = RGB(225, 225, 225)
    End With
    
    ' Add the "Total" column header
    With wsMain.Cells(2, sheetCount + 12)
        .Value = "Total"
        .Font.Color = RGB(154, 0, 54)
    End With
    
    ' Set sheet names as column headers and format
    For i = 2 To sheetCount
        sourceColumn = i + 11
        wsMain.Cells(2, sourceColumn) = Sheets(i).Name
        wsMain.Cells(2, sourceColumn).NumberFormat = "0000"
    Next i
    
    ' Process data rows
    lastRow = 100 ' Adjust if needed to actual data range
    For i = 3 To lastRow
        cellValue = wsMain.Range("L" & i).Value
        If cellValue = "" Then
            wsMain.Cells(i, sheetCount + 12) = ""
            Exit For
        End If
        
        ' Populate data from each sheet
        For j = 2 To sheetCount
            sourceColumn = j + 11
            Set lookupRange = Sheets(j).Range("U4:AK89")
            
            On Error Resume Next
            wsMain.Cells(i, sourceColumn) = Application.WorksheetFunction.VLookup(cellValue, lookupRange, 17, False)
            On Error GoTo 0
        Next j
        
        ' Calculate row total
        Set resultRange = wsMain.Range(wsMain.Cells(i, 13), wsMain.Cells(i, sheetCount + 11))
        wsMain.Cells(i, sheetCount + 12) = Application.WorksheetFunction.Sum(resultRange)
        wsMain.Cells(i, sheetCount + 12).Font.Color = RGB(225, 0, 0)
    Next i
    
    ' Apply alternate row formatting
    For i = 4 To lastRow Step 2
        If wsMain.Range("L" & i).Value = "" Then Exit For
        wsMain.Range(wsMain.Cells(i, "L"), wsMain.Cells(i, sheetCount + 12)).Interior.Color = RGB(194, 211, 232)
    Next i
    
    Application.ScreenUpdating = True
End Sub
