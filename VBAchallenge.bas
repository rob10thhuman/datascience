Attribute VB_Name = "Module1"
Sub testWorkboox()

    For Each ws In Worksheets
    
    'Declare variables
    Dim tickerSym As String
    Dim i, sumData As Integer
    Dim lastRow, endRow As Long
    Dim vol, totalVol, closeVal, startVal As Double
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Find last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'Start the summative data
    sumData = 2
    
    'Set field types (%)
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Columns("J").NumberFormat = "0.00"
    
    'startVal is pulling the last value of the ticker
    'this is an attempt to find the correct value
    'startVal = ws.Cells(2, 3).Value
    
    For i = 2 To lastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Add list of ticker symbols
            tickerSym = ws.Cells(i, 1).Value
            ws.Range("I" & sumData).Value = tickerSym
            
            'Get Yearly Change
            closeVal = ws.Cells(i, 6).Value
            ws.Range("J" & sumData).Value = closeVal - startVal
            
            'Get Percent of Change
            startVal = ws.Cells(i, 3).Value
            closeVal = ws.Cells(i, 6).Value
            ws.Range("K" & sumData).Value = closeVal / startVal
            
            'Get total volume- PROBLEM: It's not summing the value
            totalVol = totalVol + ws.Cells(i, 7).Value
            ws.Range("L" & sumData).Value = totalVol
            
            'Placeholders to see if I can get correct value
            ws.Range("M" & sumData).Value = startVal
            ws.Range("N" & sumData).Value = closeVal
        
            sumData = sumData + 1
        End If
           
        'Set Conditional Formatting
        If (ws.Cells(i, 10).Value > 0) Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf (ws.Cells(i, 10).Value < 0) Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        If (ws.Cells(i, 11).Value > 0) Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
        ElseIf (ws.Cells(i, 11).Value < 0) Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
        End If
            
        Next i
        
        Next ws
End Sub
