Sub wallstreet()

Dim ticker As String 'Set initial to hold the ticker
Dim yearOpen As Double 'To store the value of each opening year
Dim yearClose As Double 'To store the value of each closing year
Dim percentChange As Double 'To calculate the value of percentage change
Dim yearlyChange As Long 'To hold the value of the yearly change
Dim volume As Long 'To hold the value of stock each year

'Loop through all the worksheet
For Each ws In Worksheets
    'setting the header in each worksheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'Hold the value of the integer
    summary_table_row = 2
    'Determine the last row of each worksheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    yearStart = 2 'starting position for each worksheet - second row

For i = 2 To lastRow
    'Debug.Print i

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Set the ticket value
        ticker = ws.Cells(i, 1).Value
        'set the volume
        volume = ws.Cells(i, 7).Value
                
        'check if opening value is zero, if so pick the next non zero
        If ws.Cells(yearStart, 6).Value = 0 Then
            For j = yearStart To i
            'Debug.Print i, yearStart, ws.Name
            If ws.Cells(j, 6).Value <> 0 Then
                'Debug.Print ws.Cells(j, 6).Value
                yearOpen = ws.Cells(j, 6).Value
                Exit For
            Else
                yearOpen = 1
            End If
        Next j
    Else
    yearOpen = ws.Cells(yearStart, 6).Value
      
     End If
        
        'assign year end value
        yearClose = ws.Cells(i, 6).Value
        
        'Calculation for yearly change
        'Debug.Print yearOpen
        yearlyChange = yearClose - yearOpen
        
    
        'Calculate the percent change
        percentChange = (yearClose - yearOpen) / yearOpen
        
        
        ws.Cells(summary_table_row, 9).Value = ticker
        ws.Cells(summary_table_row, 10).Value = yearlyChange
        ws.Cells(summary_table_row, 11).Value = percentChange
        ws.Cells(summary_table_row, 12).Value = volume
        summary_table_row = summary_table_row + 1

        'reset the volume counter to 0
        volume = 0
        yearStart = i + 1
        
    End If

  Next i
    ws.Columns("K").NumberFormat = "0.00%"

'Setting up column colors
'Determine the last row
lastRowsChange = ws.Cells(Rows.Count, 10).End(xlUp).Row

'looping through the entire column
For Color = 2 To lastRowsChange
'if negative number turn into red color

    If ws.Cells(Color, 10).Value < 0 Then
        ws.Cells(Color, 10).Interior.ColorIndex = 3
    Else
'if positive turn into green cells
ws.Cells(Color, 10).Interior.ColorIndex = 4
    End If
Next Color


Next ws

End Sub

