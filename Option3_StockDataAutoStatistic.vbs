Sub StockStatistic()

'Define variables needed for summary table
Dim Rng As Range
Dim rng2 As Range
Dim dblMin As Double
Dim dblMax As Double
Dim stkMax As Double

'Define yearly change value for conditional formatting
Dim Val As Double




'--------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
For Each ws In Worksheets


    'Do initial cleaning up before a new cycle starts
    ws.Range("I:Q").Delete
    

    ' Determine the Last Row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Add the word Ticker to the Summary First Column Header
    ws.Cells(1, 9).Value = "Ticker"
    'Add the word Yearly Change to the Summary Second Column Header
    ws.Cells(1, 10).Value = "Yearly Change"
    'Add the word Percent Change to the Summary Third Column Header
    ws.Cells(1, 11).Value = "Percent Change"
    ' Add the word Total Stock Volume to the Summary Fourth Column Header
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' Set an initial variable for holding the ticker name
    Dim Ticker_Name As String
    
    ' Set an initial variable for holding the total per ticker name
    Dim Ticker_Total As Double
    Ticker_Total = 0

    'Set an initial variable for holding each ticker's open and close price in a year
    Dim Ticker_Open As Double
    Ticker_Open = 0
    Dim Ticker_Close As Double
    Ticker_Close = 0

    'Set an initial variable for holding the price change
    Dim Yearly_Change As Double
    Yearly_Change = 0

    'Set an initial variable for holding the percentage change
    Dim Percent_Change As Double
    Percent_Change = 0
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    
     
    
    ' Loop through all ticker transactions
    For i = 2 To lastRow
    
       '----------------------------------------------------------
       'This part will get each ticker's begining row and ending row position
       'by using a, b counter
       '----------------------------------------------------------
        'Check if ticker value changed or not
        'If ticker name has no change, get the open price postion row count number
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            If a = 0 Then
                a = i
            End If
        Else
        
        
            'Get next new ticker name beginning row position
            If a > 0 And b = 0 Then
                b = i
            End If
            
            'This part handles the case of only one row ticker case, not the case in sample data
            If (a = 0 And b = 0) Then
                a = i
                b = i
            End If
            
        End If
        
        'Populate the final Ticker year openning value
        If (b > 0 And a > 0) Then
            Ticker_Open = ws.Cells(i - (b - a), 3).Value
            
            'Reset row beginning and ending counter to be zero
            a = 0
            b = 0
        
        End If
        '-----------------------------
        'Ending of year open price handling
        '-----------------------------

        ' Check if we are still within the same ticker name, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the ticker name
             Ticker_Name = ws.Cells(i, 1).Value

            ' Add to the Ticker Total Stock Volume
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            
            ' Add to the Ticker Close Price
            Ticker_Close = ws.Cells(i, 6).Value

            ' Calculate the price change
            Yearly_Change = Ticker_Close - Ticker_Open
            
            'This part handles Ticker_Open zero case
            If Ticker_Open = 0 Then
                Percent_Change = 0
            Else
            

            'Calculate percentage change
            Percent_Change = (Ticker_Close - Ticker_Open) / Ticker_Open

            ' Print the Ticker Name in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

            ' Print the Ticker Total volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = Ticker_Total

            ' Print the price change to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

            ' Print the percentage change to the Summary Table
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the Ticker Total
            Ticker_Total = 0
            
            End If

        ' If the cell immediately following a row is the same ticker...
        Else

            ' Add to the Ticker Total
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

            'Pass the beginning open price
            Ticker_Open = ws.Cells(i, 3).Value

        End If
        
      

    Next i
    
      
        'Put conditional setting on 'Yearly Change' Column
    Dim j As Long
    
    For j = 2 To lastRow
            Val = ws.Cells(j, 10).Value
        
            With ws.Cells(j, 10).Interior
             'Checking the condition and assigning the appropriate color
                If Val > 0 Then
                    .ColorIndex = 4
                ElseIf Val < 0 Then
                    .ColorIndex = 3
                
                End If
            End With
        Next j
    
    'Change percent change column format to be percent
    ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"



    'Below section handles the summary table for max and min value
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    




    'Set range from which to determine value
    Set Rng = ws.Range("K2:K" & lastRow)
    Set rng2 = ws.Range("L2:L" & lastRow)

    'Worksheet function MIN, MAX returns the values in a range
    dblMax = Application.WorksheetFunction.Max(Rng)
    dblMin = Application.WorksheetFunction.Min(Rng)
    stkMax = Application.WorksheetFunction.Max(rng2)

    ws.Cells(2, 17).Value = dblMax
    ws.Cells(3, 17).Value = dblMin
    ws.Cells(4, 17).Value = stkMax

    'Change percent change column format to be percent
    ws.Range("Q2:Q3").NumberFormat = "0.00%"



    'This part will get the max, min change and postion informaiton
    Dim n, m, p, q, maxcell, maxval, minval, mincell, maxstkvlm, maxstkcell As Long
    m = Range("K2:K" & lastRow).Value
    p = Range("K2:K" & lastRow).Value
    q = Range("L2:L" & lastRow).Value


    maxval = ws.Cells(2, 11)
    minval = ws.Cells(2, 11)
    maxstkvlm = ws.Cells(2, 12)

    For n = 2 To lastRow
        If ws.Cells(n, 11) > maxval Then
            maxval = ws.Cells(n, 11).Value
            maxcell = n
        End If
    
    Next n

    'This part will get the min change and postion information
    ws.Cells(2, 16).Value = ws.Cells(maxcell, 9)

    For p = 2 To lastRow
        If ws.Cells(p, 11) < minval Then
            minval = ws.Cells(p, 11).Value
            mincell = p
        End If
    Next p

    ws.Cells(3, 16).Value = ws.Cells(mincell, 9)


    'this part will ge the max stock volumn and position
    For q = 2 To lastRow
        If ws.Cells(q, 12) > maxstkvlm Then
            maxstkvlm = ws.Cells(q, 12).Value
            maxstkcell = q
        End If
    Next q
    'Populate the ticker name based on max stock volumn row
    ws.Cells(4, 16).Value = ws.Cells(maxstkcell, 9)


    'Adjust Columns O,Q based on value
    
    ws.Columns("O:Q").AutoFit



'--------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
Next ws

End Sub




