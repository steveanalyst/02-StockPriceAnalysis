Sub CalculateVolume()

'--------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
For Each ws In Worksheets

    ' Determine the Last Row
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

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
    For i = 2 To Lastrow
    
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

            'Calculate percentage change
            If Ticker_Open = 0 Then
                Percent_Change = 0
            Else
            Percent_Change = (Ticker_Close - Ticker_Open) / Ticker_Open
            End If
            
    

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

        ' If the cell immediately following a row is the same ticker...
        Else

            ' Add to the Ticker Total
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

            'Pass the beginning open price
            Ticker_Open = ws.Cells(i, 3).Value

        End If
        
        'Put conditional setting on 'Yearly Change' Column
        Dim Val As Double
        
        Val = ws.Cells(i, 10).Value
        
        With ws.Cells(i, 10).Interior
        'Checking the condition and assigning the appropriate color
                If Val > 0 Then
                .ColorIndex = 4
                ElseIf Val < 0 Then
                .ColorIndex = 3
                
                End If
        End With

    Next i
    
'Change percent change column format to be percent
ws.Range("K2:K" & Lastrow).NumberFormat = "0.00%"




'--------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
Next ws

End Sub




