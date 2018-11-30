Sub CalculateVolume()

'--------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
For Each ws In Worksheets

    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Add the word Ticker to the Summary First Column Header
    ws.Cells(1, 9).Value = "Ticker"
    ' Add the word Total Stock Volume to the Summary Second Column Header
    ws.Cells(1, 10).Value = "Total Stock Volume"

    ' Set an initial variable for holding the ticker name
    Dim Ticker_Name As String
    
    ' Set an initial variable for holding the total per ticker name
    Dim Ticker_Total As Double
    Ticker_Total = 0
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Loop through all ticker transactions
    For i = 2 To LastRow

        ' Check if we are still within the same ticker name, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the ticker name
            Ticker_Name = ws.Cells(i, 1).Value

            ' Add to the Ticker Total Stock Volume
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

            ' Print the Ticker Name in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

            ' Print the Ticker Total volume to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Ticker_Total

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the Ticker Total
            Ticker_Total = 0

        ' If the cell immediately following a row is the same ticker...
        Else

        ' Add to the Ticker Total
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

        End If

    Next i

'--------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
Next ws

End Sub



