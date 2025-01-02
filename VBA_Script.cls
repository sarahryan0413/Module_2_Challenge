VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Instructions: Create a script that loops through all the stocks for each quarter and outputs the following information:
    'A: The ticker symbol
    'B: Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
    'C: The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
    'D: The total stock volume of the stock.
    'E: Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
    
Sub StockChallenge()

    'define variables
    Dim i As Long
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    Dim TotalStockVolume As Double
    Dim GreatestIncrease As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecrease As Double
    Dim GreatestDecreaseTicker As String
    Dim GreatestTotalVolume As Double
    Dim GreatestTotalVolumeTicker As String
    
    'loop through all sheets
    For Each ws In Worksheets
    
        'determine the last row with data in the current sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'insert headers for the new columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'A: copy ticker symbols from column A to I (new column for ticker)
        ws.Range("I2:I" & LastRow).Value = ws.Range("A2:A" & LastRow).Value


        'B: loop through each row to perform calculations
        For i = 2 To LastRow
            
            'B: retrieve opening and closing prices for each stock
            OpeningPrice = ws.Cells(i, 3).Value
            ClosingPrice = ws.Cells(i, 6).Value
                
            'B: calculate the quarterly change (closing - opening)
            QuarterlyChange = ClosingPrice - OpeningPrice
                
            'B: store the quarterly change in column J
            ws.Cells(i, 10).Value = QuarterlyChange
                
            'B: color the cell based on the quarterly change
            If QuarterlyChange > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4 'green for positive change
                
            ElseIf QuarterlyChange = 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 2 'white for no change
                
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3 'red for negative change
            End If
            
            
            'C: calculate the percent change (Quarterly Change / Opening Price * 100)
            If OpeningPrice <> 0 Then  'potential error if opening price is zero
                PercentChange = (QuarterlyChange / OpeningPrice) * 100
            Else
                PercentChange = 0 'Assign 0 if opening price is 0
            End If
                
            'C: store the calculated percent change in column K
            ws.Cells(i, 11).Value = PercentChange
            
            'C: format the percent change to show two decimal places
            ws.Cells(i, 11).NumberFormat = "0.00%"
            
            
            'D: retrieve the volume for each stock
            Volume = ws.Cells(i, 7).Value
            
            'D: calculate the total stock volume (Volume * Closing Price)
            TotalStockVolume = Volume * ClosingPrice
            
            'D: store the calculated total stock volume in column L
            ws.Cells(i, 12).Value = TotalStockVolume
            
        Next i
        
        
         'E: Loop through for greatest %
        For i = 2 To LastRow
            
            'E: find the greatest % increase change
            If ws.Cells(i, 11).Value > GreatestIncrease Then
            
                'assigns the values in column 9 and 11 into the variables below.
                GreatestIncrease = ws.Cells(i, 11).Value
                GreatestIncreaseTicker = ws.Cells(i, 9).Value
            
            'E: find the greatest % decrease change
            ElseIf ws.Cells(i, 11).Value < GreatestDecrease Then
            
                'assigns the values in column 9 and 11 into the variables below.
                GreatestDecrease = ws.Cells(i, 11).Value
                GreatestDecreaseTicker = ws.Cells(i, 9).Value
                
            End If
        Next i
                    
        'E: final answers in the designated cells
        ws.Range("P2").Value = GreatestIncreaseTicker
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("P3").Value = GreatestDecreaseTicker
        ws.Range("Q3").Value = GreatestDecrease
        
        'E: format percentage with two decimal places
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        
        'E: Loop through for greatest total volume
        For i = 2 To LastRow
        
            'E: find the greatest total volume
            If ws.Cells(i, 12).Value > GreatestTotalVolume Then
            
                'assigns the values in column 9 and 12 into the variables below.
                GreatestTotalVolume = ws.Cells(i, 12).Value
                GreatestTotalVolumeTicker = ws.Cells(i, 9).Value
            
            End If
        Next i
        
        'E: final answers in the designated cells
        ws.Range("P4").Value = GreatestTotalVolumeTicker
        ws.Range("Q4").Value = GreatestTotalVolume

        'autofit columns to adjust to the content
        ws.Columns("I:Q").AutoFit
        
    Next ws
    
End Sub

