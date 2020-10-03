Attribute VB_Name = "Module2"


Sub StockData():
 
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    'Declaration of variables
    
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryTableRow As Long
    Dim LastRow As Long
    Dim MaxIncrease_ticker As String
    Dim MaxDecrease_ticker As String
    Dim GTVolume_ticker As String
    Dim MaxIncrease_value As Double
    MaxIncrease_value = 0
    Dim MaxDecrease_value As Double
    MaxDecrease_value = 0
    Dim GTVolume_value As Double
    GTVolume_value = 0
    
     'Set headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("Q1").Value = "Value"

        'Define the postion of SummaryTableRow
        SummaryTableRow = 2
    
        'Find the last cell that is not empty
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Find the initial value of beginning stock value for the first ticker
        OpeningPrice = ws.Cells(2, 3).Value
    
        ' Find the range of the rows of the worksheet
         For i = 2 To LastRow
    
            'Check if the ticker name is the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
    
            ' Calculation
            ClosingPrice = ws.Cells(i, 6).Value
            YearlyChange = ClosingPrice - OpeningPrice
    
             ' Set conditions
             If OpeningPrice <> 0 Then
                PercentChange = (YearlyChange / OpeningPrice) * 100
                
             End If
             
             ' Color fill yearly price change: red for negative and green for positive
             If (YearlyChange > 0) Then
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
                
             ElseIf (YearlyChange <= 0) Then
                ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
             
             End If
            
            ' Add to the ticker total volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            'Next opening price
            OpeningPrice = ws.Cells(i + 1, 3).Value
        
            'fill in values into summary table
            ws.Cells(SummaryTableRow, 9).Value = Ticker
            ws.Cells(SummaryTableRow, 10).Value = YearlyChange
            ws.Cells(SummaryTableRow, 11).Value = (CStr(PercentChange) & "%")
            ws.Cells(SummaryTableRow, 12).Value = TotalVolume
    
            ' Add 1 to the summary table row count
            SummaryTableRow = SummaryTableRow + 1
            
            ' Bonus
                If (PercentChange > MaxIncrease_value) Then
                    MaxIncrease_value = PercentChange
                    MaxIncrease_ticker = Ticker
            
                ElseIf (PercentChange < MaxDecrease_value) Then
                    MaxDecrease_value = PercentChange
                    MaxDecrease_ticker = Ticker
                
                End If
            
                If (TotalVolume > GTVolume_value) Then
                    GTVolume_value = TotalVolume
                    GTVolume_ticker = Ticker
                
                End If
            
             ' Reset the value
            PercentChange = 0
            TotalVolume = 0
            
            Else
           
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
           
            End If
            
            Next i
            
    'Print all values into table
    ws.Range("P2").Value = MaxIncrease_ticker
    ws.Range("Q2").Value = (CStr(MaxIncrease_value) & "%")
    ws.Range("P3").Value = MaxDecrease_ticker
    ws.Range("Q3").Value = (CStr(MaxDecrease_value) & "%")
    ws.Range("P4").Value = GTVolume_ticker
    ws.Range("Q4").Value = GTVolume_value
    
    Next ws



End Sub




