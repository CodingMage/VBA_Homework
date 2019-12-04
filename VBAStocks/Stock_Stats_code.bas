Attribute VB_Name = "Module1"
Sub Stock_Stats()
    
    'Populate the cell titles
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    Range("n2").Value = "Greatest % Increase"
    Range("n3").Value = "Greatest % Decrease"
    Range("n4").Value = "Greatest Total Volume"
    Range("o1").Value = "Ticker"
    Range("p1").Value = "Value"
    
    'Define variables
    ' "current_ticker" is used to store the ticker value in the loop
    Dim ticker, current_ticker As String
    Dim open_price, close_price As Double
    Dim yearly_change, percent_change As Double
    Dim total As Double         ' Total Stock Volume
    Dim ticker_count As Long    ' The row index of column i
    Dim lastrow As Long
    
    ' Variables for the challenges question
    Dim greatest_increase, greatest_decrease As Double
    Dim greatest_total As Double
    Dim ticker_increase, ticker_decrease, ticker_total As String
    
    
    'Initialize some values
    ticker = Cells(2, 1).Value
    open_price = Cells(2, 3).Value
    total = 0
    ticker_count = 2             'Starting from 2 in order to skip the title line
                        
    greatest_increase = 0
    greatest_decrease = 0
    greatest_total = 0
    
    'Calculate lastrow
    
    lastrow = Cells(2, 1).End(xlDown).Row
    
    'Start the for loop
    
    For i = 2 To lastrow + 1   ' lastrow+1 is used for the purpose of
                               ' displaying the last ticker
        
        '  Set current_ticker at the start of the each loop
        
        current_ticker = Cells(i, 1).Value
        
        ' If the same ticker, collect the close_price and add the increase the total volume
        If ticker = current_ticker Then
        
            close_price = Cells(i, 6).Value
            total = total + Cells(i, 7).Value
            
        Else    ' When ticker changes
                ' Calculate all the changes
                ' Update ticker name
            
            yearly_change = close_price - open_price
            If open_price = 0 Then  ' As open_price used as denominator, it should not be 0
                percent_change = 0
            Else
                percent_change = yearly_change / open_price
            
            End If
            
            ' Populate the cells
            
            Cells(ticker_count, 9).Value = ticker
            Cells(ticker_count, 10).Value = yearly_change
            Cells(ticker_count, 11).Value = percent_change
            Cells(ticker_count, 12).Value = total
            
            ' Looking for the greatest increase/decrease and total
            If greatest_increase < percent_change Then
                greatest_increase = percent_change
                ticker_increase = ticker
            End If
            
            If percent_change < greatest_decrease Then
                greatest_decrease = percent_change
                ticker_decrease = ticker
            End If
            
            If total > greatest_total Then
                greatest_total = total
                ticker_total = ticker
            End If
            
            'formatting the cells as it goes
            Cells(ticker_count, 11).NumberFormat = "0.00%"
            
            'Color formatting for Percentage Change
            If percent_change < 0 Then
                Cells(ticker_count, 10).Interior.Color = RGB(255, 0, 0)
            ElseIf percent_change = 0 Then
            
            Else
                Cells(ticker_count, 10).Interior.Color = RGB(0, 255, 0)
            End If
            
            
            ' Reset total and open_price
            total = 0
            open_price = Cells(i, 3).Value
            
            ' Change the ticker, ticker_count
            ticker = current_ticker
            ticker_count = ticker_count + 1
            
        End If
    
    
    Next i
    
    ' Populate the cells with results
    Range("O2").Value = ticker_increase
    Range("O3").Value = ticker_decrease
    Range("O4").Value = ticker_total
    Range("P2").Value = greatest_increase
    Range("P3").Value = greatest_decrease
    Range("P4").Value = greatest_total
    
    'formatting cell P2, P3 and P4
    Range("P2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"
    Range("P4").NumberFormat = "0.0000E+00"
    
End Sub

