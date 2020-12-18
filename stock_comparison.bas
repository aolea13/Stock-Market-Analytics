Attribute VB_Name = "Module_Main"
Sub stock_comparison()

Dim sht As Worksheet
For Each sht In Worksheets

    'Create Headers
    sht.Range("I1").Value = "Ticker Symbol"
    sht.Range("J1").Value = "Yearly Change"
    sht.Range("K1").Value = "Percent Change"
    sht.Range("L1").Value = "Total Volume"

    'Def all Var
    Dim stock_start As Double
    stock_start = 0
    Dim total_volume As Double

    'Set j and Last Row for i
    j = 2
    LastRow = sht.Cells(Rows.Count, 1).End(xlUp).Row


    For i = 2 To LastRow

    'Total Volume Calc
    total_volume = total_volume + sht.Cells(i, 7)

    If sht.Cells(i, 1).Value <> sht.Cells(i - 1, 1).Value Then
        stock_start = sht.Cells(i, 3).Value
    End If

    'Cell Population
    If sht.Cells(i, 1).Value <> sht.Cells(i + 1, 1).Value Then

        'This will fill the ticker symbol
        sht.Cells(j, 9).Value = sht.Cells(i, 1).Value

        'This will fill the Total Volume
        sht.Cells(j, 12).Value = total_volume

        'Stock End values
        stock_end = sht.Cells(i, 6).Value

       'Setting Yearly Change
        yearly_change = stock_end - stock_start
        sht.Cells(j, 10).Value = yearly_change

        'Color Formatting
        If yearly_change >= 0 Then
            sht.Cells(j, 10).Interior.ColorIndex = 4
        Else
            sht.Cells(j, 10).Interior.ColorIndex = 3
        End If

        'Calc for Percent Change Accounting for Zero values and Formatting for %
        If stock_start = 0 Or stock_end = 0 Then
            percent_change = 0
            sht.Cells(j, 11).Value = percent_change
            sht.Cells(j, 11).NumberFormat = "0.00%"
        Else
            percent_change = yearly_change / stock_start
            sht.Cells(j, 11).Value = percent_change
            sht.Cells(j, 11).NumberFormat = "0.00%"
        End If

        'Moving j by 1 as needed
        j = j + 1

        'Reset all values to 0
        stock_start = 0
        total_volume = 0
        stock_end = 0
        yearly_change = 0
        percent_change = 0
        End If

   Next i
        
        'BONUS

        'Set Lastrow for j
        LastRow = sht.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Setting Start Point for Greatest Increase, Greatest Decrease, and Greatest Volume
        greatest_increase = sht.Cells(2, 11).Value
        greatest_decrease = sht.Cells(2, 11).Value
        high_volume = sht.Cells(2, 12).Value
        
        For j = 2 To LastRow
            'Greatest Increase Calc and Ticker
            If sht.Cells(j, 11).Value > greatest_increase Then
                greatest_increase = sht.Cells(j, 11).Value
                greatest_increase_ticker = sht.Cells(j, 9).Value
            End If
            
            'Greatest Decrease Calc and Ticker
            If sht.Cells(j, 11).Value < greatest_decrease Then
                greatest_decrease = sht.Cells(j, 11).Value
                greatest_decrease_ticker = sht.Cells(j, 9).Value
            End If
            
            'Greatest Volume  Calc and Ticker
            If sht.Cells(j, 12).Value > high_volume Then
                high_volume = sht.Cells(j, 12).Value
                high_vol_ticker = sht.Cells(j, 9).Value
            End If
            
        Next j
               
        'Axis Headers Summary Table
        sht.Range("O2").Value = "Greatest % Increase"
        sht.Range("O3").Value = "Greatest % Decrease"
        sht.Range("O4").Value = "Greatest Total Volume"
        sht.Range("P1").Value = "Ticker"
        sht.Range("Q1").Value = "Value"
        
        'Table Population for Summary Table
        sht.Range("P2").Value = greatest_increase_ticker
        sht.Range("Q2").Value = greatest_increase
        sht.Range("P3").Value = greatest_decrease_ticker
        sht.Range("Q3").Value = greatest_decrease
        sht.Range("P4").Value = high_vol_ticker
        sht.Range("Q4").Value = high_volume
        sht.Range("Q2").NumberFormat = "0.00%"
        sht.Range("Q3").NumberFormat = "0.00%"
 
    Next sht
    
End Sub


