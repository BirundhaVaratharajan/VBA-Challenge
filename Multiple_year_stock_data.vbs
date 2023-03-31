Sub Alphabetical_testing()
    
    
    Dim MainSht As Worksheet
   ' Dim Wb As Workbook
    
    'Set Wb = ActiveWorkbook
    
    For Each MainSht In Worksheets
    
    'Defining Row and Column
        Dim Ticker_Col As Integer
        Dim Summ_Table_Row As Integer
        Dim Sum_Last_Row As Integer
        Dim Percent_Change_Col As Integer
        
        
        Dim Yearly_Change_Col As Integer
        
        Dim Yearly_Change As Double
        Dim GreatestPerInc As Double
        Dim GreatestPerDec As Double
        
        Dim ClosingPrice As Double
        Dim OpenPrice As Double
        Dim PercentChange As Double
        Dim TotalStockVolume As LongLong
        Dim GreatestTotalVolume As Double
        Dim CellRange As Range
        Dim TickerName As String
        Dim ticker_max As String
        Dim ticker_min As String
        Dim GreatVol_TickerName As String
        
        Dim TotalStock_Volume_Col As Integer
        Dim Stock_Vol_Col As Integer
        Dim i As Long
        MainSht.Range("I1").Value = "Ticker"
        MainSht.Range("J1").Value = "Yearly Change"
        MainSht.Range("k1").Value = "Percent Change"
        MainSht.Range("L1").Value = "Total Stock Volume"
        MainSht.Range("O2").Value = "Greatest % increase"
        MainSht.Range("O3").Value = "Greatest % decrease"
        MainSht.Range("O4").Value = "Greatest Total Volume"
        MainSht.Range("P1").Value = "Ticker"
        MainSht.Range("Q1").Value = "Value"
 
        
        Summ_Table_Row = 2
        Sum_Last_Row = 0
        Ticker_Col = 9
        Beg_Year = 2
        Yearly_Change_Col = 10
        Percent_Change_Col = 11
        
        TotalStock_Volume_Col = 12
        Stock_Vol_Col = 7
        
        ' Finding Last Row
        Last_Row = MainSht.Cells(Rows.Count, 1).End(xlUp).Row
        
        TotalStockVolume = 0
        GreatestTotalVolume = 0
        
        
        
        
        For i = 2 To Last_Row
        TotalStockVolume = TotalStockVolume + Cells(i, Stock_Vol_Col).Value
        
        If (MainSht.Cells(i + 1, 1).Value <> MainSht.Cells(i, 1).Value) Then
        
           
            
            'getting Open Price fron the column 3
            OpenPrice = MainSht.Cells(i, 3).Value
            
            'Getting Closing price from the column 6
            
            ClosingPrice = MainSht.Cells(i, 6).Value
            
            Yearly_Change = ClosingPrice - OpenPrice
             
                
             
            If (OpenPrice <> 0) Then
                PercentChange = (Yearly_Change / OpenPrice) * 100
            End If
            
            TickerName = MainSht.Cells(i, 1).Value
                 
             'Printing Ticker Lable
            MainSht.Cells(Summ_Table_Row, Ticker_Col).Value = TickerName
            
            'Printing Yearly Change
           MainSht.Cells(Summ_Table_Row, Yearly_Change_Col).Value = Yearly_Change
           
           If (Yearly_Change > 0) Then
               MainSht.Cells(Summ_Table_Row, Yearly_Change_Col).Interior.ColorIndex = 4
           ElseIf (Yearly_Change <= 0) Then
               MainSht.Cells(Summ_Table_Row, Yearly_Change_Col).Interior.ColorIndex = 3
            

           End If
         
           'Printing Percent Change
            MainSht.Cells(Summ_Table_Row, Percent_Change_Col).Value = Format(PercentChange, "Percent")
            
            'Total Stock Volume
            MainSht.Cells(Summ_Table_Row, TotalStock_Volume_Col).Value = TotalStockVolume
            
                       
            Summ_Table_Row = Summ_Table_Row + 1
            Beg_Year = i + 1
           
            TotalStockVolume = 0
        Else
            'TotalStockVolume = TotalStockVolume + Cells(i, Stock_Vol_Col).Value
            
        
        End If
        
        
        Next
        
        
         ' Finding Last Row
        Sum_Last_Row = MainSht.Cells(Rows.Count, Percent_Change_Col).End(xlUp).Row
        Summ_Table_Row = 2
    GreatestPerInc = MainSht.Cells(Summ_Table_Row, Percent_Change_Col).Value
    GreatestPerDec = MainSht.Cells(Summ_Table_Row, Percent_Change_Col).Value
    GreatestTotalVolume = MainSht.Cells(Summ_Table_Row, TotalStock_Volume_Col).Value
   ticker_max = MainSht.Cells(2, Ticker_Col).Value
     
           
        For i = 2 To Sum_Last_Row
       
       PercentChange = MainSht.Cells(i, Percent_Change_Col).Value
       TotalStockVolume = MainSht.Cells(i, TotalStock_Volume_Col).Value
       TickerName = MainSht.Cells(i, Ticker_Col).Value
         'Finding Greatest % increase,gratest % decrease and greates total volume
                      
            If (PercentChange > GreatestPerInc) Then
                GreatestPerInc = PercentChange
                ticker_max = TickerName
                      
            Else
                GreatestPerInc = GreatestPerInc
                'ticker_max = TickerName
                
             End If
                
            If (PercentChange < GreatestPerDec) Then
                GreatestPerDec = PercentChange
                ticker_min = TickerName
            Else
            GreatestPerDec = GreatestPerDec
                           
            End If
            
            If (TotalStockVolume > GreatestTotalVolume) Then
                GreatestTotalVolume = TotalStockVolume
                GreatVol_TickerName = TickerName
            End If

            Next
            
            MainSht.Cells(2, 16).Value = ticker_max
            MainSht.Cells(3, 16).Value = ticker_min
            MainSht.Cells(4, 16).Value = GreatVol_TickerName
            
            'Total Stock Volume
            MainSht.Cells(2, 17).Value = Format(GreatestPerInc, "Percent")
            MainSht.Cells(3, 17).Value = Format(GreatestPerDec, "Percent")
            MainSht.Cells(4, 17).Value = GreatestTotalVolume
            

             
        
    Next MainSht

End Sub



