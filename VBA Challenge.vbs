Sub TickerCounter():
    
    'set up variable that is a type of worksheet
    Dim sheet As Worksheet
    
    For Each sheet In ThisWorkbook.Worksheets
    
        sheetName = sheet.Name
        
         'variable to hold ticker name
        ticker = ""
    
        'variable to hold the total volume stock volume
        totalStockVolume = 0
        
        'variable to hold yearlyChange
        yearlyChange = 0
        
        'variable to hold percent change
        percentChange = 0
        
        'variable to hold the summary table starter row
        summaryTableRow = 2
    
        'variable lastrow will hold the value of the lastrow number in our sheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'variable to hold the initial opening price for each each sheet
        openingPrice = sheet.Cells(2, 3).Value
        
            'Loop through each row
            For Row = 2 To lastrow
    
            'Check to see if ticker changes
                 If sheet.Cells(Row + 1, 1).Value <> sheet.Cells(Row, 1).Value Then
                
                'first set the ticker
                 ticker = sheet.Cells(Row, 1).Value
            
                 'add the vol from the row
                totalStockVolume = totalStockVolume + sheet.Cells(Row, 7).Value
                
                'calculate the closingPrice
                closingPrice = sheet.Cells(Row, 6).Value
                
                'calculate the yearlyChange (close of last date - open of first date)
                yearlyChange = yearlyChange + closingPrice - openingPrice
                
                'calculate the percent change (yearlyChange/opening Price %)
                percentChange = yearlyChange / openingPrice
                
                 'add the ticker to the i column in the summary table row
                sheet.Cells(summaryTableRow, 9).Value = ticker
            
                 'add the total volume to the L column in the summary table row
                sheet.Cells(summaryTableRow, 12).Value = totalStockVolume
                
                'add the yearly change to column J in the summary table row
                sheet.Cells(summaryTableRow, 10).Value = yearlyChange
                
                If yearlyChange <= 0 Then
                    sheet.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
                Else
                    sheet.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
                End If


                'add the percent change to column K column in summary table row
                sheet.Cells(summaryTableRow, 11).Value = percentChange
                
                'change the number format to percent
                sheet.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
            
                'go to the next summary table row (add 1 on to the value of the summary table row)
                summaryTableRow = summaryTableRow + 1
            
                'reset the total stock volume to 0
                totalStockVolume = 0
                
                'reset the opening Price
                openingPrice = sheet.Cells(Row + 1, 3).Value
                
                'reset the yearlyChange to 0
                yearlyChange = 0
            
            Else
                'if brand stays the same, add on to the total charges from the C column
                totalStockVolume = totalStockVolume + sheet.Cells(Row, 7).Value
        
            End If
                      
            Next Row
            
        Next sheet
              
End Sub