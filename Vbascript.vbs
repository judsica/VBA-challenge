Attribute VB_Name = "Module1"
Sub stockdata()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Multiple_year_stock_data As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim Yearlychange As Double
   
    
For Each ws In Worksheets

' ---------------------------
'Loop through all the sheets
'create new columns

        
' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        Multiple_year_stock_data = ws.Name
        'MsgBox WorksheetName

        ' Add a Column for the Ticker
        ws.Range("I1").EntireColumn.Insert

        ' Add the new column names
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
       
        ' Add the Year to all rows
      
        ' Find the opening price at the beginning of the year
        openingPrice = ws.Cells(2, 3).Value
        
       
    Dim summary_table_row As Integer
    Dim total_volume As LongLong
        summary_table_row = 2
    
    For cell = 2 To LastRow
        
        If Cells(cell + 1, 1).Value <> Cells(cell, 1).Value Then
        
        ' get the ticker value from the current row
        ticker = Cells(cell, 1).Value
        
        ' output the ticker in the summary table
        ws.Cells(summary_table_row, 9).Value = ticker
        
        ' Find the closing price at the end of the year
        closingPrice = ws.Cells(cell, "F").Value
        
        ' Calculate the yearly change
        Yearlychange = closingPrice - openingPrice
        
        ' conditional formatting for positive values being green
         If (Yearlychange > 0) Then
            
            ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
            'Otherwish color it red when negative
            Else
            ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
            
        End If
    
        
        ' Output the yearly change in the adjacent column
        ws.Cells(summary_table_row, 10).Value = Yearlychange
        
        ' add last closing price total volume
        total_volume = total_volume + Cells(cell, 7).Value
        ' output total volume to summary table
        ws.Cells(summary_table_row, 12).Value = total_volume
        'Calculating percentage change
        ws.Cells(summary_table_row, 11).Value = (Yearlychange / openingPrice) * 100
             
             
        'reset the opening value for next ticker
        openingPrice = ws.Cells(cell + 1, 3).Value
        
        summary_table_row = summary_table_row + 1
        
        total_volume = 0
        
        Else 'else if the tickers are the same
        
        'add volume to total volume
        total_volume = total_volume + Cells(cell, 7).Value
        
        End If
        

        'Adding Functionality , with Greatest % increase, Greatest % decrease, Greatest Total volume, Ticker and Value
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest Total volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
       
     
   
    Next cell
        
    Next ws

End Sub
