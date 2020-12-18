Attribute VB_Name = "Module1"
Sub Alpha()

    ' Create variables for Ticker, yearly change, percent change,
    ' and total stock volume and table display

    Dim Ticker As String
    Dim Yearly As Double
    Dim Percent As Double
    
    'set an initial variable for holding the total stock value
    Dim Tot_S_Vol As LongLong
    Tot_S_Vol = 0
    Dim ws As Worksheet
    Dim opening_price As Double
    Dim closing_price As Double
    
    'ws.cells(1,3) = opening_price
    
    Dim lastrow As Double
    
   ' Keep track of the location for each ticker in the table
    Dim Ticker_Row As Long
    
    Ticker_Row = 2
             
    For Each ws In Worksheets
        
           ' Variable for the Summary Table row
        Dim summary_row As LongLong
        opening_price = ws.Cells(2, 3).Value
        summary_row = 2
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 10) = "Ticker"
        ws.Cells(1, 11) = "Yearly Change"
        ws.Cells(1, 12) = "Percent Change"
        ws.Cells(1, 13) = "Total Stock Volume"
        
   
   ' Loop through all stocks purchases
        For i = 2 To lastrow
          
        ' Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
              ' Set the Ticker name
              Ticker = ws.Cells(i, 1).Value
              opening_price = ws.Cells(i, 3).Value
              closing_price = ws.Cells(i, 6).Value
               ' Reset the Brand Total
              
              ' Add to the Total Stock Volume
              Tot_S_Vol = Tot_S_Vol + ws.Cells(i, 7)
              ws.Cells(summary_row, 13).Value = Tot_S_Vol
              Yearly = (ws.Cells(i, 6).Value) - (opening_price)
              
              
              
              
              
              
              
              
              ws.Cells(summary_row, 11).Value = Yearly
              If Tot_S_Vol > 0 Then
              
                Percent = ws.Cells(summary_row, 11).Value / (opening_price) * 100
               End If
              
              ws.Cells(summary_row, 12).Value = Percent
              opening_price = ws.Cells(i + 1, 3).Value
                'Add to the Ticker Total Stock Volume
              'Tot_S_Vol = Tot_S_Vol + ws.Cells(I, 7).Value
        
              ' Print the Ticker to the Table
              ws.Cells(summary_row, 10).Value = Ticker
              
           ' Print the Total Stock Volume to the Summary Table
              summary_row = summary_row + 1
              Tot_S_Vol = 0
                  
            ' If the cell immediately following a row is the same brand...
        Else
        
        Tot_S_Vol = Tot_S_Vol + ws.Cells(i, 7)
    
    
         
    
        End If
    
      Next i
      
    Next ws
  
End Sub
    
    






