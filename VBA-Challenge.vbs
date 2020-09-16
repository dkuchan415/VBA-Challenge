Attribute VB_Name = "Module1"
Sub VBAWallStreet()

    ' Define Variables
    
    Dim Ticker As String
    Dim LastRow As Long
    Dim SummaryLastRow As Integer
    Dim Opening_Price As Double
    Dim Stock_Volume As Double
    Dim ws As Worksheet
    Dim j As Integer
    
    ' Worksheet Loop
    
    For Each ws In Worksheets
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "YearlyChange"
        ws.Cells(1, 11).Value = "PercentChange"
        ws.Cells(1, 12).Value = "TotalStockVolume"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Create a script that will loop through all the stocks for one year
    
        j = 2
        Opening_Price = ws.Cells(2, 3).Value
        Stock_Volume = 0
        
        'Ticker
        
         For i = 2 To LastRow
             If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                 ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                 
                
             ' Yearly Change
             
                 Closing_Price = ws.Cells(i, 6).Value
                 
                 ws.Cells(j, 10).Value = Closing_Price - Opening_Price
                 
              ' Percent Change = change over opening price
              
                 If Opening_Price = 0 Then
                 
                     ws.Cells(j, 11).Value = 0
                 
                 Else:
                 
                     ws.Cells(j, 11).Value = (Closing_Price - Opening_Price) / Opening_Price
                     
                 End If
                 
                 Opening_Price = ws.Cells(i + 1, 3).Value
                 
              'Record Total Volume
              
                 Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
              
                 ws.Cells(j, 12).Value = Stock_Volume
                 
                 Stock_Volume = 0
                 
                 j = j + 1
                 
             Else:
                 
                 ' Total Stock Volume
                  
                 Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                 
             End If
             
        
         Next i
         
         ' Format Percent Change Cells
         
         SummaryLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
         For k = 2 To SummaryLastRow
         
         ws.Cells(k, 11).NumberFormat = "0.00%"
         
             If (ws.Cells(k, 11) >= 0) Then
                 ws.Cells(k, 11).Interior.ColorIndex = 4
             Else: ws.Cells(k, 11).Interior.ColorIndex = 3
             
             End If
             
         Next k
        
    Next ws
        
        
    
    ' Summary on each sheet including the ticker symbol, yearly change, percent change, total stock volume
    
    
    
    
    
    ' Challenge. Add stocks with greatest % increase, % decrease, Greatest volume
    
    ' Replicate for each page
    
    
    
    
    

End Sub
