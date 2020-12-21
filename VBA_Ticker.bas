Attribute VB_Name = "VBA_Ticker"
Sub prices()
Dim i As Double
Dim lastrow_1 As Long 'set variable to hold open price
Dim open_price As Double 'set variable to hold close price
Dim close_price As Double 'set variable for ticker
Dim ticker As String 'set variable to keep track of location of each ticker
Dim Ticker_i As Long 'sets row where results will display with worksheet
Dim volume As Double 'sets variable for stock volume
Dim yearly_change As Double 'sets variable for yearly change in stock price
Dim percent_change As Long 'sets variable for percent change in stock price
Dim ws As Worksheet 'defines integer to run on all worksheets

For Each ws In ThisWorkbook.Worksheets 'begin doing the below on all workbook sheets
    
    ws.Range("I1").Value = "Ticker" 'adds Ticker header to all workbook sheets
    ws.Range("J1").Value = "Yearly Change" 'adds Yearly Change header to all workbook sheets
    ws.Range("K1").Value = "Percent Change" 'adds Percent Change header to all workbook sheets
    ws.Range("L1").Value = "Total Volume" 'adds Total Volume header to all workbook sheets
    ws.Range("N1").Value = "Opening Price" 'adds Opening Price header to all workbook sheets OPTIONAL
    ws.Range("O1").Value = "Closing Price" 'adds Closing Price header to all workbook sheets OPTIONAL
    
    Ticker_i = 2 'sets initial row for displaying results
    lastrow_1 = ws.Cells(Rows.Count, 1).End(xlUp).Row 'defines the last row of data for the macro to run through
    open_price = ws.Cells(2, 3).Value 'sets the initial opening price
    volume = 0 'sets the initial stock volume

    For i = 2 To lastrow_1 'create loop for worksheet
    
        ticker = ws.Cells(i, 1).Value 'set the ticker
        close_price = ws.Cells(i, 6).Value 'set the closing price for each ticker
        volume = volume + ws.Cells(i, 7).Value 'sets the total stock volum for each ticker
    
        If (ticker <> Cells(i + 1, 1).Value) Then ' IF the value underneath current cell at column 1 is different
            ws.Cells(Ticker_i, 9).Value = ticker 'print ticker in new column
            ws.Cells(Ticker_i, 14).Value = open_price 'print opening price in new column OPTIONAL
            ws.Cells(Ticker_i, 15).Value = close_price 'print closing price in new column OPTIONAL
            ws.Cells(Ticker_i, 10).Value = close_price - open_price 'define the yearly change
            
            ws.Cells(Ticker_i, 12).Value = volume 'print stock volume in new column
            
            open_price = ws.Cells(i + 1, 3).Value 'set new opening price
            Ticker_i = Ticker_i + 1 'add one to the summary table
            volume = 0 'resets the stock volume to zero for each ticker
            
        End If
    
    Next i
Next ws
End Sub
