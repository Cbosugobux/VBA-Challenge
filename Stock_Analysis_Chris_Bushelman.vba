Attribute VB_Name = "Module1"
Sub stock_analysis()

Dim ticker_name As String
Dim Open_Price As Double
Dim Close_Price As Double
Dim QDelta As Double
Dim Perc_chng As Double
Dim volTotal As Double
Dim lastRow As Long
Dim tickerRow As Integer
Dim firstOpenRow As Long
Dim ws As Worksheet
Dim processedTickers As Collection

 




For Each ws In ActiveWorkbook.Worksheets
    tickerRow = 2
    volTotal = 0
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set processedTickers = New Collection ' Code assistance provided by Xpert, EdX AI Learning Assistant
 
    'New Headers
    ws.Columns("I:Q").AutoFit
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(2, 15).Value = "Greatest %Increase"
    ws.Cells(3, 15).Value = "Greatest %Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
    
    For Row = 2 To lastRow
        'check if the ticker has already been processed
        On Error Resume Next
        processedTickers.Add ws.Cells(Row, 1).Value, CStr(ws.Cells(Row, 1).Value) ' Code assistance provided by Xpert, EdX AI Learning Assistant
        
        'check if current row is the first occurrence of the ticker
        If ws.Cells(Row - 1, 1).Value <> ws.Cells(Row, 1).Value Then
        firstOpenRow = Row 'First Occurence for Opening Price
        Open_Price = ws.Cells(firstOpenRow, 3).Value 'Opening Price
        End If
    
        ' Check to see if ticker symbol changes
        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
        
            ' Display name of Ticker
            ticker_name = ws.Cells(Row, 1).Value
        
            ' Put the ticker name in the row
            ws.Range("I" & tickerRow).Value = ticker_name
                 
            ' Find the Closing Price
            Close_Price = ws.Cells(Row, 6).Value
        
            ' Calculate QDelta
            If Open_Price <> 0 Then
                QDelta = (Close_Price - Open_Price)
            ' List QDelta in Column 10
            ws.Cells(tickerRow, 10).Value = QDelta
        
            ' Format QDelta Cells
            If QDelta > 0 Then
                ws.Cells(tickerRow, 10).Interior.ColorIndex = 4 ' Green for positive
            ElseIf QDelta < 0 Then
                ws.Cells(tickerRow, 10).Interior.ColorIndex = 3 ' Red for negative
            Else
                ws.Cells(tickerRow, 10).Interior.ColorIndex = 2 ' White for zero
            End If
         
            If Open_Price <> 0 Then
                ' Find Percent Change
                Perc_chg = ((Close_Price - Open_Price) / Open_Price)
  
          
                ' List Percent Change
                ws.Cells(tickerRow, 11).Value = Perc_chg
                ws.Cells(tickerRow, 11).NumberFormat = "0.00%"
                
            End If
            
            ' Add the total Volume into column L
            ws.Cells(tickerRow, 12).Value = volTotal + Cells(Row, 7).Value
            
            ' Increment tickerRow for the next ticker
            tickerRow = tickerRow + 1
            
            ' Reset volTotal
            volTotal = 0
         
        End If
    
        Else
        ' If the ticker does not change, add volume to volTotal
        If IsNumeric(ws.Cells(Row, 7).Value) Then
            volTotal = volTotal + ws.Cells(Row, 7).Value
        End If
        
    End If
        
  Next Row
  
    'After settling, find the min and max percent change in column I
    Dim r As Range
    Dim minTicker As String
    Dim maxTicker As String
    Dim maxV As Double
    Dim vR As Range
    Dim maxvolTicker As String
    
    'Define r
    Set r = ws.Range("K2:K" & ws.Cells(Rows.Count, "K").End(xlUp).Row)
    Set vR = ws.Range("L2:L" & ws.Cells(Rows.Count, "L").End(xlUp).Row)
    
    'Find Min and Max in our range
    xmin = Application.WorksheetFunction.Min(r)
    xmax = Application.WorksheetFunction.Max(r)
    maxV = Application.WorksheetFunction.Max(vR)
    
    'Find the corresponding ticker for our values
    minTicker = ws.Cells(Application.WorksheetFunction.Match(xmin, r, 0) + 1, "I")
    maxTicker = ws.Cells(Application.WorksheetFunction.Match(xmax, r, 0) + 1, "I")
    maxvolTicker = ws.Cells(Application.WorksheetFunction.Match(maxV, vR, 0) + 1, "I")
    
    'Print Values
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 16).Value = maxTicker
    ws.Cells(3, 16).Value = minTicker
    ws.Cells(4, 16).Value = maxvolTicker
    ws.Cells(2, 17).Value = xmax
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).Value = xmin
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = maxV
    
    
       
    
Next ws

End Sub

