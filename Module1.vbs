Attribute VB_Name = "Module1"
Sub yearly_change()
'Module 2 HomeWork using Stock Info

   'declare worksheet
   Dim ws As Worksheet
   
     'loop through each worksheet
     For Each ws In Worksheets
   
        'Create the column headings
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
    
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
    
    
        'set variable for ticker symbol
        Dim ticker As String
    
        'set variable for opening price
        Dim ticker_open As Double
        ticker_open = 0
    
        'set variable for closing price
        Dim ticker_close As Double
        ticker_close = 0
    
        'set variable for actual change in price
        Dim yearly_change As Double
        yearly_change = 0
       
        'set variable for percentage change in price
        Dim percentage_change As Double
        percentage_change = 0
    
        'set variable for total volume
        Dim total_volume As Double
        total_volume = 0
    
        'set variable to hold the row
        Dim row As Long
    
        'set ticker row
        Dim tkrRow As Long
        tkrRow = 2
    
        ' define the last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
        'loop through all the stock entries
        For row = 2 To lastRow
                 
            'check for changes in the ticker symbol
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                         
            'set ticker symbol  - grab symbol before change
            ticker = ws.Cells(row, 1).Value
            
            'grab opening price value
            ticker_open = ws.Cells(tkrRow, 3).Value
            
            'grab closing price value
            ticker_close = ws.Cells(row, 6).Value
            
            'calculate yearly change
            yearly_change = (ticker_close - ticker_open)
            
            'calculate percentage change with div by 0 exception
            If ticker_open <> 0 Then
            percentage_change = (yearly_change / ticker_open)
            Else: percentage_change = "0"
            End If
            
            
                'display the ticker symbol
                ws.Range("J" & tkrRow).Value = ticker
            
                'display the Yearly change ticker open
                ws.Range("K" & tkrRow).Value = yearly_change
            
                    'set display to background green if positive red if positive
                    If ws.Range("K" & tkrRow).Value > 0 Then
                    ws.Range("K" & tkrRow).Interior.Color = vbGreen
                
                    'set display to background red if negative
                    ElseIf ws.Range("K" & tkrRow).Value < 0 Then
                    ws.Range("K" & tkrRow).Interior.Color = vbRed
               
                    'set display to background yellow if no change
                    ElseIf ws.Range("K" & tkrRow).Value = 0 Then
                    ws.Range("K" & tkrRow).Interior.Color = vbYellow
                               
                    End If
            
                'display the pecentage change
                ws.Range("L" & tkrRow).Value = percentage_change
            
                ' add volume count
                total_volume = (total_volume + (ws.Cells(tkrRow, 7).Value))
            
                'display the total volume
                ws.Range("M" & tkrRow).Value = total_volume
            
                tkrRow = tkrRow + 1
                       
                total_volume = 0
            
            Else
                ' if there is no change in the volume, keep adding to the total
                total_volume = (total_volume + (ws.Cells(row, 7).Value))
                                 
            End If
            
        Next row
        
        ' find and state the maximum percentage increase
        Dim max_percent_increase As Double
        max_percent_increase = WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & lastRow))
        ws.Range("Q2").Value = max_percent_increase
        
        ' match ticker with max percentage increase
        Dim ticker_increase As String
        ticker_increase = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("L2" & ":" & "L" & lastRow), 0)
        ws.Range("P2").Value = ws.Range("J" & ticker_increase + 1).Value
                
        ' find and state the maximum percentage decrease
        Dim max_percent_decrease As Double
        max_percent_decrease = WorksheetFunction.Min(ws.Range("L2" & ":" & "L" & lastRow))
        ws.Range("Q3").Value = max_percent_decrease
        
        ' match ticker with max percentage decrease
        Dim ticker_decrease As String
        ticker_decrease = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("L2" & ":" & "L" & lastRow), 0)
        ws.Range("P3").Value = ws.Range("J" & ticker_decrease + 1).Value
        
        ' find and state the maximum total volume
        Dim max_total_volume As Double
        max_total_volume = WorksheetFunction.Max(ws.Range("M2" & ":" & "M" & lastRow))
        ws.Range("Q4").Value = max_total_volume
        
        ' match ticker with max percentage decrease
        Dim ticker_volume As String
        ticker_volume = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("m2" & ":" & "m" & lastRow), 0)
        ws.Range("P4").Value = ws.Range("J" & ticker_volume + 1).Value
                 
        'autofit the columns
        ws.Range("J1:Q5").Columns.AutoFit
        
        'format output cells
        ws.Range("L2" & ":" & "L" & lastRow).NumberFormat = "0.00%"
        ws.Range("Q2" & ":" & "Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0000000"
        ws.Range("J1:Q1").Font.FontStyle = "bold"
        
        
    Next ws
     
End Sub
