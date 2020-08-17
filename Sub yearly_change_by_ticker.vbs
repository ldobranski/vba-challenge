Sub yearly_change_by_ticker()

Dim ws As Worksheet

'Loop through all worksheets
    For Each ws In Worksheets

'Set an initial variable for holding the ticker name
    Dim Ticker_Name As String
    
'Set an initial variable for holding the total stock volume
    Dim Stock_Volume As LongLong
    Stock_Volume = 0
    
'Set initial variables for holding yearly change in value and percent change
    Dim Yearly_Change As Double
    Dim Opening_Value As Double
    Dim Closing_Value As Double
    Dim Percent_Change As Double
    
'Keep track of the location for summarizing each ticker
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
'Find last row (XXX - as ws.Cells)
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (last_row)
    
  
'Loop through all ticker volumes
    For i = 2 To lastrow
       
'Get the opening stock value
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        Opening_Value = ws.Cells(i, 3).Value
    
    End If
    
    'Check if still within the same ticker name
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set the ticker name
            Ticker_Name = ws.Cells(i, 1).Value
            'MsgBox (Ticker_Name)
        
        'Set closing stock value
            Closing_Value = ws.Cells(i, 6).Value
            'MsgBox (Closing_Value)
        
        'Add to total stock volume
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            'MsgBox (Stock_Volume)
        
        'Calculate the difference between opening and closing value
            Yearly_Change = Closing_Value - Opening_Value
            'MsgBox (Opening_Value)
            'MsgBox (Closing_Value)
            
        'Calculate percent change and format columns with % sign
            If Opening_Value = 0 Then
                Percent_Change = 0
                
                Else
                Percent_Change = (Closing_Value - Opening_Value) / Closing_Value
                ws.Columns("K").NumberFormat = "0.00%"
                
            End If
        
        'Print percent change in the summary table
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
        'Format cell as green for positive change; red for negative
            If Percent_Change > 0 Then
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                
                Else
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
        
        'Print the yearly change in Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
        'Print the ticker name in Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
        
        'Print the Total stock volume for the ticker in summary table
            ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
        
        'Add one to the summary table row to move to next ticker results
            Summary_Table_Row = Summary_Table_Row + 1
              
        'Reset total stock volume to zero for next ticker
        Stock_Volume = 0
        
    Else
    
        'Add to total stock volume
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        'MsgBox (Stock_Volume)
           
    End If
        
    Next i
    
    'Print Headers in the Summary Table
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
    'Autofit cell widths on Summary Table
       ws.Range("I:L").Columns.AutoFit
       
Next ws

	MsgBox ("moderate solution complete")

End Sub
