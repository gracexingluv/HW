VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockMarket()

    'loop through all worksheet
    For Each ws In Worksheets
    
    Dim Ticker_name As String
    Dim Yearly_change As Double
    Dim Percentage_change As Double
    Dim Total_stock_volume As Double
    Dim Summary_table_row As Double
    Dim Yearly_open_row As Double
    Dim row_counter As Double
    

    
    'add titles in summary table1
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'add titles of summary table2
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest% Increase"
    ws.Cells(3, 15).Value = "Greatest% Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    
    'set lastrow
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'initialize parameter
    Summary_table_row = 2
    Total_stock_volume = 0
    row_counter = 0
    Max = 0
    Min = 0
    max_volume = 0
    
    'loop through original data,find reslut for summary1
        For i = 2 To Lastrow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            
                '1st row of current ticker
                Yearly_open_row = i - row_counter
                
                'close value in last row of current ticker - open value in 1st row of current ticker
                Yearly_change = ws.Cells(i, 6).Value - ws.Cells(Yearly_open_row, 3).Value
                
                'change percentage calculation
                Percentage_change = Yearly_change / (ws.Cells(Yearly_open_row, 3).Value + 0.000001)
                
                'total volume of curent ticker calculation
                Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
                
                'fill values to summary table1
                ws.Range("I" & Summary_table_row).Value = ws.Cells(i, 1).Value
                ws.Range("J" & Summary_table_row).Value = Yearly_change
                ws.Range("K" & Summary_table_row).Value = Percentage_change
                ws.Range("L" & Summary_table_row).Value = Total_stock_volume
                
                'format cell in percentage
                ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
                ws.Range("L" & Summary_table_row).Value = Total_stock_volume
                
                'summary table1 - conditional format
                If ws.Range("K" & Summary_table_row).Value > 0 Then
                    ws.Range("K" & Summary_table_row).Interior.ColorIndex = 4
            
                ElseIf ws.Range("K" & Summary_table_row).Value < 0 Then
                    ws.Range("K" & Summary_table_row).Interior.ColorIndex = 3
                End If
                
                'summary table2 row1 - greatest increase percentage
                If ws.Range("K" & Summary_table_row).Value > Max Then
                    Max = ws.Range("K" & Summary_table_row).Value
                    ws.Cells(2, 17).Value = Max
                    ws.Cells(2, 16).Value = ws.Range("I" & Summary_table_row).Value
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                    
                'summary table2 row2 - greatest decrease percentage
                ElseIf ws.Range("K" & Summary_table_row).Value < Min Then
                    Min = ws.Range("K" & Summary_table_row).Value
                    ws.Cells(3, 17).Value = Min
                    ws.Cells(3, 16).Value = ws.Range("I" & Summary_table_row).Value
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                    
                'summary table2 row3 - greatest total volume
                ElseIf ws.Range("L" & Summary_table_row).Value > max_volume Then
                    max_volume = ws.Range("L" & Summary_table_row).Value
                    ws.Cells(4, 17).Value = max_volume
                    ws.Cells(4, 16).Value = ws.Range("I" & Summary_table_row).Value
                       
                End If
                
                'reset for next ticker
                Summary_table_row = Summary_table_row + 1
                Total_stock_volume = 0
                row_counter = 0
            
            'same ticker
            Else
                Total_stock_volume = Total_stock_volume + ws.Cells(i, 7).Value
                row_counter = row_counter + 1
            
            End If
            
        Next i
        
        
    
    Next ws
    
        
End Sub

