Attribute VB_Name = "Module4"
Sub VBA_Homework_Final()

For Each ws In Worksheets
    
    Dim Ticker As String
    
    Dim Yearly_Change As Double
    
    Dim Stock_Volume_Total As Double
        Stock_Volume_Total = 0
        
    Dim Opening_Value As Double
    
    Dim Closing_Value As Double
    
    Dim Percent_Change As Double
    
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
    'Header Text
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
    'Height and width
        ws.Rows("1:1").RowHeight = 16
        ws.Range("I1").ColumnWidth = 7
        ws.Range("J1").ColumnWidth = 13
        ws.Range("K1").ColumnWidth = 14
        ws.Range("L1").ColumnWidth = 17
        
    'Formatting Styles
       ws.Columns("K:K").NumberFormat = "0.00%"
       ws.Columns("L:L").Style = "Comma"
           
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
        If ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
        
            'Value of opening of first day of the year
            Opening_Value = ws.Cells(i, 3)
            
            'ws.Range("N" & Summary_Table_Row).Value = Opening_Value
        
        End If
        
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker = ws.Cells(i, 1).Value
            
            'Value of closing the last day of the year
            Closing_Value = ws.Cells(i, 6).Value
            
            Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value
            
            Yearly_Change = Closing_Value - Opening_Value
            
                If Opening_Value = 0 Then
                
                    Percent_Change = 0
                
                Else
                
                    Percent_Change = Yearly_Change / Opening_Value
                
                End If
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
                    'Formatting color based on value
                    If Yearly_Change < 0 Then

                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    
                    Else

                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
                    End If
            
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            ws.Range("L" & Summary_Table_Row).Value = Stock_Volume_Total
            
            'ws.Range("O" & Summary_Table_Row).Value = Closing_Value
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Stock_Volume_Total = 0
            
        Else
        
            Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value
            
        End If

    Next i
    
Next ws

End Sub

