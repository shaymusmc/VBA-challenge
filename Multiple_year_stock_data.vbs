'Had trouble finding what was needed from class activities, required a lot of "google-fu" - much code found on mrexcel.com, stackoverlow, and github user Jtuttle314

Sub wallstreet()


'Loop Worksheets
For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        Sheets(ws.Name).Select
    
'Set variables
Dim date_open As Variant
Dim date_close As Variant
Dim i As Double
i = 2
Dim j As Double
j = 2
    
'Write Headings
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Volume"
    
'Set starting values
Cells(j, 9).Value = Cells(j, 1).Value
date_open = Cells(i, 3).Value

'Start Sheet Loop for origin
last_row = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To last_row
            
        'If statement for Ticker Symbols and Volume
        If Cells(i, 1).Value = Cells(j, 9).Value Then
                    
            volume = volume + Cells(i, 7).Value
            
            date_close = Cells(i, 6).Value
            
            Else
            
            Cells(j, 10).Value = date_close - date_open
            
                If date_close <= 0 Then
                    Cells(j, 11).Value = 0
                    
                    Else
                    
                        If date_open <= 0 Then
                            Cells(j, 11).Value = 100
                            
                            Else
                                Cells(j, 11).Value = (date_close / date_open) - 1
                        
                        End If
                
                End If
            
                'Format Column to two decimal places
                Columns("K").NumberFormat = "0.00%"
            
                If Cells(j, 10).Value >= 0 Then
                    Cells(j, 10).Interior.ColorIndex = 4
                    
                    Else
                        Cells(j, 10).Interior.ColorIndex = 3
                
                End If
                

        Cells(j, 12).Value = volume
        
        'Reset Ticker and Volume
        date_open = Cells(i, 3).Value
        volume = Cells(i, 7).Value
        
        'Increase j
        j = j + 1
        
        Cells(j, 9).Value = Cells(i, 1).Value
        End If
        
    Next i
    
    'Relooping for all other lines
    Cells(j, 10).Value = date_close - date_open
    
        If date_close <= 0 Then
            Cells(j, 11).Value = 0
            
            Else
            
                If date_open <= 0 Then
                    Cells(j, 11).Value = 100
                    
                    Else
                        Cells(j, 11).Value = (date_close / date_open) - 1
                        
                 End If
                 
        End If
        
        'Format Column to two decimal places
        Columns("K").NumberFormat = "0.00%"
    
        If Cells(j, 10).Value >= 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
    
            Else
                Cells(j, 10).Interior.ColorIndex = 3
        
        End If
    
    Cells(j, 12).Value = volume
         
    'Finding Greatest measures
        
        last_row = Cells(Rows.Count, "I").End(xlUp).Row
        
        For j = 2 To last_row
        
            If Cells(j, 11).Value > volume_greatest_increase Then
            
                ticker_greatest_increase = Cells(j, 9).Value
                volume_greatest_increase = Cells(j, 11).Value
        
            End If
        
        
            If Cells(j, 11).Value < volume_greatest_decrease Then
            
                ticker_greatest_decrease = Cells(j, 9).Value
                volume_greatest_decrease = Cells(j, 11).Value
        
            End If
        
        
            If Cells(j, 12).Value > volume_greatest_total_volume Then
            
                ticker_greatest_total_volume = Cells(j, 9).Value
                volume_greatest_total_volume = Cells(j, 12).Value
        
            End If
        
        Next j
        
    Cells(2, 16).Value = ticker_greatest_increase
    Cells(2, 17).Value = volume_greatest_increase
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = ticker_greatest_decrease
    Cells(3, 17).Value = volume_greatest_decrease
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 16).Value = ticker_greatest_total_volume
    Cells(4, 17).Value = volume_greatest_total_volume
    
    'Resize Columns
    Columns("I:Q").EntireColumn.AutoFit
    
    'Force page back to top.
    Cells(1, 1).Select
            
Next ws

End Sub