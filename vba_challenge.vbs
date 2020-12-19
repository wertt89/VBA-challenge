Sub vba_challenge()

    'Setting current sheet for looping purposes
    Dim current_sheet As Worksheet
        
    'Loop through all worksheets
    For Each current_sheet In Worksheets
    
        'Setting all needed variables for basic and bonus tables and setting values to zero if number
        Dim ticker_name As String
        Dim ticker_volume As Double
        ticker_volume = 0
        Dim open_value As Double
        open_value = 0
        Dim close_value As Double
        close_value = 0
        Dim value_change As Double
        value_change = 0
        Dim percentage_change As Double
        percentage_change = 0
        Dim biggest_increase As String
        Dim biggest_percentage_increase As Double
        biggest_percentage_increase = 0
        Dim biggest_decrease As String
        Dim biggest_percentage_decrease As Double
        biggest_percentage_decrease = 0
        Dim biggest_volume As String
        Dim biggest_volume_number As Double
        biggest_volume_number = 0
    
        'Keeping track of ticker name locations in the summary table for the current worksheet
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
    
        'Setting last row so that data structure accommodates to varying data ranges
        Dim last_row As Long
            last_row = current_sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Setting headers for both basic and bonus tables
        current_sheet.Range("I1").Value = "Ticker"
        current_sheet.Range("J1").Value = "Yearly Change"
        current_sheet.Range("K1").Value = "Percent Change"
        current_sheet.Range("L1").Value = "Total Stock Volume"
        current_sheet.Range("O2").Value = "Greatest % Increase"
        current_sheet.Range("O3").Value = "Greatest % Decrease"
        current_sheet.Range("O4").Value = "Greatest Total Volume"
        current_sheet.Range("P1").Value = "Ticker"
        current_sheet.Range("Q1").Value = "Value"
    
        'Changing open_value from zero to the open of the first tracker
        open_value = current_sheet.Cells(2, 3).Value
    
        'Initializing loop from where data starts to appear until last row
        For i = 2 To last_row
    
            'Seeing when ticker changes
            If current_sheet.Cells(i + 1, 1).Value <> current_sheet.Cells(i, 1).Value Then
            
                'Inserting ticker data in array
                ticker_name = current_sheet.Cells(i, 1).Value
            
                'Calculating value_change and percentage_change
                close_value = current_sheet.Cells(i, 6).Value
                value_change = close_value - open_value
                'Avoid division by zero
                If open_value <> 0 Then
                    percentage_change = (value_change / open_value) * 100
                End If
                
                'Adding ticker volume
                ticker_volume = ticker_volume + current_sheet.Cells(i, 7).Value
                
                'Printing ticker name in basic table Column I
                current_sheet.Range("I" & Summary_Table_Row).Value = ticker_name
                'Printing yearly change in basic table Column J
                current_sheet.Range("J" & Summary_Table_Row).Value = value_change
                'Setting coloring based on negative or positive yearly change
                If (value_change > 0) Then
                    current_sheet.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                ElseIf (value_change <= 0) Then
                    current_sheet.Range("J" & Summary_Table_Row).Interior.Color = vbRed
                End If
                
                'Printing percentage change in basic table Column K
                current_sheet.Range("K" & Summary_Table_Row).Value = (CStr(percentage_change) & "%")
                'Printing stock volume in basic table Column K
                current_sheet.Range("L" & Summary_Table_Row).Value = ticker_volume
                
                'Adding one row to the set value of summary table defined earlier
                Summary_Table_Row = Summary_Table_Row + 1
                'Resetting value_change for new ticker
                value_change = 0
                'Resetting close_value for new ticker
                close_value = 0
                'Add next ticker's open value
                open_value = current_sheet.Cells(i + 1, 3).Value
                
                'Calculating new counters for bonus table
                If (percentage_change > biggest_percentage_increase) Then
                    biggest_percentage_increase = percentage_change
                    biggest_increase = ticker_name
                ElseIf (percentage_change < biggest_percentage_decrease) Then
                    biggest_percentage_decrease = percentage_change
                    biggest_decrease = ticker_name
                End If
                       
                If (ticker_volume > biggest_volume_number) Then
                    biggest_volume_number = ticker_volume
                    biggest_volume = ticker_name
                End If
                
                'Resetting counters
                percentage_change = 0
                ticker_volume = 0
                
            'Keep adding to ticker volume until change happens
            Else
                ticker_volume = ticker_volume + current_sheet.Cells(i, 7).Value
            End If
        
    Next i
        
            'Populate bonus table
            current_sheet.Range("Q2").Value = (CStr(biggest_percentage_increase) & "%")
            current_sheet.Range("Q3").Value = (CStr(biggest_percentage_decrease) & "%")
            current_sheet.Range("P2").Value = biggest_increase
            current_sheet.Range("P3").Value = biggest_decrease
            current_sheet.Range("Q4").Value = biggest_volume_number
            current_sheet.Range("P4").Value = biggest_volume
        
     Next current_sheet
                
End Sub