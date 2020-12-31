Attribute VB_Name = "Module1"
Sub Stock_Analysis()



'____________________________Setting Up____________________________'

'Set variables used for looping through all worksheets
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

'Count the number of rows
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Set up Ticker to compare with
LastTicker = ""

'Set up sum to record the sum of the stock volume
Sum = 0



'____________________________Starting Loop____________________________'

'Loop through each worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

'Set up header of the table on each worksheet
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"

'Set up Max, Min, Count and their respective Index to record the greatest ticker
Max = 0
MaxIndex = 0
Min = 0
MinIndex = 0
Count = 0
CountIndex = 0

'Set up counter to keep track of rows in the table to write to, and will reset on each worksheet
Counter = 1

    'Loop through each row till a row after the end of list
    For Row = 2 To (LastRow + 1)
    
        'If the ticker is same:
        If Cells(Row, 1).Value = LastTicker Then
            
            'Add the stock volume to sum
            Sum = Sum + Cells(Row, 7)
            
        'If it is the very first ticker:
        ElseIf Counter = 1 Then
            
            'Break this condition so it will go to the following
            Counter = Counter + 1
            
            'Record ticker symbol and write it to the table just below the header
            LastTicker = Cells(Row, 1).Value
            Cells(Counter, 9).Value = LastTicker
            
            'Record OpenInitial for calculating Change
            OpenInitial = Cells(Row, 3).Value
        
        'If the ticker is different:
        Else
            
            'Increment the counter
            Counter = Counter + 1
            
            'Record new ticker symbol and write it to the table
            LastTicker = Cells(Row, 1).Value
            Cells(Counter, 9).Value = LastTicker
            
            'Calculate change and record in Change
            CloseFinal = Cells(Row - 1, 6).Value
            Change = CloseFinal - OpenInitial
                
                'Use a conditional to avoid divide by zero error
                'If OpenInitial is zero, then set percent to 0
                If OpenInitial = 0 Then
                    
                    Percent = 0
                
                'Otherwise, save the value to Percent
                Else
                
                    Percent = Change / OpenInitial
                                            
                End If
            
            'Use the Percent to compare to the Max, Min and Sum to Count to find the greatest throughout the worksheet
            If Percent > Max Then
                
                Max = Percent
                MaxIndex = Counter
                
            ElseIf Percent < Min Then
                
                Min = Percent
                MinIndex = Counter
                
            End If
            
            If Sum > Count Then
                
                Count = Sum
                CountIndex = Counter
                
            End If
                            
            'Write Change and Percent to the table
            Percent = FormatPercent(Percent, 2)
            Cells(Counter - 1, 10) = Change
            Cells(Counter - 1, 11) = Percent
                
            'If Change is positive, then make it green
            If Cells(Counter - 1, 10).Value > 0 Then
            
                Cells(Counter - 1, 10).Interior.ColorIndex = 4
            
            'If Change is negative, then make it red
            ElseIf Cells(Counter - 1, 10).Value < 0 Then
            
                Cells(Counter - 1, 10).Interior.ColorIndex = 3
                
            End If
                            
            'Record new OpenInitial
            OpenInitial = Cells(Row, 3).Value
            
            'Write Sum to the table and reset it
            Cells(Counter - 1, 12).Value = Sum
            Sum = 0
            
        End If
        
    Next Row
    
    'Write the "Greatest" value to the second table on first page
    Cells(2, 15).Value = Cells(MaxIndex - 1, 9).Value
    Cells(3, 15).Value = Cells(MinIndex - 1, 9).Value
    Cells(4, 15).Value = Cells(CountIndex - 1, 9).Value
    Cells(2, 16).Value = FormatPercent(Max, 2)
    Cells(3, 16).Value = FormatPercent(Min, 2)
    Cells(4, 16).Value = Count
    
Next

starting_ws.Activate



End Sub


