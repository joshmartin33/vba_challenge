Attribute VB_Name = "Module1"
Sub CreditCard()

    '1Declare all my variables
    '2Setup workbook loop
    '3 add Colum headers
    '4 add bonus column/row headers
    '5 Start worksheet loop
        '5a zero out variables
        '5b Function to find last cell
        '5c Grab "Opening" variable value
        '5d Start for loop
            '1 if cell and cell + 1 match add to running total variable
            '2 Else
                'add to running total
                'add ticker to data grid
                'add running total to data grid
                'Grab closing variable value
                'calculate change variable value
                'If change variable > 0 cell turns green
                    'Else red
                'add change variable value to data grid
                'calculate change_percent variable value
                'add change_precent variable to data grid
                '''''Bonus
                'If to see if change_percent > than current. Place if true
                'If to see if change_percent < than current. Place if true
                'If to see if RunningTotal > than current. Place if true
                ''''''Closing
                'Counter + 1
                'Grab new Opening variable value
                'Running total = 0
            'End if
        'Next i
    'Next workbook
'End Sub
    
    'Declaring my variables
    
    Dim RunningTotal As LongLong
    Dim Opening As Double
    Dim Closing As Double
    Dim LastRow As Long
    Dim Counter As Integer
    Dim Change As Double
    Dim ChangePercent As Double
    
    ''''''Setup worksheet loop'''''''''
    For Each ws In Worksheets
        
        'Adding in Column Headers for data grid
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Adding Bonus Column/Row Headers
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Format cell/column number format to %
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        
        'Opening values for variables
        
        Opening = ws.Cells(2, 3).Value
        Counter = 2
        RunningTotal = 0
        
        'Find last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'For loop
        'Designed to look at i value and compare it against i + 1
        'if values match add to running total
        
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            RunningTotal = RunningTotal + ws.Cells(i, 7).Value
            
            Else
            
            'Calculate/declare variables values
            
            RunningTotal = RunningTotal + ws.Cells(i, 7).Value
            Closing = ws.Cells(i, 6).Value
            Change = Closing - Opening
            ChangePercent = Change / Opening
            
            'Add values to data grid
            
                    'Add Ticker
            ws.Range("I" & Counter).Value = ws.Cells(i, 1).Value
                    'Add Yearly change
            ws.Range("J" & Counter).Value = Change
                    'Add percent change
            ws.Range("K" & Counter).Value = ChangePercent
                    'Add total stock volume
            ws.Range("L" & Counter).Value = RunningTotal
            
            'Add conditionals to change cells showing red for negatice and green for positive
            
            If Change < 0 Then
                ws.Range("J" & Counter).Interior.ColorIndex = 3
            Else
                ws.Range("J" & Counter).Interior.ColorIndex = 4
            End If
            
            '''BONUS'''
            
            'check if ChangePercent is greater
            If ChangePercent > ws.Cells(2, 17).Value Then
                    'Overrides Ticker cell value
                ws.Cells(2, 16).Value = ws.Cells(i, 1).Value
                    'Overrides ChangePercent cell value
                ws.Cells(2, 17).Value = ChangePercent
            End If
            
             'check if ChangePercent is lower
            If ChangePercent < ws.Cells(3, 17).Value Then
                    'Overrides Ticker cell value
                ws.Cells(3, 16).Value = ws.Cells(i, 1).Value
                    'Overrides ChangePercent cell value
                ws.Cells(3, 17).Value = ChangePercent
            End If
            
             'check if RunningTotal is greater
            If RunningTotal > ws.Cells(4, 17).Value Then
                    'Overrides Ticker cell value
                ws.Cells(4, 16).Value = ws.Cells(i, 1).Value
                    'Overrides greatest total volume cell value
                ws.Cells(4, 17).Value = RunningTotal
            End If
            
            '''''Closing'''''
            
            'Set new Variable values
            RunningTotal = 0
            Counter = Counter + 1
            Opening = ws.Cells(i + 1, 3).Value
            
        End If
    Next i
    
    Next ws
    
MsgBox("Process Complete")

End Sub


