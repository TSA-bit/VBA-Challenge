Attribute VB_Name = "Module1"
Sub Stock_Analysis()
'Loop through all sheets
For Each ws In Worksheets

'Setting Dimensions
    Dim Total As Double
    Dim i As Long
    Dim Change As Single
    Dim j As Integer
    Dim Start As Long
    Dim PercentageChange As Single
    Dim Days As Integer
    Dim DailyChange As Single
    Dim AverageChange As Single

'Inserting Row titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

'Initial Values
    j = 0
    Total = 0
    Change = 0
    Start = 2

'What is the last row number of the last row with data
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To RowCount

    'If ticker code changes then print results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'store results in variables
            Total = Total + ws.Cells(i, 7).Value
        
        'Handle zero total volume
            If Total = 0 Then
            'print the results
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
            
            Else
            'find first non-zero starting value
                If ws.Cells(Start, 3) = 0 Then
                    For find_value = Start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            Start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
            
            'Calculate change
                Change = (ws.Cells(i, 6) - ws.Cells(Start, 3))
                PercentChange = Round((Change / ws.Cells(Start, 3) * 100), 2)
            
            'Beginning of next Stock Ticker
                Start = i + 1
            
            'print results
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = Round(Change, 2)
                ws.Range("K" & 2 + j).Value = "%" & PercentChange
                ws.Range("L" & 2 + j).Value = Total
            
            'colors positive change green and negative change red
                Select Case Change
                    Case Is > 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
            
            End If
        
        'reset variables for new Stock Ticker
            Total = 0
            Change = 0
            j = j + 1
            Days = 0
        
    'If ticker is the same - add results
        Else
            Total = Total + ws.Cells(i, 7).Value
        
        End If
    
    Next i

'Find the max and min then place in new cell
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))

'Return one less since header row is not a factord
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)

'final ticker symbol for total, greatest %increase, greatest %decrease and greatest total volume
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)
  
Next ws

End Sub
