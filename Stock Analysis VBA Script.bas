Attribute VB_Name = "Module1"
Sub stock_analysis()

  'identify your variables- there is a lot - so i'm giving those to you
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        j = 0
        total = 0
        change = 0
        start = 2
        
        'Set title row - columns with new names
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'find the row number of the last row with data
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' go through the whole data set starting at row 2 until the last row
        For i = 2 To rowCount
            
            ' If ticker changes then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'store results in variables
                total = total + ws.Cells(i, 7).Value
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value

        'handle zero total volume
                If total = 0 Then
                    percentChange = 0
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
         ws.Range("J" & 2 + j).Value = change
        ws.Range("K" & 2 + j).Value = percentChange
                ws.Range("L" & 2 + j).Value = total
        Else
                    percentChange = change / total
                End If

                
                ' Find First non zero starting value
                If ws.Cells(start, 3).Value = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
                
                ' Calculate Change
                change = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
                percentChange = change / ws.Cells(start, 3).Value
                
                ' start of the next stock ticker
                start = i + 1
        
          'print the results
                ws.Range("J" & 2 + j).Value = change
        ws.Range("K" & 2 + j).Value = percentChange
        ws.Range("J" & 2 + j).NumberFormat = "0.00"
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
        ws.Range("L" & 2 + j).Value = total

         ' colors positives green and negative numbers red
                Select Case change
                    Case Is > 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
                
                ' reset variables for new stock ticker
                total = 0
                change = 0
                j = j + 1
                start = i + 1
        
         ' If ticker is still the same add results
            Else
                total = total + ws.Cells(i, 7).Value

            End If
        Next i
        
         'Take the max and min and place them in a separate part in the worksheet
' Examples of max function. You need a Min too, which works similarly
ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

' Returns one less because header row is not a factor
' Another function - Match
increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

' Final ticker symbol for total, greatest % of increase and decrease, and average
ws.Range("P2") = ws.Cells(increase_number + 1, 9)
ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
ws.Range("P4") = ws.Cells(volume_number + 1, 9)

' Row names for summary table
ws.Range("O2") = "Greatest % increase"
ws.Range("O3") = "Greatest % decrease"
ws.Range("O4") = "Greatest Total Stock Volume"

' Column headers for summary table
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

    Next ws


End Sub
