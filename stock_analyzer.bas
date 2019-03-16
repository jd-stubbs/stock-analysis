Attribute VB_Name = "Module1"
Sub stock_analyzer()
    
    'declare variables
    Dim ws As Worksheet
    Dim total As Double
    Dim first As Double
    Dim last As Double
    Dim percent_change As Double
    Dim sRow As Integer
    Dim ticker As String
    
    'iterate through each worksheet
    For Each ws In Worksheets
        With ws
            'reset counters for each new sheet
            sRow = 2
            total = 0
            first = .Cells(2, 6).Value
            
            'Set up headers for sheet
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
    
            'iterate through the entire sheet
            For i = 2 To .Cells(Rows.Count, 1).End(xlUp).Row
                'sum the total volume
                total = total + .Cells(i, 7).Value
        
                'check to see if next stock ticker is different that the current stock ticker
                'if yes, print results for current stock ticker
                If .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Then
            
                    'save end of year value for later calculations
                    last = .Cells(i, 6).Value
            
                    'print current stock ticker in next row of list
                    ticker = .Cells(i, 1).Value
                    .Cells(sRow, 9).Value = ticker
            
                    'calulate, print and format yearly change in next row of list
                    .Cells(sRow, 10).Value = last - first
                    .Cells(sRow, 10).NumberFormat = "0.00"
                
                    'check if yearly change is positive or negative
                    'positive = green  |  negative = red
                    If last - first < 0 Then
                        .Cells(sRow, 10).Interior.ColorIndex = 3
                    Else
                        .Cells(sRow, 10).Interior.ColorIndex = 4
                    End If
                
                    'calculate, print and format percent change in next row of list
                    If first = 0 Then
                        percent_change = 0
                    Else
                        percent_change = (last - first) / first
                    End If
                    .Cells(sRow, 11).Value = percent_change
                    .Cells(sRow, 11).NumberFormat = "0.00%"
            
                    'check if current ticker percent change is greater than
                    'the current max percent change. if so, replace it and
                    'change the ticker to the current ticker
                    If percent_change > .Range("Q2").Value Then
                        .Range("Q2").Value = percent_change
                        .Range("Q2").NumberFormat = "0.00%"
                        .Range("P2").Value = ticker
                    End If
                
                    'check if current ticker percent change is less than
                    'the current min percent change. if so, replace it and
                    'change the ticker to the current ticker
                    If percent_change < .Range("Q3").Value Then
                        .Range("Q3").Value = percent_change
                        .Range("Q3").NumberFormat = "0.00%"
                        .Range("P3").Value = ticker
                    End If
            
                    'print total volume in next row of list
                    .Cells(sRow, 12).Value = total
            
                    'check if current ticker total volume is greater than
                    'the current max total volume. if so, replace it and
                    'change the ticker to the current ticker
                    If total > .Range("Q4").Value Then
                        .Range("Q4").Value = total
                        .Range("P4").Value = ticker
                    End If
            
                    'reset variables for next stock ticker
                    sRow = sRow + 1
                    total = 0
                    first = .Cells(i + 1, 6)
                End If
            Next i
    
            'Adjust column widths for easy reading
            .Columns("I:L").AutoFit
            .Columns("O:Q").AutoFit
            
        End With
    Next
End Sub

Sub clear()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Columns("I:Q").clear
    Next
End Sub
