Sub LoopThroughStocks()
    Dim ws As Worksheet
    Dim ClosePrice As Double
    Dim Volume As LongLong
    Dim FirstRow As Long
    Dim LastRow As Long
    Dim Ticker As String
    Dim InitialOpenPrice As Double
    Dim FinalClosePrice As Double
    Dim TotalVolume As LongLong
    Dim SummaryRow As Long
    Dim TickerCounter As Long
    Dim i As Long
    Dim QuarterlyChange As Double
    Dim PercentageChange As Double
    Dim MaxPercentageIncrease As Double
    Dim MaxPercentageDecrease As Double
    Dim MaxVolume As LongLong
    Dim TickerMaxIncrease As String
    Dim TickerMaxDecrease As String
    Dim TickerMaxVolume As String
    
    TickerCounter = 2  ' Starting row for summary
    
    ' Loop through each worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ' Reset variables for each worksheet
        InitialOpenPrice = 0
        MaxPercentageIncrease = 0
        MaxPercentageDecrease = 0
        MaxVolume = 0
        TickerMaxIncrease = ""
        TickerMaxDecrease = ""
        TickerMaxVolume = ""
        
        ' Check if the sheet name starts with "Q" (assuming quarters are named Q1, Q2, etc.)
        If Left(ws.Name, 1) = "Q" Then
            ' Get the first and last rows of data
            FirstRow = 2
            LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
            
            ' Initialize variables
            Ticker = ""
            Volume = 0
            
            ' Write headers for the summary
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Volume"
            
            ' Loop through each row of stock data in the worksheet
            For i = FirstRow To LastRow
                If ws.Cells(i, 1).Value = Ticker Then
                    Volume = Volume + ws.Cells(i, 7).Value
                Else
                    If Ticker <> "" Then
                        ' Calculate the final values for the previous ticker
                        FinalClosePrice = ws.Cells(i - 1, 6).Value
                        QuarterlyChange = FinalClosePrice - InitialOpenPrice
                        If InitialOpenPrice <> 0 Then
                            PercentageChange = (QuarterlyChange / InitialOpenPrice) * 100
                        Else
                            PercentageChange = 0
                        End If
                        
                        ' Write the summary for the previous ticker
                        ws.Cells(TickerCounter, 9).Value = Ticker
                        ws.Cells(TickerCounter, 10).Value = QuarterlyChange
                        ws.Cells(TickerCounter, 11).Value = Application.WorksheetFunction.Round(PercentageChange, 2)
                        ws.Cells(TickerCounter, 12).Value = Volume
                        
                        ' Check for max percentage increase, decrease, and volume
                        If PercentageChange > MaxPercentageIncrease Then
                            MaxPercentageIncrease = PercentageChange
                            TickerMaxIncrease = Ticker
                        End If
                        
                        If PercentageChange < MaxPercentageDecrease Then
                            MaxPercentageDecrease = PercentageChange
                            TickerMaxDecrease = Ticker
                        End If
                        
                        If Volume > MaxVolume Then
                            MaxVolume = Volume
                            TickerMaxVolume = Ticker
                        End If
                        
                        TickerCounter = TickerCounter + 1
                    End If
                    
                    ' Reset values for the new ticker
                    Ticker = ws.Cells(i, 1).Value
                    If IsNumeric(ws.Cells(i, 3).Value) Then
                        InitialOpenPrice = ws.Cells(i, 3).Value  ' Initialize for the new ticker if numeric
                    Else
                        InitialOpenPrice = 0  ' Set to 0 if not numeric (handle error case)
                    End If
                    Volume = ws.Cells(i, 7).Value
                End If
            Next i
            
            ' Calculate the final values for the last ticker in the worksheet
            FinalClosePrice = ws.Cells(LastRow, 6).Value
            QuarterlyChange = FinalClosePrice - InitialOpenPrice
            If InitialOpenPrice <> 0 Then
                PercentageChange = (QuarterlyChange / InitialOpenPrice) * 100
            Else
                PercentageChange = 0
            End If
            
            ' Write the summary for the last ticker
            ws.Cells(TickerCounter, 9).Value = Ticker
            ws.Cells(TickerCounter, 10).Value = QuarterlyChange
            ws.Cells(TickerCounter, 11).Value = Application.WorksheetFunction.Round(PercentageChange, 2)
            ws.Cells(TickerCounter, 12).Value = Volume
            
            ' Check for max percentage increase, decrease, and volume for the last ticker
            If PercentageChange > MaxPercentageIncrease Then
                MaxPercentageIncrease = PercentageChange
                TickerMaxIncrease = Ticker
            End If
            
            If PercentageChange < MaxPercentageDecrease Then
                MaxPercentageDecrease = PercentageChange
                TickerMaxDecrease = Ticker
            End If
            
            If Volume > MaxVolume Then
                MaxVolume = Volume
                TickerMaxVolume = Ticker
            End If
            
            ' Apply conditional formatting for percentage change
            Dim rng As Range
            Set rng = ws.Range(ws.Cells(2, 11), ws.Cells(TickerCounter, 11))
            
            ' Clear existing conditional formatting
            rng.FormatConditions.Delete
            
            ' Format cells with conditions
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                .Interior.Color = RGB(146, 208, 80) ' Green for positive change
            End With
            
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                .Interior.Color = RGB(255, 0, 0) ' Red for negative change
            End With
        End If
    Next ws
    
    ' Output the stocks with greatest % increase, % decrease, and total volume
    MsgBox "Stock data summary has been added to the bottom of each quarterly sheet." & vbCrLf & _
           "Greatest % Increase: " & TickerMaxIncrease & " (" & Format(MaxPercentageIncrease, "0.00") & "%)" & vbCrLf & _
           "Greatest % Decrease: " & TickerMaxDecrease & " (" & Format(MaxPercentageDecrease, "0.00") & "%)" & vbCrLf & _
           "Greatest Total Volume: " & TickerMaxVolume & " (" & MaxVolume & " units)"
End Sub



